#!/usr/bin/env bun
import { mkdirSync, writeFileSync } from 'node:fs'
import { basename, join } from 'node:path'

type ContractName = 'first-view' | 'edit-verify' | 'repeated-inspection'
type StepStatus = 'ok' | 'timeout' | 'failed' | 'skipped'

interface Args {
	readonly inputFile: string
	readonly sheet: string
	readonly range: string
	readonly editInputFile: string
	readonly rowLimit: number
	readonly mutations: number
	readonly repeat: number
	readonly warmup: number
	readonly timeoutMs: number
	readonly outDir: string
	readonly contract: ContractName | 'all'
	readonly dryRun: boolean
	readonly json: boolean
}

interface StepSpec {
	readonly contract: ContractName
	readonly id: string
	readonly label: string
	readonly command: readonly string[]
	readonly timeoutMs: number
	readonly profileCommand?: readonly string[]
	readonly required: boolean
}

interface StepResult {
	readonly contract: ContractName
	readonly id: string
	readonly label: string
	readonly status: StepStatus
	readonly command: string
	readonly outputPath?: string
	readonly stderrPath?: string
	readonly profileCommand?: string
	readonly elapsedMs: number
	readonly summary?: unknown
	readonly error?: string
}

const DEFAULT_INPUT = 'research/excel-corpus/NYC_311_SR_2010-2020-sample-1M.xlsx'
const DEFAULT_SHEET = 'NYC_311_SR_2010-2020-sample-1M'
const DEFAULT_RANGE = 'A1:AO1000001'
const DEFAULT_EDIT_INPUT = 'fixtures/xlsx/stress/dense-100k.xlsx'
const DEFAULT_OUT_DIR = join('/private/tmp', `ascend-practical-contracts-${Date.now()}`)
const PROFILE_SCRIPT = 'fixtures/benchmarks/profile-bun.ts'

function readOption(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	return index >= 0 ? process.argv[index + 1] : undefined
}

function hasFlag(name: string): boolean {
	return process.argv.includes(name)
}

function positiveInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

function parseContract(raw: string | undefined): Args['contract'] {
	if (
		raw === undefined ||
		raw === 'all' ||
		raw === 'first-view' ||
		raw === 'edit-verify' ||
		raw === 'repeated-inspection'
	) {
		return raw ?? 'all'
	}
	throw new Error('--contract must be all, first-view, edit-verify, or repeated-inspection')
}

function parseArgs(): Args {
	return {
		inputFile: readOption('--input-file') ?? DEFAULT_INPUT,
		sheet: readOption('--sheet') ?? DEFAULT_SHEET,
		range: readOption('--range') ?? DEFAULT_RANGE,
		editInputFile: readOption('--edit-input-file') ?? DEFAULT_EDIT_INPUT,
		rowLimit: positiveInt(readOption('--row-limit'), 500),
		mutations: positiveInt(readOption('--mutations'), 25),
		repeat: positiveInt(readOption('--repeat'), 3),
		warmup: nonNegativeInt(readOption('--warmup'), 1),
		timeoutMs: positiveInt(readOption('--timeout-ms'), 600_000),
		outDir: readOption('--out-dir') ?? DEFAULT_OUT_DIR,
		contract: parseContract(readOption('--contract')),
		dryRun: hasFlag('--dry-run'),
		json: hasFlag('--json'),
	}
}

function shellQuote(value: string): string {
	return /^[A-Za-z0-9_./:=@+-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function commandString(args: readonly string[]): string {
	return args.map(shellQuote).join(' ')
}

function profileCommand(label: string, command: readonly string[], outDir: string): string[] {
	return [
		'bun',
		'run',
		PROFILE_SCRIPT,
		'--mode',
		'all-md',
		'--label',
		label,
		'--out-dir',
		outDir,
		'--',
		...command,
	]
}

function readPhaseArgs(args: Args, phase: string): string[] {
	return [
		'bun',
		'run',
		'fixtures/benchmarks/xlsx-read-phase.ts',
		'--input-file',
		args.inputFile,
		'--rows',
		'1000001',
		'--cols',
		'41',
		'--workload',
		'dense-values',
		'--read-source',
		'raw-ooxml',
		'--phase',
		phase,
		'--repeat',
		String(args.repeat),
		'--warmup',
		String(args.warmup),
		'--validation-mode',
		phase === 'zip' ? 'none' : 'sample',
		'--json',
	]
}

function firstWindowArgs(args: Args, only: string): string[] {
	return [
		'bun',
		'run',
		'fixtures/benchmarks/agent-first-window.ts',
		'--input-file',
		args.inputFile,
		'--sheet',
		args.sheet,
		'--range',
		args.range,
		'--row-limit',
		String(args.rowLimit),
		'--only',
		only,
		'--repeat',
		String(args.repeat),
		'--warmup',
		String(args.warmup),
		'--no-gc-between-samples',
		'--json',
	]
}

function workflowArgs(args: Args): string[] {
	return [
		'bun',
		'run',
		'fixtures/benchmarks/agent-workflow.ts',
		'--input-file',
		args.editInputFile,
		'--surface',
		'api',
		'--row-limit',
		String(args.rowLimit),
		'--mutations',
		String(args.mutations),
		'--repeat',
		String(args.repeat),
		'--warmup',
		String(args.warmup),
		'--json',
	]
}

function postWriteArgs(args: Args): string[] {
	return [
		'bun',
		'run',
		'fixtures/benchmarks/post-write-breakdown.ts',
		'--input-file',
		args.editInputFile,
		'--updates',
		String(args.mutations),
		'--repeat',
		String(args.repeat),
		'--warmup',
		String(args.warmup),
		'--json',
	]
}

function buildSteps(args: Args): StepSpec[] {
	const profileDir = join(args.outDir, 'profiles')
	const steps: StepSpec[] = [
		{
			contract: 'first-view',
			id: 'package-scan-zip-xml',
			label: 'Cold package scan, ZIP inflate, worksheet decode',
			command: readPhaseArgs(args, 'zip'),
			timeoutMs: args.timeoutMs,
			required: true,
		},
		{
			contract: 'first-view',
			id: 'capped-read-window',
			label: 'Capped XLSX parse, workbook hydration, first agent window',
			command: readPhaseArgs(args, 'capped-agent-window'),
			timeoutMs: args.timeoutMs,
			required: true,
		},
		{
			contract: 'first-view',
			id: 'api-first-view',
			label: 'API first agent view payload shaping',
			command: firstWindowArgs(args, 'api'),
			timeoutMs: args.timeoutMs,
			profileCommand: profileCommand(
				'contract-first-view-api',
				firstWindowArgs(args, 'api'),
				profileDir,
			),
			required: true,
		},
		{
			contract: 'edit-verify',
			id: 'workflow-commit',
			label: 'Edit plan to commit with write timings',
			command: workflowArgs(args),
			timeoutMs: args.timeoutMs,
			profileCommand: profileCommand(
				'contract-edit-verify-workflow',
				workflowArgs(args),
				profileDir,
			),
			required: true,
		},
		{
			contract: 'edit-verify',
			id: 'post-write-breakdown',
			label: 'Write, reopen, check, lint, preservation, package graph verify',
			command: postWriteArgs(args),
			timeoutMs: args.timeoutMs,
			required: true,
		},
		{
			contract: 'repeated-inspection',
			id: 'cached-agent-window',
			label: 'Repeated first-window cache hit for agent reads',
			command: firstWindowArgs(args, 'capped'),
			timeoutMs: args.timeoutMs,
			required: true,
		},
		{
			contract: 'repeated-inspection',
			id: 'tui-first-paint-cache',
			label: 'Repeated TUI open/render after cache warmup',
			command: firstWindowArgs(args, 'tui'),
			timeoutMs: args.timeoutMs,
			required: true,
		},
	]
	return args.contract === 'all' ? steps : steps.filter((step) => step.contract === args.contract)
}

async function runStep(step: StepSpec, outDir: string, dryRun: boolean): Promise<StepResult> {
	const command = commandString(step.command)
	const outputPath = join(outDir, `${step.contract}-${step.id}.json`)
	const stderrPath = join(outDir, `${step.contract}-${step.id}.stderr.txt`)
	if (dryRun) {
		return {
			contract: step.contract,
			id: step.id,
			label: step.label,
			status: 'skipped',
			command,
			outputPath,
			stderrPath,
			...(step.profileCommand ? { profileCommand: commandString(step.profileCommand) } : {}),
			elapsedMs: 0,
		}
	}
	console.error(`[contracts] start ${step.contract}/${step.id}: ${step.label}`)
	const start = performance.now()
	const proc = Bun.spawn(step.command, {
		cwd: process.cwd(),
		stdout: 'pipe',
		stderr: 'pipe',
	})
	const timeout = setTimeout(() => {
		proc.kill()
	}, step.timeoutMs)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	clearTimeout(timeout)
	const elapsedMs = performance.now() - start
	writeFileSync(outputPath, stdout)
	writeFileSync(stderrPath, stderr)
	const timedOut = elapsedMs >= step.timeoutMs && exitCode !== 0
	const status: StepStatus = exitCode === 0 ? 'ok' : timedOut ? 'timeout' : 'failed'
	console.error(
		`[contracts] ${status} ${step.contract}/${step.id}: ${(elapsedMs / 1000).toFixed(1)}s`,
	)
	let summary: unknown
	let error: string | undefined
	if (status === 'ok') {
		try {
			summary = (JSON.parse(stdout) as { readonly summary?: unknown }).summary
		} catch (err) {
			error = err instanceof Error ? err.message : String(err)
		}
	} else {
		error = stderr.trim().split('\n').slice(0, 8).join('\n') || `exit code ${exitCode}`
	}
	return {
		contract: step.contract,
		id: step.id,
		label: step.label,
		status,
		command,
		outputPath,
		stderrPath,
		...(step.profileCommand ? { profileCommand: commandString(step.profileCommand) } : {}),
		elapsedMs,
		...(summary !== undefined ? { summary } : {}),
		...(error ? { error } : {}),
	}
}

function metric(summary: unknown, key: string): string {
	if (!summary || typeof summary !== 'object') return ''
	const value = (summary as Record<string, unknown>)[key]
	return typeof value === 'number' ? value.toFixed(value >= 100 ? 1 : 3) : ''
}

function buildMarkdown(args: Args, results: readonly StepResult[]): string {
	const lines = [
		'# Ascend Practical Latency Contracts',
		'',
		`Generated: ${new Date().toISOString()}`,
		`Input workbook: \`${args.inputFile}\``,
		`Edit workbook: \`${args.editInputFile}\``,
		`Timeout: ${args.timeoutMs}ms per step`,
		'',
		'## Contract 1: Unknown Workbook To Safe First Agent View',
		'',
		'Guardrail: report cold package/ZIP/XML work separately from cached first-window results; do not call hot-cache numbers end-to-end.',
		'',
		firstViewTable(results),
		'',
		'## Contract 2: Edit Plan To Commit/Reopen/Verify',
		'',
		'Guardrail: write latency is not complete unless output bytes are written, reopened, checked, linted, and package graph verified.',
		'',
		editVerifyTable(results),
		'',
		'## Contract 3: Repeated Viewport/Table Inspection After First Load',
		'',
		'Guardrail: separate cold hydration from warm cache hit, window shaping, render bytes, and payload bytes. Table-specific inspection remains a required follow-up when the selected workbook has real table metadata.',
		'',
		repeatedInspectionTable(results),
		'',
		'## Profile Commands',
		'',
		...results
			.filter((result) => result.profileCommand)
			.map((result) => `- ${result.contract}/${result.id}: \`${result.profileCommand}\``),
		'',
		'## Next Bottleneck Choice',
		'',
		nextBottleneck(results),
		'',
	]
	return `${lines.join('\n')}\n`
}

function firstViewTable(results: readonly StepResult[]): string {
	const zip = results.find((entry) => entry.id === 'package-scan-zip-xml')
	const capped = results.find((entry) => entry.id === 'capped-read-window')
	const api = results.find((entry) => entry.id === 'api-first-view')
	return [
		'| Phase | Median ms | Memory / payload | Evidence |',
		'|---|---:|---:|---|',
		`| ZIP open | ${metric(zip?.summary, 'zipOpenMedianMs')} | | ${statusLink(zip)} |`,
		`| Worksheet inflate | ${metric(zip?.summary, 'worksheetInflateMedianMs')} | | ${statusLink(zip)} |`,
		`| Worksheet decode | ${metric(zip?.summary, 'worksheetDecodeMedianMs')} | | ${statusLink(zip)} |`,
		`| Capped read/hydrate | ${metric(capped?.summary, 'cappedReadWindowMedianMs')} | peak RSS ${metric(capped?.summary, 'peakRssBytes')} bytes | ${statusLink(capped)} |`,
		`| Window shaping | ${metric(capped?.summary, 'cappedAgentWindowMedianMs')} | | ${statusLink(capped)} |`,
		`| API payload first view | ${metric(api?.summary, 'apiFirstWindowMedianMs')} | ${metric(api?.summary, 'payloadBytesMedian')} bytes | ${statusLink(api)} |`,
	].join('\n')
}

function editVerifyTable(results: readonly StepResult[]): string {
	const workflow = results.find((entry) => entry.id === 'workflow-commit')
	const postWrite = results.find((entry) => entry.id === 'post-write-breakdown')
	return [
		'| Phase | Median ms | Payload/bytes | Evidence |',
		'|---|---:|---:|---|',
		`| Plan | ${metric(workflow?.summary, 'planMedianMs')} | ${metric(workflow?.summary, 'planPayloadBytesMedian')} bytes | ${statusLink(workflow)} |`,
		`| Commit verified total | ${metric(workflow?.summary, 'commitVerifiedTotalMedianMs')} | ${metric(workflow?.summary, 'commitVerifiedPayloadBytesMedian')} bytes | ${statusLink(workflow)} |`,
		`| Write | ${metric(postWrite?.summary, 'commitWriteMedianMs')} | output ${metric(postWrite?.summary, 'outputBytesMedian')} bytes | ${statusLink(postWrite)} |`,
		`| Reopen output | ${metric(postWrite?.summary, 'commitPostWriteReopenMedianMs')} | | ${statusLink(postWrite)} |`,
		`| Check | ${metric(postWrite?.summary, 'commitPostWriteCheckMedianMs')} | | ${statusLink(postWrite)} |`,
		`| Formula lint | ${metric(postWrite?.summary, 'commitPostWriteLintMedianMs')} | | ${statusLink(postWrite)} |`,
		`| Package graph verify | ${metric(postWrite?.summary, 'commitPostWritePackageGraphAuditMedianMs')} | | ${statusLink(postWrite)} |`,
	].join('\n')
}

function repeatedInspectionTable(results: readonly StepResult[]): string {
	const cached = results.find((entry) => entry.id === 'cached-agent-window')
	const tui = results.find((entry) => entry.id === 'tui-first-paint-cache')
	return [
		'| Phase | Cold median ms | Warm median ms | Payload/render | Evidence |',
		'|---|---:|---:|---:|---|',
		`| Cached document first window | ${metric(cached?.summary, 'cappedOpenWindowMedianMs')} | ${metric(cached?.summary, 'cappedWarmOpenWindowMedianMs')} | cells ${metric(cached?.summary, 'cellsMedian')} | ${statusLink(cached)} |`,
		`| TUI open/render first paint | ${metric(tui?.summary, 'tuiFirstPaintMedianMs')} | ${metric(tui?.summary, 'tuiWarmFirstPaintMedianMs')} | frame ${metric(tui?.summary, 'tuiFrameBytesMedian')} bytes | ${statusLink(tui)} |`,
	].join('\n')
}

function statusLink(result: StepResult | undefined): string {
	if (!result) return 'not run'
	const path = result.outputPath ? basename(result.outputPath) : ''
	return result.status === 'ok' ? `ok (${path})` : `${result.status}: ${result.error ?? path}`
}

function nextBottleneck(results: readonly StepResult[]): string {
	const postWrite = results.find((entry) => entry.id === 'post-write-breakdown')
	const capped = results.find((entry) => entry.id === 'capped-read-window')
	const api = results.find((entry) => entry.id === 'api-first-view')
	const write = Number(metric(postWrite?.summary, 'commitWriteMedianMs'))
	const reopen = Number(metric(postWrite?.summary, 'commitPostWriteReopenMedianMs'))
	const cappedRead = Number(metric(capped?.summary, 'cappedReadWindowMedianMs'))
	const apiFirst = Number(metric(api?.summary, 'apiFirstWindowMedianMs'))
	const candidates = [
		['edit-verify/write', write],
		['edit-verify/reopen', reopen],
		['first-view/capped-read', cappedRead],
		['first-view/api-payload', apiFirst],
	] as const
	const [name, value] = candidates
		.filter(([, candidate]) => Number.isFinite(candidate) && candidate > 0)
		.sort((a, b) => b[1] - a[1])[0] ?? ['not-measured', 0]
	if (name === 'not-measured') {
		return 'No optimization selected yet: one or more required measurements failed or timed out.'
	}
	return `Choose \`${name}\` first. It is the largest measured envelope phase in this report (${value.toFixed(1)}ms median). Max plausible win is bounded by that phase share of its named envelope; changes outside this phase are discarded unless they move the same contract.`
}

async function run() {
	const args = parseArgs()
	mkdirSync(args.outDir, { recursive: true })
	mkdirSync(join(args.outDir, 'profiles'), { recursive: true })
	const steps = buildSteps(args)
	const results: StepResult[] = []
	for (const step of steps) {
		results.push(await runStep(step, args.outDir, args.dryRun))
	}
	const payload = { tool: 'practical-latency-contracts', args, results }
	const jsonPath = join(args.outDir, 'summary.json')
	const markdownPath = join(args.outDir, 'summary.md')
	writeFileSync(jsonPath, JSON.stringify(payload, null, 2))
	writeFileSync(markdownPath, buildMarkdown(args, results))
	const output = { ...payload, summaryJson: jsonPath, summaryMarkdown: markdownPath }
	if (args.json) console.log(JSON.stringify(output, null, 2))
	else console.log(`summary: ${markdownPath}`)
}

await run()
