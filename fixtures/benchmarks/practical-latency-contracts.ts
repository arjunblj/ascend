#!/usr/bin/env bun
import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
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

interface EnvelopeDecision {
	readonly contract: ContractName
	readonly envelope: string
	readonly envelopeMedianMs: number
	readonly largestPhase: string
	readonly phaseMedianMs: number
	readonly profileCommand: string
	readonly guardrail: string
}

interface DiagnosticCeiling {
	readonly name: string
	readonly medianMs: number
	readonly profileCommand: string
	readonly guardrail: string
}

interface PhaseCandidate {
	readonly name: string
	readonly medianMs: number
	readonly profileCommand: string
}

interface WorktreeState {
	readonly branchLine: string
	readonly dirty: boolean
	readonly trackedDirty: boolean
	readonly trackedDirtyFiles: readonly string[]
	readonly untrackedCount: number
	readonly status: readonly string[]
}

interface InputProvenance {
	readonly path: string
	readonly tracked: boolean
	readonly exists: boolean
	readonly releaseClaimable: boolean
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

function tableInspectionArgs(args: Args): string[] {
	return [
		'bun',
		'run',
		'fixtures/benchmarks/table-inspection-cache.ts',
		'--row-limit',
		String(args.rowLimit),
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
			profileCommand: profileCommand(
				'contract-first-view-package-scan',
				readPhaseArgs(args, 'zip'),
				profileDir,
			),
			required: true,
		},
		{
			contract: 'first-view',
			id: 'capped-read-window',
			label: 'Capped XLSX parse, workbook hydration, first agent window',
			command: readPhaseArgs(args, 'capped-agent-window'),
			timeoutMs: args.timeoutMs,
			profileCommand: profileCommand(
				'contract-first-view-capped-window',
				readPhaseArgs(args, 'capped-agent-window'),
				profileDir,
			),
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
			profileCommand: profileCommand(
				'contract-edit-verify-post-write',
				postWriteArgs(args),
				profileDir,
			),
			required: true,
		},
		{
			contract: 'repeated-inspection',
			id: 'cached-agent-window',
			label: 'Repeated first-window cache hit for agent reads',
			command: firstWindowArgs(args, 'capped'),
			timeoutMs: args.timeoutMs,
			profileCommand: profileCommand(
				'contract-repeated-agent-window',
				firstWindowArgs(args, 'capped'),
				profileDir,
			),
			required: true,
		},
		{
			contract: 'repeated-inspection',
			id: 'tui-first-paint-cache',
			label: 'Repeated TUI open/render after cache warmup',
			command: firstWindowArgs(args, 'tui'),
			timeoutMs: args.timeoutMs,
			profileCommand: profileCommand(
				'contract-repeated-tui-first-paint',
				firstWindowArgs(args, 'tui'),
				profileDir,
			),
			required: true,
		},
		{
			contract: 'repeated-inspection',
			id: 'table-inspection-cache',
			label: 'Repeated table metadata and row-window inspection after cache warmup',
			command: tableInspectionArgs(args),
			timeoutMs: args.timeoutMs,
			profileCommand: profileCommand(
				'contract-repeated-table-inspection',
				tableInspectionArgs(args),
				profileDir,
			),
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
	let killedByTimeout = false
	const timeout = setTimeout(() => {
		killedByTimeout = true
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
	const timedOut = killedByTimeout && exitCode !== 0
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
	const value = numericMetric(summary, key)
	return value === undefined ? '' : value.toFixed(value >= 100 ? 1 : 3)
}

function numericMetric(summary: unknown, key: string): number | undefined {
	if (!summary || typeof summary !== 'object') return undefined
	const value = (summary as Record<string, unknown>)[key]
	return typeof value === 'number' && Number.isFinite(value) ? value : undefined
}

function buildMarkdown(
	args: Args,
	results: readonly StepResult[],
	worktree: WorktreeState,
	inputs: readonly InputProvenance[],
): string {
	const lines = [
		'# Ascend Practical Latency Contracts',
		'',
		`Generated: ${new Date().toISOString()}`,
		`Input workbook: \`${args.inputFile}\``,
		`Edit workbook: \`${args.editInputFile}\``,
		`Input provenance: ${inputProvenanceSummary(inputs)}`,
		`Timeout: ${args.timeoutMs}ms per step`,
		`Worktree: ${claimabilityLabel(worktree, inputs)} (${worktree.branchLine}; untracked entries ${worktree.untrackedCount})`,
		'',
		'## Contract Phase Map',
		'',
		'- Unknown workbook to safe first agent view: cold package scan, ZIP open, worksheet inflate/decode, capped parse and workbook hydration, first-window shaping, API payload shaping.',
		'- Edit plan to commit/reopen/verify: plan payload, write, output byte read/hash, reopened output hydration, structural check, formula lint, preservation summary, package graph verification.',
		'- Repeated viewport/table inspection after first load: process-local cache hit, window shaping, TUI render bytes, table metadata lookup, table row-window materialization, payload bytes.',
		'',
		'## Timeout And Progress Behavior',
		'',
		`Each step emits \`[contracts] start ...\` and a terminal status to stderr, writes stdout to a \`${basename(args.outDir)}/<contract>-<step>.json\` artifact, and is killed after ${args.timeoutMs}ms. A timeout or failed required step remains visible in this report instead of being silently dropped.`,
		'',
		'## Worktree Guardrail',
		'',
		worktreeGuardrail(worktree, inputs),
		'',
		'## Decision Matrix',
		'',
		'Only user-visible envelope phases are eligible for the production target. Diagnostic ceilings show max parse cost, but cannot win this ranking unless the matching envelope moves.',
		'',
		decisionMatrix(results),
		'',
		'## Production Target',
		'',
		productionTarget(results),
		'',
		'## Diagnostic Ceilings',
		'',
		diagnosticCeilingTable(results),
		'',
		'## Contract 1: Unknown Workbook To Safe First Agent View',
		'',
		'Guardrail: report cold package/ZIP/XML work separately from capped first-window results; do not call hot-cache numbers end-to-end, and do not treat the full worksheet scan diagnostic as first-view latency unless the first-view command actually performs it.',
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
		'## Guardrails Preventing Fake Speed',
		'',
		'- A hot-cache result must name its cache assumption and cannot be reported as end-to-end unknown workbook latency.',
		'- Full worksheet ZIP/XML scan measurements are diagnostic ceilings; they are not optimization targets for first view unless the capped first-view envelope moves too.',
		'- A write result is incomplete unless output bytes are written, reopened from bytes or disk, structurally checked, linted, and package graph verified.',
		'- Synthetic table or writer results must be labeled as generated/cached when workload generation is included or excluded.',
		'- Any optimization below 2% is noise; 2-5% requires remeasurement; above 5% still needs a profile-backed root cause and no memory/correctness regression.',
		'',
	]
	return `${lines.join('\n')}\n`
}

function buildProfileMarkdown(
	args: Args,
	results: readonly StepResult[],
	worktree: WorktreeState,
	inputs: readonly InputProvenance[],
): string {
	const lines = [
		'# Ascend Practical Contract Profile Summary',
		'',
		`Generated: ${new Date().toISOString()}`,
		`Output directory: \`${args.outDir}\``,
		`Input provenance: ${inputProvenanceSummary(inputs)}`,
		`Worktree: ${claimabilityLabel(worktree, inputs)} (${worktree.branchLine}; untracked entries ${worktree.untrackedCount})`,
		'',
		'## Worktree Guardrail',
		'',
		worktreeGuardrail(worktree, inputs),
		'',
		'## User-Visible Envelope Decisions',
		'',
		decisionMatrix(results),
		'',
		'## Diagnostic Ceilings',
		'',
		diagnosticCeilingTable(results),
		'',
		'## Memory And Payload',
		'',
		memoryPayloadTable(results),
		'',
		'## Production Target',
		'',
		productionTarget(results),
		'',
		'## Profile Commands To Run Before Code Changes',
		'',
		...results
			.filter((result) => result.profileCommand)
			.map((result) => `- ${result.contract}/${result.id}: \`${result.profileCommand}\``),
		'',
		'## Guardrail',
		'',
		'This artifact is a contract report, not an optimization claim. A production change must re-run the same contract subset on a compatible worktree state and show a median movement in one named envelope before it is kept.',
		'',
	]
	return `${lines.join('\n')}\n`
}

function claimabilityLabel(worktree: WorktreeState, inputs: readonly InputProvenance[]): string {
	if (worktree.trackedDirty) return 'tracked dirty; not release-claimable'
	return inputs.every((input) => input.releaseClaimable)
		? 'tracked clean with tracked inputs'
		: 'tracked clean with local/private inputs; diagnostic only'
}

function inputProvenanceSummary(inputs: readonly InputProvenance[]): string {
	return inputs
		.map((input) => {
			const status = input.releaseClaimable
				? 'tracked'
				: input.exists
					? input.tracked
						? 'tracked'
						: 'local/private'
					: 'missing'
			return `\`${input.path}\` (${status})`
		})
		.join(', ')
}

function worktreeGuardrail(worktree: WorktreeState, inputs: readonly InputProvenance[]): string {
	const localInputs = inputs.filter((input) => !input.releaseClaimable)
	if (!worktree.trackedDirty && localInputs.length === 0) {
		return `Tracked files and benchmark inputs were clean when this report was generated, so benchmark deltas can be compared against the recorded commit state. Untracked entries are recorded separately (${worktree.untrackedCount}) and must be treated as artifact context, not code changes.`
	}
	const tracked =
		worktree.trackedDirtyFiles.length > 0
			? worktree.trackedDirtyFiles.map((file) => `\`${file}\``).join(', ')
			: 'none'
	const inputsText =
		localInputs.length > 0 ? localInputs.map((input) => `\`${input.path}\``).join(', ') : 'none'
	return `Treat numbers as diagnostic only until rerun from a tracked-clean code state with tracked benchmark inputs, or until local inputs are explicitly documented as private diagnostics. Tracked dirty files: ${tracked}. Local/private or missing inputs: ${inputsText}. Untracked entries: ${worktree.untrackedCount}.`
}

function firstViewTable(results: readonly StepResult[]): string {
	const zip = results.find((entry) => entry.id === 'package-scan-zip-xml')
	const capped = results.find((entry) => entry.id === 'capped-read-window')
	const api = results.find((entry) => entry.id === 'api-first-view')
	return [
		'| Phase | Median ms | Memory / payload | Evidence |',
		'|---|---:|---:|---|',
		`| ZIP open | ${metric(zip?.summary, 'zipOpenMedianMs')} | | ${statusLink(zip)} |`,
		`| Full worksheet inflate diagnostic | ${metric(zip?.summary, 'worksheetInflateMedianMs')} | | ${statusLink(zip)} |`,
		`| Full worksheet decode diagnostic | ${metric(zip?.summary, 'worksheetDecodeMedianMs')} | | ${statusLink(zip)} |`,
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
		`| Prepared plan | ${metric(workflow?.summary, 'preparedPlanMedianMs')} | ${metric(workflow?.summary, 'preparedPlanPayloadBytesMedian')} bytes | ${statusLink(workflow)} |`,
		`| Commit verified total | ${metric(workflow?.summary, 'commitVerifiedTotalMedianMs')} | ${metric(workflow?.summary, 'commitVerifiedPayloadBytesMedian')} bytes | ${statusLink(workflow)} |`,
		`| Prepared commit verified total | ${metric(workflow?.summary, 'preparedCommitVerifiedTotalMedianMs')} | ${metric(workflow?.summary, 'preparedCommitVerifiedPayloadBytesMedian')} bytes | ${statusLink(workflow)} |`,
		`| Prepared write-policy snapshot | ${metric(workflow?.summary, 'preparedCommitWritePolicySnapshotMedianMs')} | | ${statusLink(workflow)} |`,
		`| Prepared write-policy check | ${metric(workflow?.summary, 'preparedCommitWritePolicyCheckMedianMs')} | | ${statusLink(workflow)} |`,
		`| Prepared reopen output | ${metric(workflow?.summary, 'preparedCommitPostWriteReopenMedianMs')} | | ${statusLink(workflow)} |`,
		`| Commit write-policy snapshot | ${metric(postWrite?.summary, 'commitWritePolicySnapshotMedianMs')} | | ${statusLink(postWrite)} |`,
		`| Commit package graph | ${metric(postWrite?.summary, 'commitPackageGraphMedianMs')} | | ${statusLink(postWrite)} |`,
		`| Commit write-plan summary | ${metric(postWrite?.summary, 'commitWritePlanSummaryMedianMs')} | | ${statusLink(postWrite)} |`,
		`| Commit write-policy check | ${metric(postWrite?.summary, 'commitWritePolicyCheckMedianMs')} | | ${statusLink(postWrite)} |`,
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
	const table = results.find((entry) => entry.id === 'table-inspection-cache')
	return [
		'| Phase | Cold median ms | Warm median ms | Payload/render | Evidence |',
		'|---|---:|---:|---:|---|',
		`| Cached document first window | ${metric(cached?.summary, 'cappedOpenWindowMedianMs')} | ${metric(cached?.summary, 'cappedWarmOpenWindowMedianMs')} | cells ${metric(cached?.summary, 'cellsMedian')} | ${statusLink(cached)} |`,
		`| TUI open/render first paint | ${metric(tui?.summary, 'tuiFirstPaintMedianMs')} | ${metric(tui?.summary, 'tuiWarmFirstPaintMedianMs')} | frame ${metric(tui?.summary, 'tuiFrameBytesMedian')} bytes | ${statusLink(tui)} |`,
		`| Table metadata and row window | ${metric(table?.summary, 'coldOpenMedianMs')} + inspect ${metric(table?.summary, 'coldInspectMedianMs')} | ${metric(table?.summary, 'warmOpenMedianMs')} + inspect ${metric(table?.summary, 'warmInspectMedianMs')} | payload ${metric(table?.summary, 'payloadBytesMedian')} bytes, tables ${metric(table?.summary, 'tableCountMedian')} | ${statusLink(table)} |`,
	].join('\n')
}

function statusLink(result: StepResult | undefined): string {
	if (!result) return 'not run'
	const path = result.outputPath ? basename(result.outputPath) : ''
	return result.status === 'ok' ? `ok (${path})` : `${result.status}: ${result.error ?? path}`
}

function productionTarget(results: readonly StepResult[]): string {
	const selected = envelopeDecisions(results)
		.filter((decision) => decision.phaseMedianMs > 0)
		.sort((a, b) => b.envelopeMedianMs - a.envelopeMedianMs || b.phaseMedianMs - a.phaseMedianMs)[0]
	if (!selected) {
		return 'No optimization selected yet: one or more required measurements failed or timed out.'
	}
	const share =
		selected.envelopeMedianMs > 0 ? (selected.phaseMedianMs / selected.envelopeMedianMs) * 100 : 100
	return `Choose exactly one production target: \`${selected.largestPhase}\` in \`${selected.contract}\`. It belongs to the largest true user-visible envelope (${selected.envelopeMedianMs.toFixed(1)}ms median), and users wait on ${selected.phaseMedianMs.toFixed(1)}ms inside it (${share.toFixed(1)}% max plausible win). Required profile before code changes: \`${selected.profileCommand}\`. Guardrail: ${selected.guardrail}`
}

function decisionMatrix(results: readonly StepResult[]): string {
	const rows = envelopeDecisions(results).map((decision) => {
		const share =
			decision.envelopeMedianMs > 0
				? `${((decision.phaseMedianMs / decision.envelopeMedianMs) * 100).toFixed(1)}%`
				: ''
		return `| ${decision.envelope} | ${decision.envelopeMedianMs.toFixed(3)} | ${decision.largestPhase} | ${decision.phaseMedianMs.toFixed(3)} | ${share} | \`${decision.profileCommand}\` | ${decision.guardrail} |`
	})
	return [
		'| Envelope | Envelope median ms | Largest user-wait phase | Phase median ms | Max plausible win | Required profile command | Guardrail |',
		'|---|---:|---|---:|---:|---|---|',
		...rows,
	].join('\n')
}

function diagnosticCeilingTable(results: readonly StepResult[]): string {
	const rows = diagnosticCeilings(results)
	return [
		'| Diagnostic phase | Median ms | Profile command | Target eligibility |',
		'|---|---:|---|---|',
		...rows.map(
			(row) =>
				`| ${row.name} | ${row.medianMs ? row.medianMs.toFixed(3) : ''} | ${row.profileCommand ? `\`${row.profileCommand}\`` : ''} | ${row.guardrail} |`,
		),
	].join('\n')
}

function diagnosticCeilings(results: readonly StepResult[]): DiagnosticCeiling[] {
	const zip = results.find((entry) => entry.id === 'package-scan-zip-xml')
	return [
		{
			name: 'Cold ZIP open',
			medianMs: numericMetric(zip?.summary, 'zipOpenMedianMs') ?? 0,
			profileCommand: zip?.profileCommand ?? '',
			guardrail: 'diagnostic only; not an envelope target unless first-view latency moves',
		},
		{
			name: 'Full worksheet XML inflate',
			medianMs: numericMetric(zip?.summary, 'worksheetInflateMedianMs') ?? 0,
			profileCommand: zip?.profileCommand ?? '',
			guardrail: 'diagnostic ceiling; do not chase for first view by itself',
		},
		{
			name: 'Full worksheet XML decode',
			medianMs: numericMetric(zip?.summary, 'worksheetDecodeMedianMs') ?? 0,
			profileCommand: zip?.profileCommand ?? '',
			guardrail: 'diagnostic ceiling; eligible only if a user-visible envelope performs this scan',
		},
	]
}

function envelopeDecisions(results: readonly StepResult[]): EnvelopeDecision[] {
	const capped = results.find((entry) => entry.id === 'capped-read-window')
	const api = results.find((entry) => entry.id === 'api-first-view')
	const workflow = results.find((entry) => entry.id === 'workflow-commit')
	const postWrite = results.find((entry) => entry.id === 'post-write-breakdown')
	const cached = results.find((entry) => entry.id === 'cached-agent-window')
	const tui = results.find((entry) => entry.id === 'tui-first-paint-cache')
	const table = results.find((entry) => entry.id === 'table-inspection-cache')

	const decisions: EnvelopeDecision[] = []
	const firstViewEnvelope = numericMetric(api?.summary, 'apiFirstWindowMedianMs')
	if (firstViewEnvelope !== undefined) {
		const largest = largestPhase([
			{
				name: 'API first-view payload/open',
				medianMs: firstViewEnvelope,
				profileCommand: api?.profileCommand ?? '',
			},
			{
				name: 'Capped read/hydrate feeding first view',
				medianMs: numericMetric(capped?.summary, 'cappedReadWindowMedianMs') ?? 0,
				profileCommand: capped?.profileCommand ?? api?.profileCommand ?? '',
			},
		])
		decisions.push({
			contract: 'first-view',
			envelope: 'Unknown workbook to safe first agent view',
			envelopeMedianMs: Math.max(firstViewEnvelope, largest.medianMs),
			largestPhase: largest.name,
			phaseMedianMs: largest.medianMs,
			profileCommand: largest.profileCommand,
			guardrail: 'must improve the first-view command, not only the full worksheet XML diagnostic',
		})
	}

	const editEnvelope =
		numericMetric(workflow?.summary, 'preparedCommitVerifiedTotalMedianMs') ??
		numericMetric(workflow?.summary, 'commitVerifiedTotalMedianMs') ??
		numericMetric(workflow?.summary, 'totalMedianMs')
	if (editEnvelope !== undefined) {
		const workflowProfile = workflow?.profileCommand ?? ''
		const postWriteProfile = postWrite?.profileCommand ?? workflowProfile
		const preparedDirtyWrite =
			(numericMetric(workflow?.summary, 'preparedCommitToBytesMedianMs') ??
				numericMetric(postWrite?.summary, 'commitToBytesMedianMs') ??
				0) +
			(numericMetric(workflow?.summary, 'preparedCommitWriteFileMedianMs') ??
				numericMetric(postWrite?.summary, 'commitWriteFileMedianMs') ??
				0)
		const largest = largestPhase([
			{
				name: 'Prepared plan/open',
				medianMs: numericMetric(workflow?.summary, 'preparedPlanMedianMs') ?? 0,
				profileCommand: workflowProfile,
			},
			{
				name: 'Prepared write-policy snapshot',
				medianMs:
					numericMetric(workflow?.summary, 'preparedCommitWritePolicySnapshotMedianMs') ??
					numericMetric(postWrite?.summary, 'commitWritePolicySnapshotMedianMs') ??
					0,
				profileCommand: workflowProfile || postWriteProfile,
			},
			{
				name: 'Prepared commit write-policy check',
				medianMs:
					numericMetric(workflow?.summary, 'preparedCommitWritePolicyCheckMedianMs') ??
					numericMetric(postWrite?.summary, 'commitWritePolicyCheckMedianMs') ??
					0,
				profileCommand: workflowProfile || postWriteProfile,
			},
			{
				name: 'Prepared dirty write bytes',
				medianMs: preparedDirtyWrite,
				profileCommand: workflowProfile,
			},
			{
				name: 'Prepared reopen written output',
				medianMs:
					numericMetric(workflow?.summary, 'preparedCommitPostWriteReopenMedianMs') ??
					numericMetric(postWrite?.summary, 'commitPostWriteReopenMedianMs') ??
					0,
				profileCommand: workflowProfile || postWriteProfile,
			},
			{
				name: 'Prepared structural check after reopen',
				medianMs:
					numericMetric(workflow?.summary, 'preparedCommitPostWriteCheckMedianMs') ??
					numericMetric(postWrite?.summary, 'commitPostWriteCheckMedianMs') ??
					0,
				profileCommand: workflowProfile || postWriteProfile,
			},
			{
				name: 'Prepared preservation/package verification',
				medianMs: Math.max(
					numericMetric(workflow?.summary, 'preparedCommitPostWritePreservationMedianMs') ??
						numericMetric(postWrite?.summary, 'commitPostWritePreservationMedianMs') ??
						0,
					numericMetric(workflow?.summary, 'preparedCommitPostWritePackageGraphAuditMedianMs') ??
						numericMetric(postWrite?.summary, 'commitPostWritePackageGraphAuditMedianMs') ??
						0,
				),
				profileCommand: workflowProfile || postWriteProfile,
			},
			{
				name: 'Prepared commit unassigned/finalize overhead',
				medianMs: preparedCommitUnassignedMs(workflow?.summary),
				profileCommand: workflowProfile,
			},
		])
		decisions.push({
			contract: 'edit-verify',
			envelope: 'Edit plan to commit/reopen/verify',
			envelopeMedianMs: editEnvelope,
			largestPhase: largest.name,
			phaseMedianMs: largest.medianMs,
			profileCommand: largest.profileCommand,
			guardrail: 'must preserve write, reopen, structural check, lint, and package verification',
		})
	}

	const repeatedCandidates: PhaseCandidate[] = [
		{
			name: 'Warm cached agent window',
			medianMs: numericMetric(cached?.summary, 'cappedWarmOpenWindowMedianMs') ?? 0,
			profileCommand: cached?.profileCommand ?? '',
		},
		{
			name: 'Warm TUI first paint',
			medianMs: numericMetric(tui?.summary, 'tuiWarmFirstPaintMedianMs') ?? 0,
			profileCommand: tui?.profileCommand ?? '',
		},
		{
			name: 'Warm table inspection',
			medianMs: numericMetric(table?.summary, 'warmInspectMedianMs') ?? 0,
			profileCommand: table?.profileCommand ?? '',
		},
	]
	const repeatedLargest = largestPhase(repeatedCandidates)
	if (repeatedLargest.medianMs > 0) {
		decisions.push({
			contract: 'repeated-inspection',
			envelope: 'Repeated viewport/table inspection after first load',
			envelopeMedianMs: repeatedLargest.medianMs,
			largestPhase: repeatedLargest.name,
			phaseMedianMs: repeatedLargest.medianMs,
			profileCommand: repeatedLargest.profileCommand,
			guardrail: 'must name the hot-cache assumption and include cold-open context separately',
		})
	}

	return decisions
}

function preparedCommitUnassignedMs(summary: unknown): number {
	const preparedCommit = numericMetric(summary, 'preparedCommitMedianMs')
	if (preparedCommit === undefined) return 0
	const known =
		sumMetrics(summary, [
			'preparedCommitWritePolicySnapshotMedianMs',
			'preparedCommitPackageGraphMedianMs',
			'preparedCommitPackageGraphAuditMedianMs',
			'preparedCommitApplyMedianMs',
			'preparedCommitWritePlanSummaryMedianMs',
			'preparedCommitWritePolicyCheckMedianMs',
			'preparedCommitWritePolicyBuildMedianMs',
			'preparedCommitToBytesMedianMs',
			'preparedCommitWriteFileMedianMs',
			'preparedCommitOutputByteReadMedianMs',
			'preparedCommitOutputHashMedianMs',
			'preparedCommitPostWriteReopenMedianMs',
			'preparedCommitPostWriteCheckMedianMs',
			'preparedCommitPostWriteLintMedianMs',
			'preparedCommitPostWritePreservationMedianMs',
			'preparedCommitPostWritePackageGraphMedianMs',
			'preparedCommitPostWritePackageGraphAuditMedianMs',
		]) ?? 0
	return Math.max(0, preparedCommit - known)
}

function sumMetrics(summary: unknown, keys: readonly string[]): number | undefined {
	let total = 0
	let found = false
	for (const key of keys) {
		const value = numericMetric(summary, key)
		if (value === undefined) continue
		total += value
		found = true
	}
	return found ? total : undefined
}

function largestPhase(candidates: readonly PhaseCandidate[]): PhaseCandidate {
	return (
		candidates
			.filter((candidate) => candidate.medianMs > 0)
			.sort((a, b) => b.medianMs - a.medianMs)[0] ?? {
			name: 'not measured',
			medianMs: 0,
			profileCommand: '',
		}
	)
}

function memoryPayloadTable(results: readonly StepResult[]): string {
	const capped = results.find((entry) => entry.id === 'capped-read-window')
	const api = results.find((entry) => entry.id === 'api-first-view')
	const workflow = results.find((entry) => entry.id === 'workflow-commit')
	const postWrite = results.find((entry) => entry.id === 'post-write-breakdown')
	const table = results.find((entry) => entry.id === 'table-inspection-cache')
	return [
		'| Item | Value | Guardrail |',
		'|---|---:|---|',
		`| First-view peak RSS | ${metric(capped?.summary, 'peakRssBytes')} bytes | cold capped read/hydrate only |`,
		`| API first-view payload | ${metric(api?.summary, 'payloadBytesMedian')} bytes | payload shaping is separate from cold XML parse |`,
		`| Commit payload | ${metric(workflow?.summary, 'commitVerifiedPayloadBytesMedian')} bytes | verified commit payload, not just write bytes |`,
		`| Post-write output | ${metric(postWrite?.summary, 'outputBytesMedian')} bytes | reopened and verified output |`,
		`| Table inspection payload | ${metric(table?.summary, 'payloadBytesMedian')} bytes | warm cache result requires cold-open context |`,
		`| Table inspection peak RSS | ${metric(table?.summary, 'peakRssBytes')} bytes | process-local table benchmark |`,
	].join('\n')
}

async function run() {
	const args = parseArgs()
	mkdirSync(args.outDir, { recursive: true })
	mkdirSync(join(args.outDir, 'profiles'), { recursive: true })
	const worktree = readWorktreeState()
	const inputs = [inputProvenance(args.inputFile), inputProvenance(args.editInputFile)]
	const steps = buildSteps(args)
	const results: StepResult[] = []
	for (const step of steps) {
		results.push(await runStep(step, args.outDir, args.dryRun))
	}
	const payload = {
		tool: 'practical-latency-contracts',
		args,
		worktree,
		inputs,
		results,
		decisions: envelopeDecisions(results),
		diagnosticCeilings: diagnosticCeilings(results),
	}
	const jsonPath = join(args.outDir, 'summary.json')
	const markdownPath = join(args.outDir, 'summary.md')
	const profileMarkdownPath = join(args.outDir, 'profile-summary.md')
	writeFileSync(jsonPath, JSON.stringify(payload, null, 2))
	writeFileSync(markdownPath, buildMarkdown(args, results, worktree, inputs))
	writeFileSync(profileMarkdownPath, buildProfileMarkdown(args, results, worktree, inputs))
	const output = {
		...payload,
		summaryJson: jsonPath,
		summaryMarkdown: markdownPath,
		profileSummaryMarkdown: profileMarkdownPath,
	}
	if (args.json) console.log(JSON.stringify(output, null, 2))
	else console.log(`summary: ${markdownPath}`)
}

function readWorktreeState(): WorktreeState {
	const result = Bun.spawnSync(['git', 'status', '--short', '--branch'], {
		cwd: process.cwd(),
		stdout: 'pipe',
		stderr: 'pipe',
	})
	if (result.exitCode !== 0) {
		return {
			branchLine: 'git-status-unavailable',
			dirty: true,
			trackedDirty: true,
			trackedDirtyFiles: [],
			untrackedCount: 0,
			status: [],
		}
	}
	const text = new TextDecoder().decode(result.stdout)
	const lines = text.trimEnd().split('\n').filter(Boolean)
	const branchLine = lines[0] ?? 'unknown'
	const entries = lines.slice(1)
	const trackedDirtyFiles = entries
		.filter((line) => !line.startsWith('?? '))
		.map((line) => line.slice(3).trim())
	const untrackedCount = entries.filter((line) => line.startsWith('?? ')).length
	return {
		branchLine,
		dirty: entries.length > 0,
		trackedDirty: trackedDirtyFiles.length > 0,
		trackedDirtyFiles,
		untrackedCount,
		status: entries,
	}
}

function inputProvenance(path: string): InputProvenance {
	const exists = existsSync(path)
	const tracked = isGitTracked(path)
	return {
		path,
		tracked,
		exists,
		releaseClaimable: exists && tracked,
	}
}

function isGitTracked(path: string): boolean {
	const result = Bun.spawnSync(['git', 'ls-files', '--error-unmatch', path], {
		cwd: process.cwd(),
		stdout: 'ignore',
		stderr: 'ignore',
	})
	return result.exitCode === 0
}

await run()
