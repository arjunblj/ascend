#!/usr/bin/env bun
import { fileURLToPath } from 'node:url'

type Surface = 'api' | 'both'
type TargetStatus = 'ok' | 'approval-required' | 'failed'

interface Args {
	readonly repeat: number
	readonly warmup: number
	readonly surface: Surface
	readonly target?: string
	readonly json: boolean
}

interface AgentWorkflowTarget {
	readonly name: string
	readonly file: string
	readonly rowLimit: number
	readonly mutations: number
	readonly approvals?: readonly string[] | 'all'
}

interface TargetResult {
	readonly name: string
	readonly file: string
	readonly status: TargetStatus
	readonly reproCommand: string
	readonly profileCommand: string
	readonly approvals?: readonly string[] | 'all'
	readonly summary?: {
		readonly totalMedianMs?: number
		readonly preparedCommitVerifiedTotalMedianMs?: number
		readonly payloadBytesMedian?: number
		readonly preparedCommitVerifiedPayloadBytesMedian?: number
		readonly compactHydratedOpenCountMedian?: number
		readonly preparedCommitVerifiedHydratedOpenCountMedian?: number
		readonly mcpPreparedCommitVerifiedTotalMedianMs?: number
		readonly mcpPreparedCommitVerifiedPayloadBytesMedian?: number
		readonly mcpPreparedCommitVerifiedHydratedOpenCountMedian?: number
		readonly valid?: boolean
		readonly preparedValid?: boolean
		readonly mcpValid?: boolean
		readonly mcpPreparedValid?: boolean
	}
	readonly error?: string
}

const WORKFLOW_SCRIPT = 'fixtures/benchmarks/agent-workflow.ts'
const WORKFLOW_SCRIPT_PATH = fileURLToPath(new URL('./agent-workflow.ts', import.meta.url))
const PROFILE_SCRIPT = 'fixtures/benchmarks/profile-bun.ts'

const TARGETS: readonly AgentWorkflowTarget[] = [
	{
		name: 'stress-dense-100k',
		file: 'fixtures/xlsx/stress/dense-100k.xlsx',
		rowLimit: 500,
		mutations: 1000,
	},
	{
		name: 'poi-with-various-data-approved',
		file: 'fixtures/xlsx/poi/WithVariousData.xlsx',
		rowLimit: 100,
		mutations: 25,
		approvals: 'all',
	},
	{
		name: 'poi-sample-ss',
		file: 'fixtures/xlsx/poi/SampleSS.xlsx',
		rowLimit: 5,
		mutations: 1,
	},
]

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

function parseArgs(): Args {
	const surface = readOption('--surface') ?? 'api'
	if (surface !== 'api' && surface !== 'both') {
		throw new Error(`Unsupported --surface "${surface}". Use api or both.`)
	}
	const target = readOption('--target')
	return {
		repeat: positiveInt(readOption('--repeat'), 3),
		warmup: nonNegativeInt(readOption('--warmup'), 1),
		surface,
		...(target !== undefined ? { target } : {}),
		json: hasFlag('--json'),
	}
}

function selectedTargets(args: Args): readonly AgentWorkflowTarget[] {
	if (args.target === undefined || args.target === 'all') return TARGETS
	const selected = TARGETS.filter((target) => target.name === args.target)
	if (selected.length === 0) throw new Error(`Unknown --target "${args.target}"`)
	return selected
}

function approvalArgs(approvals: readonly string[] | 'all' | undefined): string[] {
	if (approvals === undefined) return []
	return ['--approval', approvals === 'all' ? 'all' : approvals.join(',')]
}

function workflowScriptArgs(script: string, target: AgentWorkflowTarget, args: Args): string[] {
	return [
		script,
		'--input-file',
		target.file,
		'--surface',
		args.surface,
		'--row-limit',
		String(target.rowLimit),
		'--mutations',
		String(target.mutations),
		'--repeat',
		String(args.repeat),
		'--warmup',
		String(args.warmup),
		...approvalArgs(target.approvals),
		'--json',
	]
}

function workflowArgs(target: AgentWorkflowTarget, args: Args): string[] {
	return workflowScriptArgs(WORKFLOW_SCRIPT_PATH, target, args)
}

function workflowCommandArgs(target: AgentWorkflowTarget, args: Args): string[] {
	return ['bun', 'run', ...workflowScriptArgs(WORKFLOW_SCRIPT, target, args)]
}

function targetCommands(
	target: AgentWorkflowTarget,
	args: Args,
): Pick<TargetResult, 'reproCommand' | 'profileCommand'> {
	const reproArgs = workflowCommandArgs(target, args)
	return {
		reproCommand: commandString(reproArgs),
		profileCommand: commandString([
			'bun',
			'run',
			PROFILE_SCRIPT,
			'--mode',
			'all-md',
			'--label',
			`agent-workflow-${target.name}`,
			'--',
			...reproArgs,
		]),
	}
}

function commandString(args: readonly string[]): string {
	return args.map(shellQuote).join(' ')
}

function shellQuote(value: string): string {
	return /^[A-Za-z0-9_./:=@+-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function classifyFailure(stderr: string): TargetStatus {
	return /Commit requires explicit approval|approval/i.test(stderr) ? 'approval-required' : 'failed'
}

async function runTarget(target: AgentWorkflowTarget, args: Args): Promise<TargetResult> {
	const commands = targetCommands(target, args)
	const proc = Bun.spawn([Bun.argv[0], ...workflowArgs(target, args)], {
		cwd: process.cwd(),
		stderr: 'pipe',
		stdout: 'pipe',
	})
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		return {
			name: target.name,
			file: target.file,
			status: classifyFailure(stderr),
			...commands,
			...(target.approvals !== undefined ? { approvals: target.approvals } : {}),
			error: stderr.trim().split('\n').slice(0, 6).join('\n'),
		}
	}
	const payload = JSON.parse(stdout) as { readonly summary?: TargetResult['summary'] }
	return {
		name: target.name,
		file: target.file,
		status: 'ok',
		...commands,
		...(target.approvals !== undefined ? { approvals: target.approvals } : {}),
		summary: payload.summary,
	}
}

function summarize(results: readonly TargetResult[]) {
	const ok = results.filter((result) => result.status === 'ok')
	return {
		targetCount: results.length,
		okCount: ok.length,
		approvalRequiredCount: results.filter((result) => result.status === 'approval-required').length,
		failedCount: results.filter((result) => result.status === 'failed').length,
		validCount: ok.filter((result) => result.summary?.valid === true).length,
		preparedValidCount: ok.filter((result) => result.summary?.preparedValid === true).length,
		mcpValidCount: ok.filter((result) => result.summary?.mcpValid === true).length,
		mcpPreparedValidCount: ok.filter((result) => result.summary?.mcpPreparedValid === true).length,
	}
}

async function run() {
	const args = parseArgs()
	const targets = selectedTargets(args)
	const results: TargetResult[] = []
	for (const target of targets) results.push(await runTarget(target, args))
	const payload = {
		tool: 'agent-workflow-corpus',
		args,
		targets: targets.map((target) => target.name),
		summary: summarize(results),
		results,
	}
	if (args.json) console.log(JSON.stringify(payload, null, 2))
	else console.log(payload.summary)
}

await run()
