#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { runCli } from '../../apps/cli/src/index.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly previewRows: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface CliSample {
	readonly tuiFullMs: number
	readonly tuiFullOutputBytes: number
	readonly tuiPreviewMs: number
	readonly tuiPreviewOutputBytes: number
	readonly openDefaultMs: number
	readonly openDefaultOutputBytes: number
}

const WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'plain-text',
	'string-heavy',
	'sparse-wide',
])

function readOption(args: readonly string[], name: string): string | undefined {
	const index = args.indexOf(name)
	return index >= 0 ? args[index + 1] : undefined
}

function hasFlag(args: readonly string[], name: string): boolean {
	return args.includes(name)
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
	const workload = readOption(process.argv, '--workload') ?? 'mixed-10pct-text'
	if (!WORKLOADS.has(workload)) throw new Error(`Unsupported --workload "${workload}"`)
	return {
		rows: positiveInt(readOption(process.argv, '--rows'), 65_536),
		cols: positiveInt(readOption(process.argv, '--cols'), 10),
		previewRows: positiveInt(readOption(process.argv, '--preview-rows'), 500),
		workload: workload as WorkloadName,
		repeat: positiveInt(readOption(process.argv, '--repeat'), 5),
		warmup: nonNegativeInt(readOption(process.argv, '--warmup'), 1),
		json: hasFlag(process.argv, '--json'),
	}
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const middle = Math.floor(sorted.length / 2)
	const upper = sorted[middle] ?? 0
	return sorted.length % 2 === 1 ? upper : ((sorted[middle - 1] ?? upper) + upper) / 2
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

async function captureCli(args: readonly string[]) {
	const originalLog = console.log
	const originalError = console.error
	const stdout: string[] = []
	const stderr: string[] = []
	console.log = (...values: unknown[]) => {
		stdout.push(values.map(String).join(' '))
	}
	console.error = (...values: unknown[]) => {
		stderr.push(values.map(String).join(' '))
	}
	const start = performance.now()
	try {
		const exitCode = await runCli([...args])
		const ms = performance.now() - start
		const output = stdout.join('\n')
		const error = stderr.join('\n')
		if (exitCode !== 0) throw new Error(`${args.join(' ')} failed: ${error || output}`)
		return {
			ms,
			output,
			outputBytes: output.length,
		}
	} finally {
		console.log = originalLog
		console.error = originalError
	}
}

async function runSample(path: string, previewRows: number): Promise<CliSample> {
	const tuiFull = await captureCli(['tui', path])
	runGc()
	const tuiPreview = await captureCli(['tui', path, '--preview-rows', String(previewRows)])
	runGc()
	const openDefault = await captureCli(['open', path])
	runGc()
	return {
		tuiFullMs: tuiFull.ms,
		tuiFullOutputBytes: tuiFull.outputBytes,
		tuiPreviewMs: tuiPreview.ms,
		tuiPreviewOutputBytes: tuiPreview.outputBytes,
		openDefaultMs: openDefault.ms,
		openDefaultOutputBytes: openDefault.outputBytes,
	}
}

function summarize(samples: readonly CliSample[]) {
	const tuiFullMedianMs = median(samples.map((sample) => sample.tuiFullMs))
	const openDefaultMedianMs = median(samples.map((sample) => sample.openDefaultMs))
	const tuiPreviewMedianMs = median(samples.map((sample) => sample.tuiPreviewMs))
	return {
		tuiFullMedianMs,
		tuiPreviewMedianMs,
		openDefaultMedianMs,
		openDefaultSpeedupVsTuiFull: tuiFullMedianMs / openDefaultMedianMs,
		tuiPreviewSpeedupVsTuiFull: tuiFullMedianMs / tuiPreviewMedianMs,
		tuiFullOutputBytesMedian: median(samples.map((sample) => sample.tuiFullOutputBytes)),
		tuiPreviewOutputBytesMedian: median(samples.map((sample) => sample.tuiPreviewOutputBytes)),
		openDefaultOutputBytesMedian: median(samples.map((sample) => sample.openDefaultOutputBytes)),
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const samples: CliSample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runSample(data.xlsxPath, args.previewRows)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(await runSample(data.xlsxPath, args.previewRows))
			runGc()
		}
		const payload = {
			tool: 'cli-open-preview',
			args,
			summary: summarize(samples),
			samples,
		}
		if (args.json) console.log(JSON.stringify(payload, null, 2))
		else console.log(payload.summary)
	} finally {
		await rm(data.xlsxPath, { force: true })
	}
}

await run()
