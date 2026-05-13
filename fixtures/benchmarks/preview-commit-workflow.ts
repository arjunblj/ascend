#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import type { Operation } from '@ascend/schema'
import {
	commitAgentPlan,
	createAgentPlan,
	indexToColumn,
	WorkbookDocument,
} from '../../packages/sdk/src/index.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly previewRows: number
	readonly updates: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface Sample {
	readonly directFullFirstWindowMs: number
	readonly directFullPlanMs: number
	readonly directFullCommitMs: number
	readonly directFullTotalMs: number
	readonly previewFirstWindowMs: number
	readonly previewPlanMs: number
	readonly previewCommitMs: number
	readonly previewTotalMs: number
	readonly previewPromoteFirstWindowMs: number
	readonly previewPromoteMs: number
	readonly previewPromotePlanMs: number
	readonly previewPromoteCommitMs: number
	readonly previewPromoteTotalMs: number
	readonly directFullHydratedCells: number
	readonly previewHydratedCells: number
	readonly promotedHydratedCells: number
	readonly windowCells: number
	readonly directValid: boolean
	readonly previewValid: boolean
	readonly previewPromoteValid: boolean
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
		updates: positiveInt(readOption(process.argv, '--updates'), 1_000),
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

async function timed<T>(
	fn: () => Promise<T> | T,
): Promise<{ readonly ms: number; readonly value: T }> {
	const start = performance.now()
	const value = await fn()
	return { ms: performance.now() - start, value }
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function buildSetCellsOperation(count: number, rows: number, cols: number): Operation {
	const updates: Extract<Operation, { op: 'setCells' }>['updates'] = []
	for (let index = 0; index < count; index++) {
		const row = Math.floor(index / cols) % rows
		const col = index % cols
		updates.push({
			ref: `${indexToColumn(col)}${row + 1}`,
			value: 300_000 + index,
		})
	}
	return { op: 'setCells', sheet: 'Data', updates }
}

async function readFirstWindow(
	path: string,
	range: string,
	previewRows: number,
	mode: 'full' | 'preview',
) {
	WorkbookDocument.clearCache()
	const document = await timed(() =>
		WorkbookDocument.open(
			path,
			mode === 'full' ? { mode: 'full' } : { mode: 'values', maxRows: previewRows },
		),
	)
	const window = await timed(() =>
		document.value.readWindowCompact(document.value.sheets[0] ?? 'Data', range, {
			rowLimit: previewRows,
			includeRefs: false,
			omitEmpty: true,
			flatValues: true,
		}),
	)
	return {
		document: document.value,
		ms: document.ms + window.ms,
		hydratedCells: document.value.inspect().cellCount,
		windowCells: window.value?.cells.length ?? 0,
	}
}

async function planAndCommit(inputPath: string, outputPath: string, ops: readonly Operation[]) {
	await rm(outputPath, { force: true })
	const plan = await timed(() => createAgentPlan(inputPath, ops))
	const commit = await timed(() =>
		commitAgentPlan(inputPath, ops, { output: outputPath, approvals: [] }),
	)
	return {
		planMs: plan.ms,
		commitMs: commit.ms,
		valid: commit.value.postWrite.valid,
	}
}

async function runDirectFull(
	inputPath: string,
	outputPath: string,
	range: string,
	previewRows: number,
	ops: readonly Operation[],
) {
	const firstWindow = await readFirstWindow(inputPath, range, previewRows, 'full')
	const committed = await planAndCommit(inputPath, outputPath, ops)
	return {
		firstWindowMs: firstWindow.ms,
		hydratedCells: firstWindow.hydratedCells,
		windowCells: firstWindow.windowCells,
		...committed,
		totalMs: firstWindow.ms + committed.planMs + committed.commitMs,
	}
}

async function runPreviewThenCommit(
	inputPath: string,
	outputPath: string,
	range: string,
	previewRows: number,
	ops: readonly Operation[],
) {
	const firstWindow = await readFirstWindow(inputPath, range, previewRows, 'preview')
	const committed = await planAndCommit(inputPath, outputPath, ops)
	return {
		firstWindowMs: firstWindow.ms,
		hydratedCells: firstWindow.hydratedCells,
		windowCells: firstWindow.windowCells,
		...committed,
		totalMs: firstWindow.ms + committed.planMs + committed.commitMs,
	}
}

async function runPreviewPromoteThenCommit(
	inputPath: string,
	outputPath: string,
	range: string,
	previewRows: number,
	ops: readonly Operation[],
) {
	const firstWindow = await readFirstWindow(inputPath, range, previewRows, 'preview')
	const promoted = await timed(() => firstWindow.document.withLoad({ mode: 'full' }))
	const promotedCells = promoted.value.inspect().cellCount
	const committed = await planAndCommit(inputPath, outputPath, ops)
	return {
		firstWindowMs: firstWindow.ms,
		promoteMs: promoted.ms,
		previewHydratedCells: firstWindow.hydratedCells,
		promotedHydratedCells: promotedCells,
		windowCells: firstWindow.windowCells,
		...committed,
		totalMs: firstWindow.ms + promoted.ms + committed.planMs + committed.commitMs,
	}
}

async function runSample(
	inputPath: string,
	outputBasePath: string,
	range: string,
	previewRows: number,
	ops: readonly Operation[],
): Promise<Sample> {
	const direct = await runDirectFull(
		inputPath,
		`${outputBasePath}.direct.xlsx`,
		range,
		previewRows,
		ops,
	)
	runGc()
	const preview = await runPreviewThenCommit(
		inputPath,
		`${outputBasePath}.preview.xlsx`,
		range,
		previewRows,
		ops,
	)
	runGc()
	const previewPromote = await runPreviewPromoteThenCommit(
		inputPath,
		`${outputBasePath}.preview-promote.xlsx`,
		range,
		previewRows,
		ops,
	)
	runGc()
	return {
		directFullFirstWindowMs: direct.firstWindowMs,
		directFullPlanMs: direct.planMs,
		directFullCommitMs: direct.commitMs,
		directFullTotalMs: direct.totalMs,
		previewFirstWindowMs: preview.firstWindowMs,
		previewPlanMs: preview.planMs,
		previewCommitMs: preview.commitMs,
		previewTotalMs: preview.totalMs,
		previewPromoteFirstWindowMs: previewPromote.firstWindowMs,
		previewPromoteMs: previewPromote.promoteMs,
		previewPromotePlanMs: previewPromote.planMs,
		previewPromoteCommitMs: previewPromote.commitMs,
		previewPromoteTotalMs: previewPromote.totalMs,
		directFullHydratedCells: direct.hydratedCells,
		previewHydratedCells: preview.hydratedCells,
		promotedHydratedCells: previewPromote.promotedHydratedCells,
		windowCells: preview.windowCells,
		directValid: direct.valid,
		previewValid: preview.valid,
		previewPromoteValid: previewPromote.valid,
	}
}

function summarize(samples: readonly Sample[]) {
	const directFullTotalMedianMs = median(samples.map((sample) => sample.directFullTotalMs))
	const previewTotalMedianMs = median(samples.map((sample) => sample.previewTotalMs))
	const previewPromoteTotalMedianMs = median(samples.map((sample) => sample.previewPromoteTotalMs))
	return {
		directFullTotalMedianMs,
		previewTotalMedianMs,
		previewPromoteTotalMedianMs,
		previewWorkflowSpeedupVsDirectFull: directFullTotalMedianMs / previewTotalMedianMs,
		previewPromoteOverheadVsPreview: previewPromoteTotalMedianMs / previewTotalMedianMs,
		directFullFirstWindowMedianMs: median(samples.map((sample) => sample.directFullFirstWindowMs)),
		previewFirstWindowMedianMs: median(samples.map((sample) => sample.previewFirstWindowMs)),
		previewPromoteMedianMs: median(samples.map((sample) => sample.previewPromoteMs)),
		directFullPlanMedianMs: median(samples.map((sample) => sample.directFullPlanMs)),
		previewPlanMedianMs: median(samples.map((sample) => sample.previewPlanMs)),
		directFullCommitMedianMs: median(samples.map((sample) => sample.directFullCommitMs)),
		previewCommitMedianMs: median(samples.map((sample) => sample.previewCommitMs)),
		directFullHydratedCellsMedian: median(samples.map((sample) => sample.directFullHydratedCells)),
		previewHydratedCellsMedian: median(samples.map((sample) => sample.previewHydratedCells)),
		promotedHydratedCellsMedian: median(samples.map((sample) => sample.promotedHydratedCells)),
		windowCellsMedian: median(samples.map((sample) => sample.windowCells)),
		valid: samples.every(
			(sample) => sample.directValid && sample.previewValid && sample.previewPromoteValid,
		),
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const outputBasePath = `${data.xlsxPath}.preview-commit`
	const range = `A1:${indexToColumn(args.cols - 1)}${args.rows}`
	const ops = [buildSetCellsOperation(args.updates, args.rows, args.cols)]
	const samples: Sample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runSample(data.xlsxPath, outputBasePath, range, args.previewRows, ops)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			samples.push(await runSample(data.xlsxPath, outputBasePath, range, args.previewRows, ops))
			runGc()
		}
		const payload = {
			tool: 'preview-commit-workflow',
			args,
			summary: summarize(samples),
			samples,
		}
		if (args.json) console.log(JSON.stringify(payload, null, 2))
		else console.log(payload.summary)
	} finally {
		await rm(data.xlsxPath, { force: true })
		await rm(`${outputBasePath}.direct.xlsx`, { force: true })
		await rm(`${outputBasePath}.preview.xlsx`, { force: true })
		await rm(`${outputBasePath}.preview-promote.xlsx`, { force: true })
	}
}

await run()
