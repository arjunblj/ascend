#!/usr/bin/env bun
import { createHash } from 'node:crypto'
import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { indexToColumn } from '../../packages/core/src/index.ts'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import { readXlsx } from '../../packages/io-xlsx/src/index.ts'
import type { CellValue } from '../../packages/schema/src/index.ts'
import { topLeftScalar } from '../../packages/schema/src/index.ts'
import {
	type CorpusBenchmarkTier,
	type CorpusManifestEntry,
	type NormalizedCorpusManifestEntry,
	normalizeManifest,
	selectManifestEntries,
} from '../corpus/manifest.ts'

type OracleMode = 'cached-values'

interface Args {
	readonly corpusRoot: string
	readonly manifest: string
	readonly file?: string
	readonly tags: readonly string[]
	readonly tiers: readonly CorpusBenchmarkTier[]
	readonly maxWorkbooks?: number
	readonly sampleSeed: number
	readonly oracle: OracleMode
	readonly json: boolean
	readonly maxMismatches?: number
}

interface FormulaSnapshot {
	readonly ref: string
	readonly formula: string
	readonly cached: string
}

interface WorkbookFormulaResult {
	readonly file: string
	readonly source?: string
	readonly sourceUrl?: string
	readonly oracle: OracleMode
	readonly formulaCount: number
	readonly comparedCount: number
	readonly mismatchCount: number
	readonly errorCount: number
	readonly beforeHash: string
	readonly afterHash: string
	readonly mismatches: readonly FormulaMismatch[]
	readonly readMs: number
	readonly recalcMs: number
}

interface FormulaMismatch {
	readonly ref: string
	readonly formula: string
	readonly cached: string
	readonly calculated: string
}

interface SuitePayload {
	readonly formatVersion: 1
	readonly suite: 'ascend-formula-corpus-correctness'
	readonly generatedAt: string
	readonly oracle: OracleMode
	readonly manifest: string
	readonly corpusRoot: string
	readonly selection: {
		readonly file?: string
		readonly tags: readonly string[]
		readonly tiers: readonly CorpusBenchmarkTier[]
		readonly maxWorkbooks?: number
		readonly sampleSeed: number
	}
	readonly results: readonly WorkbookFormulaResult[]
	readonly summary: {
		readonly workbookCount: number
		readonly formulaCount: number
		readonly comparedCount: number
		readonly mismatchCount: number
		readonly errorCount: number
		readonly perfectWorkbookCount: number
	}
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	return index >= 0 ? process.argv[index + 1] : undefined
}

function readRepeatedFlag(name: string): string[] {
	const values: string[] = []
	for (let i = 0; i < process.argv.length; i++) {
		if (process.argv[i] === name && process.argv[i + 1]) values.push(process.argv[i + 1] as string)
	}
	return values
}

function hasFlag(name: string): boolean {
	return process.argv.includes(name)
}

function positiveInt(raw: string | undefined): number | undefined {
	if (raw === undefined) return undefined
	const value = Number.parseInt(raw, 10)
	if (!Number.isFinite(value) || value <= 0) throw new Error(`Expected positive integer for ${raw}`)
	return value
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	if (raw === undefined) return fallback
	const value = Number.parseInt(raw, 10)
	if (!Number.isFinite(value) || value < 0)
		throw new Error(`Expected non-negative integer for ${raw}`)
	return value
}

function readArgs(): Args {
	const corpusRoot = resolve(readFlag('--corpus-root') ?? 'research/excel-corpus')
	return {
		corpusRoot,
		manifest: resolve(readFlag('--manifest') ?? `${corpusRoot}/manifest.json`),
		...(readFlag('--file') ? { file: readFlag('--file') } : {}),
		tags: readRepeatedFlag('--tag'),
		tiers: readRepeatedFlag('--tier') as CorpusBenchmarkTier[],
		maxWorkbooks: positiveInt(readFlag('--max-workbooks')),
		sampleSeed: nonNegativeInt(readFlag('--sample-seed'), 1),
		oracle: 'cached-values',
		json: hasFlag('--json'),
		maxMismatches: positiveInt(readFlag('--max-mismatches')),
	}
}

function hashLines(lines: readonly string[]): string {
	const hash = createHash('sha256')
	for (const line of lines) hash.update(line).update('\n')
	return hash.digest('hex')
}

function serializeValue(value: CellValue): string {
	const scalar = topLeftScalar(value)
	switch (scalar.kind) {
		case 'empty':
			return 'empty'
		case 'number':
			return `n:${canonicalNumber(scalar.value)}`
		case 'date':
			return `d:${canonicalNumber(scalar.serial)}`
		case 'string':
			return `s:${scalar.value}`
		case 'boolean':
			return `b:${scalar.value ? 'true' : 'false'}`
		case 'error':
			return `e:${scalar.value}`
		case 'richText':
			return `s:${scalar.runs.map((run) => run.text).join('')}`
	}
}

function canonicalNumber(value: number): string {
	if (Object.is(value, -0)) return '0'
	if (Number.isInteger(value)) return String(value)
	return value.toPrecision(15).replace(/(?:\.0+|(\.\d*?)0+)$/, '$1')
}

function formatRef(sheetName: string, row: number, col: number): string {
	return `${quoteSheet(sheetName)}!${indexToColumn(col)}${row + 1}`
}

function quoteSheet(sheetName: string): string {
	return /^[A-Za-z_]\w*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`
}

function collectFormulaSnapshots(workbook: {
	readonly sheets: readonly {
		readonly name: string
		readonly cells: {
			iterate(): Iterable<
				readonly [number, number, { readonly value: CellValue; readonly formula: string | null }]
			>
		}
	}[]
}): FormulaSnapshot[] {
	const snapshots: FormulaSnapshot[] = []
	for (const sheet of workbook.sheets) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formula) continue
			snapshots.push({
				ref: formatRef(sheet.name, row, col),
				formula: cell.formula,
				cached: serializeValue(cell.value),
			})
		}
	}
	return snapshots.sort((a, b) => a.ref.localeCompare(b.ref))
}

function compareSnapshots(
	before: readonly FormulaSnapshot[],
	after: readonly FormulaSnapshot[],
): readonly FormulaMismatch[] {
	const afterByRef = new Map(after.map((entry) => [entry.ref, entry]))
	const mismatches: FormulaMismatch[] = []
	for (const entry of before) {
		const calculated = afterByRef.get(entry.ref)
		if (!calculated) {
			mismatches.push({ ...entry, calculated: 'missing' })
			continue
		}
		if (entry.cached !== calculated.cached) {
			mismatches.push({ ...entry, calculated: calculated.cached })
		}
	}
	return mismatches
}

async function loadManifest(path: string): Promise<readonly NormalizedCorpusManifestEntry[]> {
	const raw = await readFile(path, 'utf-8')
	return normalizeManifest(JSON.parse(raw) as CorpusManifestEntry[])
}

function seededShuffle<T>(items: readonly T[], seed: number): T[] {
	const result = [...items]
	let state = seed >>> 0
	for (let i = result.length - 1; i > 0; i--) {
		state = (Math.imul(state, 1664525) + 1013904223) >>> 0
		const j = state % (i + 1)
		const tmp = result[i] as T
		result[i] = result[j] as T
		result[j] = tmp
	}
	return result
}

async function runWorkbook(
	entry: NormalizedCorpusManifestEntry,
	args: Pick<Args, 'corpusRoot' | 'oracle'>,
): Promise<WorkbookFormulaResult> {
	const path = resolve(args.corpusRoot, entry.file)
	const bytes = new Uint8Array(await readFile(path))
	const readStart = performance.now()
	const read = readXlsx(bytes)
	const readMs = performance.now() - readStart
	if (!read.ok) {
		return {
			file: entry.file,
			...(entry.source ? { source: entry.source } : {}),
			...(entry.sourceUrl ? { sourceUrl: entry.sourceUrl } : {}),
			oracle: args.oracle,
			formulaCount: 0,
			comparedCount: 0,
			mismatchCount: 0,
			errorCount: 1,
			beforeHash: hashLines([]),
			afterHash: hashLines([]),
			mismatches: [],
			readMs,
			recalcMs: 0,
		}
	}
	const before = collectFormulaSnapshots(read.value.workbook)
	const recalcStart = performance.now()
	const recalc = recalculate(read.value.workbook, defaultCalcContext())
	const recalcMs = performance.now() - recalcStart
	const after = collectFormulaSnapshots(read.value.workbook)
	const mismatches = compareSnapshots(before, after)
	return {
		file: entry.file,
		...(entry.source ? { source: entry.source } : {}),
		...(entry.sourceUrl ? { sourceUrl: entry.sourceUrl } : {}),
		oracle: args.oracle,
		formulaCount: before.length,
		comparedCount: before.length,
		mismatchCount: mismatches.length,
		errorCount: recalc.errors.length,
		beforeHash: hashLines(before.map((entry) => `${entry.ref}\t${entry.formula}\t${entry.cached}`)),
		afterHash: hashLines(after.map((entry) => `${entry.ref}\t${entry.formula}\t${entry.cached}`)),
		mismatches: mismatches.slice(0, 50),
		readMs,
		recalcMs,
	}
}

function summarize(results: readonly WorkbookFormulaResult[]): SuitePayload['summary'] {
	return {
		workbookCount: results.length,
		formulaCount: results.reduce((sum, result) => sum + result.formulaCount, 0),
		comparedCount: results.reduce((sum, result) => sum + result.comparedCount, 0),
		mismatchCount: results.reduce((sum, result) => sum + result.mismatchCount, 0),
		errorCount: results.reduce((sum, result) => sum + result.errorCount, 0),
		perfectWorkbookCount: results.filter(
			(result) => result.mismatchCount === 0 && result.errorCount === 0,
		).length,
	}
}

export async function runFormulaCorpusCorrectness(args: Args): Promise<SuitePayload> {
	if (!existsSync(args.manifest)) {
		throw new Error(`Missing formula corpus manifest: ${args.manifest}`)
	}
	const manifest = await loadManifest(args.manifest)
	const selected = selectManifestEntries(manifest, {
		...(args.file ? { file: args.file } : {}),
		...(args.tags.length > 0 ? { tags: args.tags } : {}),
		...(args.tiers.length > 0 ? { tiers: args.tiers } : {}),
	})
	const sampled = seededShuffle(selected, args.sampleSeed).slice(0, args.maxWorkbooks)
	const results: WorkbookFormulaResult[] = []
	for (const entry of sampled) {
		results.push(await runWorkbook(entry, args))
	}
	return {
		formatVersion: 1,
		suite: 'ascend-formula-corpus-correctness',
		generatedAt: new Date().toISOString(),
		oracle: args.oracle,
		manifest: args.manifest,
		corpusRoot: args.corpusRoot,
		selection: {
			...(args.file ? { file: args.file } : {}),
			tags: args.tags,
			tiers: args.tiers,
			...(args.maxWorkbooks ? { maxWorkbooks: args.maxWorkbooks } : {}),
			sampleSeed: args.sampleSeed,
		},
		results,
		summary: summarize(results),
	}
}

function render(payload: SuitePayload): void {
	console.log(
		`formula corpus correctness: workbooks=${payload.summary.workbookCount} formulas=${payload.summary.formulaCount} mismatches=${payload.summary.mismatchCount} errors=${payload.summary.errorCount}`,
	)
	for (const result of payload.results) {
		console.log(
			`${result.file}: formulas=${result.formulaCount} mismatches=${result.mismatchCount} errors=${result.errorCount} read=${result.readMs.toFixed(2)}ms recalc=${result.recalcMs.toFixed(2)}ms`,
		)
	}
}

if (import.meta.main) {
	const args = readArgs()
	const payload = await runFormulaCorpusCorrectness(args)
	if (args.json) console.log(JSON.stringify(payload, null, 2))
	else render(payload)
	if (args.maxMismatches !== undefined && payload.summary.mismatchCount > args.maxMismatches) {
		console.error(
			`formula corpus correctness failed: mismatches ${payload.summary.mismatchCount} exceeded ${args.maxMismatches}`,
		)
		process.exit(1)
	}
}
