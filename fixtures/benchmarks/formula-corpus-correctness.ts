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
	loadCorpusManifestEntries,
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
	readonly maxReportedMismatches?: number
	readonly sampleSeed: number
	readonly oracle: OracleMode
	readonly json: boolean
	readonly maxMismatches?: number
	readonly maxAcceptedMismatches?: number
	readonly maxUnacceptedMismatches?: number
	readonly maxSemanticMismatches?: number
	readonly maxVolatileOracleSkips?: number
	readonly maxErrors?: number
	readonly minWorkbooks?: number
	readonly minFormulas?: number
	readonly minComparedFormulas?: number
	readonly minPerfectWorkbooks?: number
	readonly minSemanticPerfectWorkbooks?: number
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
	readonly noCachedFormulaCount: number
	readonly volatileOracleSkipCount: number
	readonly mismatchCount: number
	readonly acceptedMismatchCount: number
	readonly unacceptedMismatchCount: number
	readonly semanticMismatchCount: number
	readonly numericDriftMismatchCount: number
	readonly errorCount: number
	readonly beforeHash: string
	readonly afterHash: string
	readonly mismatches: readonly FormulaMismatch[]
	readonly volatileOracleSkips: readonly FormulaOracleSkip[]
	readonly readMs: number
	readonly recalcMs: number
}

interface FormulaMismatch {
	readonly ref: string
	readonly formula: string
	readonly cached: string
	readonly calculated: string
	readonly classification: FormulaMismatchClassification
	readonly reason: string
}

interface FormulaOracleSkip {
	readonly ref: string
	readonly formula: string
	readonly cached: string
	readonly calculated: string
	readonly reason: string
}

interface FormulaComparison {
	readonly mismatches: readonly FormulaMismatch[]
	readonly volatileOracleSkips: readonly FormulaOracleSkip[]
}

type FormulaMismatchClassification = 'semantic' | 'numeric-drift'

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
		readonly maxReportedMismatches?: number
		readonly sampleSeed: number
	}
	readonly results: readonly WorkbookFormulaResult[]
	readonly summary: {
		readonly workbookCount: number
		readonly formulaCount: number
		readonly comparedCount: number
		readonly noCachedFormulaCount: number
		readonly volatileOracleSkipCount: number
		readonly mismatchCount: number
		readonly acceptedMismatchCount: number
		readonly unacceptedMismatchCount: number
		readonly semanticMismatchCount: number
		readonly numericDriftMismatchCount: number
		readonly errorCount: number
		readonly perfectWorkbookCount: number
		readonly semanticPerfectWorkbookCount: number
	}
}

const NUMERIC_DRIFT_ABS_TOLERANCE = 1e-12
const NUMERIC_DRIFT_REL_TOLERANCE = 2e-8
const NON_DETERMINISTIC_FUNCTION_RE =
	/(?:^|[^A-Z0-9_.])(?:NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY)\s*\(/i

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

function nonNegativeIntOptional(raw: string | undefined): number | undefined {
	if (raw === undefined) return undefined
	const value = Number.parseInt(raw, 10)
	if (!Number.isFinite(value) || value < 0)
		throw new Error(`Expected non-negative integer for ${raw}`)
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
		maxReportedMismatches: nonNegativeInt(readFlag('--max-reported-mismatches'), 50),
		sampleSeed: nonNegativeInt(readFlag('--sample-seed'), 1),
		oracle: 'cached-values',
		json: hasFlag('--json'),
		maxMismatches: nonNegativeIntOptional(readFlag('--max-mismatches')),
		maxAcceptedMismatches: nonNegativeIntOptional(readFlag('--max-accepted-mismatches')),
		maxUnacceptedMismatches: nonNegativeIntOptional(readFlag('--max-unaccepted-mismatches')),
		maxSemanticMismatches: nonNegativeIntOptional(readFlag('--max-semantic-mismatches')),
		maxVolatileOracleSkips: nonNegativeIntOptional(readFlag('--max-volatile-oracle-skips')),
		maxErrors: nonNegativeIntOptional(readFlag('--max-errors')),
		minWorkbooks: positiveInt(readFlag('--min-workbooks')),
		minFormulas: positiveInt(readFlag('--min-formulas')),
		minComparedFormulas: positiveInt(readFlag('--min-compared-formulas')),
		minPerfectWorkbooks: positiveInt(readFlag('--min-perfect-workbooks')),
		minSemanticPerfectWorkbooks: positiveInt(readFlag('--min-semantic-perfect-workbooks')),
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
			return `s:${canonicalString(scalar.value)}`
		case 'boolean':
			return `b:${scalar.value ? 'true' : 'false'}`
		case 'error':
			return `e:${scalar.value}`
		case 'richText':
			return `s:${canonicalString(scalar.runs.map((run) => run.text).join(''))}`
	}
}

function canonicalString(value: string): string {
	let result = ''
	for (let index = 0; index < value.length; index++) {
		const char = value[index] ?? ''
		const code = char.charCodeAt(0)
		if ((code >= 0 && code <= 8) || code === 11 || code === 12 || (code >= 14 && code <= 31)) {
			result += `_x${code.toString(16).toUpperCase().padStart(4, '0')}_`
		} else {
			result += char
		}
	}
	return result
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
): FormulaComparison {
	const afterByRef = new Map(after.map((entry) => [entry.ref, entry]))
	const mismatches: FormulaMismatch[] = []
	const volatileOracleSkips: FormulaOracleSkip[] = []
	for (const entry of before) {
		if (entry.cached === 'empty') continue
		const calculated = afterByRef.get(entry.ref)
		if (!calculated) {
			mismatches.push({
				...entry,
				calculated: 'missing',
				classification: 'semantic',
				reason: 'formula cell missing after recalculation',
			})
			continue
		}
		if (entry.cached !== calculated.cached) {
			if (NON_DETERMINISTIC_FUNCTION_RE.test(entry.formula)) {
				volatileOracleSkips.push({
					...entry,
					calculated: calculated.cached,
					reason:
						'formula uses time or random functions whose historical cached values are not reproducible',
				})
				continue
			}
			mismatches.push({
				...entry,
				calculated: calculated.cached,
				...classifyMismatch(entry, calculated),
			})
		}
	}
	return classifyDownstreamComparisons(mismatches, volatileOracleSkips)
}

function classifyMismatch(
	cached: FormulaSnapshot,
	calculated: FormulaSnapshot,
): Pick<FormulaMismatch, 'classification' | 'reason'> {
	const cachedNumber = parseSerializedNumber(cached.cached)
	const calculatedNumber = parseSerializedNumber(calculated.cached)
	if (cachedNumber && calculatedNumber && cachedNumber.kind === calculatedNumber.kind) {
		const diff = Math.abs(cachedNumber.value - calculatedNumber.value)
		const scale = Math.max(1, Math.abs(cachedNumber.value), Math.abs(calculatedNumber.value))
		if (diff <= NUMERIC_DRIFT_ABS_TOLERANCE || diff <= scale * NUMERIC_DRIFT_REL_TOLERANCE) {
			return {
				classification: 'numeric-drift',
				reason: `numeric difference ${diff} within abs=${NUMERIC_DRIFT_ABS_TOLERANCE} or rel=${NUMERIC_DRIFT_REL_TOLERANCE}`,
			}
		}
	}

	const cachedComplex = parseSerializedComplex(cached.cached)
	const calculatedComplex = parseSerializedComplex(calculated.cached)
	if (cachedComplex && calculatedComplex && cachedComplex.suffix === calculatedComplex.suffix) {
		const realDiff = Math.abs(cachedComplex.real - calculatedComplex.real)
		const imagDiff = Math.abs(cachedComplex.imag - calculatedComplex.imag)
		const realScale = Math.max(1, Math.abs(cachedComplex.real), Math.abs(calculatedComplex.real))
		const imagScale = Math.max(1, Math.abs(cachedComplex.imag), Math.abs(calculatedComplex.imag))
		if (
			(realDiff <= NUMERIC_DRIFT_ABS_TOLERANCE ||
				realDiff <= realScale * NUMERIC_DRIFT_REL_TOLERANCE) &&
			(imagDiff <= NUMERIC_DRIFT_ABS_TOLERANCE ||
				imagDiff <= imagScale * NUMERIC_DRIFT_REL_TOLERANCE)
		) {
			return {
				classification: 'numeric-drift',
				reason: `complex numeric difference real=${realDiff} imag=${imagDiff} within abs=${NUMERIC_DRIFT_ABS_TOLERANCE} or rel=${NUMERIC_DRIFT_REL_TOLERANCE}`,
			}
		}
	}

	return {
		classification: 'semantic',
		reason: 'cached and calculated values differ beyond tolerance',
	}
}

function classifyDownstreamComparisons(
	mismatches: readonly FormulaMismatch[],
	volatileOracleSkips: readonly FormulaOracleSkip[],
): FormulaComparison {
	let current = [...mismatches]
	const volatile = [...volatileOracleSkips]
	let changed = true
	while (changed) {
		changed = false
		const byRef = new Map(current.map((mismatch) => [mismatch.ref, mismatch]))
		const volatileRefs = new Set(volatile.map((skip) => skip.ref))
		const next: FormulaMismatch[] = []
		for (const mismatch of current) {
			if (mismatch.classification !== 'semantic') {
				next.push(mismatch)
				continue
			}
			let replacement: FormulaMismatch | null = null
			let movedToVolatile = false
			for (const ref of extractFormulaReferences(mismatch.formula, mismatch.ref)) {
				if (volatileRefs.has(ref)) {
					changed = true
					volatile.push({
						ref: mismatch.ref,
						formula: mismatch.formula,
						cached: mismatch.cached,
						calculated: mismatch.calculated,
						reason: `downstream of volatile oracle skip ${ref}`,
					})
					movedToVolatile = true
					break
				}
				const precedent = byRef.get(ref)
				if (!precedent) continue
				if (precedent.classification === 'numeric-drift') {
					changed = true
					replacement = {
						...mismatch,
						classification: 'numeric-drift',
						reason: `downstream of numeric-drift precedent ${ref}`,
					}
					break
				}
			}
			if (movedToVolatile) continue
			next.push(replacement ?? mismatch)
		}
		current = next
	}
	return { mismatches: current, volatileOracleSkips: volatile }
}

function extractFormulaReferences(formula: string, currentRef: string): readonly string[] {
	const currentSheet = sheetNameFromRef(currentRef)
	const refs = new Set<string>()
	const refPattern =
		/(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_ .]*))!)?(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?/g
	for (const match of formula.matchAll(refPattern)) {
		const rawSheet = match[1]?.replace(/''/g, "'") ?? match[2] ?? currentSheet
		const first = normalizeA1Ref(match[3] ?? '')
		const second = normalizeA1Ref(match[4] ?? '')
		if (rawSheet && first) refs.add(`${quoteSheet(rawSheet)}!${first}`)
		if (rawSheet && second) refs.add(`${quoteSheet(rawSheet)}!${second}`)
	}
	return [...refs]
}

function sheetNameFromRef(ref: string): string {
	const separator = ref.lastIndexOf('!')
	const rawSheet = separator >= 0 ? ref.slice(0, separator) : ''
	if (rawSheet.startsWith("'") && rawSheet.endsWith("'")) {
		return rawSheet.slice(1, -1).replace(/''/g, "'")
	}
	return rawSheet
}

function normalizeA1Ref(ref: string): string {
	return ref.replace(/\$/g, '').toUpperCase()
}

function parseSerializedNumber(serialized: string): { kind: 'n' | 'd'; value: number } | null {
	const match = /^(n|d):(.+)$/.exec(serialized)
	if (!match) return null
	const value = Number(match[2])
	if (!Number.isFinite(value)) return null
	return { kind: match[1] as 'n' | 'd', value }
}

function parseSerializedComplex(
	serialized: string,
): { real: number; imag: number; suffix: 'i' | 'j' } | null {
	const numberPattern = String.raw`[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?`
	const match = new RegExp(`^s:(${numberPattern})(${numberPattern})([ij])$`).exec(serialized)
	if (!match) return null
	const real = Number(match[1])
	const imag = Number(match[2])
	if (!Number.isFinite(real) || !Number.isFinite(imag)) return null
	return { real, imag, suffix: match[3] as 'i' | 'j' }
}

async function loadManifest(path: string): Promise<readonly NormalizedCorpusManifestEntry[]> {
	return normalizeManifest(await loadCorpusManifestEntries(path))
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
	args: Pick<Args, 'corpusRoot' | 'maxReportedMismatches' | 'oracle'>,
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
			noCachedFormulaCount: 0,
			volatileOracleSkipCount: 0,
			mismatchCount: 0,
			acceptedMismatchCount: 0,
			unacceptedMismatchCount: 0,
			semanticMismatchCount: 0,
			numericDriftMismatchCount: 0,
			errorCount: 1,
			beforeHash: hashLines([]),
			afterHash: hashLines([]),
			mismatches: [],
			volatileOracleSkips: [],
			readMs,
			recalcMs: 0,
		}
	}
	const before = collectFormulaSnapshots(read.value.workbook)
	const comparableBefore = before.filter((entry) => entry.cached !== 'empty')
	const recalcStart = performance.now()
	const recalc = recalculate(
		read.value.workbook,
		defaultCalcContext({
			dateSystem: read.value.workbook.calcSettings.dateSystem,
			iterativeCalc: read.value.workbook.calcSettings.iterativeCalc,
		}),
	)
	const recalcMs = performance.now() - recalcStart
	const after = collectFormulaSnapshots(read.value.workbook)
	const comparableRefs = new Set(comparableBefore.map((entry) => entry.ref))
	const comparableAfter = after.filter((entry) => comparableRefs.has(entry.ref))
	const comparison = compareSnapshots(before, after)
	const { mismatches, volatileOracleSkips } = comparison
	const semanticMismatchCount = countMismatches(mismatches, 'semantic')
	const numericDriftMismatchCount = countMismatches(mismatches, 'numeric-drift')
	const acceptedMismatchCount = numericDriftMismatchCount
	return {
		file: entry.file,
		...(entry.source ? { source: entry.source } : {}),
		...(entry.sourceUrl ? { sourceUrl: entry.sourceUrl } : {}),
		oracle: args.oracle,
		formulaCount: before.length,
		comparedCount: comparableBefore.length,
		noCachedFormulaCount: before.length - comparableBefore.length,
		volatileOracleSkipCount: volatileOracleSkips.length,
		mismatchCount: mismatches.length,
		acceptedMismatchCount,
		unacceptedMismatchCount: semanticMismatchCount,
		semanticMismatchCount,
		numericDriftMismatchCount,
		errorCount: recalc.errors.length,
		beforeHash: hashLines(
			comparableBefore.map((entry) => `${entry.ref}\t${entry.formula}\t${entry.cached}`),
		),
		afterHash: hashLines(
			comparableAfter.map((entry) => `${entry.ref}\t${entry.formula}\t${entry.cached}`),
		),
		mismatches: mismatches.slice(0, args.maxReportedMismatches ?? 50),
		volatileOracleSkips: volatileOracleSkips.slice(0, args.maxReportedMismatches ?? 50),
		readMs,
		recalcMs,
	}
}

function countMismatches(
	mismatches: readonly FormulaMismatch[],
	classification: FormulaMismatchClassification,
): number {
	return mismatches.filter((mismatch) => mismatch.classification === classification).length
}

function summarize(results: readonly WorkbookFormulaResult[]): SuitePayload['summary'] {
	return {
		workbookCount: results.length,
		formulaCount: results.reduce((sum, result) => sum + result.formulaCount, 0),
		comparedCount: results.reduce((sum, result) => sum + result.comparedCount, 0),
		noCachedFormulaCount: results.reduce((sum, result) => sum + result.noCachedFormulaCount, 0),
		volatileOracleSkipCount: results.reduce(
			(sum, result) => sum + result.volatileOracleSkipCount,
			0,
		),
		mismatchCount: results.reduce((sum, result) => sum + result.mismatchCount, 0),
		acceptedMismatchCount: results.reduce((sum, result) => sum + result.acceptedMismatchCount, 0),
		unacceptedMismatchCount: results.reduce(
			(sum, result) => sum + result.unacceptedMismatchCount,
			0,
		),
		semanticMismatchCount: results.reduce((sum, result) => sum + result.semanticMismatchCount, 0),
		numericDriftMismatchCount: results.reduce(
			(sum, result) => sum + result.numericDriftMismatchCount,
			0,
		),
		errorCount: results.reduce((sum, result) => sum + result.errorCount, 0),
		perfectWorkbookCount: results.filter(
			(result) => result.mismatchCount === 0 && result.errorCount === 0,
		).length,
		semanticPerfectWorkbookCount: results.filter(
			(result) => result.semanticMismatchCount === 0 && result.errorCount === 0,
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
			maxReportedMismatches: args.maxReportedMismatches,
			sampleSeed: args.sampleSeed,
		},
		results,
		summary: summarize(results),
	}
}

export function formulaCorpusCorrectnessAssertionFailures(
	payload: SuitePayload,
	args: Pick<
		Args,
		| 'maxErrors'
		| 'maxMismatches'
		| 'maxAcceptedMismatches'
		| 'maxUnacceptedMismatches'
		| 'maxSemanticMismatches'
		| 'maxVolatileOracleSkips'
		| 'minWorkbooks'
		| 'minFormulas'
		| 'minComparedFormulas'
		| 'minPerfectWorkbooks'
		| 'minSemanticPerfectWorkbooks'
	>,
): readonly string[] {
	const failures: string[] = []
	if (args.maxMismatches !== undefined && payload.summary.mismatchCount > args.maxMismatches) {
		failures.push(`mismatches ${payload.summary.mismatchCount} exceeded ${args.maxMismatches}`)
	}
	if (
		args.maxAcceptedMismatches !== undefined &&
		payload.summary.acceptedMismatchCount > args.maxAcceptedMismatches
	) {
		failures.push(
			`accepted mismatches ${payload.summary.acceptedMismatchCount} exceeded ${args.maxAcceptedMismatches}`,
		)
	}
	if (
		args.maxUnacceptedMismatches !== undefined &&
		payload.summary.unacceptedMismatchCount > args.maxUnacceptedMismatches
	) {
		failures.push(
			`unaccepted mismatches ${payload.summary.unacceptedMismatchCount} exceeded ${args.maxUnacceptedMismatches}`,
		)
	}
	if (
		args.maxSemanticMismatches !== undefined &&
		payload.summary.semanticMismatchCount > args.maxSemanticMismatches
	) {
		failures.push(
			`semantic mismatches ${payload.summary.semanticMismatchCount} exceeded ${args.maxSemanticMismatches}`,
		)
	}
	if (
		args.maxVolatileOracleSkips !== undefined &&
		payload.summary.volatileOracleSkipCount > args.maxVolatileOracleSkips
	) {
		failures.push(
			`volatile oracle skips ${payload.summary.volatileOracleSkipCount} exceeded ${args.maxVolatileOracleSkips}`,
		)
	}
	if (args.maxErrors !== undefined && payload.summary.errorCount > args.maxErrors) {
		failures.push(`errors ${payload.summary.errorCount} exceeded ${args.maxErrors}`)
	}
	if (args.minWorkbooks !== undefined && payload.summary.workbookCount < args.minWorkbooks) {
		failures.push(`workbooks ${payload.summary.workbookCount} below ${args.minWorkbooks}`)
	}
	if (args.minFormulas !== undefined && payload.summary.formulaCount < args.minFormulas) {
		failures.push(`formulas ${payload.summary.formulaCount} below ${args.minFormulas}`)
	}
	if (
		args.minComparedFormulas !== undefined &&
		payload.summary.comparedCount < args.minComparedFormulas
	) {
		failures.push(
			`compared formulas ${payload.summary.comparedCount} below ${args.minComparedFormulas}`,
		)
	}
	if (
		args.minPerfectWorkbooks !== undefined &&
		payload.summary.perfectWorkbookCount < args.minPerfectWorkbooks
	) {
		failures.push(
			`perfect workbooks ${payload.summary.perfectWorkbookCount} below ${args.minPerfectWorkbooks}`,
		)
	}
	if (
		args.minSemanticPerfectWorkbooks !== undefined &&
		payload.summary.semanticPerfectWorkbookCount < args.minSemanticPerfectWorkbooks
	) {
		failures.push(
			`semantic-perfect workbooks ${payload.summary.semanticPerfectWorkbookCount} below ${args.minSemanticPerfectWorkbooks}`,
		)
	}
	return failures
}

function render(payload: SuitePayload): void {
	console.log(
		`formula corpus correctness: workbooks=${payload.summary.workbookCount} formulas=${payload.summary.formulaCount} compared=${payload.summary.comparedCount} noCached=${payload.summary.noCachedFormulaCount} volatileOracleSkips=${payload.summary.volatileOracleSkipCount} mismatches=${payload.summary.mismatchCount} accepted=${payload.summary.acceptedMismatchCount} unaccepted=${payload.summary.unacceptedMismatchCount} semantic=${payload.summary.semanticMismatchCount} numericDrift=${payload.summary.numericDriftMismatchCount} errors=${payload.summary.errorCount}`,
	)
	for (const result of payload.results) {
		console.log(
			`${result.file}: formulas=${result.formulaCount} compared=${result.comparedCount} noCached=${result.noCachedFormulaCount} volatileOracleSkips=${result.volatileOracleSkipCount} mismatches=${result.mismatchCount} accepted=${result.acceptedMismatchCount} unaccepted=${result.unacceptedMismatchCount} semantic=${result.semanticMismatchCount} numericDrift=${result.numericDriftMismatchCount} errors=${result.errorCount} read=${result.readMs.toFixed(2)}ms recalc=${result.recalcMs.toFixed(2)}ms`,
		)
	}
}

if (import.meta.main) {
	const args = readArgs()
	const payload = await runFormulaCorpusCorrectness(args)
	if (args.json) console.log(JSON.stringify(payload, null, 2))
	else render(payload)
	const failures = formulaCorpusCorrectnessAssertionFailures(payload, args)
	if (failures.length > 0) {
		console.error(`formula corpus correctness failed: ${failures.join('; ')}`)
		process.exit(1)
	}
}
