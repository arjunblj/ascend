#!/usr/bin/env bun
import { HyperFormula } from 'hyperformula'
import { createWorkbook, type StyleId, type Workbook } from '../../packages/core/src/index.ts'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import {
	type CellValue,
	EMPTY,
	numberValue,
	stringValue,
	topLeftScalar,
} from '../../packages/schema/src/index.ts'

const SID = 0 as StyleId

type ProfileName =
	| 'hf-prefix-range-sum'
	| 'hf-prefix-range-dirty-head'
	| 'hf-prefix-range-dirty-tail'
	| 'hf-indexed-index-match'
	| 'hf-indexed-index-match-dirty-key'
	| 'hf-indexed-index-match-dirty-value'
type ProfileSelection = ProfileName | 'all'
type EngineName = 'ascend' | 'hyperformula'
type PrefixAggregate = 'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX'
type PrefixAggregateSelection = PrefixAggregate | 'ALL'

const PREFIX_AGGREGATES: readonly PrefixAggregate[] = ['SUM', 'COUNT', 'AVERAGE', 'MIN', 'MAX']
const HYPERFORMULA_EXCEL_LIMITS = {
	licenseKey: 'gpl-v3',
	maxRows: 1_048_576,
	maxColumns: 16_384,
} as const

interface Args {
	readonly profile: ProfileSelection
	readonly rows: number
	readonly formulas: number
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
	readonly aggregate: PrefixAggregateSelection
	readonly minOperationSpeedup?: number
	readonly minTotalSpeedup?: number
	readonly assertCorrectness: boolean
}

type ResolvedArgs = Omit<Args, 'profile' | 'aggregate'> & {
	readonly profile: ProfileName
	readonly aggregate: PrefixAggregate
}

interface Sample {
	readonly setupMs: number
	readonly operationMs: number
	readonly totalMs: number
}

interface EngineCase {
	readonly engine: EngineName
	readonly samples: readonly Sample[]
	readonly correctness: Record<string, string | number | boolean>
}

interface Profile {
	readonly name: ProfileName
	readonly sourceBenchmark: string
	readonly sourceUrl: string
	readonly notes: string
	runAscend(args: ResolvedArgs): EngineRun
	runHyperFormula(args: ResolvedArgs): EngineRun
}

interface EngineRun {
	readonly sample: Sample
	readonly correctness: Record<string, string | number | boolean>
}

interface ProfilePayload {
	readonly formatVersion: 1
	readonly suite: 'ascend-formula-sota'
	readonly generatedAt: string
	readonly profile: {
		readonly name: ProfileName
		readonly sourceBenchmark: string
		readonly sourceUrl: string
		readonly notes: string
		readonly rows: number
		readonly formulas: number
		readonly aggregate?: PrefixAggregate
		readonly repeat: number
		readonly warmup: number
	}
	readonly cases: readonly {
		readonly engine: EngineName
		readonly metrics: ReturnType<typeof summarize>
		readonly samples: readonly Sample[]
		readonly correctness: Record<string, string | number | boolean>
	}[]
	readonly comparison: {
		readonly totalSpeedupVsHyperFormula: number
		readonly operationSpeedupVsHyperFormula: number
	}
}

const PROFILES: Record<ProfileName, Profile> = {
	'hf-prefix-range-sum': {
		name: 'hf-prefix-range-sum',
		sourceBenchmark: 'HyperFormula dependency-graph range composition example',
		sourceUrl: 'https://hyperformula.handsontable.com/guide/dependency-graph.html',
		notes:
			'Growing SUM(A$1:A<n>) formulas test whether the engine avoids repeatedly materializing nearly identical prefix ranges.',
		runAscend: runAscendPrefixRangeSum,
		runHyperFormula: runHyperFormulaPrefixRangeSum,
	},
	'hf-prefix-range-dirty-head': {
		name: 'hf-prefix-range-dirty-head',
		sourceBenchmark: 'HyperFormula incremental recalculation over optimized growing ranges',
		sourceUrl: 'https://hyperformula.handsontable.com/guide/dependency-graph.html',
		notes:
			'After an initial full recalc, edit A1 and measure dirty propagation through growing aggregate prefix formulas.',
		runAscend: (args) => runAscendPrefixRangeDirty(args, 'head'),
		runHyperFormula: (args) => runHyperFormulaPrefixRangeDirty(args, 'head'),
	},
	'hf-prefix-range-dirty-tail': {
		name: 'hf-prefix-range-dirty-tail',
		sourceBenchmark: 'HyperFormula incremental recalculation over optimized growing ranges',
		sourceUrl: 'https://hyperformula.handsontable.com/guide/dependency-graph.html',
		notes:
			'After an initial full recalc, edit the last source row covered by formulas and measure dirty propagation through growing aggregate prefix formulas.',
		runAscend: (args) => runAscendPrefixRangeDirty(args, 'tail'),
		runHyperFormula: (args) => runHyperFormulaPrefixRangeDirty(args, 'tail'),
	},
	'hf-indexed-index-match': {
		name: 'hf-indexed-index-match',
		sourceBenchmark: 'HyperFormula useColumnIndex exact MATCH/INDEX optimization',
		sourceUrl: 'https://hyperformula.handsontable.com/guide/performance.html',
		notes:
			'Exact INDEX/MATCH formulas over a keyed table compare Ascend lookup caching against HyperFormula with useColumnIndex enabled.',
		runAscend: runAscendIndexedVlookup,
		runHyperFormula: runHyperFormulaIndexedVlookup,
	},
	'hf-indexed-index-match-dirty-key': {
		name: 'hf-indexed-index-match-dirty-key',
		sourceBenchmark: 'HyperFormula useColumnIndex exact MATCH/INDEX incremental lookup-key edit',
		sourceUrl: 'https://hyperformula.handsontable.com/guide/performance.html',
		notes:
			'After initial INDEX/MATCH calculation over a keyed table, edit one lookup key and measure incremental recalculation.',
		runAscend: (args) => runAscendIndexedVlookupDirty(args, 'key'),
		runHyperFormula: (args) => runHyperFormulaIndexedVlookupDirty(args, 'key'),
	},
	'hf-indexed-index-match-dirty-value': {
		name: 'hf-indexed-index-match-dirty-value',
		sourceBenchmark: 'HyperFormula useColumnIndex exact MATCH/INDEX incremental return-value edit',
		sourceUrl: 'https://hyperformula.handsontable.com/guide/performance.html',
		notes:
			'After initial INDEX/MATCH calculation over a keyed table, edit one indexed return value and measure incremental recalculation.',
		runAscend: (args) => runAscendIndexedVlookupDirty(args, 'value'),
		runHyperFormula: (args) => runHyperFormulaIndexedVlookupDirty(args, 'value'),
	},
}

function readOption(args: readonly string[], name: string): string | undefined {
	const index = args.indexOf(name)
	return index >= 0 ? args[index + 1] : undefined
}

function hasFlag(args: readonly string[], name: string): boolean {
	return args.includes(name)
}

function renderHelp(): string {
	return [
		'Ascend formula SOTA benchmark runner',
		'',
		'Usage:',
		'  bun run fixtures/benchmarks/formula-sota.ts --profile <name> [--rows N] [--formulas N] [--repeat N] [--warmup N] [--aggregate SUM|COUNT|AVERAGE|MIN|MAX|ALL] [--json]',
		'',
		'Options:',
		'  --profile <name|all> Public comparator profile to run.',
		'  --rows N             Source row count. Defaults depend on profile.',
		'  --formulas N         Formula count. Defaults depend on profile.',
		'  --repeat N           Number of measured samples. Defaults to 5.',
		'  --warmup N           Number of warmup samples. Defaults to 1.',
		'  --aggregate <name>   Prefix aggregate for hf-prefix-range profiles. Defaults to ALL for --profile all, otherwise SUM.',
		'  --min-operation-speedup N',
		'                       Fail unless Ascend operation median is at least N x HyperFormula.',
		'  --min-total-speedup N',
		'                       Fail unless Ascend total median is at least N x HyperFormula.',
		'  --assert-correctness Fail unless all correctness match flags are true and Ascend errors are 0.',
		'  --json               Emit JSON instead of a text summary.',
		'  --help, -h           Show this help without running benchmarks.',
		'',
		'Profiles:',
		...Object.values(PROFILES)
			.map((profile) => `  ${profile.name} - ${profile.sourceBenchmark} (${profile.sourceUrl})`)
			.sort(),
	].join('\n')
}

function positiveInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

function optionalPositiveNumber(raw: string | undefined): number | undefined {
	if (raw === undefined) return undefined
	const value = Number(raw)
	if (!Number.isFinite(value) || value <= 0) {
		throw new Error(`Expected a positive number, received "${raw}"`)
	}
	return value
}

function parseArgs(): Args {
	const argv = process.argv.slice(2)
	if (hasFlag(argv, '--help') || hasFlag(argv, '-h')) {
		console.log(renderHelp())
		process.exit(0)
	}
	const rawProfile = readOption(argv, '--profile') ?? 'hf-prefix-range-sum'
	if (rawProfile !== 'all' && !(rawProfile in PROFILES)) {
		throw new Error(
			`Unsupported --profile "${rawProfile}". Expected all, ${Object.keys(PROFILES).join(', ')}`,
		)
	}
	const profile = rawProfile as ProfileSelection
	const rawAggregate = (
		readOption(argv, '--aggregate') ?? (profile === 'all' ? 'ALL' : 'SUM')
	).toUpperCase()
	if (!isPrefixAggregateSelection(rawAggregate)) {
		throw new Error('Unsupported --aggregate. Expected SUM, COUNT, AVERAGE, MIN, MAX, or ALL')
	}
	return {
		profile,
		rows: positiveInt(
			readOption(argv, '--rows'),
			profile === 'all' || profile.startsWith('hf-prefix-range') ? 5000 : 8000,
		),
		formulas: positiveInt(
			readOption(argv, '--formulas'),
			profile === 'all' || profile.startsWith('hf-prefix-range') ? 5000 : 1000,
		),
		repeat: positiveInt(readOption(argv, '--repeat'), 5),
		warmup: nonNegativeInt(readOption(argv, '--warmup'), 1),
		json: hasFlag(argv, '--json'),
		aggregate: rawAggregate,
		minOperationSpeedup: optionalPositiveNumber(readOption(argv, '--min-operation-speedup')),
		minTotalSpeedup: optionalPositiveNumber(readOption(argv, '--min-total-speedup')),
		assertCorrectness: hasFlag(argv, '--assert-correctness'),
	}
}

function isPrefixAggregateSelection(value: string): value is PrefixAggregateSelection {
	return value === 'ALL' || (PREFIX_AGGREGATES as readonly string[]).includes(value)
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const middle = Math.floor(sorted.length / 2)
	const upper = sorted[middle] ?? 0
	return sorted.length % 2 === 1 ? upper : ((sorted[middle - 1] ?? upper) + upper) / 2
}

function summarize(samples: readonly Sample[]) {
	return {
		setupMedianMs: median(samples.map((sample) => sample.setupMs)),
		operationMedianMs: median(samples.map((sample) => sample.operationMs)),
		totalMedianMs: median(samples.map((sample) => sample.totalMs)),
	}
}

function numberCell(value: number): CellValue {
	return numberValue(value)
}

function setNumber(workbook: Workbook, row: number, col: number, value: number): void {
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Missing sheet')
	sheet.cells.set(row, col, { value: numberCell(value), formula: null, styleId: SID })
}

function setString(workbook: Workbook, row: number, col: number, value: string): void {
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Missing sheet')
	sheet.cells.set(row, col, { value: stringValue(value), formula: null, styleId: SID })
}

function setFormula(workbook: Workbook, row: number, col: number, formula: string): void {
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Missing sheet')
	sheet.cells.set(row, col, { value: EMPTY, formula, styleId: SID })
}

function readNumber(workbook: Workbook, row: number, col: number): number {
	const value = workbook.sheets[0]?.cells.readValue(row, col)
	const scalar = topLeftScalar(value)
	if (scalar.kind !== 'number') {
		throw new Error(`Expected number at row=${row} col=${col}; got ${scalar.kind}`)
	}
	return scalar.value
}

function expectedPrefixSum(row: number): number {
	return ((row + 1) * (row + 2)) / 2
}

function expectedPrefixAggregate(aggregate: PrefixAggregate, row: number): number {
	switch (aggregate) {
		case 'SUM':
			return expectedPrefixSum(row)
		case 'COUNT':
			return row + 1
		case 'AVERAGE':
			return (row + 2) / 2
		case 'MIN':
			return 1
		case 'MAX':
			return row + 1
	}
}

function numbersClose(actual: number, expected: number): boolean {
	return Math.abs(actual - expected) <= Math.max(1e-9, Math.abs(expected) * 1e-10)
}

function mutationValueForAggregate(aggregate: PrefixAggregate): number {
	switch (aggregate) {
		case 'MIN':
			return -1000003
		case 'MAX':
			return 1000003
		default:
			return 1000003
	}
}

function expectedDirtyPrefixAggregate(
	aggregate: PrefixAggregate,
	formulaRow: number,
	editedRow: number,
	newSourceValue: number,
): number {
	const base = expectedPrefixAggregate(aggregate, formulaRow)
	if (formulaRow < editedRow) return base
	const delta = newSourceValue - (editedRow + 1)
	switch (aggregate) {
		case 'SUM':
			return expectedPrefixSum(formulaRow) + delta
		case 'COUNT':
			return formulaRow + 1
		case 'AVERAGE':
			return (expectedPrefixSum(formulaRow) + delta) / (formulaRow + 1)
		case 'MIN':
			return Math.min(base, newSourceValue)
		case 'MAX':
			return Math.max(base, newSourceValue)
	}
}

function runAscendPrefixRangeSum(args: ResolvedArgs): EngineRun {
	const setupStart = performance.now()
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let row = 0; row < args.rows; row++) {
		setNumber(workbook, row, 0, row + 1)
	}
	for (let row = 0; row < args.formulas; row++) {
		setFormula(workbook, row, 1, `${args.aggregate}(A$1:A${row + 1})`)
	}
	const setupMs = performance.now() - setupStart
	const operationStart = performance.now()
	const result = recalculate(workbook, defaultCalcContext())
	const operationMs = performance.now() - operationStart
	const lastFormulaRow = args.formulas - 1
	const lastValue = readNumber(workbook, lastFormulaRow, 1)
	const expectedLastValue = expectedPrefixAggregate(args.aggregate, lastFormulaRow)
	return {
		sample: { setupMs, operationMs, totalMs: setupMs + operationMs },
		correctness: {
			changed: result.changed.length,
			errors: result.errors.length,
			aggregate: args.aggregate,
			lastValue,
			expectedLastValue,
			lastValueMatches: numbersClose(lastValue, expectedLastValue),
		},
	}
}

function runHyperFormulaPrefixRangeSum(args: ResolvedArgs): EngineRun {
	const setupStart = performance.now()
	const rows = Math.max(args.rows, args.formulas)
	const data: (number | string | null)[][] = []
	for (let row = 0; row < rows; row++) {
		data.push([
			row < args.rows ? row + 1 : null,
			row < args.formulas ? `=${args.aggregate}(A$1:A${row + 1})` : null,
		])
	}
	const hf = HyperFormula.buildEmpty(HYPERFORMULA_EXCEL_LIMITS)
	hf.addSheet('Sheet1')
	hf.suspendEvaluation()
	hf.setCellContents({ sheet: 0, row: 0, col: 0 }, data)
	const setupMs = performance.now() - setupStart
	const operationStart = performance.now()
	hf.resumeEvaluation()
	const operationMs = performance.now() - operationStart
	const lastFormulaRow = args.formulas - 1
	const lastValue = hf.getCellValue({ sheet: 0, row: lastFormulaRow, col: 1 })
	const expectedLastValue = expectedPrefixAggregate(args.aggregate, lastFormulaRow)
	hf.destroy()
	return {
		sample: { setupMs, operationMs, totalMs: setupMs + operationMs },
		correctness: {
			aggregate: args.aggregate,
			lastValue: typeof lastValue === 'number' ? lastValue : String(lastValue),
			expectedLastValue,
			lastValueMatches: typeof lastValue === 'number' && numbersClose(lastValue, expectedLastValue),
		},
	}
}

function prefixDirtyEditRow(args: ResolvedArgs, position: 'head' | 'tail'): number {
	return position === 'head' ? 0 : Math.max(0, Math.min(args.rows, args.formulas) - 1)
}

function runAscendPrefixRangeDirty(args: ResolvedArgs, position: 'head' | 'tail'): EngineRun {
	const setupStart = performance.now()
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let row = 0; row < args.rows; row++) {
		setNumber(workbook, row, 0, row + 1)
	}
	for (let row = 0; row < args.formulas; row++) {
		setFormula(workbook, row, 1, `${args.aggregate}(A$1:A${row + 1})`)
	}
	const initial = recalculate(workbook, defaultCalcContext())
	const setupMs = performance.now() - setupStart
	const editRow = prefixDirtyEditRow(args, position)
	const newSourceValue = mutationValueForAggregate(args.aggregate)
	setNumber(workbook, editRow, 0, newSourceValue)
	const operationStart = performance.now()
	const result = recalculate(workbook, defaultCalcContext(), {
		dirtyOnly: true,
		dirtyRefs: [`Sheet1!A${editRow + 1}`],
	})
	const operationMs = performance.now() - operationStart
	const probeRow = args.formulas - 1
	const probeValue = readNumber(workbook, probeRow, 1)
	const expectedProbeValue = expectedDirtyPrefixAggregate(
		args.aggregate,
		probeRow,
		editRow,
		newSourceValue,
	)
	return {
		sample: { setupMs, operationMs, totalMs: setupMs + operationMs },
		correctness: {
			aggregate: args.aggregate,
			editPosition: position,
			editRow,
			newSourceValue,
			initialChanged: initial.changed.length,
			initialErrors: initial.errors.length,
			changed: result.changed.length,
			errors: result.errors.length,
			probeValue,
			expectedProbeValue,
			probeValueMatches: numbersClose(probeValue, expectedProbeValue),
		},
	}
}

function runHyperFormulaPrefixRangeDirty(args: ResolvedArgs, position: 'head' | 'tail'): EngineRun {
	const setupStart = performance.now()
	const rows = Math.max(args.rows, args.formulas)
	const data: (number | string | null)[][] = []
	for (let row = 0; row < rows; row++) {
		data.push([
			row < args.rows ? row + 1 : null,
			row < args.formulas ? `=${args.aggregate}(A$1:A${row + 1})` : null,
		])
	}
	const hf = HyperFormula.buildEmpty(HYPERFORMULA_EXCEL_LIMITS)
	hf.addSheet('Sheet1')
	hf.suspendEvaluation()
	hf.setCellContents({ sheet: 0, row: 0, col: 0 }, data)
	const initialChanges = hf.resumeEvaluation()
	const setupMs = performance.now() - setupStart
	const editRow = prefixDirtyEditRow(args, position)
	const newSourceValue = mutationValueForAggregate(args.aggregate)
	const operationStart = performance.now()
	const changes = hf.setCellContents({ sheet: 0, row: editRow, col: 0 }, newSourceValue)
	const operationMs = performance.now() - operationStart
	const probeRow = args.formulas - 1
	const probeValue = hf.getCellValue({ sheet: 0, row: probeRow, col: 1 })
	const expectedProbeValue = expectedDirtyPrefixAggregate(
		args.aggregate,
		probeRow,
		editRow,
		newSourceValue,
	)
	hf.destroy()
	return {
		sample: { setupMs, operationMs, totalMs: setupMs + operationMs },
		correctness: {
			aggregate: args.aggregate,
			editPosition: position,
			editRow,
			newSourceValue,
			initialChanged: initialChanges.length,
			changed: changes.length,
			probeValue: typeof probeValue === 'number' ? probeValue : String(probeValue),
			expectedProbeValue,
			probeValueMatches:
				typeof probeValue === 'number' && numbersClose(probeValue, expectedProbeValue),
		},
	}
}

function lookupKey(index: number): string {
	return `key-${String(index + 1).padStart(6, '0')}`
}

function lookupIndex(formulaIndex: number, rows: number): number {
	return rows - ((formulaIndex * 37) % rows) - 1
}

function alternateLookupIndex(formulaIndex: number, rows: number): number {
	return (lookupIndex(formulaIndex, rows) + Math.max(1, Math.floor(rows / 2))) % rows
}

function mutatedLookupValue(formulaIndex: number): number {
	return 9000000 + formulaIndex
}

function runAscendIndexedVlookup(args: ResolvedArgs): EngineRun {
	const setupStart = performance.now()
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let row = 0; row < args.rows; row++) {
		setString(workbook, row, 0, lookupKey(row))
		setNumber(workbook, row, 1, (row + 1) * 3)
		setNumber(workbook, row, 2, (row + 1) * 7)
	}
	for (let row = 0; row < args.formulas; row++) {
		const keyIndex = lookupIndex(row, args.rows)
		setString(workbook, row, 4, lookupKey(keyIndex))
		setFormula(
			workbook,
			row,
			5,
			`INDEX(C$1:C$${args.rows},MATCH(E${row + 1},A$1:A$${args.rows},0))`,
		)
	}
	const setupMs = performance.now() - setupStart
	const operationStart = performance.now()
	const result = recalculate(workbook, defaultCalcContext())
	const operationMs = performance.now() - operationStart
	const probeRow = args.formulas - 1
	const expectedValue = (lookupIndex(probeRow, args.rows) + 1) * 7
	const value = readNumber(workbook, probeRow, 5)
	return {
		sample: { setupMs, operationMs, totalMs: setupMs + operationMs },
		correctness: {
			changed: result.changed.length,
			errors: result.errors.length,
			probeValue: value,
			expectedProbeValue: expectedValue,
			probeValueMatches: value === expectedValue,
		},
	}
}

function runAscendIndexedVlookupDirty(args: ResolvedArgs, editKind: 'key' | 'value'): EngineRun {
	const setupStart = performance.now()
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let row = 0; row < args.rows; row++) {
		setString(workbook, row, 0, lookupKey(row))
		setNumber(workbook, row, 1, (row + 1) * 3)
		setNumber(workbook, row, 2, (row + 1) * 7)
	}
	for (let row = 0; row < args.formulas; row++) {
		const keyIndex = lookupIndex(row, args.rows)
		setString(workbook, row, 4, lookupKey(keyIndex))
		setFormula(
			workbook,
			row,
			5,
			`INDEX(C$1:C$${args.rows},MATCH(E${row + 1},A$1:A$${args.rows},0))`,
		)
	}
	const initial = recalculate(workbook, defaultCalcContext())
	const setupMs = performance.now() - setupStart
	const probeRow = args.formulas - 1
	const originalKeyIndex = lookupIndex(probeRow, args.rows)
	let expectedValue: number
	let dirtyRef: string
	if (editKind === 'key') {
		const newKeyIndex = alternateLookupIndex(probeRow, args.rows)
		setString(workbook, probeRow, 4, lookupKey(newKeyIndex))
		expectedValue = (newKeyIndex + 1) * 7
		dirtyRef = `Sheet1!E${probeRow + 1}`
	} else {
		expectedValue = mutatedLookupValue(probeRow)
		setNumber(workbook, originalKeyIndex, 2, expectedValue)
		dirtyRef = `Sheet1!C${originalKeyIndex + 1}`
	}
	const operationStart = performance.now()
	const result = recalculate(workbook, defaultCalcContext(), {
		dirtyOnly: true,
		dirtyRefs: [dirtyRef],
	})
	const operationMs = performance.now() - operationStart
	const value = readNumber(workbook, probeRow, 5)
	return {
		sample: { setupMs, operationMs, totalMs: setupMs + operationMs },
		correctness: {
			editKind,
			dirtyRef,
			initialChanged: initial.changed.length,
			initialErrors: initial.errors.length,
			changed: result.changed.length,
			errors: result.errors.length,
			probeValue: value,
			expectedProbeValue: expectedValue,
			probeValueMatches: value === expectedValue,
		},
	}
}

function runHyperFormulaIndexedVlookup(args: ResolvedArgs): EngineRun {
	const setupStart = performance.now()
	const data: (number | string | null)[][] = []
	for (let row = 0; row < args.rows; row++) {
		data.push([lookupKey(row), (row + 1) * 3, (row + 1) * 7, null, null, null])
	}
	for (let row = 0; row < args.formulas; row++) {
		const target = data[row]
		if (!target) throw new Error(`Missing target row ${row}`)
		const keyIndex = lookupIndex(row, args.rows)
		target[4] = lookupKey(keyIndex)
		target[5] = `=INDEX(C$1:C$${args.rows},MATCH(E${row + 1},A$1:A$${args.rows},0))`
	}
	const hf = HyperFormula.buildEmpty({ ...HYPERFORMULA_EXCEL_LIMITS, useColumnIndex: true })
	hf.addSheet('Sheet1')
	hf.suspendEvaluation()
	hf.setCellContents({ sheet: 0, row: 0, col: 0 }, data)
	const setupMs = performance.now() - setupStart
	const operationStart = performance.now()
	hf.resumeEvaluation()
	const operationMs = performance.now() - operationStart
	const probeRow = args.formulas - 1
	const expectedValue = (lookupIndex(probeRow, args.rows) + 1) * 7
	const value = hf.getCellValue({ sheet: 0, row: probeRow, col: 5 })
	hf.destroy()
	return {
		sample: { setupMs, operationMs, totalMs: setupMs + operationMs },
		correctness: {
			probeValue: typeof value === 'number' ? value : String(value),
			expectedProbeValue: expectedValue,
			probeValueMatches: value === expectedValue,
		},
	}
}

function runHyperFormulaIndexedVlookupDirty(
	args: ResolvedArgs,
	editKind: 'key' | 'value',
): EngineRun {
	const setupStart = performance.now()
	const data: (number | string | null)[][] = []
	for (let row = 0; row < args.rows; row++) {
		data.push([lookupKey(row), (row + 1) * 3, (row + 1) * 7, null, null, null])
	}
	for (let row = 0; row < args.formulas; row++) {
		const target = data[row]
		if (!target) throw new Error(`Missing target row ${row}`)
		const keyIndex = lookupIndex(row, args.rows)
		target[4] = lookupKey(keyIndex)
		target[5] = `=INDEX(C$1:C$${args.rows},MATCH(E${row + 1},A$1:A$${args.rows},0))`
	}
	const hf = HyperFormula.buildEmpty({ ...HYPERFORMULA_EXCEL_LIMITS, useColumnIndex: true })
	hf.addSheet('Sheet1')
	hf.suspendEvaluation()
	hf.setCellContents({ sheet: 0, row: 0, col: 0 }, data)
	const initialChanges = hf.resumeEvaluation()
	const setupMs = performance.now() - setupStart
	const probeRow = args.formulas - 1
	const originalKeyIndex = lookupIndex(probeRow, args.rows)
	let expectedValue: number
	let editAddress: { sheet: number; row: number; col: number }
	if (editKind === 'key') {
		const newKeyIndex = alternateLookupIndex(probeRow, args.rows)
		expectedValue = (newKeyIndex + 1) * 7
		editAddress = { sheet: 0, row: probeRow, col: 4 }
		const operationStart = performance.now()
		const changes = hf.setCellContents(editAddress, lookupKey(newKeyIndex))
		const operationMs = performance.now() - operationStart
		const value = hf.getCellValue({ sheet: 0, row: probeRow, col: 5 })
		hf.destroy()
		return {
			sample: { setupMs, operationMs, totalMs: setupMs + operationMs },
			correctness: {
				editKind,
				initialChanged: initialChanges.length,
				changed: changes.length,
				probeValue: typeof value === 'number' ? value : String(value),
				expectedProbeValue: expectedValue,
				probeValueMatches: value === expectedValue,
			},
		}
	}
	expectedValue = mutatedLookupValue(probeRow)
	editAddress = { sheet: 0, row: originalKeyIndex, col: 2 }
	const operationStart = performance.now()
	const changes = hf.setCellContents(editAddress, expectedValue)
	const operationMs = performance.now() - operationStart
	const value = hf.getCellValue({ sheet: 0, row: probeRow, col: 5 })
	hf.destroy()
	return {
		sample: { setupMs, operationMs, totalMs: setupMs + operationMs },
		correctness: {
			editKind,
			initialChanged: initialChanges.length,
			changed: changes.length,
			probeValue: typeof value === 'number' ? value : String(value),
			expectedProbeValue: expectedValue,
			probeValueMatches: value === expectedValue,
		},
	}
}

function runEngine(engine: EngineName, profile: Profile, args: ResolvedArgs): EngineCase {
	for (let i = 0; i < args.warmup; i++) {
		runGc()
		if (engine === 'ascend') profile.runAscend(args)
		else profile.runHyperFormula(args)
	}
	const samples: Sample[] = []
	let correctness: Record<string, string | number | boolean> = {}
	for (let i = 0; i < args.repeat; i++) {
		runGc()
		const run = engine === 'ascend' ? profile.runAscend(args) : profile.runHyperFormula(args)
		samples.push(run.sample)
		correctness = run.correctness
	}
	return { engine, samples, correctness }
}

function collectAssertionFailures(
	args: Args,
	comparison: {
		readonly totalSpeedupVsHyperFormula: number
		readonly operationSpeedupVsHyperFormula: number
	},
	cases: readonly EngineCase[],
): string[] {
	const failures: string[] = []
	if (
		args.minOperationSpeedup !== undefined &&
		comparison.operationSpeedupVsHyperFormula < args.minOperationSpeedup
	) {
		failures.push(
			`operation speedup ${comparison.operationSpeedupVsHyperFormula.toFixed(3)}x is below required ${args.minOperationSpeedup}x`,
		)
	}
	if (
		args.minTotalSpeedup !== undefined &&
		comparison.totalSpeedupVsHyperFormula < args.minTotalSpeedup
	) {
		failures.push(
			`total speedup ${comparison.totalSpeedupVsHyperFormula.toFixed(3)}x is below required ${args.minTotalSpeedup}x`,
		)
	}
	if (args.assertCorrectness) {
		for (const entry of cases) {
			const matchFlags = Object.entries(entry.correctness).filter(([key]) =>
				key.endsWith('Matches'),
			)
			if (matchFlags.length === 0) {
				failures.push(`${entry.engine} exposed no correctness match flags`)
			}
			for (const [key, value] of matchFlags) {
				if (value !== true) failures.push(`${entry.engine}.${key} was ${String(value)}`)
			}
			if (entry.engine === 'ascend' && (entry.correctness.errors ?? 0) !== 0) {
				failures.push(`ascend errors=${String(entry.correctness.errors)}`)
			}
		}
	}
	return failures
}

function runProfile(
	profile: Profile,
	args: Args,
	generatedAt: string,
	aggregate: PrefixAggregate,
): ProfilePayload {
	const profileArgs: ResolvedArgs = { ...args, profile: profile.name, aggregate }
	const ascend = runEngine('ascend', profile, profileArgs)
	const hyperformula = runEngine('hyperformula', profile, profileArgs)
	const ascendSummary = summarize(ascend.samples)
	const hyperformulaSummary = summarize(hyperformula.samples)
	const comparison = {
		totalSpeedupVsHyperFormula:
			hyperformulaSummary.totalMedianMs / Math.max(ascendSummary.totalMedianMs, Number.EPSILON),
		operationSpeedupVsHyperFormula:
			hyperformulaSummary.operationMedianMs /
			Math.max(ascendSummary.operationMedianMs, Number.EPSILON),
	}
	return {
		formatVersion: 1,
		suite: 'ascend-formula-sota',
		generatedAt,
		profile: {
			name: profile.name,
			sourceBenchmark: profile.sourceBenchmark,
			sourceUrl: profile.sourceUrl,
			notes: profile.notes,
			rows: profileArgs.rows,
			formulas: profileArgs.formulas,
			...(isPrefixProfile(profile) ? { aggregate: profileArgs.aggregate } : {}),
			repeat: profileArgs.repeat,
			warmup: profileArgs.warmup,
		},
		cases: [
			{
				engine: ascend.engine,
				metrics: ascendSummary,
				samples: ascend.samples,
				correctness: ascend.correctness,
			},
			{
				engine: hyperformula.engine,
				metrics: hyperformulaSummary,
				samples: hyperformula.samples,
				correctness: hyperformula.correctness,
			},
		],
		comparison,
	}
}

function selectedProfiles(selection: ProfileSelection): readonly Profile[] {
	return selection === 'all' ? Object.values(PROFILES) : [PROFILES[selection]]
}

function isPrefixProfile(profile: Profile): boolean {
	return profile.name.startsWith('hf-prefix-range')
}

function selectedProfileRuns(
	selection: ProfileSelection,
	aggregate: PrefixAggregateSelection,
): readonly { profile: Profile; aggregate: PrefixAggregate }[] {
	const runs: { profile: Profile; aggregate: PrefixAggregate }[] = []
	for (const profile of selectedProfiles(selection)) {
		if (!isPrefixProfile(profile)) {
			runs.push({ profile, aggregate: 'SUM' })
			continue
		}
		const aggregates = aggregate === 'ALL' ? PREFIX_AGGREGATES : [aggregate]
		for (const selectedAggregate of aggregates) {
			runs.push({ profile, aggregate: selectedAggregate })
		}
	}
	return runs
}

function geometricMean(values: readonly number[]): number {
	if (values.length === 0) return 0
	return Math.exp(
		values.reduce((sum, value) => sum + Math.log(Math.max(value, Number.EPSILON)), 0) /
			values.length,
	)
}

function suiteSummary(profiles: readonly ProfilePayload[]) {
	const operationSpeedups = profiles.map(
		(profile) => profile.comparison.operationSpeedupVsHyperFormula,
	)
	const totalSpeedups = profiles.map((profile) => profile.comparison.totalSpeedupVsHyperFormula)
	return {
		profileCount: profiles.length,
		minOperationSpeedupVsHyperFormula: Math.min(...operationSpeedups),
		minTotalSpeedupVsHyperFormula: Math.min(...totalSpeedups),
		geomeanOperationSpeedupVsHyperFormula: geometricMean(operationSpeedups),
		geomeanTotalSpeedupVsHyperFormula: geometricMean(totalSpeedups),
	}
}

function renderProfile(payload: ProfilePayload): void {
	console.log(`${payload.profile.name}: ${payload.profile.notes}`)
	for (const entry of payload.cases) {
		console.log(
			`${entry.engine}: setup=${entry.metrics.setupMedianMs.toFixed(2)}ms operation=${entry.metrics.operationMedianMs.toFixed(2)}ms total=${entry.metrics.totalMedianMs.toFixed(2)}ms`,
		)
	}
	console.log(
		`speedup vs HyperFormula: operation=${payload.comparison.operationSpeedupVsHyperFormula.toFixed(2)}x total=${payload.comparison.totalSpeedupVsHyperFormula.toFixed(2)}x`,
	)
}

const args = parseArgs()
const generatedAt = new Date().toISOString()
const profilePayloads = selectedProfileRuns(args.profile, args.aggregate).map((run) =>
	runProfile(run.profile, args, generatedAt, run.aggregate),
)
const assertionFailures = profilePayloads.flatMap((payload) =>
	collectAssertionFailures(args, payload.comparison, payload.cases).map(
		(failure) =>
			`${payload.profile.name}${payload.profile.aggregate ? `/${payload.profile.aggregate}` : ''}: ${failure}`,
	),
)

if (args.profile === 'all' || profilePayloads.length > 1) {
	const payload = {
		formatVersion: 1,
		suite: 'ascend-formula-sota',
		generatedAt,
		selection: {
			profile: args.profile,
			rows: args.rows,
			formulas: args.formulas,
			aggregate: args.aggregate,
			repeat: args.repeat,
			warmup: args.warmup,
		},
		profiles: profilePayloads,
		summary: suiteSummary(profilePayloads),
	}
	if (args.json) {
		console.log(JSON.stringify(payload, null, 2))
	} else {
		for (const profile of profilePayloads) {
			renderProfile(profile)
			console.log('')
		}
		console.log(
			`formula SOTA suite: profiles=${payload.summary.profileCount} minOperation=${payload.summary.minOperationSpeedupVsHyperFormula.toFixed(2)}x minTotal=${payload.summary.minTotalSpeedupVsHyperFormula.toFixed(2)}x geomeanOperation=${payload.summary.geomeanOperationSpeedupVsHyperFormula.toFixed(2)}x geomeanTotal=${payload.summary.geomeanTotalSpeedupVsHyperFormula.toFixed(2)}x`,
		)
	}
} else {
	const payload = profilePayloads[0]
	if (!payload) throw new Error('No formula SOTA profile selected')
	if (args.json) {
		console.log(JSON.stringify(payload, null, 2))
	} else {
		renderProfile(payload)
	}
}

if (assertionFailures.length > 0) {
	for (const failure of assertionFailures)
		console.error(`formula-sota assertion failed: ${failure}`)
	process.exit(1)
}
