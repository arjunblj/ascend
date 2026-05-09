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
type EngineName = 'ascend' | 'hyperformula'
type PrefixAggregate = 'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX'

interface Args {
	readonly profile: ProfileName
	readonly rows: number
	readonly formulas: number
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
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
	runAscend(args: Args): EngineRun
	runHyperFormula(args: Args): EngineRun
}

interface EngineRun {
	readonly sample: Sample
	readonly correctness: Record<string, string | number | boolean>
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

function positiveInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

function parseArgs(): Args {
	const argv = process.argv.slice(2)
	const rawProfile = readOption(argv, '--profile') ?? 'hf-prefix-range-sum'
	if (!(rawProfile in PROFILES)) {
		throw new Error(
			`Unsupported --profile "${rawProfile}". Expected ${Object.keys(PROFILES).join(', ')}`,
		)
	}
	const profile = rawProfile as ProfileName
	const rawAggregate = (readOption(argv, '--aggregate') ?? 'SUM').toUpperCase()
	if (!isPrefixAggregate(rawAggregate)) {
		throw new Error('Unsupported --aggregate. Expected SUM, COUNT, AVERAGE, MIN, or MAX')
	}
	return {
		profile,
		rows: positiveInt(
			readOption(argv, '--rows'),
			profile.startsWith('hf-prefix-range') ? 5000 : 8000,
		),
		formulas: positiveInt(
			readOption(argv, '--formulas'),
			profile.startsWith('hf-prefix-range') ? 5000 : 1000,
		),
		repeat: positiveInt(readOption(argv, '--repeat'), 5),
		warmup: nonNegativeInt(readOption(argv, '--warmup'), 1),
		json: hasFlag(argv, '--json'),
		aggregate: rawAggregate,
	}
}

function isPrefixAggregate(value: string): value is PrefixAggregate {
	return (
		value === 'SUM' ||
		value === 'COUNT' ||
		value === 'AVERAGE' ||
		value === 'MIN' ||
		value === 'MAX'
	)
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

function runAscendPrefixRangeSum(args: Args): EngineRun {
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

function runHyperFormulaPrefixRangeSum(args: Args): EngineRun {
	const setupStart = performance.now()
	const rows = Math.max(args.rows, args.formulas)
	const data: (number | string)[][] = []
	for (let row = 0; row < rows; row++) {
		data.push([row + 1, `=${args.aggregate}(A$1:A${row + 1})`])
	}
	const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' })
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

function prefixDirtyEditRow(args: Args, position: 'head' | 'tail'): number {
	return position === 'head' ? 0 : Math.max(0, Math.min(args.rows, args.formulas) - 1)
}

function runAscendPrefixRangeDirty(args: Args, position: 'head' | 'tail'): EngineRun {
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

function runHyperFormulaPrefixRangeDirty(args: Args, position: 'head' | 'tail'): EngineRun {
	const setupStart = performance.now()
	const rows = Math.max(args.rows, args.formulas)
	const data: (number | string)[][] = []
	for (let row = 0; row < rows; row++) {
		data.push([row + 1, `=${args.aggregate}(A$1:A${row + 1})`])
	}
	const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' })
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

function runAscendIndexedVlookup(args: Args): EngineRun {
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

function runAscendIndexedVlookupDirty(args: Args, editKind: 'key' | 'value'): EngineRun {
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

function runHyperFormulaIndexedVlookup(args: Args): EngineRun {
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
	const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useColumnIndex: true })
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

function runHyperFormulaIndexedVlookupDirty(args: Args, editKind: 'key' | 'value'): EngineRun {
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
	const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useColumnIndex: true })
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

function runEngine(engine: EngineName, profile: Profile, args: Args): EngineCase {
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

const args = parseArgs()
const profile = PROFILES[args.profile]
const ascend = runEngine('ascend', profile, args)
const hyperformula = runEngine('hyperformula', profile, args)
const ascendSummary = summarize(ascend.samples)
const hyperformulaSummary = summarize(hyperformula.samples)
const payload = {
	formatVersion: 1,
	suite: 'ascend-formula-sota',
	generatedAt: new Date().toISOString(),
	profile: {
		name: profile.name,
		sourceBenchmark: profile.sourceBenchmark,
		sourceUrl: profile.sourceUrl,
		notes: profile.notes,
		rows: args.rows,
		formulas: args.formulas,
		...(args.profile.startsWith('hf-prefix-range') ? { aggregate: args.aggregate } : {}),
		repeat: args.repeat,
		warmup: args.warmup,
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
	comparison: {
		totalSpeedupVsHyperFormula:
			hyperformulaSummary.totalMedianMs / Math.max(ascendSummary.totalMedianMs, Number.EPSILON),
		operationSpeedupVsHyperFormula:
			hyperformulaSummary.operationMedianMs /
			Math.max(ascendSummary.operationMedianMs, Number.EPSILON),
	},
}

if (args.json) {
	console.log(JSON.stringify(payload, null, 2))
} else {
	console.log(`${profile.name}: ${profile.notes}`)
	for (const entry of payload.cases) {
		console.log(
			`${entry.engine}: setup=${entry.metrics.setupMedianMs.toFixed(2)}ms operation=${entry.metrics.operationMedianMs.toFixed(2)}ms total=${entry.metrics.totalMedianMs.toFixed(2)}ms`,
		)
	}
	console.log(
		`speedup vs HyperFormula: operation=${payload.comparison.operationSpeedupVsHyperFormula.toFixed(2)}x total=${payload.comparison.totalSpeedupVsHyperFormula.toFixed(2)}x`,
	)
}
