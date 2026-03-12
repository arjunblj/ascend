import { createWorkbook, type StyleId, type Workbook } from '../../packages/core/src/index.ts'
import {
	applyOperations,
	defaultCalcContext,
	recalculate,
} from '../../packages/engine/src/index.ts'
import { readCsv, writeCsv } from '../../packages/io-csv/src/index.ts'
import { readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { booleanValue, EMPTY, numberValue, stringValue } from '../../packages/schema/src/index.ts'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'
import {
	type BenchmarkCaseResult,
	createBenchmarkSuite,
	formatBytes,
	formatRate,
	summarizeSamples,
} from './results.ts'

const SID = 0 as StyleId

interface ScenarioInput {
	readonly workbook?: Workbook
	readonly bytes?: Uint8Array
	readonly content?: string
	readonly rows: number
	readonly cols: number
	readonly cells: number
}

interface ScenarioRunResult {
	readonly assertions?: Record<string, string | number | boolean | null>
}

interface Scenario {
	readonly name: string
	readonly category: 'read' | 'write' | 'calc' | 'workflow'
	build(): ScenarioInput
	run(input: ScenarioInput): Promise<ScenarioRunResult | undefined> | ScenarioRunResult | undefined
}

function requireBytes(input: ScenarioInput): Uint8Array {
	if (!input.bytes) throw new Error('Scenario bytes were not built')
	return input.bytes
}

function requireWorkbook(input: ScenarioInput): Workbook {
	if (!input.workbook) throw new Error('Scenario workbook was not built')
	return input.workbook
}

function requireContent(input: ScenarioInput): string {
	if (input.content === undefined) throw new Error('Scenario content was not built')
	return input.content
}

function mustWrite(workbook: Workbook): Uint8Array {
	const result = writeXlsx(workbook)
	if (!result.ok) throw new Error(result.error.message)
	return result.value
}

function setNumberCell(workbook: Workbook, row: number, col: number, value: number): void {
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	sheet.cells.set(row, col, { value: numberValue(value), formula: null, styleId: SID })
}

function setStringCell(workbook: Workbook, row: number, col: number, value: string): void {
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	sheet.cells.set(row, col, { value: stringValue(value), formula: null, styleId: SID })
}

function setBooleanCell(workbook: Workbook, row: number, col: number, value: boolean): void {
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	sheet.cells.set(row, col, { value: booleanValue(value), formula: null, styleId: SID })
}

function setFormulaCell(workbook: Workbook, row: number, col: number, formula: string): void {
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	sheet.cells.set(row, col, { value: EMPTY, formula, styleId: SID })
}

function createDeterministicRandom(seed: number): () => number {
	let state = seed >>> 0
	return () => {
		state = (Math.imul(state, 1664525) + 1013904223) >>> 0
		return state / 0x1_0000_0000
	}
}

function buildDenseWorkbook(rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			setNumberCell(workbook, r, c, r * cols + c + 1)
		}
	}
	return workbook
}

function buildStringDenseWorkbook(rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			const value = c % 2 === 0 ? `label-${r}-${c}` : `shared-${c % 5}`
			sheet.cells.set(r, c, {
				value: { kind: 'string', value },
				formula: null,
				styleId: SID,
			})
		}
	}
	return workbook
}

function buildSparseWorkbook(rows: number, cols: number, step: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < rows; r += step) {
		for (let c = 0; c < cols; c++) {
			setNumberCell(workbook, r, c, r + c + 1)
		}
	}
	return workbook
}

function buildMultiSheetWorkbook(sheetCount: number, rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	for (let s = 0; s < sheetCount; s++) {
		const sheet = workbook.addSheet(`Sheet${s + 1}`)
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				sheet.cells.set(r, c, {
					value: numberValue((s + 1) * (r + 1) * (c + 1)),
					formula: null,
					styleId: SID,
				})
			}
		}
	}
	return workbook
}

function buildFormulaChainWorkbook(length: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	setNumberCell(workbook, 0, 0, 1)
	for (let r = 1; r < length; r++) {
		setFormulaCell(workbook, r, 0, `A${r}+1`)
	}
	return workbook
}

function buildRangeAggregationWorkbook(length: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < length; r++) {
		setNumberCell(workbook, r, 0, r + 1)
		setFormulaCell(workbook, r, 1, `SUM(A1:A${r + 1})`)
	}
	return workbook
}

function buildIfShortCircuitWorkbook(length: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < length; r++) {
		setNumberCell(workbook, r, 0, 1)
		setNumberCell(workbook, r, 1, r + 1)
		setFormulaCell(workbook, r, 2, `IF(A${r + 1}>0,B${r + 1},SUM(B1:B${length}))`)
	}
	return workbook
}

function buildLookupExactWorkbook(rowCount: number, queryCount: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < rowCount; r++) {
		setStringCell(workbook, r, 0, `key-${String(r + 1).padStart(5, '0')}`)
		setNumberCell(workbook, r, 1, (r + 1) * 10)
	}
	for (let r = 0; r < queryCount; r++) {
		const keyIndex = rowCount - ((r * 37) % rowCount) - 1
		setStringCell(workbook, r, 3, `key-${String(keyIndex + 1).padStart(5, '0')}`)
		setFormulaCell(
			workbook,
			r,
			4,
			`XLOOKUP(D${r + 1},A$1:A$${rowCount},B$1:B$${rowCount},"missing",0)`,
		)
	}
	return workbook
}

function buildSpillChurnWorkbook(length: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < length; r++) {
		setNumberCell(workbook, r, 0, length - r)
		setBooleanCell(workbook, r, 2, true)
	}
	setFormulaCell(workbook, 0, 1, `FILTER(A1:A${length},C1:C${length},0)`)
	return workbook
}

function buildStructuralInsertRowsWorkbook(rowCount: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < rowCount; r++) {
		setNumberCell(workbook, r, 0, r + 1)
		setFormulaCell(workbook, r, 1, `SUM(A1:A${r + 1})`)
	}
	return workbook
}

function buildSdkEditCycleWorkbook(rows: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	for (let r = 0; r < rows; r++) {
		setNumberCell(workbook, r, 0, r + 1)
		setFormulaCell(workbook, r, 1, `A${r + 1}*2`)
	}
	recalculate(workbook, defaultCalcContext())
	return workbook
}

function buildDefinedNameHeavyWorkbook(nameCount: number, formulaCount: number): Workbook {
	const workbook = createWorkbook()
	const sheet = workbook.addSheet('Sheet1')
	for (let i = 0; i < nameCount; i++) {
		const row = i
		sheet.cells.set(row, 0, {
			value: numberValue(i + 1),
			formula: null,
			styleId: SID,
		})
		workbook.definedNames.set(`Name${i + 1}`, `Sheet1!A${row + 1}`)
	}
	for (let i = 0; i < formulaCount; i++) {
		const left = (i % nameCount) + 1
		const right = ((i + 1) % nameCount) + 1
		sheet.cells.set(i, 2, {
			value: EMPTY,
			formula: `Name${left}+Name${right}`,
			styleId: SID,
		})
	}
	return workbook
}

function buildSdkDefinedNameEditWorkbook(nameCount: number): Workbook {
	const workbook = createWorkbook()
	const sheet = workbook.addSheet('Sheet1')
	for (let i = 0; i < nameCount; i++) {
		sheet.cells.set(i, 0, {
			value: numberValue(i + 1),
			formula: null,
			styleId: SID,
		})
		workbook.definedNames.set(`Name${i + 1}`, `Sheet1!A${i + 1}`)
	}
	sheet.cells.set(0, 2, { value: EMPTY, formula: 'Name1+Name2', styleId: SID })
	recalculate(workbook, defaultCalcContext())
	return workbook
}

const scenarios: readonly Scenario[] = [
	{
		name: 'write-dense-40k',
		category: 'write',
		build() {
			const workbook = buildDenseWorkbook(2000, 20)
			return { workbook, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			mustWrite(requireWorkbook(input))
		},
	},
	{
		name: 'write-large-100k',
		category: 'write',
		build() {
			const workbook = buildDenseWorkbook(5000, 20)
			return { workbook, rows: 5000, cols: 20, cells: 100_000 }
		},
		run(input) {
			mustWrite(requireWorkbook(input))
		},
	},
	{
		name: 'write-multi-sheet',
		category: 'write',
		build() {
			const workbook = buildMultiSheetWorkbook(8, 1000, 10)
			return { workbook, rows: 1000, cols: 10, cells: 80_000 }
		},
		run(input) {
			mustWrite(requireWorkbook(input))
		},
	},
	{
		name: 'roundtrip-dense-40k',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(2000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
			mustWrite(result.value.workbook)
		},
	},
	{
		name: 'read-write-shared-formulas',
		category: 'read',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			for (let r = 0; r < 10_000; r++) {
				sheet.cells.set(r, 0, {
					value: numberValue(r + 1),
					formula: null,
					styleId: SID,
				})
				sheet.cells.set(r, 1, {
					value: EMPTY,
					formula: `A${r + 1}*2`,
					styleId: SID,
				})
			}
			const bytes = mustWrite(workbook)
			return { bytes, rows: 10_000, cols: 2, cells: 20_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
			mustWrite(result.value.workbook)
		},
	},
	{
		name: 'read-large-200k',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(10_000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 10_000, cols: 20, cells: 200_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-metadata-dense',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(2000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input), { mode: 'metadata-only' })
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-full-dense',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(2000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-full-string-dense',
		category: 'read',
		build() {
			const workbook = buildStringDenseWorkbook(2000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-full-sparse',
		category: 'read',
		build() {
			const workbook = buildSparseWorkbook(50_000, 8, 200)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 50_000, cols: 8, cells: 2_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-selective-sheet',
		category: 'read',
		build() {
			const workbook = buildMultiSheetWorkbook(4, 800, 10)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 800, cols: 10, cells: 32_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input), { sheets: ['Sheet3'] })
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-window-dense-values',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(5000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 5000, cols: 20, cells: 100_000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input), { mode: 'values' })
			const window = wb.readWindow('Sheet1', 'A1:T5000', { rowLimit: 250 })
			if (!window) throw new Error('Dense window benchmark failed to read Sheet1')
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
				},
			}
		},
	},
	{
		name: 'read-window-dense-values-compact',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(5000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 5000, cols: 20, cells: 100_000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input), { mode: 'values' })
			const window = wb.readWindowCompact('Sheet1', 'A1:T5000', {
				rowLimit: 250,
				includeRefs: false,
			})
			if (!window) throw new Error('Compact dense window benchmark failed to read Sheet1')
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
				},
			}
		},
	},
	{
		name: 'read-window-formula-chain-compact',
		category: 'read',
		build() {
			const workbook = buildFormulaChainWorkbook(6000)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 6000, cols: 1, cells: 6000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input), { mode: 'formula' })
			const window = wb.readWindowCompact('Sheet1', 'A1:A6000', {
				rowLimit: 500,
				includeRefs: false,
			})
			if (!window) throw new Error('Formula chain window benchmark failed to read Sheet1')
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
					formulaCount: window.cells.filter((cell) => cell.formula !== null).length,
				},
			}
		},
	},
	{
		name: 'read-window-sparse-wide',
		category: 'read',
		build() {
			const workbook = buildSparseWorkbook(100_000, 20, 500)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 100_000, cols: 20, cells: 4_000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input), { mode: 'values' })
			const window = wb.readWindow('Sheet1', 'A1:T100000', { rowLimit: 5000 })
			if (!window) throw new Error('Sparse window benchmark failed to read Sheet1')
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
				},
			}
		},
	},
	{
		name: 'workflow-reopen-values-window',
		category: 'workflow',
		build() {
			const workbook = buildDenseWorkbook(4000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 4000, cols: 20, cells: 80_000 }
		},
		async run(input) {
			let totalCells = 0
			for (let i = 0; i < 3; i++) {
				const wb = await AscendWorkbook.open(requireBytes(input), { mode: 'values' })
				const window = wb.readWindow('Sheet1', 'A1:T4000', { rowLimit: 200 })
				if (!window) throw new Error('Workflow benchmark failed to read Sheet1')
				totalCells += window.cells.length
			}
			return {
				assertions: {
					iterations: 3,
					totalCellsRead: totalCells,
				},
			}
		},
	},
	{
		name: 'workflow-sdk-edit-cycle',
		category: 'workflow',
		build() {
			const workbook = buildSdkEditCycleWorkbook(5000)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 5000, cols: 2, cells: 10_000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input))
			const updates = Array.from({ length: 100 }, (_, index) => ({
				ref: `A${index + 1}`,
				value: (index + 1) * 3,
			}))
			const ops = [{ op: 'setCells', sheet: 'Sheet1', updates }] as const
			const preview = wb.preview(ops)
			const apply = wb.apply(ops)
			const recalc = wb.recalc()
			const bytes = wb.toBytes()
			return {
				assertions: {
					previewChanges: preview.cellChanges.length,
					applyErrors: apply.errors.length,
					recalcChanged: recalc.changed.length,
					bytes: bytes.byteLength,
				},
			}
		},
	},
	{
		name: 'workflow-sdk-defined-names-edit-cycle',
		category: 'workflow',
		build() {
			const workbook = buildSdkDefinedNameEditWorkbook(5000)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 5000, cols: 3, cells: 15_000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input))
			const preview = wb.preview([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 99 }] },
			])
			const apply = wb.apply([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 99 }] },
			])
			const recalc = wb.recalc()
			return {
				assertions: {
					previewChanges: preview.cellChanges.length,
					applyErrors: apply.errors.length,
					recalcChanged: recalc.changed.length,
				},
			}
		},
	},
	{
		name: 'formula-inspect-chain',
		category: 'workflow',
		build() {
			const workbook = buildFormulaChainWorkbook(6000)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 6000, cols: 1, cells: 6000 }
		},
		async run(input) {
			const wb = await AscendWorkbook.open(requireBytes(input))
			const info = wb.formula('Sheet1!A6000')
			if (!info) throw new Error('Formula inspect benchmark could not load formula target')
			return {
				assertions: {
					volatile: info.volatile,
					refCount: info.refs.length,
					functionCount: info.functions.length,
				},
			}
		},
	},
	{
		name: 'recalc-formula-chain',
		category: 'calc',
		build() {
			const workbook = buildFormulaChainWorkbook(6000)
			return { workbook, rows: 6000, cols: 1, cells: 6000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-range-aggregation',
		category: 'calc',
		build() {
			const workbook = buildRangeAggregationWorkbook(800)
			return { workbook, rows: 800, cols: 2, cells: 1600 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-if-short-circuit',
		category: 'calc',
		build() {
			const workbook = buildIfShortCircuitWorkbook(5000)
			return { workbook, rows: 5000, cols: 3, cells: 15_000 }
		},
		run(input) {
			const workbook = requireWorkbook(input)
			recalculate(workbook, defaultCalcContext())
			const lastValue = workbook.sheets[0]?.cells.get(4999, 2)?.value
			return {
				assertions: {
					lastValue: lastValue?.kind === 'number' ? lastValue.value : null,
				},
			}
		},
	},
	{
		name: 'recalc-sparse-aggregation',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			const random = createDeterministicRandom(0xc0ffee)
			for (let r = 0; r < 100_000; r += 100) {
				sheet.cells.set(r, 0, {
					value: numberValue(random() * 1000),
					formula: null,
					styleId: SID,
				})
			}
			sheet.cells.set(0, 1, { value: EMPTY, formula: 'SUM(A1:A100000)', styleId: SID })
			return { workbook, rows: 100_000, cols: 2, cells: 1001 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-lookup-exact-large',
		category: 'calc',
		build() {
			const workbook = buildLookupExactWorkbook(20_000, 2_000)
			return { workbook, rows: 20_000, cols: 5, cells: 44_000 }
		},
		run(input) {
			const workbook = requireWorkbook(input)
			recalculate(workbook, defaultCalcContext())
			const lastValue = workbook.sheets[0]?.cells.get(1999, 4)?.value
			return {
				assertions: {
					lastValue: lastValue?.kind === 'number' ? lastValue.value : null,
				},
			}
		},
	},
	{
		name: 'recalc-lookup-exact-incremental',
		category: 'calc',
		build() {
			const workbook = buildLookupExactWorkbook(20_000, 2_000)
			recalculate(workbook, defaultCalcContext())
			for (let r = 0; r < 2_000; r++) {
				setStringCell(workbook, r, 3, `key-${String(((r * 53) % 20_000) + 1).padStart(5, '0')}`)
			}
			return { workbook, rows: 20_000, cols: 5, cells: 44_000 }
		},
		run(input) {
			const workbook = requireWorkbook(input)
			recalculate(workbook, defaultCalcContext(), {
				dirtyOnly: true,
				dirtyRefs: ['Sheet1!D1:D2000'],
			})
			const lastValue = workbook.sheets[0]?.cells.get(1999, 4)?.value
			return {
				assertions: {
					lastValue: lastValue?.kind === 'number' ? lastValue.value : null,
				},
			}
		},
	},
	{
		name: 'recalc-defined-names-heavy',
		category: 'calc',
		build() {
			const workbook = buildDefinedNameHeavyWorkbook(2000, 2000)
			return { workbook, rows: 2000, cols: 3, cells: 6000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-incremental',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			for (let r = 0; r < 1000; r++) {
				sheet.cells.set(r, 0, {
					value: numberValue(r + 1),
					formula: null,
					styleId: SID,
				})
				sheet.cells.set(r, 1, {
					value: EMPTY,
					formula: `A${r + 1}*2`,
					styleId: SID,
				})
			}
			sheet.cells.set(0, 2, { value: EMPTY, formula: 'SUM(B1:B1000)', styleId: SID })
			recalculate(workbook, defaultCalcContext())
			sheet.cells.set(499, 0, {
				value: numberValue(999),
				formula: null,
				styleId: SID,
			})
			return { workbook, rows: 1000, cols: 3, cells: 2002 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext(), {
				dirtyOnly: true,
				dirtyRefs: ['Sheet1!A500'],
			})
		},
	},
	{
		name: 'recalc-cross-sheet',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			const sheet1 = workbook.addSheet('Sheet1')
			for (let r = 0; r < 1000; r++) {
				sheet1.cells.set(r, 0, {
					value: numberValue(r + 1),
					formula: null,
					styleId: SID,
				})
			}
			for (let s = 1; s < 5; s++) {
				const sheet = workbook.addSheet(`Sheet${s + 1}`)
				for (let r = 0; r < 1000; r++) {
					sheet.cells.set(r, 0, {
						value: EMPTY,
						formula: `Sheet1!A${r + 1}*2`,
						styleId: SID,
					})
				}
			}
			return { workbook, rows: 1000, cols: 1, cells: 5000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-dynamic-array',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			for (let r = 0; r < 1000; r++) {
				setNumberCell(workbook, r, 0, r + 1)
			}
			setFormulaCell(workbook, 0, 1, 'SEQUENCE(1000)')
			setFormulaCell(workbook, 0, 2, 'SORT(A1:A1000)')
			return { workbook, rows: 1000, cols: 3, cells: 3000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-dynamic-spill-churn',
		category: 'calc',
		build() {
			const workbook = buildSpillChurnWorkbook(2000)
			return { workbook, rows: 2000, cols: 3, cells: 6000 }
		},
		run(input) {
			const workbook = requireWorkbook(input)
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			recalculate(workbook, defaultCalcContext())
			for (let r = 0; r < 2000; r++) {
				sheet.cells.set(r, 2, {
					value: booleanValue(r % 10 === 0),
					formula: null,
					styleId: SID,
				})
			}
			recalculate(workbook, defaultCalcContext())
			for (let r = 0; r < 2000; r++) {
				sheet.cells.set(r, 2, {
					value: booleanValue(true),
					formula: null,
					styleId: SID,
				})
			}
			recalculate(workbook, defaultCalcContext())
			const anchorValue = sheet.cells.get(0, 1)?.value
			return {
				assertions: {
					anchorKind: anchorValue?.kind ?? 'missing',
					anchorValue: anchorValue?.kind === 'number' ? anchorValue.value : null,
				},
			}
		},
	},
	{
		name: 'recalc-sumifs-large',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			const categories = ['cat1', 'cat2', 'cat3', 'cat4', 'cat5']
			const random = createDeterministicRandom(0x5eed1234)
			for (let r = 0; r < 10_000; r++) {
				sheet.cells.set(r, 0, {
					value: { kind: 'string', value: categories[r % 5] ?? 'cat1' },
					formula: null,
					styleId: SID,
				})
				sheet.cells.set(r, 1, {
					value: numberValue(random() * 1000),
					formula: null,
					styleId: SID,
				})
			}
			for (let i = 0; i < 5; i++) {
				setFormulaCell(workbook, i, 3, `SUMIFS(B1:B10000,A1:A10000,"${categories[i]}")`)
			}
			return { workbook, rows: 10_000, cols: 4, cells: 20_005 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-criteria-caching',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			const categories = ['cat1', 'cat2', 'cat3', 'cat4', 'cat5']
			const random = createDeterministicRandom(0xdeadbeef)
			for (let r = 0; r < 10_000; r++) {
				sheet.cells.set(r, 0, {
					value: { kind: 'string', value: categories[r % 5] ?? 'cat1' },
					formula: null,
					styleId: SID,
				})
				sheet.cells.set(r, 1, {
					value: numberValue(random() * 1000),
					formula: null,
					styleId: SID,
				})
			}
			for (let i = 0; i < 1000; i++) {
				const cat = categories[i % 5] ?? 'cat1'
				setFormulaCell(workbook, i, 3, `SUMIFS(B$1:B$10000,A$1:A$10000,"${cat}")`)
			}
			return { workbook, rows: 10_000, cols: 4, cells: 21_000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-quickselect',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			const random = createDeterministicRandom(0x12345678)
			for (let r = 0; r < 10_000; r++) {
				sheet.cells.set(r, 0, {
					value: numberValue(random() * 10000),
					formula: null,
					styleId: SID,
				})
			}
			for (let i = 0; i < 50; i++) {
				setFormulaCell(workbook, i, 1, `LARGE(A$1:A$10000,${i + 1})`)
				setFormulaCell(workbook, i, 2, `SMALL(A$1:A$10000,${i + 1})`)
			}
			return { workbook, rows: 10_000, cols: 3, cells: 10_100 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'structural-insert-rows-recalc',
		category: 'calc',
		build() {
			const workbook = buildStructuralInsertRowsWorkbook(4000)
			return { workbook, rows: 4000, cols: 2, cells: 8000 }
		},
		run(input) {
			const workbook = requireWorkbook(input)
			const result = applyOperations(workbook, [
				{ op: 'insertRows', sheet: 'Sheet1', at: 2000, count: 200 },
			])
			if (!result.ok) throw new Error(result.error.message)
			recalculate(workbook, defaultCalcContext())
			return {
				assertions: {
					changedRefs: result.value.affectedCells.length,
					sheetsModified: result.value.sheetsModified.length,
				},
			}
		},
	},
	{
		name: 'read-csv-large',
		category: 'read',
		build() {
			const rows: string[] = []
			for (let r = 0; r < 50_000; r++) {
				const row: string[] = []
				for (let c = 0; c < 10; c++) {
					row.push(String(r * 10 + c + 1))
				}
				rows.push(row.join(','))
			}
			const content = rows.join('\n')
			return { content, rows: 50_000, cols: 10, cells: 500_000 }
		},
		run(input) {
			const result = readCsv(requireContent(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'write-csv-large',
		category: 'write',
		build() {
			const workbook = buildDenseWorkbook(50_000, 10)
			return { workbook, rows: 50_000, cols: 10, cells: 500_000 }
		},
		run(input) {
			const result = writeCsv(requireWorkbook(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
]

const scenarioSets = {
	smoke: [
		'read-full-dense',
		'workflow-sdk-edit-cycle',
		'workflow-sdk-defined-names-edit-cycle',
		'recalc-incremental',
		'recalc-if-short-circuit',
		'recalc-lookup-exact-incremental',
		'recalc-dynamic-spill-churn',
		'recalc-criteria-caching',
		'recalc-quickselect',
		'structural-insert-rows-recalc',
		'read-csv-large',
	],
} as const

function getRssBytes(): number | undefined {
	try {
		const usage = process.memoryUsage()
		return typeof (process.memoryUsage as { rss?: () => number }).rss === 'function'
			? (process.memoryUsage as { rss: () => number }).rss()
			: typeof usage.rss === 'number'
				? usage.rss
				: undefined
	} catch {
		return undefined
	}
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		// Best effort only.
	}
}

function renderSummary(results: readonly BenchmarkCaseResult[]): string {
	const headers = [
		'scenario',
		'category',
		'median-ms',
		'p95-ms',
		'cells',
		'bytes',
		'throughput',
		'rss-delta',
		'retained-rss',
	]
	const rows = results.map((result) => [
		result.name,
		result.category,
		result.metrics.medianMs.toFixed(2),
		result.metrics.p95Ms.toFixed(2),
		String(result.dimensions.cells ?? 'n/a'),
		typeof result.dimensions.bytes === 'number' ? formatBytes(result.dimensions.bytes) : 'n/a',
		result.metrics.throughputPerSec !== undefined
			? formatRate(result.metrics.throughputPerSec)
			: 'n/a',
		result.metrics.rssDeltaBytes !== undefined ? formatBytes(result.metrics.rssDeltaBytes) : 'n/a',
		result.metrics.retainedRssDeltaBytes !== undefined
			? formatBytes(result.metrics.retainedRssDeltaBytes)
			: 'n/a',
	])

	const widths = headers.map((header, index) =>
		Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)),
	)
	const pad = (value: string, width: number) =>
		value + ' '.repeat(Math.max(0, width - value.length))
	const line = (cells: readonly string[]) =>
		cells.map((cell, index) => pad(cell, widths[index] ?? 0)).join('  ')

	return [
		line(headers),
		widths.map((width) => '─'.repeat(width)).join('──'),
		...rows.map(line),
	].join('\n')
}

async function runScenario(
	scenario: Scenario,
	repeat: number,
	warmup: number,
): Promise<BenchmarkCaseResult> {
	const samples: Array<{
		readonly durationMs: number
		readonly throughputPerSec: number
		readonly rssDeltaBytes?: number
		readonly retainedRssDeltaBytes?: number
	}> = []
	let firstInput: ScenarioInput | undefined
	let assertions: Record<string, string | number | boolean | null> | undefined
	for (let i = 0; i < warmup; i++) {
		const input = scenario.build()
		firstInput ??= input
		await scenario.run(input)
	}
	runGc()
	for (let i = 0; i < repeat; i++) {
		const input = scenario.build()
		firstInput ??= input
		runGc()
		const rssBefore = getRssBytes()
		const start = performance.now()
		const runResult = await scenario.run(input)
		const durationMs = performance.now() - start
		const rssAfter = getRssBytes()
		runGc()
		const rssAfterGc = getRssBytes()
		samples.push({
			durationMs,
			throughputPerSec:
				durationMs > 0 ? (input.cells / durationMs) * 1000 : Number.POSITIVE_INFINITY,
			rssDeltaBytes:
				rssBefore !== undefined && rssAfter !== undefined
					? Math.max(0, rssAfter - rssBefore)
					: undefined,
			retainedRssDeltaBytes:
				rssBefore !== undefined && rssAfterGc !== undefined
					? Math.max(0, rssAfterGc - rssBefore)
					: undefined,
		})
		assertions ??= runResult?.assertions
	}
	const input = firstInput
	if (!input) throw new Error(`Scenario "${scenario.name}" did not produce input`)
	return {
		name: scenario.name,
		category: scenario.category,
		dimensions: {
			rows: input.rows,
			cols: input.cols,
			cells: input.cells,
			bytes: input.bytes?.byteLength ?? 0,
			repeat,
		},
		metrics: summarizeSamples(samples),
		...(repeat > 1 ? { samples } : {}),
		...(assertions ? { assertions } : {}),
	}
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	return index >= 0 ? process.argv[index + 1] : undefined
}

function hasFlag(name: string): boolean {
	return process.argv.includes(name)
}

async function runScenarioIsolated(
	scenario: Scenario,
	repeat: number,
	warmup: number,
	json: boolean,
): Promise<BenchmarkCaseResult> {
	const proc = Bun.spawn(
		[
			'bun',
			'run',
			process.argv[1] ?? import.meta.path,
			'--scenario',
			scenario.name,
			'--repeat',
			String(repeat),
			'--warmup',
			String(warmup),
			'--json',
		],
		{
			stdout: 'pipe',
			stderr: 'pipe',
			cwd: process.cwd(),
		},
	)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(stderr.trim() || `Synthetic benchmark scenario "${scenario.name}" failed`)
	}
	const parsed = JSON.parse(stdout) as BenchmarkCaseResult
	if (!json) {
		console.log(`completed ${scenario.name}`)
	}
	return parsed
}

async function main(): Promise<void> {
	const json = process.argv.includes('--json')
	const ci = hasFlag('--ci')
	const scenarioName = readFlag('--scenario')
	const scenarioSetName = readFlag('--set')
	const repeat = Math.max(1, Number.parseInt(readFlag('--repeat') ?? '1', 10) || 1)
	const warmup = Math.max(0, Number.parseInt(readFlag('--warmup') ?? '1', 10) || 1)
	const outputJson = json || ci
	if (scenarioName) {
		const scenario = scenarios.find((entry) => entry.name === scenarioName)
		if (!scenario) throw new Error(`Unknown synthetic benchmark scenario "${scenarioName}"`)
		const result = await runScenario(scenario, repeat, warmup)
		if (outputJson) {
			console.log(JSON.stringify(result, null, 2))
			return
		}
		console.log(renderSummary([result]))
		return
	}
	const selectedScenarios = scenarioSetName
		? (() => {
				const names = scenarioSets[scenarioSetName as keyof typeof scenarioSets]
				if (!names) throw new Error(`Unknown synthetic benchmark set "${scenarioSetName}"`)
				const selected = names
					.map((name) => scenarios.find((entry) => entry.name === name))
					.filter((entry): entry is Scenario => entry !== undefined)
				if (selected.length !== names.length) {
					throw new Error(`Synthetic benchmark set "${scenarioSetName}" is out of sync`)
				}
				return selected
			})()
		: scenarios
	const results: BenchmarkCaseResult[] = []
	for (const scenario of selectedScenarios) {
		results.push(await runScenarioIsolated(scenario, repeat, warmup, outputJson))
	}
	const suite = createBenchmarkSuite({
		suite: scenarioSetName
			? `ascend-synthetic-benchmarks-${scenarioSetName}`
			: 'ascend-synthetic-benchmarks',
		kind: 'synthetic',
		cases: results,
		metadata: {
			repeat,
			warmup,
			...(scenarioSetName ? { set: scenarioSetName } : {}),
		},
	})
	if (outputJson) {
		console.log(JSON.stringify(suite, null, 2))
		return
	}

	console.log('Ascend benchmark summary')
	console.log(renderSummary(results))
}

await main()
