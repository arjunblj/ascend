import { readFileSync } from 'node:fs'
import {
	createWorkbook,
	parseRange,
	type StyleId,
	type Workbook,
} from '../../packages/core/src/index.ts'
import {
	applyOperations,
	defaultCalcContext,
	recalculate,
} from '../../packages/engine/src/index.ts'
import { clearGlobalParseCache, parseFormula } from '../../packages/formulas/src/index.ts'
import { readCsv, writeCsv } from '../../packages/io-csv/src/index.ts'
import {
	extractZip,
	readXlsx,
	writeXlsx,
	writeXlsxStreaming,
} from '../../packages/io-xlsx/src/index.ts'
import {
	booleanValue,
	dateValue,
	EMPTY,
	numberValue,
	stringValue,
} from '../../packages/schema/src/index.ts'
import {
	AscendSession,
	AscendWorkbook,
	SheetHandle,
	WorkbookDocument,
	WorkbookSession,
} from '../../packages/sdk/src/index.ts'
import {
	type BenchmarkCaseResult,
	createBenchmarkSuite,
	formatBytes,
	formatRate,
	summarizeSamples,
} from './results.ts'

const SID = 0 as StyleId
const xlsxSharedStringUsageCache = new WeakMap<Uint8Array, SharedStringUsageAssertions>()
const xlsxWorksheetScalarUsageCache = new WeakMap<Uint8Array, WorksheetScalarUsageAssertions>()
const WORKSHEET_SCALAR_USAGE_CELL_SAMPLE_LIMIT = 5_000

interface ScenarioInput {
	readonly workbook?: Workbook
	readonly bytes?: Uint8Array
	readonly byteCount?: number
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

function makeBenchmarkSheetHandle(workbook: Workbook, name = 'Sheet1'): SheetHandle {
	const sheet = workbook.getSheet(name)
	if (!sheet) throw new Error(`Benchmark workbook missing ${name}`)
	return new SheetHandle(
		name,
		() => sheet,
		(_row, _col, cell) => cell.formula,
		(_row, _col, cell) => {
			switch (cell.value.kind) {
				case 'number':
				case 'string':
				case 'boolean':
					return String(cell.value.value)
				default:
					return ''
			}
		},
		() => ({
			token: 'benchmark',
			generations: { workbook: 0, sheetMetadata: 0, formulas: 0, styles: 0 },
			load: {
				mode: 'full',
				isPartial: false,
				cellsHydrated: true,
				richSheetMetadataHydrated: true,
				hasAllSheets: true,
				partialReasons: [],
				sourceSheets: [name],
				loadedSheets: [name],
			},
		}),
	)
}

function mustWrite(workbook: Workbook, options?: Parameters<typeof writeXlsx>[2]): Uint8Array {
	const result = writeXlsx(workbook, undefined, options)
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

function buildPlainNumericWorkbook(rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	const sheet = workbook.addSheet('Sheet1')
	const values = Array.from({ length: cols }, (_, col) => col + 1)
	for (let row = 0; row < rows; row++) {
		sheet.cells.setPlainNumberSpan(row, 0, values)
	}
	return workbook
}

function buildMixedPlainWorkbook(rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	const sheet = workbook.addSheet('Sheet1')
	for (let row = 0; row < rows; row++) {
		for (let col = 0; col < cols; col++) {
			if ((col & 1) === 0) {
				sheet.cells.setPlainNumber(row, col, row * cols + col + 1)
			} else {
				sheet.cells.set(row, col, {
					value: stringValue(`label-${col % 8}`),
					formula: null,
					styleId: SID,
				})
			}
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

function buildUniqueStringWorkbook(rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			sheet.cells.set(r, c, {
				value: stringValue(`unique-${r}-${c}`),
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

function buildMixedDataWorkbook(rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			const kind = (r + c) % 4
			if (kind === 0) {
				sheet.cells.set(r, c, {
					value: numberValue(r * cols + c + 1),
					formula: null,
					styleId: SID,
				})
			} else if (kind === 1) {
				sheet.cells.set(r, c, {
					value: stringValue(`row-${r}-col-${c}`),
					formula: null,
					styleId: SID,
				})
			} else if (kind === 2) {
				sheet.cells.set(r, c, {
					value: booleanValue((r + c) % 2 === 0),
					formula: null,
					styleId: SID,
				})
			} else {
				sheet.cells.set(r, c, {
					value: dateValue(44927 + (r % 365)),
					formula: null,
					styleId: SID,
				})
			}
		}
	}
	return workbook
}

function buildReadStyleHeavyWorkbook(rows: number, uniqueStyles: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	const fonts = ['Arial', 'Calibri', 'Helvetica', 'Times New Roman', 'Courier New']
	const formats = ['General', '0.00', '#,##0', '0.0%', '[$-409]m/d/yy', '@']
	const styleIds: StyleId[] = []
	for (let i = 0; i < uniqueStyles; i++) {
		const styleId = workbook.styles.register({
			font: {
				name: fonts[i % fonts.length],
				size: 9 + (i % 8),
				bold: i % 5 === 0,
				color: {
					kind: 'rgb',
					rgb: `FF${String((i * 37) % 256).padStart(2, '0')}${String((i * 17) % 256).padStart(2, '0')}${String((i * 7) % 256).padStart(2, '0')}`,
				},
			},
			fill: {
				pattern: 'solid',
				fgColor: {
					kind: 'rgb',
					rgb: `FF${String((i * 13) % 256).padStart(2, '0')}${String((i * 31) % 256).padStart(2, '0')}${String((i * 19) % 256).padStart(2, '0')}`,
				},
			},
			numberFormat: formats[i % formats.length],
		}) as StyleId
		styleIds.push(styleId)
		sheet.cells.set(i, 0, {
			value: numberValue(i + 1),
			formula: null,
			styleId,
		})
	}
	for (let r = uniqueStyles; r < rows; r++) {
		const styleId = styleIds[r % uniqueStyles] ?? SID
		sheet.cells.set(r, 0, {
			value: numberValue(r + 1),
			formula: null,
			styleId,
		})
	}
	return workbook
}

function buildReadStringHeavyWorkbook(rows: number, uniqueStrings: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	for (let r = 0; r < rows; r++) {
		const idx = r % uniqueStrings
		sheet.cells.set(r, 0, {
			value: stringValue(`unique-string-${idx}-${'x'.repeat((idx % 20) + 1)}`),
			formula: null,
			styleId: SID,
		})
		sheet.cells.set(r, 1, {
			value: stringValue(`another-${idx}-${String(r).padStart(5, '0')}`),
			formula: null,
			styleId: SID,
		})
	}
	return workbook
}

function buildReadFormulaDenseWorkbook(rows: number, _cols: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	for (let r = 0; r < rows; r++) {
		const row1 = r + 1
		sheet.cells.set(r, 0, {
			value: numberValue(row1),
			formula: null,
			styleId: SID,
		})
		setFormulaCell(workbook, r, 1, `A${row1}*2`)
		setFormulaCell(workbook, r, 2, `SUM(A$1:A${row1})`)
		setFormulaCell(workbook, r, 3, `B${row1}+C${row1}`)
		setFormulaCell(workbook, r, 4, `IF(A${row1}>100,"big","small")`)
	}
	return workbook
}

function buildStyleHeavyWorkbook(cellCount: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	const fonts = ['Arial', 'Calibri', 'Helvetica', 'Times New Roman', 'Courier New']
	for (let i = 0; i < cellCount; i++) {
		const row = Math.floor(i / 100)
		const col = i % 100
		const styleId = workbook.styles.register({
			font: {
				name: fonts[i % fonts.length],
				size: 10 + (i % 6),
				bold: i % 3 === 0,
				color: {
					kind: 'rgb',
					rgb: `FF${String((i * 37) % 256).padStart(2, '0')}${String((i * 17) % 256).padStart(2, '0')}${String((i * 7) % 256).padStart(2, '0')}`,
				},
			},
			fill: {
				pattern: 'solid',
				fgColor: {
					kind: 'rgb',
					rgb: `FF${String((i * 13) % 256).padStart(2, '0')}${String((i * 31) % 256).padStart(2, '0')}${String((i * 19) % 256).padStart(2, '0')}`,
				},
			},
			numberFormat: ['0', '0.00', '#,##0', '0.0%', '[$-409]m/d/yy', 'General'][i % 6],
		}) as StyleId
		sheet.cells.set(row, col, {
			value: numberValue(i + 1),
			formula: null,
			styleId,
		})
	}
	return workbook
}

function buildStructuralInsertLargeWorkbook(rows: number, cols: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			sheet.cells.set(r, c, {
				value: numberValue(r * cols + c + 1),
				formula: null,
				styleId: SID,
			})
		}
	}
	const colLetter = (n: number) =>
		n < 26
			? String.fromCharCode(65 + n)
			: String.fromCharCode(64 + Math.floor(n / 26)) + String.fromCharCode(65 + (n % 26))
	sheet.cells.set(0, cols, {
		value: EMPTY,
		formula: `SUM(A1:${colLetter(cols - 1)}${rows})`,
		styleId: SID,
	})
	sheet.cells.set(1, cols, { value: EMPTY, formula: `AVERAGE(A1:A${rows})`, styleId: SID })
	sheet.cells.set(2, cols, { value: EMPTY, formula: `SUM(A1:B${rows})`, styleId: SID })
	return workbook
}

function buildFormulaCompilationUniqueWorkbook(formulaCount: number): Workbook {
	const workbook = createWorkbook()
	workbook.addSheet('Sheet1')
	const sheet = workbook.sheets[0]
	if (!sheet) throw new Error('Benchmark workbook missing first sheet')
	for (let r = 0; r < formulaCount; r++) {
		sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: SID })
		sheet.cells.set(r, 1, {
			value: EMPTY,
			formula: `A${r + 1}+${r + 1}`,
			styleId: SID,
		})
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

const interactiveSessionCache = new WeakMap<Uint8Array, Promise<AscendSession>>()
const patchStreamStateCache = new WeakMap<
	Uint8Array,
	Promise<{
		readonly session: AscendSession
		lastToken: string
		iteration: number
	}>
>()
let cachedDenseViewportBytes: Uint8Array | undefined
let cachedSparseWideViewportBytes: Uint8Array | undefined
let cachedSemanticViewportBytes: Uint8Array | undefined
let cachedManyOverlayViewportBytes: Uint8Array | undefined
let cachedRealRichViewportBytes: Uint8Array | undefined
const realInteractivePatchCorpusBytes = new Map<string, Uint8Array>()

interface RealInteractivePatchTarget {
	readonly label: string
	readonly path: string
	readonly sheet: string
	readonly editRef: string
	readonly topRow: number
	readonly leftCol: number
	readonly rowCount: number
	readonly colCount: number
}

const REAL_INTERACTIVE_PATCH_CORPUS: readonly RealInteractivePatchTarget[] = [
	{
		label: 'xlsxwriter-styles-formulas',
		path: 'fixtures/xlsx/xlsxwriter/styles_formulas.xlsx',
		sheet: 'Data',
		editRef: 'C2',
		topRow: 0,
		leftCol: 0,
		rowCount: 100,
		colCount: 20,
	},
	{
		label: 'xlsxwriter-strings-links',
		path: 'fixtures/xlsx/xlsxwriter/strings_links.xlsx',
		sheet: 'Strings',
		editRef: 'A2',
		topRow: 0,
		leftCol: 0,
		rowCount: 50,
		colCount: 12,
	},
	{
		label: 'poi-with-table',
		path: 'fixtures/xlsx/poi/WithTable.xlsx',
		sheet: 'Foglio1',
		editRef: 'A2',
		topRow: 0,
		leftCol: 0,
		rowCount: 50,
		colCount: 12,
	},
	{
		label: 'poi-structured-references',
		path: 'fixtures/xlsx/poi/StructuredReferences.xlsx',
		sheet: 'Formulas',
		editRef: 'A2',
		topRow: 0,
		leftCol: 0,
		rowCount: 50,
		colCount: 12,
	},
	{
		label: 'poi-comments',
		path: 'fixtures/xlsx/poi/comments.xlsx',
		sheet: 'Sheet1',
		editRef: 'A1',
		topRow: 0,
		leftCol: 0,
		rowCount: 50,
		colCount: 12,
	},
	{
		label: 'poi-merge-cells',
		path: 'fixtures/xlsx/poi/merge_cells.xlsx',
		sheet: 'Merge',
		editRef: 'A1',
		topRow: 0,
		leftCol: 0,
		rowCount: 50,
		colCount: 12,
	},
	{
		label: 'stress-dense-100k',
		path: 'fixtures/xlsx/stress/dense-100k.xlsx',
		sheet: 'Data',
		editRef: 'A1',
		topRow: 0,
		leftCol: 0,
		rowCount: 100,
		colCount: 20,
	},
]
const DENSE_READINESS_THINK_DELAYS_MS = [0, 50, 150, 300] as const

function realInteractivePatchCorpusTarget(label: string): RealInteractivePatchTarget {
	const target = REAL_INTERACTIVE_PATCH_CORPUS.find((candidate) => candidate.label === label)
	if (!target) throw new Error(`Missing real interactive patch corpus target: ${label}`)
	return target
}

function cachedDenseViewportWorkbookBytes(): Uint8Array {
	cachedDenseViewportBytes ??= mustWrite(buildDenseWorkbook(5000, 20))
	return cachedDenseViewportBytes
}

function cachedSparseWideViewportWorkbookBytes(): Uint8Array {
	cachedSparseWideViewportBytes ??= mustWrite(buildSparseWorkbook(100_000, 20, 500))
	return cachedSparseWideViewportBytes
}

function cachedSemanticViewportWorkbookBytes(): Uint8Array {
	if (cachedSemanticViewportBytes) return cachedSemanticViewportBytes
	const workbook = AscendWorkbook.create()
	workbook.apply([
		{
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [
				{ ref: 'A1', value: 'Region' },
				{ ref: 'B1', value: 'Status' },
				{ ref: 'C1', value: 'Amount' },
				...Array.from({ length: 100 }, (_, index) => [
					{ ref: `A${index + 2}`, value: index % 2 === 0 ? 'West' : 'East' },
					{ ref: `B${index + 2}`, value: index % 3 === 0 ? 'Open' : 'Closed' },
					{ ref: `C${index + 2}`, value: (index + 1) * 10 },
				]).flat(),
			],
		},
		{ op: 'setFormula', sheet: 'Sheet1', ref: 'D2', formula: 'C2*2' },
		{ op: 'setStyle', sheet: 'Sheet1', range: 'C2:C101', style: { numberFormat: '$0.00' } },
		{ op: 'setComment', sheet: 'Sheet1', ref: 'E5', text: 'review', author: 'agent' },
		{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'F5', url: 'https://example.com' },
		{
			op: 'setDataValidation',
			sheet: 'Sheet1',
			range: 'B2:B101',
			rule: { type: 'list', formula1: '"Open,Closed"', allowBlank: true },
		},
		{
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'C2:C101',
			rule: { type: 'cellIs', operator: 'greaterThan', formula: '100', priority: 1 },
		},
		{ op: 'mergeCells', sheet: 'Sheet1', range: 'H1:I1' },
		{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C101', column: 1, values: ['Open'] },
		{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:C101', name: 'Sales', hasHeaders: true },
		{ op: 'freezePane', sheet: 'Sheet1', row: 1, col: 1 },
	])
	cachedSemanticViewportBytes = workbook.toBytes()
	return cachedSemanticViewportBytes
}

function cachedManyOverlayViewportWorkbookBytes(): Uint8Array {
	if (cachedManyOverlayViewportBytes) return cachedManyOverlayViewportBytes
	const workbook = AscendWorkbook.create()
	const sheet = workbook.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('Benchmark workbook missing Sheet1')
	for (let row = 0; row < 250; row++) {
		for (let col = 0; col < 20; col++) {
			sheet.cells.set(row, col, {
				value: numberValue(row * 20 + col + 1),
				formula: null,
				styleId: SID,
			})
		}
		const rowRange = `A${row + 1}:T${row + 1}`
		sheet.dataValidations.push({
			sqref: rowRange,
			type: 'whole',
			operator: 'greaterThan',
			formula1: '0',
		})
		sheet.conditionalFormats.push({
			sqref: rowRange,
			rules: [{ type: 'expression', formulas: [`$A${row + 1}>0`], priority: row + 1 }],
		})
		sheet.merges.push(parseRange(`S${row + 1}:T${row + 1}`))
	}
	cachedManyOverlayViewportBytes = workbook.toBytes()
	return cachedManyOverlayViewportBytes
}

function cachedRealRichViewportWorkbookBytes(): Uint8Array {
	cachedRealRichViewportBytes ??= new Uint8Array(
		readFileSync('fixtures/xlsx/xlsxwriter/styles_formulas.xlsx'),
	)
	return cachedRealRichViewportBytes
}

function realInteractivePatchCorpusTargetBytes(target: RealInteractivePatchTarget): Uint8Array {
	const cached = realInteractivePatchCorpusBytes.get(target.path)
	if (cached) return cached
	const bytes = new Uint8Array(readFileSync(target.path))
	realInteractivePatchCorpusBytes.set(target.path, bytes)
	return bytes
}

function realInteractivePatchCorpusByteCount(): number {
	return REAL_INTERACTIVE_PATCH_CORPUS.reduce(
		(total, target) => total + realInteractivePatchCorpusTargetBytes(target).byteLength,
		0,
	)
}

function xlsxArchiveFootprint(bytes: Uint8Array): {
	readonly partCount: number
	readonly compressedBytes: number
	readonly uncompressedBytes: number
	readonly worksheetUncompressedBytes: number
	readonly largestPartPath: string
	readonly largestPartCompressedBytes: number
	readonly largestPartUncompressedBytes: number
} {
	const archive = extractZip(bytes)
	let partCount = 0
	let compressedBytes = 0
	let uncompressedBytes = 0
	let worksheetUncompressedBytes = 0
	let largestPartPath = ''
	let largestPartCompressedBytes = 0
	let largestPartUncompressedBytes = 0
	for (const entry of archive.entries()) {
		partCount += 1
		compressedBytes += entry.compressedSize
		uncompressedBytes += entry.uncompressedSize
		if (entry.path.startsWith('xl/worksheets/')) {
			worksheetUncompressedBytes += entry.uncompressedSize
		}
		if (entry.uncompressedSize > largestPartUncompressedBytes) {
			largestPartPath = entry.path
			largestPartCompressedBytes = entry.compressedSize
			largestPartUncompressedBytes = entry.uncompressedSize
		}
	}
	return {
		partCount,
		compressedBytes,
		uncompressedBytes,
		worksheetUncompressedBytes,
		largestPartPath,
		largestPartCompressedBytes,
		largestPartUncompressedBytes,
	}
}

interface SharedStringUsageAssertions {
	readonly sharedStringsPartCompressedBytes: number
	readonly sharedStringsPartUncompressedBytes: number
	readonly sharedStringCount: number
	readonly worksheetSharedStringCellRefs: number
	readonly distinctSharedStringRefs: number
	readonly maxSharedStringRef: number
	readonly unusedSharedStringCount: number
	readonly sharedStringReferenceCoverage: number
	readonly sharedStringReferenceFanout: number
}

interface WorksheetScalarUsageAssertions {
	readonly worksheetScalarScanCellLimit: number
	readonly worksheetScalarScanLimitReached: boolean
	readonly worksheetCount: number
	readonly worksheetRows: number
	readonly worksheetCells: number
	readonly worksheetValueCells: number
	readonly worksheetFormulaCells: number
	readonly worksheetNumericCells: number
	readonly worksheetDefaultStyleNumericCells: number
	readonly worksheetDefaultStyleNumericRuns: number
	readonly worksheetDefaultStyleNumericMaxRun: number
	readonly worksheetDefaultStyleNumericAvgRun: number
	readonly worksheetSharedStringCells: number
	readonly worksheetInlineStringCells: number
	readonly worksheetBooleanCells: number
	readonly worksheetStyledCells: number
}

type MutableWorksheetScalarUsageAssertions = {
	-readonly [K in keyof WorksheetScalarUsageAssertions]: WorksheetScalarUsageAssertions[K]
}

function xlsxSharedStringUsage(bytes: Uint8Array): SharedStringUsageAssertions {
	const cached = xlsxSharedStringUsageCache.get(bytes)
	if (cached) return cached
	const archive = extractZip(bytes)
	const sharedStringsEntry = archive.get('xl/sharedStrings.xml')
	const sharedStringsXml = sharedStringsEntry
		? archive.readText(sharedStringsEntry.path)
		: undefined
	const sharedStringCount = sharedStringsXml ? countElementOpens(sharedStringsXml, 'si') : 0
	const referenced = new Set<number>()
	let worksheetSharedStringCellRefs = 0
	for (const entry of archive.entries()) {
		if (!entry.path.startsWith('xl/worksheets/') || !entry.path.endsWith('.xml')) continue
		const xml = archive.readText(entry.path)
		if (!xml) continue
		worksheetSharedStringCellRefs += collectSharedStringRefs(xml, referenced)
	}
	let maxSharedStringRef = -1
	for (const index of referenced) if (index > maxSharedStringRef) maxSharedStringRef = index
	const result = {
		sharedStringsPartCompressedBytes: sharedStringsEntry?.compressedSize ?? 0,
		sharedStringsPartUncompressedBytes: sharedStringsEntry?.uncompressedSize ?? 0,
		sharedStringCount,
		worksheetSharedStringCellRefs,
		distinctSharedStringRefs: referenced.size,
		maxSharedStringRef,
		unusedSharedStringCount: Math.max(0, sharedStringCount - referenced.size),
		sharedStringReferenceCoverage: sharedStringCount > 0 ? referenced.size / sharedStringCount : 0,
		sharedStringReferenceFanout:
			referenced.size > 0 ? worksheetSharedStringCellRefs / referenced.size : 0,
	}
	xlsxSharedStringUsageCache.set(bytes, result)
	return result
}

function xlsxWorksheetScalarUsage(bytes: Uint8Array): WorksheetScalarUsageAssertions {
	const cached = xlsxWorksheetScalarUsageCache.get(bytes)
	if (cached) return cached
	const archive = extractZip(bytes)
	const result = {
		worksheetScalarScanCellLimit: WORKSHEET_SCALAR_USAGE_CELL_SAMPLE_LIMIT,
		worksheetScalarScanLimitReached: false,
		worksheetCount: 0,
		worksheetRows: 0,
		worksheetCells: 0,
		worksheetValueCells: 0,
		worksheetFormulaCells: 0,
		worksheetNumericCells: 0,
		worksheetDefaultStyleNumericCells: 0,
		worksheetDefaultStyleNumericRuns: 0,
		worksheetDefaultStyleNumericMaxRun: 0,
		worksheetDefaultStyleNumericAvgRun: 0,
		worksheetSharedStringCells: 0,
		worksheetInlineStringCells: 0,
		worksheetBooleanCells: 0,
		worksheetStyledCells: 0,
	}
	for (const entry of archive.entries()) {
		if (!entry.path.startsWith('xl/worksheets/') || !entry.path.endsWith('.xml')) continue
		const xml = archive.readText(entry.path)
		if (!xml) continue
		result.worksheetCount += 1
		collectWorksheetScalarUsage(xml, result)
		if (result.worksheetScalarScanLimitReached) break
	}
	result.worksheetDefaultStyleNumericAvgRun =
		result.worksheetDefaultStyleNumericRuns > 0
			? result.worksheetDefaultStyleNumericCells / result.worksheetDefaultStyleNumericRuns
			: 0
	xlsxWorksheetScalarUsageCache.set(bytes, result)
	return result
}

function collectWorksheetScalarUsage(
	xml: string,
	result: MutableWorksheetScalarUsageAssertions,
): void {
	let rowCursor = 0
	while (true) {
		if (result.worksheetCells >= WORKSHEET_SCALAR_USAGE_CELL_SAMPLE_LIMIT) {
			result.worksheetScalarScanLimitReached = true
			return
		}
		const rowOpen = xml.indexOf('<row', rowCursor)
		if (rowOpen === -1) return
		const rowTagEnd = xml.indexOf('>', rowOpen + 4)
		if (rowTagEnd === -1) return
		const rowClose = xml.indexOf('</row>', rowTagEnd + 1)
		if (rowClose === -1) return
		result.worksheetRows += 1
		collectWorksheetRowScalarUsage(xml, rowTagEnd + 1, rowClose, result)
		rowCursor = rowClose + 6
	}
}

function collectWorksheetRowScalarUsage(
	xml: string,
	rowStart: number,
	rowEnd: number,
	result: MutableWorksheetScalarUsageAssertions,
): void {
	let cursor = rowStart
	let nextCol = 0
	let numericRunStartCol = -1
	let numericRunLength = 0
	const flushNumericRun = () => {
		if (numericRunLength <= 0) return
		result.worksheetDefaultStyleNumericRuns += 1
		result.worksheetDefaultStyleNumericMaxRun = Math.max(
			result.worksheetDefaultStyleNumericMaxRun,
			numericRunLength,
		)
		numericRunStartCol = -1
		numericRunLength = 0
	}
	while (true) {
		if (result.worksheetCells >= WORKSHEET_SCALAR_USAGE_CELL_SAMPLE_LIMIT) {
			flushNumericRun()
			result.worksheetScalarScanLimitReached = true
			return
		}
		const cellOpen = xml.indexOf('<c', cursor)
		if (cellOpen === -1 || cellOpen >= rowEnd) {
			flushNumericRun()
			return
		}
		const next = xml.charCodeAt(cellOpen + 2)
		if (next !== 32 && next !== 9 && next !== 10 && next !== 13 && next !== 62) {
			cursor = cellOpen + 2
			continue
		}
		const cellTagEnd = xml.indexOf('>', cellOpen + 2)
		if (cellTagEnd === -1 || cellTagEnd >= rowEnd) {
			flushNumericRun()
			return
		}
		const selfClosing = xml.charCodeAt(cellTagEnd - 1) === 47
		const cellClose = selfClosing ? cellTagEnd + 1 : xml.indexOf('</c>', cellTagEnd + 1)
		if (cellClose === -1 || cellClose > rowEnd) {
			flushNumericRun()
			return
		}
		const attrs = xml.slice(cellOpen + 2, cellTagEnd)
		const type = readXmlAttribute(attrs, 't')
		const styleIdx = parseNonNegativeIntegerAttribute(attrs, 's')
		const styleIsDefault = styleIdx === undefined || styleIdx === 0
		const cellCol = parseCellRefCol(attrs) ?? nextCol
		const bodyStart = cellTagEnd + 1
		const bodyEnd = selfClosing ? cellTagEnd : cellClose
		const hasValue = xml.indexOf('<v>', bodyStart) !== -1 && xml.indexOf('<v>', bodyStart) < bodyEnd
		const hasFormula = xml.indexOf('<f', bodyStart) !== -1 && xml.indexOf('<f', bodyStart) < bodyEnd
		result.worksheetCells += 1
		if (hasValue) result.worksheetValueCells += 1
		if (hasFormula) result.worksheetFormulaCells += 1
		if (!styleIsDefault) result.worksheetStyledCells += 1
		if (type === 's') result.worksheetSharedStringCells += 1
		else if (type === 'inlineStr') result.worksheetInlineStringCells += 1
		else if (type === 'b') result.worksheetBooleanCells += 1
		const isNumeric = !hasFormula && hasValue && (type === undefined || type === 'n')
		if (isNumeric) {
			result.worksheetNumericCells += 1
			if (styleIsDefault) {
				result.worksheetDefaultStyleNumericCells += 1
				if (numericRunLength === 0) {
					numericRunStartCol = cellCol
					numericRunLength = 1
				} else if (cellCol === numericRunStartCol + numericRunLength) {
					numericRunLength += 1
				} else {
					flushNumericRun()
					numericRunStartCol = cellCol
					numericRunLength = 1
				}
			} else {
				flushNumericRun()
			}
		} else {
			flushNumericRun()
		}
		nextCol = cellCol + 1
		cursor = selfClosing ? cellTagEnd + 1 : cellClose + 4
	}
}

function countElementOpens(xml: string, tagName: string): number {
	let count = 0
	let cursor = 0
	while (true) {
		const open = xml.indexOf(`<${tagName}`, cursor)
		if (open === -1) return count
		const next = xml.charCodeAt(open + tagName.length + 1)
		if (next === 47 || next === 62 || next === 32 || next === 9 || next === 10 || next === 13) {
			count += 1
		}
		cursor = open + tagName.length + 1
	}
}

function collectSharedStringRefs(xml: string, referenced: Set<number>): number {
	let refs = 0
	let cursor = 0
	while (true) {
		const cellOpen = xml.indexOf('<c', cursor)
		if (cellOpen === -1) return refs
		const cellTagEnd = xml.indexOf('>', cellOpen + 2)
		if (cellTagEnd === -1) return refs
		const attrs = xml.slice(cellOpen + 2, cellTagEnd)
		if (!hasSharedStringCellType(attrs)) {
			cursor = cellTagEnd + 1
			continue
		}
		const cellClose = xml.indexOf('</c>', cellTagEnd + 1)
		if (cellClose === -1) return refs
		const valueOpen = xml.indexOf('<v>', cellTagEnd + 1)
		if (valueOpen !== -1 && valueOpen < cellClose) {
			const valueClose = xml.indexOf('</v>', valueOpen + 3)
			if (valueClose !== -1 && valueClose <= cellClose) {
				const index = parseNonNegativeIntText(xml, valueOpen + 3, valueClose)
				if (index !== undefined) {
					referenced.add(index)
					refs += 1
				}
			}
		}
		cursor = cellClose + 4
	}
}

function hasSharedStringCellType(attrs: string): boolean {
	return /\bt\s*=\s*(?:"s"|'s')/.test(attrs)
}

function readXmlAttribute(attrs: string, name: string): string | undefined {
	const pattern = new RegExp(`(?:^|\\s)${name}\\s*=\\s*("[^"]*"|'[^']*')`)
	const match = pattern.exec(attrs)
	if (!match?.[1]) return undefined
	return match[1].slice(1, -1)
}

function parseNonNegativeIntegerAttribute(attrs: string, name: string): number | undefined {
	const raw = readXmlAttribute(attrs, name)
	if (raw === undefined) return undefined
	return parseNonNegativeIntText(raw, 0, raw.length)
}

function parseCellRefCol(attrs: string): number | undefined {
	const ref = readXmlAttribute(attrs, 'r')
	if (!ref) return undefined
	let col = 0
	let index = 0
	while (index < ref.length) {
		const code = ref.charCodeAt(index)
		if (code >= 65 && code <= 90) col = col * 26 + (code - 64)
		else if (code >= 97 && code <= 122) col = col * 26 + (code - 96)
		else break
		index++
	}
	return col > 0 ? col - 1 : undefined
}

function parseNonNegativeIntText(xml: string, start: number, end: number): number | undefined {
	if (start >= end) return undefined
	let value = 0
	for (let index = start; index < end; index++) {
		const code = xml.charCodeAt(index)
		if (code < 48 || code > 57) return undefined
		value = value * 10 + (code - 48)
	}
	return value
}

function workbookGridStorageAssertions(workbook: Workbook): Record<string, number> {
	let gridCellCount = 0
	let gridChunkCount = 0
	let gridDenseChunkCount = 0
	let gridSparseChunkCount = 0
	let gridDenseCellCount = 0
	let gridSparseCellCount = 0
	let gridDenseCapacity = 0
	let gridSparseCapacity = 0
	let gridDenseArrayBufferBytes = 0
	let gridSparseArrayBufferBytes = 0
	let gridStyleArrayBufferBytes = 0
	let gridTotalArrayBufferBytes = 0
	for (const sheet of workbook.sheets) {
		const stats = sheet.cells.storageStats()
		gridCellCount += stats.cellCount
		gridChunkCount += stats.chunkCount
		gridDenseChunkCount += stats.denseChunkCount
		gridSparseChunkCount += stats.sparseChunkCount
		gridDenseCellCount += stats.denseCellCount
		gridSparseCellCount += stats.sparseCellCount
		gridDenseCapacity += stats.denseCapacity
		gridSparseCapacity += stats.sparseCapacity
		gridDenseArrayBufferBytes += stats.denseArrayBufferBytes
		gridSparseArrayBufferBytes += stats.sparseArrayBufferBytes
		gridStyleArrayBufferBytes += stats.styleArrayBufferBytes
		gridTotalArrayBufferBytes += stats.totalArrayBufferBytes
	}
	return {
		gridCellCount,
		gridChunkCount,
		gridDenseChunkCount,
		gridSparseChunkCount,
		gridDenseCellCount,
		gridSparseCellCount,
		gridDenseCapacity,
		gridSparseCapacity,
		gridDenseArrayBufferBytes,
		gridSparseArrayBufferBytes,
		gridStyleArrayBufferBytes,
		gridTotalArrayBufferBytes,
		gridArrayBufferBytesPerCell: gridCellCount > 0 ? gridTotalArrayBufferBytes / gridCellCount : 0,
	}
}

function workbookModelRetentionAssertions(workbook: Workbook): Record<string, number> {
	return {
		sheetCount: workbook.sheets.length,
		sourceArchiveBytesRetained: workbook.sourceArchiveBytes?.byteLength ?? 0,
		...workbookGridShapeAssertions(workbook),
		...workbookGridStorageAssertions(workbook),
	}
}

function workbookDocumentModel(document: WorkbookDocument): Workbook | null {
	const internals = document as unknown as {
		readonly view?: { getWorkbookModel(): Workbook }
	}
	return internals.view?.getWorkbookModel() ?? null
}

function sessionMutableWorkbook(session: AscendSession): AscendWorkbook | null {
	const internals = session as unknown as { readonly mutableWorkbook?: AscendWorkbook | null }
	return internals.mutableWorkbook ?? null
}

function ascendWorkbookOriginalBytes(workbook: AscendWorkbook): Uint8Array | null {
	const internals = workbook as unknown as { readonly originalBytes?: Uint8Array | null }
	return internals.originalBytes ?? null
}

function ascendWorkbookRetentionAssertions(workbook: AscendWorkbook): Record<string, number> {
	const internals = workbook as unknown as {
		readonly caps?: readonly { readonly content?: Uint8Array }[]
		readonly originalBytes?: Uint8Array | null
		readonly sourceArchive?: unknown
	}
	const capsules = internals.caps ?? []
	let capsuleContentBytes = 0
	for (const capsule of capsules) capsuleContentBytes += capsule.content?.byteLength ?? 0
	return {
		originalBytesRetained: internals.originalBytes?.byteLength ?? 0,
		capsuleCount: capsules.length,
		capsuleContentBytes,
		sourceArchiveCached: internals.sourceArchive ? 1 : 0,
		...workbookModelRetentionAssertions(workbook.getWorkbookModel()),
	}
}

function workbookGridSharingAssertions(
	readWorkbook: Workbook,
	mutableWorkbook: Workbook,
): Record<string, number> {
	let comparableSheetCount = 0
	let sharedGridMaps = 0
	let readSharedGrids = 0
	let mutableSharedGrids = 0
	let readCopyOnWriteChunkKeys = 0
	let mutableCopyOnWriteChunkKeys = 0
	for (const readSheet of readWorkbook.sheets) {
		const mutableSheet = mutableWorkbook.getSheet(readSheet.name)
		if (!mutableSheet) continue
		comparableSheetCount++
		const readCells = readSheet.cells as unknown as {
			readonly chunkRows?: unknown
			readonly _shared?: boolean
			readonly _sharedChunks?: ReadonlySet<number> | null
		}
		const mutableCells = mutableSheet.cells as unknown as {
			readonly chunkRows?: unknown
			readonly _shared?: boolean
			readonly _sharedChunks?: ReadonlySet<number> | null
		}
		if (readCells.chunkRows === mutableCells.chunkRows) sharedGridMaps++
		if (readCells._shared) readSharedGrids++
		if (mutableCells._shared) mutableSharedGrids++
		readCopyOnWriteChunkKeys += readCells._sharedChunks?.size ?? 0
		mutableCopyOnWriteChunkKeys += mutableCells._sharedChunks?.size ?? 0
	}
	return {
		comparableSheetCount,
		sharedGridMaps,
		readSharedGrids,
		mutableSharedGrids,
		readCopyOnWriteChunkKeys,
		mutableCopyOnWriteChunkKeys,
		allComparableGridsShareMaps:
			comparableSheetCount > 0 && sharedGridMaps === comparableSheetCount ? 1 : 0,
	}
}

function workbookGridShapeAssertions(workbook: Workbook): Record<string, number> {
	let nonEmptySheets = 0
	let maxUsedRows = 0
	let maxUsedCols = 0
	let formulaCells = 0
	let stringCells = 0
	let richTextCells = 0
	let arrayCells = 0
	for (const sheet of workbook.sheets) {
		const used = sheet.cells.usedRange()
		formulaCells += sheet.cells.formulaCellCount()
		stringCells += sheet.cells.stringCellCount()
		richTextCells += sheet.cells.richTextCellCount()
		arrayCells += sheet.cells.arrayCellCount()
		if (!used) continue
		nonEmptySheets += 1
		maxUsedRows = Math.max(maxUsedRows, used.end.row - used.start.row + 1)
		maxUsedCols = Math.max(maxUsedCols, used.end.col - used.start.col + 1)
	}
	return {
		nonEmptySheets,
		maxUsedRows,
		maxUsedCols,
		formulaCells,
		stringCells,
		richTextCells,
		arrayCells,
	}
}

function buildGridStorageWidthAssertions(rows: number, width: number): Record<string, number> {
	runGc()
	const baseline = phaseMemorySnapshot()
	const buildStart = performance.now()
	const workbook = buildPlainNumericWorkbook(rows, width)
	const buildMs = performance.now() - buildStart
	runGc()
	const afterBuild = phaseMemorySnapshot()
	return {
		rows,
		width,
		cells: rows * width,
		buildMs,
		heapDeltaBytes: memoryDelta(afterBuild.heapUsedBytes, baseline.heapUsedBytes),
		externalDeltaBytes: memoryDelta(afterBuild.externalBytes, baseline.externalBytes),
		arrayBuffersDeltaBytes: memoryDelta(afterBuild.arrayBuffersBytes, baseline.arrayBuffersBytes),
		...workbookGridStorageAssertions(workbook),
	}
}

function buildMixedGridStorageAssertions(rows: number, width: number): Record<string, number> {
	runGc()
	const baseline = phaseMemorySnapshot()
	const buildStart = performance.now()
	const workbook = buildMixedPlainWorkbook(rows, width)
	const buildMs = performance.now() - buildStart
	runGc()
	const afterBuild = phaseMemorySnapshot()
	return {
		rows,
		width,
		cells: rows * width,
		buildMs,
		heapDeltaBytes: memoryDelta(afterBuild.heapUsedBytes, baseline.heapUsedBytes),
		externalDeltaBytes: memoryDelta(afterBuild.externalBytes, baseline.externalBytes),
		arrayBuffersDeltaBytes: memoryDelta(afterBuild.arrayBuffersBytes, baseline.arrayBuffersBytes),
		...workbookGridStorageAssertions(workbook),
	}
}

function readXlsxGridStorageAssertions(bytes: Uint8Array): Record<string, number | boolean> {
	const readStart = performance.now()
	const result = readXlsx(bytes, { mode: 'full', richMetadata: true })
	const readMs = performance.now() - readStart
	if (!result.ok) throw new Error(`XLSX grid storage read failed: ${result.error.message}`)
	return {
		bytes: bytes.byteLength,
		readMs,
		loadIsPartial: result.value.loadInfo.isPartial,
		...workbookGridShapeAssertions(result.value.workbook),
		...workbookGridStorageAssertions(result.value.workbook),
	}
}

function prefixAssertions(
	prefix: string,
	assertions: Record<string, number>,
): Record<string, number> {
	const prefixed: Record<string, number> = {}
	for (const [key, value] of Object.entries(assertions)) {
		prefixed[`${prefix}.${key}`] = value
	}
	return prefixed
}

function interactiveSessionFor(bytes: Uint8Array): Promise<AscendSession> {
	const cached = interactiveSessionCache.get(bytes)
	if (cached) return cached
	const session = AscendSession.open(bytes, { mode: 'interactive' })
	interactiveSessionCache.set(bytes, session)
	return session
}

async function patchStreamStateFor(bytes: Uint8Array): Promise<{
	readonly session: AscendSession
	lastToken: string
	iteration: number
}> {
	const cached = patchStreamStateCache.get(bytes)
	if (cached) return cached
	const state = AscendSession.open(bytes, { mode: 'interactive', prepareEdits: true }).then(
		(session) => {
			const viewport = session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 250,
				colCount: 20,
			})
			return { session, lastToken: viewport.changeToken, iteration: 0 }
		},
	)
	patchStreamStateCache.set(bytes, state)
	return state
}

type AscendSessionOpenOptions = NonNullable<Parameters<typeof AscendSession.open>[1]>

async function measureInteractiveOpenPath(
	label: string,
	bytes: Uint8Array,
	target: RealInteractivePatchTarget,
	options: AscendSessionOpenOptions,
): Promise<Record<string, string | number | boolean | null>> {
	WorkbookDocument.clearCache()
	const openStart = performance.now()
	const session = await AscendSession.open(bytes, options)
	const openMs = performance.now() - openStart
	const load = session.inspect().load
	const viewportStart = performance.now()
	const viewport = session.readViewport({
		sheet: target.sheet,
		topRow: target.topRow,
		leftCol: target.leftCol,
		rowCount: target.rowCount,
		colCount: target.colCount,
	})
	const viewportMs = performance.now() - viewportStart
	const prepare = await session.prepareEdits()
	const readiness = session.editReadiness()
	session.close()
	WorkbookDocument.clearCache()
	return {
		[`${label}OpenMs`]: openMs,
		[`${label}LoadMode`]: load.mode,
		[`${label}LoadIsPartial`]: load.isPartial,
		[`${label}PartialReasons`]: load.partialReasons.length,
		[`${label}ViewportMs`]: viewportMs,
		[`${label}ViewportCells`]: viewport.cells.length,
		[`${label}ViewportFlatValues`]: viewport.flatValues.length,
		[`${label}PrepareEditsMs`]: prepare.timings.totalMs,
		[`${label}PrepareEnsureMutableWorkbookMs`]: prepare.timings.ensureMutableWorkbookMs,
		[`${label}PrepareReusedReadModel`]: prepare.timings.mutableWorkbookReusedReadModel ?? false,
		[`${label}PrepareMutableWorkbookOpenMs`]: prepare.timings.mutableWorkbookOpenMs ?? null,
		[`${label}PrepareCached`]: prepare.timings.mutableWorkbookCached ?? false,
		[`${label}ReadyAfterPrepare`]: readiness.ready,
		[`${label}TimeToFirstViewportMs`]: openMs + viewportMs,
		[`${label}TimeToEditReadyMs`]: openMs + viewportMs + prepare.timings.totalMs,
	}
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

function createRealDenseReadMemoryScenario(
	name: string,
	readOptions: NonNullable<Parameters<typeof readXlsx>[1]>,
): Scenario {
	return {
		name,
		category: 'read',
		build() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			return {
				rows: target.rowCount,
				cols: target.colCount,
				cells: target.rowCount * target.colCount,
				byteCount: realInteractivePatchCorpusTargetBytes(target).byteLength,
			}
		},
		run() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			const archiveFootprint = xlsxArchiveFootprint(bytes)
			WorkbookDocument.clearCache()
			runGc()
			const baseline = phaseMemorySnapshot()

			const readStart = performance.now()
			const result = readXlsx(bytes, readOptions)
			const readMs = performance.now() - readStart
			if (!result.ok) throw new Error(`Dense read failed: ${result.error.message}`)
			runGc()
			const afterRead = phaseMemorySnapshot()

			const workbook = result.value.workbook
			const loadedCells = workbook.sheets.reduce(
				(total, sheet) => total + sheet.cells.cellCount(),
				0,
			)
			const retainedArrayBuffers = memoryDelta(
				afterRead.arrayBuffersBytes,
				baseline.arrayBuffersBytes,
			)

			return {
				assertions: {
					mode: readOptions.mode ?? 'full',
					richMetadata: readOptions.richMetadata ?? null,
					bytes: bytes.byteLength,
					archivePartCount: archiveFootprint.partCount,
					archiveCompressedBytes: archiveFootprint.compressedBytes,
					archiveUncompressedBytes: archiveFootprint.uncompressedBytes,
					archiveWorksheetUncompressedBytes: archiveFootprint.worksheetUncompressedBytes,
					archiveLargestPartPath: archiveFootprint.largestPartPath,
					archiveLargestPartCompressedBytes: archiveFootprint.largestPartCompressedBytes,
					archiveLargestPartUncompressedBytes: archiveFootprint.largestPartUncompressedBytes,
					readMs,
					loadMode: result.value.loadInfo.mode,
					loadIsPartial: result.value.loadInfo.isPartial,
					cellsHydrated: result.value.loadInfo.cellsHydrated,
					richSheetMetadataHydrated: result.value.loadInfo.richSheetMetadataHydrated,
					sheetCount: workbook.sheets.length,
					loadedCells,
					...workbookGridStorageAssertions(workbook),
					capsuleCount: result.value.capsules.length,
					activeContentCount: workbook.activeContent.length,
					sourceArchiveBytesRetained: workbook.sourceArchiveBytes?.byteLength ?? 0,
					baselineRssBytes: baseline.rssBytes,
					afterReadRssDeltaBytes: memoryDelta(afterRead.rssBytes, baseline.rssBytes),
					baselineHeapUsedBytes: baseline.heapUsedBytes,
					afterReadHeapDeltaBytes: memoryDelta(afterRead.heapUsedBytes, baseline.heapUsedBytes),
					baselineExternalBytes: baseline.externalBytes,
					afterReadExternalDeltaBytes: memoryDelta(afterRead.externalBytes, baseline.externalBytes),
					baselineArrayBuffersBytes: baseline.arrayBuffersBytes,
					afterReadArrayBuffersDeltaBytes: retainedArrayBuffers,
					arrayBuffersPerLoadedCell: loadedCells > 0 ? retainedArrayBuffers / loadedCells : null,
					arrayBuffersToArchiveUncompressedRatio:
						archiveFootprint.uncompressedBytes > 0
							? retainedArrayBuffers / archiveFootprint.uncompressedBytes
							: null,
				},
			}
		},
	}
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
		name: 'write-xlsx-100k-rows',
		category: 'write',
		build() {
			const workbook = buildDenseWorkbook(100_000, 10)
			return { workbook, rows: 100_000, cols: 10, cells: 1_000_000 }
		},
		run(input) {
			mustWrite(requireWorkbook(input))
		},
	},
	{
		name: 'write-xlsx-streaming-100k-rows',
		category: 'write',
		build() {
			const workbook = buildDenseWorkbook(100_000, 10)
			return { workbook, rows: 100_000, cols: 10, cells: 1_000_000 }
		},
		async run(input) {
			const result = await writeXlsxStreaming(requireWorkbook(input), undefined, {
				streaming: true,
			})
			if (!result.ok) throw new Error(result.error.message)
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
		name: 'grid-storage-width-sweep',
		category: 'workflow',
		build() {
			const rows = 20_000
			const maxCols = 50
			return { rows, cols: maxCols, cells: rows * maxCols }
		},
		run(input) {
			const rows = input.rows
			const widths = [5, 10, 20, 50] as const
			const assertions: Record<string, number> = {}
			for (const width of widths) {
				Object.assign(
					assertions,
					prefixAssertions(`width${width}`, buildGridStorageWidthAssertions(rows, width)),
				)
			}
			return { assertions }
		},
	},
	...([5, 10, 20, 50] as const).map(
		(width): Scenario => ({
			name: `grid-storage-width-${width}`,
			category: 'workflow',
			build() {
				const rows = 20_000
				return { rows, cols: width, cells: rows * width }
			},
			run(input) {
				return { assertions: buildGridStorageWidthAssertions(input.rows, width) }
			},
		}),
	),
	...([10, 20] as const).map(
		(width): Scenario => ({
			name: `grid-storage-mixed-width-${width}`,
			category: 'workflow',
			build() {
				const rows = 20_000
				return { rows, cols: width, cells: rows * width }
			},
			run(input) {
				return { assertions: buildMixedGridStorageAssertions(input.rows, width) }
			},
		}),
	),
	...([5, 10, 20] as const).map(
		(width): Scenario => ({
			name: `xlsx-grid-storage-width-${width}`,
			category: 'read',
			build() {
				const rows = 20_000
				const bytes = mustWrite(buildPlainNumericWorkbook(rows, width))
				return { bytes, rows, cols: width, cells: rows * width, byteCount: bytes.byteLength }
			},
			run(input) {
				return { assertions: readXlsxGridStorageAssertions(requireBytes(input)) }
			},
		}),
	),
	{
		name: 'write-dirty-single-sheet',
		category: 'write',
		build() {
			const workbook = buildMultiSheetWorkbook(3, 5000, 10)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 5000, cols: 10, cells: 150_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
			const sheet = result.value.workbook.sheets[0]
			if (!sheet) throw new Error('Missing Sheet1')
			sheet.cells.set(0, 0, { value: numberValue(999), formula: null, styleId: SID })
			const written = writeXlsx(result.value.workbook, result.value.capsules, {
				dirtySheetNames: ['Sheet1'],
			})
			if (!written.ok) throw new Error(written.error.message)
		},
	},
	{
		name: 'style-heavy-workbook',
		category: 'write',
		build() {
			const workbook = buildStyleHeavyWorkbook(1000)
			return { workbook, rows: 10, cols: 100, cells: 1000 }
		},
		run(input) {
			mustWrite(requireWorkbook(input))
		},
	},
	{
		name: 'write-unique-strings-sst',
		category: 'write',
		build() {
			const workbook = buildUniqueStringWorkbook(2000, 20)
			return { workbook, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			mustWrite(requireWorkbook(input), { useSharedStrings: true })
		},
	},
	{
		name: 'write-unique-strings-inline',
		category: 'write',
		build() {
			const workbook = buildUniqueStringWorkbook(2000, 20)
			return { workbook, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			mustWrite(requireWorkbook(input), { useInlineStrings: true })
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
			const rows = 200_000
			const cols = 5
			const workbook = buildMixedDataWorkbook(rows, cols)
			const bytes = mustWrite(workbook)
			return { bytes, rows, cols, cells: rows * cols }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-style-heavy',
		category: 'read',
		build() {
			const rows = 5_000
			const uniqueStyles = 2_500
			const workbook = buildReadStyleHeavyWorkbook(rows, uniqueStyles)
			const bytes = mustWrite(workbook)
			return { bytes, rows, cols: 1, cells: rows }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-string-heavy',
		category: 'read',
		build() {
			const rows = 20_000
			const cols = 2
			const uniqueStrings = 10_500
			const workbook = buildReadStringHeavyWorkbook(rows, uniqueStrings)
			const bytes = mustWrite(workbook)
			return { bytes, rows, cols, cells: rows * cols }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-formula-dense',
		category: 'read',
		build() {
			const rows = 10_000
			const cols = 5
			const workbook = buildReadFormulaDenseWorkbook(rows, cols)
			const bytes = mustWrite(workbook)
			return { bytes, rows, cols, cells: rows * cols }
		},
		run(input) {
			const result = readXlsx(requireBytes(input))
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-large-200k-values',
		category: 'read',
		build() {
			const rows = 200_000
			const cols = 5
			const workbook = buildMixedDataWorkbook(rows, cols)
			const bytes = mustWrite(workbook)
			return { bytes, rows, cols, cells: rows * cols }
		},
		run(input) {
			const result = readXlsx(requireBytes(input), { mode: 'values' })
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
		name: 'read-values-dense',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(2000, 20)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 2000, cols: 20, cells: 40_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input), { mode: 'values' })
			if (!result.ok) throw new Error(result.error.message)
		},
	},
	{
		name: 'read-values-wide',
		category: 'read',
		build() {
			const workbook = buildDenseWorkbook(2000, 100)
			const bytes = mustWrite(workbook)
			return { bytes, rows: 2000, cols: 100, cells: 200_000 }
		},
		run(input) {
			const result = readXlsx(requireBytes(input), { mode: 'values' })
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
		name: 'prepared-open-string-heavy',
		category: 'workflow',
		build() {
			const rows = 20_000
			const cols = 2
			const uniqueStrings = 10_500
			const workbook = buildReadStringHeavyWorkbook(rows, uniqueStrings)
			const bytes = mustWrite(workbook)
			return { bytes, rows, cols, cells: rows * cols }
		},
		async run(input) {
			const bytes = requireBytes(input)
			WorkbookDocument.clearCache()

			const openStart = performance.now()
			const session = await AscendSession.open(bytes, {
				mode: 'interactive',
				prepareEdits: true,
			})
			const preparedOpenMs = performance.now() - openStart
			const readiness = session.editReadiness()
			if (!readiness.ready) throw new Error('String-heavy prepared-open session was not edit-ready')
			const viewportStart = performance.now()
			const viewport = session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 50,
				colCount: 2,
			})
			const viewportMs = performance.now() - viewportStart
			const firstCell = viewport.cells.find((cell) => cell.ref === 'A1')
			if (firstCell?.flatValue !== 'unique-string-0-x') {
				throw new Error('String-heavy prepared-open viewport returned unexpected data')
			}
			session.close()
			WorkbookDocument.clearCache()

			return {
				assertions: {
					bytes: bytes.byteLength,
					preparedOpenMs,
					viewportMs,
					viewportCells: viewport.cells.length,
					viewportFlatValues: viewport.flatValues.length,
					readinessReady: readiness.ready,
					readinessReusedReadModel: readiness.timings?.mutableWorkbookReusedReadModel ?? false,
					readinessMutableWorkbookOpenMs: readiness.timings?.mutableWorkbookOpenMs ?? null,
					readinessRebaseViewportSnapshotsMs: readiness.timings?.rebaseViewportSnapshotsMs ?? null,
				},
			}
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
		name: 'sdk-window-dense-values-compact-hot',
		category: 'read',
		build() {
			return { workbook: buildDenseWorkbook(5000, 20), rows: 250, cols: 20, cells: 5000 }
		},
		run(input) {
			const workbook = requireWorkbook(input)
			const handle = makeBenchmarkSheetHandle(workbook)
			const window = handle.readWindowCompact('A1:T5000', {
				rowLimit: 250,
				includeRefs: false,
			})
			const first = window.cells[0]
			const last = window.cells[window.cells.length - 1]
			if (
				!first ||
				!last ||
				window.cells.length !== 5000 ||
				!window.hasMore ||
				window.nextRowOffset !== 250 ||
				first?.row !== 0 ||
				first.col !== 0 ||
				first.ref !== undefined ||
				first.value.kind !== 'number' ||
				first.value.value !== 1 ||
				last?.row !== 249 ||
				last.col !== 19 ||
				last.ref !== undefined ||
				last.value.kind !== 'number' ||
				last.value.value !== 5000
			) {
				throw new Error('Dense compact-window benchmark returned an unexpected payload')
			}
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
					nextRowOffset: window.nextRowOffset ?? null,
					firstValue: first.value.kind === 'number' ? first.value.value : null,
					lastRow: last.row,
					lastCol: last.col,
					lastValue: last.value.kind === 'number' ? last.value.value : null,
				},
			}
		},
	},
	{
		name: 'sdk-window-formula-chain-compact-hot',
		category: 'read',
		build() {
			return { workbook: buildFormulaChainWorkbook(6000), rows: 500, cols: 1, cells: 500 }
		},
		run(input) {
			const workbook = requireWorkbook(input)
			const handle = makeBenchmarkSheetHandle(workbook)
			const window = handle.readWindowCompact('A1:A6000', {
				rowLimit: 500,
				includeRefs: false,
			})
			const first = window.cells[0]
			const second = window.cells[1]
			const last = window.cells[window.cells.length - 1]
			const formulaCount = window.cells.filter((cell) => cell.formula !== null).length
			if (
				!first ||
				!second ||
				!last ||
				window.cells.length !== 500 ||
				!window.hasMore ||
				window.nextRowOffset !== 500 ||
				formulaCount !== 499 ||
				first?.row !== 0 ||
				first.col !== 0 ||
				first.ref !== undefined ||
				first.value.kind !== 'number' ||
				first.value.value !== 1 ||
				first.formula !== null ||
				second?.row !== 1 ||
				second.formula !== 'A1+1' ||
				last?.row !== 499 ||
				last.col !== 0 ||
				last.ref !== undefined ||
				last.formula !== 'A499+1'
			) {
				throw new Error('Formula-chain compact-window benchmark returned an unexpected payload')
			}
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
					nextRowOffset: window.nextRowOffset ?? null,
					formulaCount,
					firstValue: first.value.kind === 'number' ? first.value.value : null,
					secondFormula: second.formula,
					lastFormula: last.formula,
				},
			}
		},
	},
	{
		name: 'sdk-window-sparse-wide-compact-hot',
		category: 'read',
		build() {
			return { workbook: buildSparseWorkbook(100_000, 20, 500), rows: 5000, cols: 20, cells: 200 }
		},
		run(input) {
			const workbook = requireWorkbook(input)
			const handle = makeBenchmarkSheetHandle(workbook)
			const window = handle.readWindowCompact('A1:T100000', {
				rowLimit: 5000,
				includeRefs: false,
			})
			const first = window.cells[0]
			const last = window.cells[window.cells.length - 1]
			if (
				!first ||
				!last ||
				window.cells.length !== 200 ||
				!window.hasMore ||
				window.nextRowOffset !== 5000 ||
				first?.row !== 0 ||
				first.col !== 0 ||
				first.ref !== undefined ||
				first.value.kind !== 'number' ||
				first.value.value !== 1 ||
				last?.row !== 4500 ||
				last.col !== 19 ||
				last.ref !== undefined ||
				last.value.kind !== 'number' ||
				last.value.value !== 4520
			) {
				throw new Error('Sparse compact-window benchmark returned an unexpected payload')
			}
			return {
				assertions: {
					returnedCells: window.cells.length,
					hasMore: window.hasMore,
					nextRowOffset: window.nextRowOffset ?? null,
					firstValue: first.value.kind === 'number' ? first.value.value : null,
					lastRow: last.row,
					lastCol: last.col,
					lastValue: last.value.kind === 'number' ? last.value.value : null,
				},
			}
		},
	},
	{
		name: 'sdk-viewport-read-dense-hot',
		category: 'read',
		build() {
			return { bytes: cachedDenseViewportWorkbookBytes(), rows: 250, cols: 20, cells: 5000 }
		},
		async run(input) {
			const session = await interactiveSessionFor(requireBytes(input))
			const viewport = session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 250,
				colCount: 20,
			})
			const first = viewport.cells[0]
			const last = viewport.cells[viewport.cells.length - 1]
			if (
				!first ||
				!last ||
				viewport.cells.length !== 5000 ||
				viewport.flatValues.length !== 5000 ||
				first.flatValue !== 1 ||
				last.flatValue !== 5000
			) {
				throw new Error('Dense interactive viewport benchmark returned an unexpected payload')
			}
			return {
				assertions: {
					returnedCells: viewport.cells.length,
					flatValues: viewport.flatValues.length,
					firstValue: first.flatValue,
					lastValue: last.flatValue,
				},
			}
		},
	},
	{
		name: 'sdk-viewport-read-sparse-wide-hot',
		category: 'read',
		build() {
			return {
				bytes: cachedSparseWideViewportWorkbookBytes(),
				rows: 5000,
				cols: 20,
				cells: 100_000,
			}
		},
		async run(input) {
			const session = await interactiveSessionFor(requireBytes(input))
			const viewport = session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 5000,
				colCount: 20,
			})
			const first = viewport.cells[0]
			const last = viewport.cells[viewport.cells.length - 1]
			if (
				!first ||
				!last ||
				viewport.cells.length !== 200 ||
				viewport.flatValues.length !== 100_000 ||
				first.flatValue !== 1 ||
				last.flatValue !== 4520
			) {
				throw new Error('Sparse interactive viewport benchmark returned an unexpected payload')
			}
			return {
				assertions: {
					returnedCells: viewport.cells.length,
					flatValues: viewport.flatValues.length,
					firstValue: first.flatValue,
					lastValue: last.flatValue,
				},
			}
		},
	},
	{
		name: 'sdk-semantic-viewport-rich',
		category: 'read',
		build() {
			return { bytes: cachedSemanticViewportWorkbookBytes(), rows: 50, cols: 10, cells: 500 }
		},
		async run(input) {
			const session = await interactiveSessionFor(requireBytes(input))
			const viewport = session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 50,
				colCount: 10,
			})
			if (
				viewport.cells.length < 150 ||
				viewport.merges.length !== 1 ||
				viewport.comments.length !== 1 ||
				viewport.hyperlinks.length !== 1 ||
				viewport.dataValidations.length !== 1 ||
				viewport.conditionalFormats.length !== 1 ||
				viewport.tables.length !== 1 ||
				viewport.autoFilter?.ref !== 'A1:C101'
			) {
				throw new Error('Semantic interactive viewport benchmark returned an unexpected payload')
			}
			return {
				assertions: {
					returnedCells: viewport.cells.length,
					flatValues: viewport.flatValues.length,
					merges: viewport.merges.length,
					comments: viewport.comments.length,
					hyperlinks: viewport.hyperlinks.length,
					validations: viewport.dataValidations.length,
					conditionalFormats: viewport.conditionalFormats.length,
					tables: viewport.tables.length,
				},
			}
		},
	},
	{
		name: 'sdk-semantic-viewport-many-overlays',
		category: 'read',
		build() {
			return { bytes: cachedManyOverlayViewportWorkbookBytes(), rows: 250, cols: 20, cells: 5000 }
		},
		async run(input) {
			const session = await interactiveSessionFor(requireBytes(input))
			const viewport = session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 250,
				colCount: 20,
			})
			if (
				viewport.cells.length !== 5000 ||
				viewport.merges.length !== 250 ||
				viewport.dataValidations.length !== 250 ||
				viewport.conditionalFormats.length !== 250 ||
				viewport.cells.some(
					(cell) =>
						!cell.flags.validation ||
						!cell.flags.conditionalFormat ||
						(cell.col >= 18 && !cell.flags.merged),
				)
			) {
				throw new Error(
					'Many-overlay interactive viewport benchmark returned an unexpected payload',
				)
			}
			return {
				assertions: {
					returnedCells: viewport.cells.length,
					flatValues: viewport.flatValues.length,
					merges: viewport.merges.length,
					validations: viewport.dataValidations.length,
					conditionalFormats: viewport.conditionalFormats.length,
				},
			}
		},
	},
	{
		name: 'patch-stream-small-edit',
		category: 'workflow',
		build() {
			return { bytes: cachedDenseViewportWorkbookBytes(), rows: 100, cols: 1, cells: 100 }
		},
		async run(input) {
			const state = await patchStreamStateFor(requireBytes(input))
			state.iteration += 1
			const updates = Array.from({ length: 100 }, (_, index) => ({
				ref: `A${index + 1}`,
				value: state.iteration * 10_000 + index,
			}))
			const edit = await state.session.apply([{ op: 'setCells', sheet: 'Sheet1', updates }], {
				recalc: false,
			})
			if (edit.apply.errors.length > 0) throw new Error('Patch-stream edit failed')
			const patchStart = performance.now()
			const patchRequest = {
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 250,
				colCount: 20,
				changedSince: state.lastToken,
			}
			const patch =
				state.session.readViewportPatch(patchRequest) ??
				state.session.readViewport(patchRequest).patch
			const patchReadMs = performance.now() - patchStart
			if (!patch || patch.changedCells.length !== 100) {
				throw new Error('Patch-stream benchmark returned an unexpected patch')
			}
			state.lastToken = patch.changeToken
			return {
				assertions: {
					changedCells: patch.changedCells.length,
					removedRefs: patch.removedRefs.length,
					patchBytes: patch.byteLength,
					sessionApplyMs: edit.timings.totalMs,
					ensureMutableWorkbookMs: edit.timings.ensureMutableWorkbookMs,
					rebaseViewportSnapshotsMs: edit.timings.rebaseViewportSnapshotsMs ?? null,
					applyMs: edit.timings.applyMs,
					recalcMs: edit.timings.recalcMs,
					generationSnapshotMs: edit.timings.generationSnapshotMs,
					inspectWriteMs: edit.timings.inspectWriteMs,
					patchReadMs,
					patchMode: 'delta',
				},
			}
		},
	},
	{
		name: 'patch-stream-first-edit',
		category: 'workflow',
		build() {
			return { bytes: cachedDenseViewportWorkbookBytes(), rows: 100, cols: 1, cells: 100 }
		},
		async run(input) {
			const bytes = requireBytes(input)
			const openStart = performance.now()
			const session = await AscendSession.open(bytes, { mode: 'interactive' })
			const openMs = performance.now() - openStart
			const previewStart = performance.now()
			const viewport = session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 250,
				colCount: 20,
			})
			const initialViewportMs = performance.now() - previewStart
			const updates = Array.from({ length: 100 }, (_, index) => ({
				ref: `A${index + 1}`,
				value: 20_000 + index,
			}))
			const edit = await session.apply([{ op: 'setCells', sheet: 'Sheet1', updates }], {
				recalc: false,
			})
			if (edit.apply.errors.length > 0) throw new Error('First-edit patch stream failed')
			const patchStart = performance.now()
			const patchRequest = {
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 250,
				colCount: 20,
				changedSince: viewport.changeToken,
			}
			const patch =
				session.readViewportPatch(patchRequest) ?? session.readViewport(patchRequest).patch
			const patchReadMs = performance.now() - patchStart
			session.close()
			if (!patch || patch.changedCells.length !== 100) {
				throw new Error('First-edit patch stream returned an unexpected patch')
			}
			return {
				assertions: {
					changedCells: patch.changedCells.length,
					removedRefs: patch.removedRefs.length,
					patchBytes: patch.byteLength,
					openMs,
					initialViewportMs,
					sessionApplyMs: edit.timings.totalMs,
					ensureMutableWorkbookMs: edit.timings.ensureMutableWorkbookMs,
					rebaseViewportSnapshotsMs: edit.timings.rebaseViewportSnapshotsMs ?? null,
					applyMs: edit.timings.applyMs,
					recalcMs: edit.timings.recalcMs,
					generationSnapshotMs: edit.timings.generationSnapshotMs,
					inspectWriteMs: edit.timings.inspectWriteMs,
					patchReadMs,
					promotedToFull: edit.load.promotedToFull,
					patchMode: 'delta',
				},
			}
		},
	},
	{
		name: 'patch-stream-prepared-first-edit',
		category: 'workflow',
		build() {
			return { bytes: cachedDenseViewportWorkbookBytes(), rows: 100, cols: 1, cells: 100 }
		},
		async run(input) {
			const bytes = requireBytes(input)
			const openStart = performance.now()
			const session = await AscendSession.open(bytes, { mode: 'interactive' })
			const openMs = performance.now() - openStart
			const previewStart = performance.now()
			const viewport = session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 250,
				colCount: 20,
			})
			const initialViewportMs = performance.now() - previewStart
			const prepare = await session.prepareEdits()
			const updates = Array.from({ length: 100 }, (_, index) => ({
				ref: `A${index + 1}`,
				value: 30_000 + index,
			}))
			const editFrameStart = performance.now()
			const edit = await session.apply([{ op: 'setCells', sheet: 'Sheet1', updates }], {
				recalc: false,
			})
			if (edit.apply.errors.length > 0) throw new Error('Prepared patch stream edit failed')
			const patchRequest = {
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 250,
				colCount: 20,
				changedSince: viewport.changeToken,
			}
			const patchStart = performance.now()
			const patch =
				session.readViewportPatch(patchRequest) ?? session.readViewport(patchRequest).patch
			const patchReadMs = performance.now() - patchStart
			const editFrameMs = performance.now() - editFrameStart
			session.close()
			if (!patch || patch.changedCells.length !== 100) {
				throw new Error('Prepared patch stream returned an unexpected patch')
			}
			return {
				assertions: {
					changedCells: patch.changedCells.length,
					removedRefs: patch.removedRefs.length,
					patchBytes: patch.byteLength,
					openMs,
					initialViewportMs,
					prepareEditsMs: prepare.timings.totalMs,
					prepareEnsureMutableWorkbookMs: prepare.timings.ensureMutableWorkbookMs,
					prepareMutableWorkbookOpenMs: prepare.timings.mutableWorkbookOpenMs ?? null,
					prepareRebaseViewportSnapshotsMs: prepare.timings.rebaseViewportSnapshotsMs ?? null,
					editFrameMs,
					sessionApplyMs: edit.timings.totalMs,
					ensureMutableWorkbookMs: edit.timings.ensureMutableWorkbookMs,
					rebaseViewportSnapshotsMs: edit.timings.rebaseViewportSnapshotsMs ?? null,
					applyMs: edit.timings.applyMs,
					recalcMs: edit.timings.recalcMs,
					generationSnapshotMs: edit.timings.generationSnapshotMs,
					inspectWriteMs: edit.timings.inspectWriteMs,
					patchReadMs,
					promotedToFull: edit.load.promotedToFull,
					patchMode: 'delta',
				},
			}
		},
	},
	{
		name: 'patch-stream-real-rich-prepared-first-edit',
		category: 'workflow',
		build() {
			return {
				bytes: cachedRealRichViewportWorkbookBytes(),
				rows: 100,
				cols: 20,
				cells: 4,
			}
		},
		async run(input) {
			const bytes = requireBytes(input)
			const openStart = performance.now()
			const session = await AscendSession.open(bytes, { mode: 'interactive' })
			const openMs = performance.now() - openStart
			const previewStart = performance.now()
			const viewport = session.readViewport({
				sheet: 'Data',
				topRow: 0,
				leftCol: 0,
				rowCount: 100,
				colCount: 20,
			})
			const initialViewportMs = performance.now() - previewStart
			const prepare = await session.prepareEdits()
			const updates = [
				{ ref: 'C2', value: 42_001 },
				{ ref: 'C3', value: 42_002 },
				{ ref: 'C4', value: 42_003 },
				{ ref: 'C5', value: 42_004 },
			]
			const editFrameStart = performance.now()
			const edit = await session.apply([{ op: 'setCells', sheet: 'Data', updates }], {
				recalc: false,
			})
			if (edit.apply.errors.length > 0) {
				throw new Error('Real rich prepared patch stream edit failed')
			}
			const patchRequest = {
				sheet: 'Data',
				topRow: 0,
				leftCol: 0,
				rowCount: 100,
				colCount: 20,
				changedSince: viewport.changeToken,
			}
			const patchStart = performance.now()
			const patch =
				session.readViewportPatch(patchRequest) ?? session.readViewport(patchRequest).patch
			const patchReadMs = performance.now() - patchStart
			const editFrameMs = performance.now() - editFrameStart
			session.close()
			if (!patch || patch.changedCells.length !== updates.length) {
				throw new Error('Real rich prepared patch stream returned an unexpected patch')
			}
			return {
				assertions: {
					changedCells: patch.changedCells.length,
					removedRefs: patch.removedRefs.length,
					patchBytes: patch.byteLength,
					openMs,
					initialViewportMs,
					prepareEditsMs: prepare.timings.totalMs,
					prepareEnsureMutableWorkbookMs: prepare.timings.ensureMutableWorkbookMs,
					editFrameMs,
					sessionApplyMs: edit.timings.totalMs,
					ensureMutableWorkbookMs: edit.timings.ensureMutableWorkbookMs,
					applyMs: edit.timings.applyMs,
					recalcMs: edit.timings.recalcMs,
					generationSnapshotMs: edit.timings.generationSnapshotMs,
					inspectWriteMs: edit.timings.inspectWriteMs,
					patchReadMs,
					viewportCells: viewport.cells.length,
					viewportFlatValues: viewport.flatValues.length,
					tables: viewport.tables.length,
					validations: viewport.dataValidations.length,
					conditionalFormats: viewport.conditionalFormats.length,
					merges: viewport.merges.length,
					promotedToFull: edit.load.promotedToFull,
					patchMode: 'delta',
				},
			}
		},
	},
	{
		name: 'patch-stream-real-corpus-prepared-first-edit',
		category: 'workflow',
		build() {
			return {
				rows: REAL_INTERACTIVE_PATCH_CORPUS.length,
				cols: 1,
				cells: REAL_INTERACTIVE_PATCH_CORPUS.length,
				byteCount: realInteractivePatchCorpusByteCount(),
			}
		},
		async run() {
			const targetAssertions: Record<string, string | number | boolean | null> = {}
			let totalBytes = 0
			let totalOpenMs = 0
			let totalInitialViewportMs = 0
			let totalPrepareEditsMs = 0
			let totalPrepareEnsureMutableWorkbookMs = 0
			let totalPrepareMutableWorkbookOpenMs = 0
			let totalPrepareRebaseViewportSnapshotsMs = 0
			let totalEditFrameMs = 0
			let totalSessionApplyMs = 0
			let totalApplyMs = 0
			let totalPatchReadMs = 0
			let totalPatchBytes = 0
			let totalChangedCells = 0
			let totalViewportCells = 0
			let totalViewportFlatValues = 0
			let totalTables = 0
			let totalComments = 0
			let totalHyperlinks = 0
			let totalValidations = 0
			let totalConditionalFormats = 0
			let totalMerges = 0
			let maxPrepareEditsMs = 0
			let slowestPrepareTarget = ''

			for (const [index, target] of REAL_INTERACTIVE_PATCH_CORPUS.entries()) {
				const bytes = realInteractivePatchCorpusTargetBytes(target)
				const assertionPrefix = target.label.replaceAll('-', '_')
				totalBytes += bytes.byteLength
				const openStart = performance.now()
				const session = await AscendSession.open(bytes, { mode: 'interactive' })
				const openMs = performance.now() - openStart
				totalOpenMs += openMs
				const previewStart = performance.now()
				const viewport = session.readViewport({
					sheet: target.sheet,
					topRow: target.topRow,
					leftCol: target.leftCol,
					rowCount: target.rowCount,
					colCount: target.colCount,
				})
				const initialViewportMs = performance.now() - previewStart
				totalInitialViewportMs += initialViewportMs
				totalViewportCells += viewport.cells.length
				totalViewportFlatValues += viewport.flatValues.length
				totalTables += viewport.tables.length
				totalComments += viewport.comments.length
				totalHyperlinks += viewport.hyperlinks.length
				totalValidations += viewport.dataValidations.length
				totalConditionalFormats += viewport.conditionalFormats.length
				totalMerges += viewport.merges.length

				const prepare = await session.prepareEdits()
				const prepareEditsMs = prepare.timings.totalMs
				const prepareEnsureMutableWorkbookMs = prepare.timings.ensureMutableWorkbookMs
				totalPrepareEditsMs += prepare.timings.totalMs
				totalPrepareEnsureMutableWorkbookMs += prepare.timings.ensureMutableWorkbookMs
				totalPrepareMutableWorkbookOpenMs += prepare.timings.mutableWorkbookOpenMs ?? 0
				totalPrepareRebaseViewportSnapshotsMs += prepare.timings.rebaseViewportSnapshotsMs ?? 0
				if (prepareEditsMs > maxPrepareEditsMs) {
					maxPrepareEditsMs = prepareEditsMs
					slowestPrepareTarget = target.label
				}

				const editFrameStart = performance.now()
				const edit = await session.apply(
					[
						{
							op: 'setCells',
							sheet: target.sheet,
							updates: [{ ref: target.editRef, value: 70_000 + index }],
						},
					],
					{ recalc: false },
				)
				if (edit.apply.errors.length > 0) {
					throw new Error(`Real patch corpus edit failed for ${target.label}`)
				}
				const patchRequest = {
					sheet: target.sheet,
					topRow: target.topRow,
					leftCol: target.leftCol,
					rowCount: target.rowCount,
					colCount: target.colCount,
					changedSince: viewport.changeToken,
				}
				const patchStart = performance.now()
				const patch =
					session.readViewportPatch(patchRequest) ?? session.readViewport(patchRequest).patch
				const patchReadMs = performance.now() - patchStart
				const editFrameMs = performance.now() - editFrameStart
				totalPatchReadMs += patchReadMs
				totalEditFrameMs += editFrameMs
				session.close()
				if (!patch || patch.changedCells.length !== 1) {
					throw new Error(`Real patch corpus returned an unexpected patch for ${target.label}`)
				}
				totalSessionApplyMs += edit.timings.totalMs
				totalApplyMs += edit.timings.applyMs
				totalPatchBytes += patch.byteLength
				totalChangedCells += patch.changedCells.length
				targetAssertions[`${assertionPrefix}.bytes`] = bytes.byteLength
				targetAssertions[`${assertionPrefix}.viewportCells`] = viewport.cells.length
				targetAssertions[`${assertionPrefix}.openMs`] = openMs
				targetAssertions[`${assertionPrefix}.initialViewportMs`] = initialViewportMs
				targetAssertions[`${assertionPrefix}.prepareEditsMs`] = prepareEditsMs
				targetAssertions[`${assertionPrefix}.prepareEnsureMutableWorkbookMs`] =
					prepareEnsureMutableWorkbookMs
				targetAssertions[`${assertionPrefix}.prepareMutableWorkbookOpenMs`] =
					prepare.timings.mutableWorkbookOpenMs ?? null
				targetAssertions[`${assertionPrefix}.prepareRebaseViewportSnapshotsMs`] =
					prepare.timings.rebaseViewportSnapshotsMs ?? null
				targetAssertions[`${assertionPrefix}.editFrameMs`] = editFrameMs
				targetAssertions[`${assertionPrefix}.sessionApplyMs`] = edit.timings.totalMs
				targetAssertions[`${assertionPrefix}.applyMs`] = edit.timings.applyMs
				targetAssertions[`${assertionPrefix}.patchReadMs`] = patchReadMs
				targetAssertions[`${assertionPrefix}.patchBytes`] = patch.byteLength
			}

			return {
				assertions: {
					workbooks: REAL_INTERACTIVE_PATCH_CORPUS.length,
					totalBytes,
					totalViewportCells,
					totalViewportFlatValues,
					tables: totalTables,
					comments: totalComments,
					hyperlinks: totalHyperlinks,
					validations: totalValidations,
					conditionalFormats: totalConditionalFormats,
					merges: totalMerges,
					changedCells: totalChangedCells,
					patchBytes: totalPatchBytes,
					openMs: totalOpenMs,
					initialViewportMs: totalInitialViewportMs,
					prepareEditsMs: totalPrepareEditsMs,
					prepareEnsureMutableWorkbookMs: totalPrepareEnsureMutableWorkbookMs,
					prepareMutableWorkbookOpenMs: totalPrepareMutableWorkbookOpenMs,
					prepareRebaseViewportSnapshotsMs: totalPrepareRebaseViewportSnapshotsMs,
					editFrameMs: totalEditFrameMs,
					sessionApplyMs: totalSessionApplyMs,
					applyMs: totalApplyMs,
					patchReadMs: totalPatchReadMs,
					maxPrepareEditsMs,
					slowestPrepareTarget,
					patchMode: 'delta',
					...targetAssertions,
				},
			}
		},
	},
	{
		name: 'real-corpus-open-phases',
		category: 'read',
		build() {
			return {
				rows: REAL_INTERACTIVE_PATCH_CORPUS.length,
				cols: 4,
				cells: REAL_INTERACTIVE_PATCH_CORPUS.length * 4,
				byteCount: realInteractivePatchCorpusByteCount(),
			}
		},
		async run() {
			const targetAssertions: Record<string, string | number | boolean | null> = {}
			let totalBytes = 0
			let totalFormulaReadXlsxMs = 0
			let totalFullReadXlsxMs = 0
			let totalSessionColdOpenMs = 0
			let totalSessionHotOpenMs = 0
			let totalFullWorkbookOpenMs = 0
			let maxFullWorkbookOpenMs = 0
			let slowestFullWorkbookOpenTarget = ''

			for (const target of REAL_INTERACTIVE_PATCH_CORPUS) {
				const bytes = realInteractivePatchCorpusTargetBytes(target)
				const assertionPrefix = target.label.replaceAll('-', '_')
				totalBytes += bytes.byteLength

				const formulaReadStart = performance.now()
				const formulaRead = readXlsx(bytes, { mode: 'formula', richMetadata: true })
				const formulaReadXlsxMs = performance.now() - formulaReadStart
				if (!formulaRead.ok) {
					throw new Error(`Formula read failed for ${target.label}: ${formulaRead.error.message}`)
				}

				const fullReadStart = performance.now()
				const fullRead = readXlsx(bytes, { mode: 'full', richMetadata: true })
				const fullReadXlsxMs = performance.now() - fullReadStart
				if (!fullRead.ok) {
					throw new Error(`Full read failed for ${target.label}: ${fullRead.error.message}`)
				}

				WorkbookDocument.clearCache()
				const sessionColdOpenStart = performance.now()
				const coldSession = await AscendSession.open(bytes, { mode: 'interactive' })
				const sessionColdOpenMs = performance.now() - sessionColdOpenStart
				coldSession.close()

				const sessionHotOpenStart = performance.now()
				const hotSession = await AscendSession.open(bytes, { mode: 'interactive' })
				const sessionHotOpenMs = performance.now() - sessionHotOpenStart
				hotSession.close()

				const fullWorkbookOpenStart = performance.now()
				const workbook = await AscendWorkbook.open(bytes, { mode: 'full', richMetadata: true })
				const fullWorkbookOpenMs = performance.now() - fullWorkbookOpenStart
				if (workbook.inspect().load.isPartial) {
					throw new Error(`Full workbook open stayed partial for ${target.label}`)
				}

				totalFormulaReadXlsxMs += formulaReadXlsxMs
				totalFullReadXlsxMs += fullReadXlsxMs
				totalSessionColdOpenMs += sessionColdOpenMs
				totalSessionHotOpenMs += sessionHotOpenMs
				totalFullWorkbookOpenMs += fullWorkbookOpenMs
				if (fullWorkbookOpenMs > maxFullWorkbookOpenMs) {
					maxFullWorkbookOpenMs = fullWorkbookOpenMs
					slowestFullWorkbookOpenTarget = target.label
				}

				targetAssertions[`${assertionPrefix}.bytes`] = bytes.byteLength
				targetAssertions[`${assertionPrefix}.formulaReadXlsxMs`] = formulaReadXlsxMs
				targetAssertions[`${assertionPrefix}.fullReadXlsxMs`] = fullReadXlsxMs
				targetAssertions[`${assertionPrefix}.sessionColdOpenMs`] = sessionColdOpenMs
				targetAssertions[`${assertionPrefix}.sessionHotOpenMs`] = sessionHotOpenMs
				targetAssertions[`${assertionPrefix}.fullWorkbookOpenMs`] = fullWorkbookOpenMs
			}

			return {
				assertions: {
					workbooks: REAL_INTERACTIVE_PATCH_CORPUS.length,
					totalBytes,
					formulaReadXlsxMs: totalFormulaReadXlsxMs,
					fullReadXlsxMs: totalFullReadXlsxMs,
					sessionColdOpenMs: totalSessionColdOpenMs,
					sessionHotOpenMs: totalSessionHotOpenMs,
					fullWorkbookOpenMs: totalFullWorkbookOpenMs,
					maxFullWorkbookOpenMs,
					slowestFullWorkbookOpenTarget,
					...targetAssertions,
				},
			}
		},
	},
	{
		name: 'real-corpus-grid-storage',
		category: 'read',
		build() {
			return {
				rows: REAL_INTERACTIVE_PATCH_CORPUS.length,
				cols: 1,
				cells: REAL_INTERACTIVE_PATCH_CORPUS.length,
				byteCount: realInteractivePatchCorpusByteCount(),
			}
		},
		run() {
			const targetAssertions: Record<string, string | number | boolean | null> = {}
			let totalBytes = 0
			let totalReadMs = 0
			let totalGridCells = 0
			let totalDenseChunks = 0
			let totalSparseChunks = 0
			let totalGridArrayBufferBytes = 0
			let densestTarget = ''
			let maxDenseChunks = 0

			for (const target of REAL_INTERACTIVE_PATCH_CORPUS) {
				const bytes = realInteractivePatchCorpusTargetBytes(target)
				const assertionPrefix = target.label.replaceAll('-', '_')
				totalBytes += bytes.byteLength
				const readStart = performance.now()
				const result = readXlsx(bytes, { mode: 'full', richMetadata: true })
				const readMs = performance.now() - readStart
				if (!result.ok) {
					throw new Error(`Grid storage read failed for ${target.label}: ${result.error.message}`)
				}
				const workbook = result.value.workbook
				const storage = workbookGridStorageAssertions(workbook)
				const shape = workbookGridShapeAssertions(workbook)
				totalReadMs += readMs
				totalGridCells += storage.gridCellCount
				totalDenseChunks += storage.gridDenseChunkCount
				totalSparseChunks += storage.gridSparseChunkCount
				totalGridArrayBufferBytes += storage.gridTotalArrayBufferBytes
				if (storage.gridDenseChunkCount > maxDenseChunks) {
					maxDenseChunks = storage.gridDenseChunkCount
					densestTarget = target.label
				}

				targetAssertions[`${assertionPrefix}.bytes`] = bytes.byteLength
				targetAssertions[`${assertionPrefix}.readMs`] = readMs
				for (const [key, value] of Object.entries(shape)) {
					targetAssertions[`${assertionPrefix}.${key}`] = value
				}
				for (const [key, value] of Object.entries(storage)) {
					targetAssertions[`${assertionPrefix}.${key}`] = value
				}
			}

			return {
				assertions: {
					workbooks: REAL_INTERACTIVE_PATCH_CORPUS.length,
					totalBytes,
					totalReadMs,
					totalGridCells,
					totalDenseChunks,
					totalSparseChunks,
					totalGridArrayBufferBytes,
					gridArrayBufferBytesPerCell:
						totalGridCells > 0 ? totalGridArrayBufferBytes / totalGridCells : 0,
					maxDenseChunks,
					densestTarget,
					...targetAssertions,
				},
			}
		},
	},
	createRealDenseReadMemoryScenario('real-dense-read-memory-metadata', {
		mode: 'metadata-only',
		richMetadata: true,
	}),
	createRealDenseReadMemoryScenario('real-dense-read-memory-values', {
		mode: 'values',
		richMetadata: true,
	}),
	createRealDenseReadMemoryScenario('real-dense-read-memory-formula', {
		mode: 'formula',
		richMetadata: true,
	}),
	createRealDenseReadMemoryScenario('real-dense-read-memory-full', {
		mode: 'full',
		richMetadata: true,
	}),
	{
		name: 'real-dense-promotion-memory',
		category: 'workflow',
		build() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			return {
				rows: target.rowCount,
				cols: target.colCount,
				cells: target.rowCount * target.colCount,
				byteCount: bytes.byteLength,
			}
		},
		async run() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			const archiveFootprint = xlsxArchiveFootprint(bytes)
			WorkbookDocument.clearCache()
			runGc()
			const baseline = phaseMemorySnapshot()

			const openStart = performance.now()
			const session = await AscendSession.open(bytes, { mode: 'interactive' })
			const openMs = performance.now() - openStart
			runGc()
			const afterOpen = phaseMemorySnapshot()

			const viewportStart = performance.now()
			const viewport = session.readViewport({
				sheet: target.sheet,
				topRow: target.topRow,
				leftCol: target.leftCol,
				rowCount: target.rowCount,
				colCount: target.colCount,
			})
			const initialViewportMs = performance.now() - viewportStart
			runGc()
			const afterViewport = phaseMemorySnapshot()

			const prepare = await session.prepareEdits()
			runGc()
			const afterPrepare = phaseMemorySnapshot()

			const editStart = performance.now()
			const edit = await session.apply(
				[
					{
						op: 'setCells',
						sheet: target.sheet,
						updates: [{ ref: target.editRef, value: 80_000 }],
					},
				],
				{ recalc: false },
			)
			if (edit.apply.errors.length > 0) throw new Error('Dense promotion memory edit failed')
			const patchRequest = {
				sheet: target.sheet,
				topRow: target.topRow,
				leftCol: target.leftCol,
				rowCount: target.rowCount,
				colCount: target.colCount,
				changedSince: viewport.changeToken,
			}
			const patchStart = performance.now()
			let patch = session.readViewportPatch(patchRequest)
			let patchMode: 'delta' | 'fresh' = 'delta'
			if (!patch) {
				const fresh = session.readViewport(patchRequest)
				patch = fresh.patch ?? null
				if (!patch) {
					patchMode = 'fresh'
					const editedCell = fresh.cells.find((cell) => cell.ref === target.editRef)
					if (editedCell?.flatValue !== 80_000) {
						throw new Error('Dense promotion memory fresh snapshot missed edited cell')
					}
				}
			}
			const patchReadMs = performance.now() - patchStart
			const editFrameMs = performance.now() - editStart
			if (patchMode === 'delta' && (!patch || patch.changedCells.length !== 1)) {
				throw new Error('Dense promotion memory patch was not delta-shaped')
			}
			runGc()
			const afterPatch = phaseMemorySnapshot()

			session.close()
			WorkbookDocument.clearCache()
			runGc()
			const afterClose = phaseMemorySnapshot()

			return {
				assertions: {
					bytes: bytes.byteLength,
					archivePartCount: archiveFootprint.partCount,
					archiveCompressedBytes: archiveFootprint.compressedBytes,
					archiveUncompressedBytes: archiveFootprint.uncompressedBytes,
					archiveWorksheetUncompressedBytes: archiveFootprint.worksheetUncompressedBytes,
					archiveLargestPartPath: archiveFootprint.largestPartPath,
					archiveLargestPartCompressedBytes: archiveFootprint.largestPartCompressedBytes,
					archiveLargestPartUncompressedBytes: archiveFootprint.largestPartUncompressedBytes,
					viewportCells: viewport.cells.length,
					viewportFlatValues: viewport.flatValues.length,
					changedCells: patch?.changedCells.length ?? 0,
					patchBytes: patch?.byteLength ?? 0,
					openMs,
					initialViewportMs,
					prepareEditsMs: prepare.timings.totalMs,
					prepareEnsureMutableWorkbookMs: prepare.timings.ensureMutableWorkbookMs,
					editFrameMs,
					sessionApplyMs: edit.timings.totalMs,
					applyMs: edit.timings.applyMs,
					patchReadMs,
					baselineRssBytes: baseline.rssBytes,
					afterOpenRssDeltaBytes: memoryDelta(afterOpen.rssBytes, baseline.rssBytes),
					afterViewportRssDeltaBytes: memoryDelta(afterViewport.rssBytes, baseline.rssBytes),
					afterPrepareRssDeltaBytes: memoryDelta(afterPrepare.rssBytes, baseline.rssBytes),
					afterPatchRssDeltaBytes: memoryDelta(afterPatch.rssBytes, baseline.rssBytes),
					afterCloseRssDeltaBytes: memoryDelta(afterClose.rssBytes, baseline.rssBytes),
					baselineHeapUsedBytes: baseline.heapUsedBytes,
					afterOpenHeapDeltaBytes: memoryDelta(afterOpen.heapUsedBytes, baseline.heapUsedBytes),
					afterViewportHeapDeltaBytes: memoryDelta(
						afterViewport.heapUsedBytes,
						baseline.heapUsedBytes,
					),
					afterPrepareHeapDeltaBytes: memoryDelta(
						afterPrepare.heapUsedBytes,
						baseline.heapUsedBytes,
					),
					afterPatchHeapDeltaBytes: memoryDelta(afterPatch.heapUsedBytes, baseline.heapUsedBytes),
					afterCloseHeapDeltaBytes: memoryDelta(afterClose.heapUsedBytes, baseline.heapUsedBytes),
					baselineExternalBytes: baseline.externalBytes,
					afterOpenExternalDeltaBytes: memoryDelta(afterOpen.externalBytes, baseline.externalBytes),
					afterViewportExternalDeltaBytes: memoryDelta(
						afterViewport.externalBytes,
						baseline.externalBytes,
					),
					afterPrepareExternalDeltaBytes: memoryDelta(
						afterPrepare.externalBytes,
						baseline.externalBytes,
					),
					afterPatchExternalDeltaBytes: memoryDelta(
						afterPatch.externalBytes,
						baseline.externalBytes,
					),
					afterCloseExternalDeltaBytes: memoryDelta(
						afterClose.externalBytes,
						baseline.externalBytes,
					),
					baselineArrayBuffersBytes: baseline.arrayBuffersBytes,
					afterOpenArrayBuffersDeltaBytes: memoryDelta(
						afterOpen.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
					afterViewportArrayBuffersDeltaBytes: memoryDelta(
						afterViewport.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
					afterPrepareArrayBuffersDeltaBytes: memoryDelta(
						afterPrepare.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
					afterPatchArrayBuffersDeltaBytes: memoryDelta(
						afterPatch.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
					afterCloseArrayBuffersDeltaBytes: memoryDelta(
						afterClose.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
					patchMode,
				},
			}
		},
	},
	{
		name: 'real-dense-progressive-readiness',
		category: 'workflow',
		build() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			return {
				rows: target.rowCount,
				cols: target.colCount,
				cells: target.rowCount * target.colCount,
				byteCount: realInteractivePatchCorpusTargetBytes(target).byteLength,
			}
		},
		async run() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)

			WorkbookDocument.clearCache()
			const currentOpenStart = performance.now()
			const currentSession = await AscendSession.open(bytes, { mode: 'interactive' })
			const currentOpenMs = performance.now() - currentOpenStart
			const currentViewportStart = performance.now()
			const currentViewport = currentSession.readViewport({
				sheet: target.sheet,
				topRow: target.topRow,
				leftCol: target.leftCol,
				rowCount: target.rowCount,
				colCount: target.colCount,
			})
			const currentInitialViewportMs = performance.now() - currentViewportStart
			const currentPrepare = await currentSession.prepareEdits()
			currentSession.close()

			WorkbookDocument.clearCache()
			const progressiveFirstWindowStart = performance.now()
			const firstWindow = await WorkbookSession.openFirstWindow(bytes, {
				sheet: target.sheet,
				range: 'A1:T100',
				rowLimit: target.rowCount,
				flatValues: true,
				omitEmpty: true,
			})
			const progressiveFirstWindowMs = performance.now() - progressiveFirstWindowStart
			firstWindow.session.close()

			WorkbookDocument.clearCache()
			const backgroundPrepareStart = performance.now()
			const backgroundSession = await AscendSession.open(bytes, {
				mode: 'interactive',
				prepareEdits: true,
			})
			const backgroundInitialReadiness = backgroundSession.editReadiness()
			const backgroundOpenMs = performance.now() - backgroundPrepareStart
			const backgroundPrepare = await backgroundSession.prepareEdits()
			const backgroundPreparedReadiness = backgroundSession.editReadiness()
			const backgroundPrepareTotalMs = performance.now() - backgroundPrepareStart
			const backgroundViewportStart = performance.now()
			const backgroundViewport = backgroundSession.readViewport({
				sheet: target.sheet,
				topRow: target.topRow,
				leftCol: target.leftCol,
				rowCount: target.rowCount,
				colCount: target.colCount,
			})
			const backgroundViewportMs = performance.now() - backgroundViewportStart
			const backgroundEditStart = performance.now()
			const backgroundEdit = await backgroundSession.apply(
				[
					{
						op: 'setCells',
						sheet: target.sheet,
						updates: [{ ref: target.editRef, value: 90_000 }],
					},
				],
				{ recalc: false },
			)
			if (backgroundEdit.apply.errors.length > 0) {
				throw new Error('Dense progressive readiness prepared edit failed')
			}
			const backgroundPatchStart = performance.now()
			const backgroundPatch =
				backgroundSession.readViewportPatch({
					sheet: target.sheet,
					topRow: target.topRow,
					leftCol: target.leftCol,
					rowCount: target.rowCount,
					colCount: target.colCount,
					changedSince: backgroundViewport.changeToken,
				}) ??
				backgroundSession.readViewport({
					sheet: target.sheet,
					topRow: target.topRow,
					leftCol: target.leftCol,
					rowCount: target.rowCount,
					colCount: target.colCount,
					changedSince: backgroundViewport.changeToken,
				}).patch
			const backgroundPatchReadMs = performance.now() - backgroundPatchStart
			const backgroundEditFrameMs = performance.now() - backgroundEditStart
			backgroundSession.close()
			if (!backgroundPatch || backgroundPatch.changedCells.length !== 1) {
				throw new Error('Dense progressive readiness patch was not delta-shaped')
			}

			const progressiveTimeToEditReadyMs = progressiveFirstWindowMs + backgroundPrepareTotalMs
			const waitBudgetAssertions: Record<string, number> = {}
			for (const delayMs of DENSE_READINESS_THINK_DELAYS_MS) {
				waitBudgetAssertions[`remainingWaitAfterFirstPaintAt${delayMs}Ms`] = Math.max(
					0,
					backgroundPrepareTotalMs - delayMs,
				)
			}

			return {
				assertions: {
					bytes: bytes.byteLength,
					currentOpenMs,
					currentInitialViewportMs,
					currentTimeToFirstViewportMs: currentOpenMs + currentInitialViewportMs,
					currentPrepareEditsMs: currentPrepare.timings.totalMs,
					currentPrepareReusedReadModel:
						currentPrepare.timings.mutableWorkbookReusedReadModel ?? false,
					currentPrepareMutableWorkbookOpenMs: currentPrepare.timings.mutableWorkbookOpenMs ?? null,
					currentPrepareRebaseViewportSnapshotsMs:
						currentPrepare.timings.rebaseViewportSnapshotsMs ?? null,
					currentTimeToEditReadyMs:
						currentOpenMs + currentInitialViewportMs + currentPrepare.timings.totalMs,
					currentViewportCells: currentViewport.cells.length,
					currentViewportFlatValues: currentViewport.flatValues.length,
					progressiveFirstWindowMs,
					progressiveWindowCells: firstWindow.window.cells.length,
					progressiveWindowFlatValues: firstWindow.window.cells.length,
					progressiveHasMore: firstWindow.window.hasMore,
					progressiveLoadIsPartial: firstWindow.load.isPartial,
					progressivePartialReasons: firstWindow.load.partialReasons.length,
					backgroundOpenMs,
					backgroundInitialReady: backgroundInitialReadiness.ready,
					backgroundInitialPreparing: backgroundInitialReadiness.preparing,
					backgroundReadinessReusedReadModel:
						backgroundInitialReadiness.timings?.mutableWorkbookReusedReadModel ?? false,
					backgroundReadinessMutableWorkbookOpenMs:
						backgroundInitialReadiness.timings?.mutableWorkbookOpenMs ?? null,
					backgroundPrepareEditsMs: backgroundPrepare.timings.totalMs,
					backgroundPrepareEnsureMutableWorkbookMs:
						backgroundPrepare.timings.ensureMutableWorkbookMs,
					backgroundPrepareCached: backgroundPrepare.timings.mutableWorkbookCached ?? false,
					backgroundPrepareReusedReadModel:
						backgroundPrepare.timings.mutableWorkbookReusedReadModel ?? false,
					backgroundPrepareMutableWorkbookOpenMs:
						backgroundPrepare.timings.mutableWorkbookOpenMs ?? null,
					backgroundPrepareRebaseViewportSnapshotsMs:
						backgroundPrepare.timings.rebaseViewportSnapshotsMs ?? null,
					backgroundPreparedReady: backgroundPreparedReadiness.ready,
					backgroundPreparedPreparing: backgroundPreparedReadiness.preparing,
					backgroundPrepareTotalMs,
					progressiveTimeToEditReadyMs,
					progressiveEditReadyDeltaVsCurrentMs:
						progressiveTimeToEditReadyMs -
						(currentOpenMs + currentInitialViewportMs + currentPrepare.timings.totalMs),
					backgroundViewportMs,
					backgroundViewportCells: backgroundViewport.cells.length,
					backgroundEditFrameMs,
					backgroundSessionApplyMs: backgroundEdit.timings.totalMs,
					backgroundApplyMs: backgroundEdit.timings.applyMs,
					backgroundPatchReadMs,
					backgroundPatchBytes: backgroundPatch.byteLength,
					backgroundPatchChangedCells: backgroundPatch.changedCells.length,
					firstPaintSpeedupVsCurrent:
						(currentOpenMs + currentInitialViewportMs) / progressiveFirstWindowMs,
					...waitBudgetAssertions,
				},
			}
		},
	},
	{
		name: 'real-dense-interactive-open-modes',
		category: 'workflow',
		build() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			return {
				rows: target.rowCount,
				cols: target.colCount,
				cells: target.rowCount * target.colCount,
				byteCount: bytes.byteLength,
			}
		},
		async run() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			const defaultInteractive = await measureInteractiveOpenPath(
				'defaultInteractive',
				bytes,
				target,
				{ mode: 'interactive' },
			)
			const explicitFull = await measureInteractiveOpenPath('explicitFull', bytes, target, {
				mode: 'full',
			})
			const preparedInteractive = await measureInteractiveOpenPath(
				'preparedInteractive',
				bytes,
				target,
				{ mode: 'interactive', prepareEdits: true },
			)
			const defaultEditReadyMs = Number(defaultInteractive.defaultInteractiveTimeToEditReadyMs)
			const explicitFullEditReadyMs = Number(explicitFull.explicitFullTimeToEditReadyMs)
			const preparedEditReadyMs = Number(preparedInteractive.preparedInteractiveTimeToEditReadyMs)
			return {
				assertions: {
					bytes: bytes.byteLength,
					...defaultInteractive,
					...explicitFull,
					...preparedInteractive,
					explicitFullEditReadyDeltaVsDefaultMs: explicitFullEditReadyMs - defaultEditReadyMs,
					preparedEditReadyDeltaVsDefaultMs: preparedEditReadyMs - defaultEditReadyMs,
					defaultVsExplicitFullEditReadyHeadroom: defaultEditReadyMs / explicitFullEditReadyMs,
					defaultVsPreparedEditReadyHeadroom: defaultEditReadyMs / preparedEditReadyMs,
				},
			}
		},
	},
	{
		name: 'real-dense-worksheet-shape',
		category: 'workflow',
		build() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			return {
				rows: target.rowCount,
				cols: target.colCount,
				cells: target.rowCount * target.colCount,
				byteCount: bytes.byteLength,
			}
		},
		run() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			return {
				assertions: {
					bytes: bytes.byteLength,
					...prefixAssertions('sharedStrings', xlsxSharedStringUsage(bytes)),
					...prefixAssertions('worksheetScalar', xlsxWorksheetScalarUsage(bytes)),
				},
			}
		},
	},
	{
		name: 'real-dense-dirty-write-lifecycle',
		category: 'workflow',
		build() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			return {
				rows: target.rowCount,
				cols: target.colCount,
				cells: target.rowCount * target.colCount,
				byteCount: bytes.byteLength,
			}
		},
		async run() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			const updates = Array.from({ length: 10 }, (_, index) => ({
				ref: `A${index + 1}`,
				value: 120_000 + index,
			}))

			WorkbookDocument.clearCache()
			const openStart = performance.now()
			const workbook = await AscendWorkbook.open(bytes)
			const openMs = performance.now() - openStart

			const applyStart = performance.now()
			const apply = workbook.apply([{ op: 'setCells', sheet: target.sheet, updates }])
			if (apply.errors.length > 0) throw new Error('Dense dirty write lifecycle apply failed')
			const applyMs = performance.now() - applyStart

			const summaryStart = performance.now()
			const plan = workbook.writePlanSummary()
			const writePlanSummaryMs = performance.now() - summaryStart

			const checkStart = performance.now()
			const check = workbook.check()
			const writePolicyCheckMs = performance.now() - checkStart
			if (!check.valid) throw new Error('Dense dirty write lifecycle check failed')

			const toBytesStart = performance.now()
			const output = workbook.toBytes()
			const toBytesMs = performance.now() - toBytesStart

			const reopenStart = performance.now()
			const reopened = await AscendWorkbook.open(output)
			const reopenMs = performance.now() - reopenStart
			const edited = reopened.sheet(target.sheet)?.cell('A1')?.value
			if (edited?.kind !== 'number' || edited.value !== 120_000) {
				throw new Error('Dense dirty write lifecycle reopen missed edited value')
			}

			WorkbookDocument.clearCache()
			const toBytesFirstOpenStart = performance.now()
			const toBytesFirstWorkbook = await AscendWorkbook.open(bytes)
			const toBytesFirstOpenMs = performance.now() - toBytesFirstOpenStart

			const toBytesFirstApplyStart = performance.now()
			const toBytesFirstApply = toBytesFirstWorkbook.apply([
				{ op: 'setCells', sheet: target.sheet, updates },
			])
			if (toBytesFirstApply.errors.length > 0) {
				throw new Error('Dense dirty write lifecycle toBytes-first apply failed')
			}
			const toBytesFirstApplyMs = performance.now() - toBytesFirstApplyStart

			const toBytesFirstStart = performance.now()
			const toBytesFirstOutput = toBytesFirstWorkbook.toBytes()
			const toBytesFirstToBytesMs = performance.now() - toBytesFirstStart

			const toBytesFirstCheckStart = performance.now()
			const toBytesFirstCheck = toBytesFirstWorkbook.check()
			const toBytesFirstCheckMs = performance.now() - toBytesFirstCheckStart
			if (!toBytesFirstCheck.valid) {
				throw new Error('Dense dirty write lifecycle toBytes-first check failed')
			}

			const toBytesFirstReopenStart = performance.now()
			const toBytesFirstReopened = await AscendWorkbook.open(toBytesFirstOutput)
			const toBytesFirstReopenMs = performance.now() - toBytesFirstReopenStart
			const toBytesFirstEdited = toBytesFirstReopened.sheet(target.sheet)?.cell('A1')?.value
			if (toBytesFirstEdited?.kind !== 'number' || toBytesFirstEdited.value !== 120_000) {
				throw new Error('Dense dirty write lifecycle toBytes-first reopen missed edited value')
			}

			const generatedSheetParts = plan.parts.filter(
				(part) => part.owner.kind === 'sheet' && part.origin === 'generated',
			).length
			const preservedSourceSheetParts = plan.parts.filter(
				(part) => part.owner.kind === 'sheet' && part.origin === 'preserved-source',
			).length

			return {
				assertions: {
					bytes: bytes.byteLength,
					outputBytes: output.byteLength,
					outputDeltaBytes: output.byteLength - bytes.byteLength,
					openMs,
					applyMs,
					writePlanSummaryMs,
					writePolicyCheckMs,
					toBytesMs,
					reopenMs,
					checkFirstTotalMs: openMs + applyMs + writePlanSummaryMs + writePolicyCheckMs + toBytesMs,
					toBytesFirstOpenMs,
					toBytesFirstApplyMs,
					toBytesFirstToBytesMs,
					toBytesFirstCheckMs,
					toBytesFirstReopenMs,
					toBytesFirstTotalMs:
						toBytesFirstOpenMs + toBytesFirstApplyMs + toBytesFirstToBytesMs + toBytesFirstCheckMs,
					toBytesFirstOutputBytes: toBytesFirstOutput.byteLength,
					writePlanParts: plan.totalParts,
					writePlanGeneratedParts: plan.byOrigin.generated,
					writePlanPreservedSourceParts: plan.byOrigin['preserved-source'],
					generatedSheetParts,
					preservedSourceSheetParts,
					checkIssues: check.issues.length,
					toBytesFirstCheckIssues: toBytesFirstCheck.issues.length,
					updatedCells: updates.length,
				},
			}
		},
	},
	{
		name: 'real-dense-prepared-open-cpu',
		category: 'workflow',
		build() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			return {
				rows: target.rowCount,
				cols: target.colCount,
				cells: target.rowCount * target.colCount,
				byteCount: bytes.byteLength,
			}
		},
		async run() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			WorkbookDocument.clearCache()

			const openStart = performance.now()
			const session = await AscendSession.open(bytes, {
				mode: 'interactive',
				prepareEdits: true,
			})
			const preparedOpenMs = performance.now() - openStart
			const readiness = session.editReadiness()
			if (!readiness.ready) throw new Error('Dense prepared-open CPU session was not edit-ready')
			const readModel = workbookDocumentModel(session.workbook())
			const mutableWorkbook = sessionMutableWorkbook(session)
			if (!readModel) throw new Error('Dense prepared-open CPU read model was unavailable')
			if (!mutableWorkbook)
				throw new Error('Dense prepared-open CPU mutable workbook was unavailable')
			const mutableModel = mutableWorkbook.getWorkbookModel()
			session.close()
			WorkbookDocument.clearCache()

			return {
				assertions: {
					bytes: bytes.byteLength,
					preparedOpenMs,
					readinessReady: readiness.ready,
					readinessReusedReadModel: readiness.timings?.mutableWorkbookReusedReadModel ?? false,
					readinessMutableWorkbookOpenMs: readiness.timings?.mutableWorkbookOpenMs ?? null,
					readinessRebaseViewportSnapshotsMs: readiness.timings?.rebaseViewportSnapshotsMs ?? null,
					readAndMutableShareSourceArchive:
						readModel.sourceArchiveBytes === mutableModel.sourceArchiveBytes ? 1 : 0,
					...prefixAssertions(
						'readMutableGridSharing',
						workbookGridSharingAssertions(readModel, mutableModel),
					),
				},
			}
		},
	},
	{
		name: 'real-dense-prepared-open-only-memory',
		category: 'workflow',
		build() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			xlsxSharedStringUsage(bytes)
			return {
				rows: target.rowCount,
				cols: target.colCount,
				cells: target.rowCount * target.colCount,
				byteCount: bytes.byteLength,
			}
		},
		async run() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			const archiveFootprint = xlsxArchiveFootprint(bytes)
			const sharedStringUsage = xlsxSharedStringUsage(bytes)
			WorkbookDocument.clearCache()
			runGc()
			const baseline = phaseMemorySnapshot()

			const openStart = performance.now()
			const session = await AscendSession.open(bytes, {
				mode: 'interactive',
				prepareEdits: true,
			})
			const preparedOpenMs = performance.now() - openStart
			const readiness = session.editReadiness()
			if (!readiness.ready) throw new Error('Dense prepared-open-only session was not edit-ready')
			const readModel = workbookDocumentModel(session.workbook())
			const mutableWorkbook = sessionMutableWorkbook(session)
			if (!readModel) throw new Error('Dense prepared-open-only read model was unavailable')
			if (!mutableWorkbook)
				throw new Error('Dense prepared-open-only mutable workbook was unavailable')
			const mutableModel = mutableWorkbook.getWorkbookModel()
			const mutableOriginalBytes = ascendWorkbookOriginalBytes(mutableWorkbook)
			runGc()
			const afterPreparedOpen = phaseMemorySnapshot()

			session.close()
			WorkbookDocument.clearCache()
			runGc()
			const afterClose = phaseMemorySnapshot()

			return {
				assertions: {
					bytes: bytes.byteLength,
					archivePartCount: archiveFootprint.partCount,
					archiveCompressedBytes: archiveFootprint.compressedBytes,
					archiveUncompressedBytes: archiveFootprint.uncompressedBytes,
					archiveWorksheetUncompressedBytes: archiveFootprint.worksheetUncompressedBytes,
					archiveLargestPartPath: archiveFootprint.largestPartPath,
					archiveLargestPartCompressedBytes: archiveFootprint.largestPartCompressedBytes,
					archiveLargestPartUncompressedBytes: archiveFootprint.largestPartUncompressedBytes,
					preparedOpenMs,
					readinessReady: readiness.ready,
					readinessReusedReadModel: readiness.timings?.mutableWorkbookReusedReadModel ?? false,
					readinessMutableWorkbookOpenMs: readiness.timings?.mutableWorkbookOpenMs ?? null,
					readinessRebaseViewportSnapshotsMs: readiness.timings?.rebaseViewportSnapshotsMs ?? null,
					...prefixAssertions('sharedStrings', sharedStringUsage),
					...prefixAssertions('readModel', workbookModelRetentionAssertions(readModel)),
					...prefixAssertions(
						'mutableWorkbook',
						ascendWorkbookRetentionAssertions(mutableWorkbook),
					),
					readAndMutableShareSourceArchive:
						readModel.sourceArchiveBytes === mutableModel.sourceArchiveBytes ? 1 : 0,
					mutableOriginalSharesSourceArchive:
						mutableOriginalBytes === mutableModel.sourceArchiveBytes ? 1 : 0,
					readSourceArchiveSharesInput: readModel.sourceArchiveBytes === bytes ? 1 : 0,
					...prefixAssertions(
						'readMutableGridSharing',
						workbookGridSharingAssertions(readModel, mutableModel),
					),
					baselineRssBytes: baseline.rssBytes,
					afterPreparedOpenRssDeltaBytes: memoryDelta(
						afterPreparedOpen.rssBytes,
						baseline.rssBytes,
					),
					afterCloseRssDeltaBytes: memoryDelta(afterClose.rssBytes, baseline.rssBytes),
					baselineHeapUsedBytes: baseline.heapUsedBytes,
					afterPreparedOpenHeapDeltaBytes: memoryDelta(
						afterPreparedOpen.heapUsedBytes,
						baseline.heapUsedBytes,
					),
					afterCloseHeapDeltaBytes: memoryDelta(afterClose.heapUsedBytes, baseline.heapUsedBytes),
					baselineExternalBytes: baseline.externalBytes,
					afterPreparedOpenExternalDeltaBytes: memoryDelta(
						afterPreparedOpen.externalBytes,
						baseline.externalBytes,
					),
					afterCloseExternalDeltaBytes: memoryDelta(
						afterClose.externalBytes,
						baseline.externalBytes,
					),
					baselineArrayBuffersBytes: baseline.arrayBuffersBytes,
					afterPreparedOpenArrayBuffersDeltaBytes: memoryDelta(
						afterPreparedOpen.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
					afterCloseArrayBuffersDeltaBytes: memoryDelta(
						afterClose.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
				},
			}
		},
	},
	{
		name: 'real-dense-prepared-open-memory',
		category: 'workflow',
		build() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			return {
				rows: target.rowCount,
				cols: target.colCount,
				cells: target.rowCount * target.colCount,
				byteCount: realInteractivePatchCorpusTargetBytes(target).byteLength,
			}
		},
		async run() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const bytes = realInteractivePatchCorpusTargetBytes(target)
			const archiveFootprint = xlsxArchiveFootprint(bytes)
			const sharedStringUsage = xlsxSharedStringUsage(bytes)
			WorkbookDocument.clearCache()
			runGc()
			const baseline = phaseMemorySnapshot()

			const openStart = performance.now()
			const session = await AscendSession.open(bytes, {
				mode: 'interactive',
				prepareEdits: true,
			})
			const preparedOpenMs = performance.now() - openStart
			const readiness = session.editReadiness()
			if (!readiness.ready) throw new Error('Dense prepared-open session was not edit-ready')
			const readModel = workbookDocumentModel(session.workbook())
			const mutableWorkbook = sessionMutableWorkbook(session)
			if (!readModel) throw new Error('Dense prepared-open read model was unavailable')
			if (!mutableWorkbook) throw new Error('Dense prepared-open mutable workbook was unavailable')
			const mutableModel = mutableWorkbook.getWorkbookModel()
			const mutableOriginalBytes = ascendWorkbookOriginalBytes(mutableWorkbook)
			runGc()
			const afterPreparedOpen = phaseMemorySnapshot()

			const viewportStart = performance.now()
			const viewport = session.readViewport({
				sheet: target.sheet,
				topRow: target.topRow,
				leftCol: target.leftCol,
				rowCount: target.rowCount,
				colCount: target.colCount,
			})
			const viewportMs = performance.now() - viewportStart
			runGc()
			const afterViewport = phaseMemorySnapshot()

			const editStart = performance.now()
			const edit = await session.apply(
				[
					{
						op: 'setCells',
						sheet: target.sheet,
						updates: [{ ref: target.editRef, value: 81_000 }],
					},
				],
				{ recalc: false },
			)
			if (edit.apply.errors.length > 0) throw new Error('Dense prepared-open edit failed')
			const patchStart = performance.now()
			const patch =
				session.readViewportPatch({
					sheet: target.sheet,
					topRow: target.topRow,
					leftCol: target.leftCol,
					rowCount: target.rowCount,
					colCount: target.colCount,
					changedSince: viewport.changeToken,
				}) ??
				session.readViewport({
					sheet: target.sheet,
					topRow: target.topRow,
					leftCol: target.leftCol,
					rowCount: target.rowCount,
					colCount: target.colCount,
					changedSince: viewport.changeToken,
				}).patch
			const patchReadMs = performance.now() - patchStart
			const editFrameMs = performance.now() - editStart
			if (!patch || patch.changedCells.length !== 1) {
				throw new Error('Dense prepared-open patch was not delta-shaped')
			}
			runGc()
			const afterPatch = phaseMemorySnapshot()

			session.close()
			WorkbookDocument.clearCache()
			runGc()
			const afterClose = phaseMemorySnapshot()

			return {
				assertions: {
					bytes: bytes.byteLength,
					archivePartCount: archiveFootprint.partCount,
					archiveCompressedBytes: archiveFootprint.compressedBytes,
					archiveUncompressedBytes: archiveFootprint.uncompressedBytes,
					archiveWorksheetUncompressedBytes: archiveFootprint.worksheetUncompressedBytes,
					archiveLargestPartPath: archiveFootprint.largestPartPath,
					archiveLargestPartCompressedBytes: archiveFootprint.largestPartCompressedBytes,
					archiveLargestPartUncompressedBytes: archiveFootprint.largestPartUncompressedBytes,
					viewportCells: viewport.cells.length,
					viewportFlatValues: viewport.flatValues.length,
					changedCells: patch.changedCells.length,
					patchBytes: patch.byteLength,
					preparedOpenMs,
					readinessReady: readiness.ready,
					readinessReusedReadModel: readiness.timings?.mutableWorkbookReusedReadModel ?? false,
					readinessMutableWorkbookOpenMs: readiness.timings?.mutableWorkbookOpenMs ?? null,
					...prefixAssertions('sharedStrings', sharedStringUsage),
					...prefixAssertions('readModel', workbookModelRetentionAssertions(readModel)),
					...prefixAssertions(
						'mutableWorkbook',
						ascendWorkbookRetentionAssertions(mutableWorkbook),
					),
					readAndMutableShareSourceArchive:
						readModel.sourceArchiveBytes === mutableModel.sourceArchiveBytes ? 1 : 0,
					mutableOriginalSharesSourceArchive:
						mutableOriginalBytes === mutableModel.sourceArchiveBytes ? 1 : 0,
					readSourceArchiveSharesInput: readModel.sourceArchiveBytes === bytes ? 1 : 0,
					...prefixAssertions(
						'readMutableGridSharing',
						workbookGridSharingAssertions(readModel, mutableModel),
					),
					viewportMs,
					editFrameMs,
					sessionApplyMs: edit.timings.totalMs,
					applyMs: edit.timings.applyMs,
					patchReadMs,
					baselineRssBytes: baseline.rssBytes,
					afterPreparedOpenRssDeltaBytes: memoryDelta(
						afterPreparedOpen.rssBytes,
						baseline.rssBytes,
					),
					afterViewportRssDeltaBytes: memoryDelta(afterViewport.rssBytes, baseline.rssBytes),
					afterPatchRssDeltaBytes: memoryDelta(afterPatch.rssBytes, baseline.rssBytes),
					afterCloseRssDeltaBytes: memoryDelta(afterClose.rssBytes, baseline.rssBytes),
					baselineHeapUsedBytes: baseline.heapUsedBytes,
					afterPreparedOpenHeapDeltaBytes: memoryDelta(
						afterPreparedOpen.heapUsedBytes,
						baseline.heapUsedBytes,
					),
					afterViewportHeapDeltaBytes: memoryDelta(
						afterViewport.heapUsedBytes,
						baseline.heapUsedBytes,
					),
					afterPatchHeapDeltaBytes: memoryDelta(afterPatch.heapUsedBytes, baseline.heapUsedBytes),
					afterCloseHeapDeltaBytes: memoryDelta(afterClose.heapUsedBytes, baseline.heapUsedBytes),
					baselineExternalBytes: baseline.externalBytes,
					afterPreparedOpenExternalDeltaBytes: memoryDelta(
						afterPreparedOpen.externalBytes,
						baseline.externalBytes,
					),
					afterViewportExternalDeltaBytes: memoryDelta(
						afterViewport.externalBytes,
						baseline.externalBytes,
					),
					afterPatchExternalDeltaBytes: memoryDelta(
						afterPatch.externalBytes,
						baseline.externalBytes,
					),
					afterCloseExternalDeltaBytes: memoryDelta(
						afterClose.externalBytes,
						baseline.externalBytes,
					),
					baselineArrayBuffersBytes: baseline.arrayBuffersBytes,
					afterPreparedOpenArrayBuffersDeltaBytes: memoryDelta(
						afterPreparedOpen.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
					afterViewportArrayBuffersDeltaBytes: memoryDelta(
						afterViewport.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
					afterPatchArrayBuffersDeltaBytes: memoryDelta(
						afterPatch.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
					afterCloseArrayBuffersDeltaBytes: memoryDelta(
						afterClose.arrayBuffersBytes,
						baseline.arrayBuffersBytes,
					),
					patchMode: 'delta',
				},
			}
		},
	},
	{
		name: 'closed-byte-session-retention-memory',
		category: 'workflow',
		build() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			return {
				rows: target.rowCount,
				cols: target.colCount,
				cells: target.rowCount * target.colCount,
				byteCount: realInteractivePatchCorpusTargetBytes(target).byteLength,
			}
		},
		async run() {
			const target = realInteractivePatchCorpusTarget('stress-dense-100k')
			const sourceBytes = realInteractivePatchCorpusTargetBytes(target)
			const closedSessions: AscendSession[] = []
			const sessionCount = 8
			WorkbookDocument.clearCache()
			runGc()
			const baseline = phaseMemorySnapshot()

			for (let index = 0; index < sessionCount; index++) {
				const bytes = new Uint8Array(sourceBytes)
				const session = await AscendSession.open(bytes, {
					mode: 'interactive',
					prepareEdits: true,
				})
				session.close()
				WorkbookDocument.clearCache()
				closedSessions.push(session)
				runGc()
			}
			runGc()
			const afterClose = phaseMemorySnapshot()
			const arrayBuffersDelta = memoryDelta(
				afterClose.arrayBuffersBytes,
				baseline.arrayBuffersBytes,
			)

			return {
				assertions: {
					bytes: sourceBytes.byteLength,
					closedSessionsRetained: closedSessions.length,
					expectedSourceBytesIfClosedSessionsRetainInput:
						closedSessions.length * sourceBytes.byteLength,
					afterCloseRssDeltaBytes: memoryDelta(afterClose.rssBytes, baseline.rssBytes),
					afterCloseHeapDeltaBytes: memoryDelta(afterClose.heapUsedBytes, baseline.heapUsedBytes),
					afterCloseExternalDeltaBytes: memoryDelta(
						afterClose.externalBytes,
						baseline.externalBytes,
					),
					afterCloseArrayBuffersDeltaBytes: arrayBuffersDelta,
					arrayBuffersDeltaPerClosedSession:
						closedSessions.length > 0 ? arrayBuffersDelta / closedSessions.length : null,
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
		name: 'recalc-sparse-wide-aggregation',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			const random = createDeterministicRandom(0x51a7e)
			for (let r = 0; r < 100_000; r += 100) {
				sheet.cells.set(r, 0, {
					value: numberValue(random() * 1000),
					formula: null,
					styleId: SID,
				})
				sheet.cells.set(r, 99, {
					value: numberValue(random() * 1000),
					formula: null,
					styleId: SID,
				})
			}
			sheet.cells.set(0, 100, { value: EMPTY, formula: 'SUM(A1:CV100000)', styleId: SID })
			return { workbook, rows: 100_000, cols: 101, cells: 2001 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-dense-rectangle-aggregation',
		category: 'calc',
		build() {
			const workbook = buildDenseWorkbook(5000, 20)
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			sheet.cells.set(0, 20, { value: EMPTY, formula: 'SUM(A1:T5000)', styleId: SID })
			return { workbook, rows: 5000, cols: 21, cells: 100_001 }
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
		name: 'structural-insert-large',
		category: 'calc',
		build() {
			const workbook = buildStructuralInsertLargeWorkbook(1000, 100)
			return { workbook, rows: 1000, cols: 101, cells: 100_000 }
		},
		run(input) {
			const workbook = requireWorkbook(input)
			const result = applyOperations(workbook, [
				{ op: 'insertRows', sheet: 'Sheet1', at: 500, count: 100 },
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
	{
		name: 'recalc-single-dirty',
		category: 'calc',
		build() {
			const workbook = buildFormulaChainWorkbook(10_000)
			recalculate(workbook, defaultCalcContext())
			setNumberCell(workbook, 0, 0, 42)
			return { workbook, rows: 10_000, cols: 1, cells: 10_000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext(), {
				dirtyOnly: true,
				dirtyRefs: ['Sheet1!A1'],
			})
		},
	},
	{
		name: 'recalc-volatile',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			for (let r = 0; r < 500; r++) {
				setFormulaCell(workbook, r, 0, 'RAND()')
			}
			for (let r = 0; r < 500; r++) {
				setFormulaCell(workbook, r, 1, 'NOW()')
			}
			return { workbook, rows: 500, cols: 2, cells: 1000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-nested-if',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			for (let r = 0; r < 2000; r++) {
				setNumberCell(workbook, r, 0, (r * 7) % 100)
				setFormulaCell(workbook, r, 1, `IF(A${r + 1}>50,IF(A${r + 1}>75,3,2),IF(A${r + 1}>25,1,0))`)
			}
			return { workbook, rows: 2000, cols: 2, cells: 4000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-shared-formula-group',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			const rows = 100_000
			for (let r = 0; r < rows; r++) {
				sheet.cells.set(r, 1, { value: numberValue(r + 1), formula: null, styleId: SID })
				sheet.cells.set(r, 2, { value: numberValue((r + 1) * 10), formula: null, styleId: SID })
			}
			sheet.cells.set(0, 0, {
				value: EMPTY,
				formula: 'B1+C1',
				styleId: SID,
				formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
			})
			for (let r = 1; r < rows; r++) {
				sheet.cells.set(r, 0, {
					value: EMPTY,
					formula: null,
					styleId: SID,
					formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
				})
			}
			return { workbook, rows, cols: 3, cells: rows * 3 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'recalc-index-match',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			for (let r = 0; r < 10_000; r++) {
				setStringCell(workbook, r, 0, `key-${String(r + 1).padStart(5, '0')}`)
				setNumberCell(workbook, r, 1, (r + 1) * 10)
			}
			for (let r = 0; r < 1000; r++) {
				const keyIndex = 10_000 - ((r * 37) % 10_000) - 1
				setStringCell(workbook, r, 3, `key-${String(keyIndex + 1).padStart(5, '0')}`)
				setFormulaCell(workbook, r, 4, `INDEX(B$1:B$10000,MATCH(D${r + 1},A$1:A$10000,0))`)
			}
			return { workbook, rows: 10_000, cols: 5, cells: 22_000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'apply-operations-batch',
		category: 'workflow',
		build() {
			const workbook = buildDenseWorkbook(500, 20)
			return { workbook, rows: 500, cols: 20, cells: 10_000 }
		},
		run(input) {
			const workbook = requireWorkbook(input)
			const updates: Array<{ ref: string; value: number }> = []
			for (let i = 0; i < 500; i++) {
				const row = ((i * 7) % 500) + 1
				const col = i % 20
				const ref = `${String.fromCharCode(65 + col)}${row}`
				updates.push({ ref, value: i * 10 })
			}
			applyOperations(workbook, [{ op: 'setCells', sheet: 'Sheet1', updates }])
		},
	},
	{
		name: 'formula-compilation-unique',
		category: 'calc',
		build() {
			const workbook = buildFormulaCompilationUniqueWorkbook(5000)
			return { workbook, rows: 5000, cols: 2, cells: 5000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	{
		name: 'parse-formulas-10k',
		category: 'calc',
		build() {
			const formulas: string[] = []
			for (let i = 0; i < 10_000; i++) {
				switch (i % 3) {
					case 0:
						formulas.push(`SUM(A${i + 1}:A${i + 100})`)
						break
					case 1:
						formulas.push(`IF(B${i + 1}>0,C${i + 1}*2,D${i + 1}+1)`)
						break
					default:
						formulas.push(`VLOOKUP(E${i + 1},A$1:D$1000,3,FALSE)`)
						break
				}
			}
			return { content: formulas.join('\n'), rows: 10_000, cols: 1, cells: 10_000 }
		},
		run(input) {
			const formulas = requireContent(input).split('\n')
			clearGlobalParseCache()
			for (const formula of formulas) {
				parseFormula(formula)
			}
		},
	},
	{
		name: 'recalc-1m-dense',
		category: 'calc',
		build() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('Benchmark workbook missing first sheet')
			for (let r = 0; r < 100_000; r++) {
				for (let c = 0; c < 10; c++) {
					sheet.cells.set(r, c, {
						value: numberValue(r * 10 + c + 1),
						formula: null,
						styleId: SID,
					})
				}
			}
			for (let i = 0; i < 1000; i++) {
				const col = i % 10
				const letter = String.fromCharCode(65 + col)
				setFormulaCell(workbook, 100_000 + i, col, `SUM(${letter}1:${letter}100000)`)
			}
			return { workbook, rows: 101_000, cols: 10, cells: 1_001_000 }
		},
		run(input) {
			recalculate(requireWorkbook(input), defaultCalcContext())
		},
	},
	...([1_000, 10_000, 100_000, 1_000_000] as const).map(
		(cellCount): Scenario => ({
			name: `memory-${cellCount >= 1_000_000 ? `${cellCount / 1_000_000}m` : `${cellCount / 1_000}k`}-cells`,
			category: 'workflow',
			build() {
				return { rows: 0, cols: 0, cells: cellCount }
			},
			run() {
				const cols = 10
				const rows = Math.ceil(cellCount / cols)
				runGc()
				const rssBefore = getRssBytes() ?? 0
				const heapBefore = process.memoryUsage().heapUsed
				const workbook = createWorkbook()
				workbook.addSheet('Sheet1')
				const sheet = workbook.sheets[0]
				if (!sheet) throw new Error('Benchmark workbook missing first sheet')
				for (let r = 0; r < rows; r++) {
					for (let c = 0; c < cols; c++) {
						sheet.cells.set(r, c, {
							value: numberValue(r * cols + c + 1),
							formula: null,
							styleId: SID,
						})
					}
				}
				runGc()
				const rssAfter = getRssBytes() ?? 0
				const heapAfter = process.memoryUsage().heapUsed
				const rssDelta = Math.max(0, rssAfter - rssBefore)
				const heapDelta = Math.max(0, heapAfter - heapBefore)
				const bytesPerCell = cellCount > 0 ? rssDelta / cellCount : 0
				return {
					assertions: {
						cellCount,
						rssDeltaBytes: rssDelta,
						heapDeltaBytes: heapDelta,
						bytesPerCell: Math.round(bytesPerCell * 100) / 100,
					},
				}
			},
		}),
	),
]

const scenarioSets = {
	smoke: [
		'read-full-dense',
		'read-values-dense',
		'sdk-window-dense-values-compact-hot',
		'sdk-window-formula-chain-compact-hot',
		'sdk-window-sparse-wide-compact-hot',
		'write-csv-large',
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
	memory: [
		'memory-1k-cells',
		'memory-10k-cells',
		'memory-100k-cells',
		'memory-1m-cells',
		'grid-storage-width-sweep',
		'grid-storage-width-5',
		'grid-storage-width-10',
		'grid-storage-width-20',
		'grid-storage-width-50',
		'grid-storage-mixed-width-10',
		'grid-storage-mixed-width-20',
		'xlsx-grid-storage-width-5',
		'xlsx-grid-storage-width-10',
		'xlsx-grid-storage-width-20',
	],
	'large-read': [
		'read-large-200k',
		'read-style-heavy',
		'read-string-heavy',
		'read-formula-dense',
		'read-large-200k-values',
	],
	'ui-latency': [
		'sdk-viewport-read-dense-hot',
		'sdk-viewport-read-sparse-wide-hot',
		'sdk-semantic-viewport-rich',
		'sdk-semantic-viewport-many-overlays',
		'patch-stream-small-edit',
		'patch-stream-first-edit',
		'patch-stream-prepared-first-edit',
		'patch-stream-real-rich-prepared-first-edit',
	],
	'ui-real-corpus': ['patch-stream-real-corpus-prepared-first-edit'],
	'real-open': ['real-corpus-open-phases', 'real-corpus-grid-storage'],
	'real-memory': [
		'real-dense-prepared-open-cpu',
		'real-dense-promotion-memory',
		'real-dense-prepared-open-only-memory',
		'real-dense-prepared-open-memory',
		'closed-byte-session-retention-memory',
	],
	'real-read-memory': [
		'real-dense-read-memory-metadata',
		'real-dense-read-memory-values',
		'real-dense-read-memory-formula',
		'real-dense-read-memory-full',
	],
	'real-readiness': ['real-dense-progressive-readiness', 'real-dense-interactive-open-modes'],
	'real-shape': ['real-dense-worksheet-shape'],
	'real-write': ['real-dense-dirty-write-lifecycle'],
	'prepared-open': ['real-dense-prepared-open-cpu', 'prepared-open-string-heavy'],
} as const

function phaseMemorySnapshot(): {
	readonly rssBytes: number
	readonly heapUsedBytes: number
	readonly externalBytes: number
	readonly arrayBuffersBytes: number
} {
	const usage = process.memoryUsage()
	return {
		rssBytes: getRssBytes() ?? 0,
		heapUsedBytes: usage.heapUsed,
		externalBytes: usage.external,
		arrayBuffersBytes: usage.arrayBuffers,
	}
}

function memoryDelta(after: number, before: number): number {
	return Math.max(0, after - before)
}

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
		'heap-delta',
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
		result.metrics.heapDeltaBytes !== undefined
			? formatBytes(result.metrics.heapDeltaBytes)
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
		readonly heapDeltaBytes: number
		readonly heapUsedBytes: number
		readonly heapTotalBytes: number
		readonly heapAfterGcBytes: number
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
		const heapBefore = process.memoryUsage().heapUsed
		const start = performance.now()
		const runResult = await scenario.run(input)
		const durationMs = performance.now() - start
		const memAfter = process.memoryUsage()
		const rssAfter = getRssBytes()
		runGc()
		const rssAfterGc = getRssBytes()
		const heapAfterGc = process.memoryUsage().heapUsed
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
			heapDeltaBytes: Math.max(0, memAfter.heapUsed - heapBefore),
			heapUsedBytes: memAfter.heapUsed,
			heapTotalBytes: memAfter.heapTotal,
			heapAfterGcBytes: heapAfterGc,
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
			bytes: input.byteCount ?? input.bytes?.byteLength ?? 0,
			repeat,
		},
		metrics: summarizeSamples(samples),
		reproCommand: commandString([
			'bun',
			'run',
			'fixtures/benchmarks/run.ts',
			'--scenario',
			scenario.name,
			'--repeat',
			String(repeat),
			'--warmup',
			String(warmup),
			'--json',
		]),
		profileCommand: commandString([
			'bun',
			'run',
			'fixtures/benchmarks/profile-bun.ts',
			'--mode',
			'all-md',
			'--label',
			`synthetic-${scenario.name}`,
			'--',
			'bun',
			'run',
			'fixtures/benchmarks/run.ts',
			'--scenario',
			scenario.name,
			'--repeat',
			String(repeat),
			'--warmup',
			String(warmup),
			'--json',
		]),
		...(repeat > 1 ? { samples } : {}),
		...(assertions ? { assertions } : {}),
	}
}

function commandString(args: readonly string[]): string {
	return args.map(shellQuote).join(' ')
}

function shellQuote(value: string): string {
	return /^[A-Za-z0-9_./:=@+-]+$/.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	return index >= 0 ? process.argv[index + 1] : undefined
}

function hasFlag(name: string): boolean {
	return process.argv.includes(name)
}

function renderHelp(): string {
	const scenarioNames = scenarios.map((scenario) => scenario.name).sort()
	const setNames = Object.keys(scenarioSets).sort()
	return [
		'Ascend synthetic benchmark runner',
		'',
		'Usage:',
		'  bun run fixtures/benchmarks/run.ts [--scenario <name> | --set <name>] [--repeat N] [--warmup N] [--json]',
		'  bun run fixtures/benchmarks/run.ts --profile [--set <name>] [--repeat N] [--warmup N]',
		'',
		'Options:',
		'  --scenario <name>  Run one benchmark scenario.',
		'  --set <name>       Run a named scenario set.',
		'  --repeat N         Number of measured samples. Defaults to 1.',
		'  --warmup N         Number of warmup samples. Defaults to 1.',
		'  --isolate-samples  Run each measured sample in a fresh process for memory-sensitive scenarios.',
		'  --json             Emit JSON instead of a text summary.',
		'  --ci               Emit JSON and CI-oriented target metadata.',
		'  --profile          Run the selected set under V8 tracing when using Node.',
		'  --help, -h         Show this help without running benchmarks.',
		'',
		'Scenario sets:',
		...setNames.map((name) => `  ${name}`),
		'',
		'Scenarios:',
		...scenarioNames.map((name) => `  ${name}`),
	].join('\n')
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

async function runScenarioSamplesIsolated(
	scenario: Scenario,
	repeat: number,
	warmup: number,
): Promise<BenchmarkCaseResult> {
	const results: BenchmarkCaseResult[] = []
	for (let i = 0; i < repeat; i++) {
		results.push(await runScenarioIsolated(scenario, 1, warmup, true))
	}
	const first = results[0]
	if (!first) throw new Error(`Synthetic benchmark scenario "${scenario.name}" produced no samples`)
	const samples = results.map((result) => ({
		durationMs: result.metrics.medianMs,
		throughputPerSec: result.metrics.throughputPerSec,
		...(result.metrics.rssDeltaBytes !== undefined
			? { rssDeltaBytes: result.metrics.rssDeltaBytes }
			: {}),
		...(result.metrics.retainedRssDeltaBytes !== undefined
			? { retainedRssDeltaBytes: result.metrics.retainedRssDeltaBytes }
			: {}),
		heapDeltaBytes: result.metrics.heapDeltaBytes ?? 0,
		heapUsedBytes: result.metrics.heapUsedBytes ?? 0,
		heapTotalBytes: result.metrics.heapTotalBytes ?? 0,
		heapAfterGcBytes: result.metrics.heapAfterGcBytes ?? 0,
	}))
	return {
		...first,
		dimensions: { ...first.dimensions, repeat },
		metrics: summarizeSamples(samples),
		reproCommand: commandString([
			'bun',
			'run',
			'fixtures/benchmarks/run.ts',
			'--scenario',
			scenario.name,
			'--repeat',
			String(repeat),
			'--warmup',
			String(warmup),
			'--isolate-samples',
			'--json',
		]),
		...(repeat > 1 ? { samples } : {}),
	}
}

function isNodeRuntime(): boolean {
	return typeof process.versions.bun === 'undefined'
}

async function runWithProfile(
	scenarioSetName: string | undefined,
	repeat: number,
	warmup: number,
): Promise<void> {
	if (!isNodeRuntime()) {
		console.log('V8 profiling is only available under Node.js (not Bun/JSC).')
		console.log('Re-run with: node --trace-deopt --trace-ic fixtures/benchmarks/run.ts ...')
		return
	}
	const setFlag = scenarioSetName ? ['--set', scenarioSetName] : []
	const deoptFile = `deopt-trace-${Date.now()}.log`
	const { execSync } = await import('node:child_process')
	const scriptPath = process.argv[1] ?? import.meta.path
	const cmd = [
		'node',
		'--trace-deopt',
		'--trace-ic',
		scriptPath,
		...setFlag,
		'--repeat',
		String(repeat),
		'--warmup',
		String(warmup),
		'--json',
	]
	console.log(`Running: ${cmd.join(' ')}`)
	console.log(`Capturing deopt/IC output to ${deoptFile}`)
	try {
		const result = execSync(cmd.join(' '), {
			encoding: 'utf-8',
			stdio: ['inherit', 'pipe', 'pipe'],
			maxBuffer: 100 * 1024 * 1024,
		})
		const stderrText = result ?? ''
		const { writeFileSync } = await import('node:fs')
		writeFileSync(deoptFile, stderrText)
		summarizeDeoptLog(stderrText)
	} catch (err: unknown) {
		const execErr = err as { stderr?: string }
		if (execErr.stderr) {
			const { writeFileSync } = await import('node:fs')
			writeFileSync(deoptFile, execErr.stderr)
			summarizeDeoptLog(execErr.stderr)
		}
		console.error('Profile run exited with error')
	}
}

function summarizeDeoptLog(log: string): void {
	const lines = log.split('\n')
	const deoptReasons = new Map<string, number>()
	const megamorphicSites = new Map<string, number>()
	for (const line of lines) {
		if (line.includes('[deoptimize')) {
			const reasonMatch = line.match(/reason: (.+?)(?:\]|$)/)
			const reason = reasonMatch?.[1]?.trim() ?? 'unknown'
			deoptReasons.set(reason, (deoptReasons.get(reason) ?? 0) + 1)
		}
		if (line.includes('megamorphic')) {
			const siteMatch = line.match(/\s(\S+:\d+:\d+)\s/)
			const site = siteMatch?.[1] ?? line.slice(0, 80)
			megamorphicSites.set(site, (megamorphicSites.get(site) ?? 0) + 1)
		}
	}
	console.log('\n--- V8 Deoptimization Summary ---')
	if (deoptReasons.size === 0) {
		console.log('No deoptimizations detected.')
	} else {
		console.log(`Total deopt reasons: ${deoptReasons.size}`)
		const sorted = [...deoptReasons.entries()].sort((a, b) => b[1] - a[1])
		for (const [reason, count] of sorted.slice(0, 20)) {
			console.log(`  ${String(count).padStart(6)} ${reason}`)
		}
	}
	console.log('\n--- Top Megamorphic Call Sites ---')
	if (megamorphicSites.size === 0) {
		console.log('No megamorphic sites detected.')
	} else {
		const sorted = [...megamorphicSites.entries()].sort((a, b) => b[1] - a[1])
		for (const [site, count] of sorted.slice(0, 20)) {
			console.log(`  ${String(count).padStart(6)} ${site}`)
		}
	}
}

async function main(): Promise<void> {
	if (hasFlag('--help') || hasFlag('-h')) {
		console.log(renderHelp())
		return
	}
	const json = process.argv.includes('--json')
	const ci = hasFlag('--ci')
	const profile = hasFlag('--profile')
	const scenarioName = readFlag('--scenario')
	const scenarioSetName = readFlag('--set')
	const repeat = Math.max(1, Number.parseInt(readFlag('--repeat') ?? '1', 10) || 1)
	const warmup = Math.max(0, Number.parseInt(readFlag('--warmup') ?? '1', 10) || 1)
	const isolateSamples = hasFlag('--isolate-samples')
	const outputJson = json || ci
	if (profile) {
		await runWithProfile(scenarioSetName, repeat, warmup)
		return
	}
	if (scenarioName) {
		const scenario = scenarios.find((entry) => entry.name === scenarioName)
		if (!scenario) throw new Error(`Unknown synthetic benchmark scenario "${scenarioName}"`)
		const result = isolateSamples
			? await runScenarioSamplesIsolated(scenario, repeat, warmup)
			: await runScenario(scenario, repeat, warmup)
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
		results.push(
			isolateSamples
				? await runScenarioSamplesIsolated(scenario, repeat, warmup)
				: await runScenarioIsolated(scenario, repeat, warmup, outputJson),
		)
		if (!outputJson && isolateSamples) console.log(`completed ${scenario.name}`)
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
			...(isolateSamples ? { isolateSamples: true } : {}),
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
