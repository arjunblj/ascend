import { createWorkbook, type StyleId, type Workbook } from '../../packages/core/src/index.ts'
import {
	applyOperations,
	defaultCalcContext,
	recalculate,
} from '../../packages/engine/src/index.ts'
import { clearGlobalParseCache, parseFormula } from '../../packages/formulas/src/index.ts'
import { readCsv, writeCsv } from '../../packages/io-csv/src/index.ts'
import { readXlsx, writeXlsx, writeXlsxStreaming } from '../../packages/io-xlsx/src/index.ts'
import {
	booleanValue,
	dateValue,
	EMPTY,
	numberValue,
	stringValue,
} from '../../packages/schema/src/index.ts'
import { AscendSession, AscendWorkbook, SheetHandle } from '../../packages/sdk/src/index.ts'
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
let cachedDenseViewportBytes: Uint8Array | undefined
let cachedSparseWideViewportBytes: Uint8Array | undefined
let cachedSemanticViewportBytes: Uint8Array | undefined

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

function interactiveSessionFor(bytes: Uint8Array): Promise<AscendSession> {
	const cached = interactiveSessionCache.get(bytes)
	if (cached) return cached
	const session = AscendSession.open(bytes, { mode: 'interactive' })
	interactiveSessionCache.set(bytes, session)
	return session
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
	memory: ['memory-1k-cells', 'memory-10k-cells', 'memory-100k-cells', 'memory-1m-cells'],
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
			bytes: input.bytes?.byteLength ?? 0,
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
	const outputJson = json || ci
	if (profile) {
		await runWithProfile(scenarioSetName, repeat, warmup)
		return
	}
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
