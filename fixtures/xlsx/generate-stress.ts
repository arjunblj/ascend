#!/usr/bin/env bun
/**
 * Generates synthetic XLSX fixtures for stress testing the Ascend spreadsheet engine.
 * Run with: bun run fixtures/xlsx/generate-stress.ts
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
	createWorkbook,
	indexToColumn,
	type StyleId,
	type Workbook,
} from '../../packages/core/src/index.ts'
import { writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { dateValue, EMPTY, numberValue, stringValue } from '../../packages/schema/src/index.ts'

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url))
const STRESS_DIR = join(SCRIPT_DIR, 'stress')

const DEFAULT_STYLE_ID = 0 as StyleId

function mustWrite(workbook: Workbook): Uint8Array {
	const result = writeXlsx(workbook)
	if (!result.ok) throw new Error(result.error.message)
	return result.value
}

function writeFixture(name: string, workbook: Workbook): void {
	const path = join(STRESS_DIR, name)
	const bytes = mustWrite(workbook)
	writeFileSync(path, bytes)
	console.log(`Wrote ${path} (${(bytes.length / 1024).toFixed(1)} KB)`)
}

function setFormulaCell(
	workbook: Workbook,
	sheetName: string,
	row: number,
	col: number,
	formula: string,
	styleId: StyleId = DEFAULT_STYLE_ID,
): void {
	const sheet = workbook.getSheet(sheetName)
	if (!sheet) throw new Error(`Sheet ${sheetName} not found`)
	sheet.cells.set(row, col, {
		value: EMPTY,
		formula,
		styleId,
	})
}

function main(): void {
	mkdirSync(STRESS_DIR, { recursive: true })

	// 1. dense-100k.xlsx - 100,000 rows x 5 cols of mixed data
	{
		const wb = createWorkbook()
		wb.addSheet('Data')
		const sheet = wb.sheets[0] as NonNullable<(typeof wb.sheets)[0]>
		const rows = 100_000
		const cols = 5
		const BATCH = 5000
		for (let r = 0; r < rows; r += BATCH) {
			for (let rr = r; rr < Math.min(r + BATCH, rows); rr++) {
				for (let c = 0; c < cols; c++) {
					const kind = (rr + c) % 3
					if (kind === 0) {
						sheet.cells.set(rr, c, {
							value: numberValue(rr * cols + c + 1),
							formula: null,
							styleId: DEFAULT_STYLE_ID,
						})
					} else if (kind === 1) {
						sheet.cells.set(rr, c, {
							value: stringValue(`row-${rr}-col-${c}`),
							formula: null,
							styleId: DEFAULT_STYLE_ID,
						})
					} else {
						sheet.cells.set(rr, c, {
							value: dateValue(44927 + (rr % 365)), // Excel serial for dates
							formula: null,
							styleId: DEFAULT_STYLE_ID,
						})
					}
				}
			}
			if ((r / BATCH) % 4 === 0) process.stdout.write(`dense-100k: ${r}/${rows} rows\r`)
		}
		console.log(`dense-100k: ${rows} rows done`)
		writeFixture('dense-100k.xlsx', wb)
	}

	// 2. many-styles.xlsx - 1000 rows with 500+ unique style combinations
	{
		const wb = createWorkbook()
		wb.addSheet('Styled')
		const sheet = wb.sheets[0] as NonNullable<(typeof wb.sheets)[0]>
		const rows = 1000
		const fonts = [
			'Arial',
			'Calibri',
			'Helvetica',
			'Times New Roman',
			'Courier New',
			'Georgia',
			'Verdana',
		]
		const formats = ['General', '0.00', '#,##0', '0.0%', '[$-409]m/d/yy', '0.00E+00', '@']
		for (let r = 0; r < rows; r++) {
			const styleId = wb.styles.register({
				font: {
					name: fonts[r % fonts.length],
					size: 9 + (r % 8),
					bold: r % 5 === 0,
					italic: r % 7 === 1,
					color: {
						kind: 'rgb',
						rgb: `FF${String((r * 37) % 256).padStart(2, '0')}${String((r * 17) % 256).padStart(2, '0')}${String((r * 7) % 256).padStart(2, '0')}`,
					},
				},
				fill: {
					pattern: 'solid',
					fgColor: {
						kind: 'rgb',
						rgb: `FF${String((r * 13) % 256).padStart(2, '0')}${String((r * 31) % 256).padStart(2, '0')}${String((r * 19) % 256).padStart(2, '0')}`,
					},
				},
				numberFormat: formats[r % formats.length],
			}) as StyleId
			sheet.cells.set(r, 0, {
				value: numberValue(r + 1),
				formula: null,
				styleId,
			})
			sheet.cells.set(r, 1, {
				value: stringValue(`styled-${r}`),
				formula: null,
				styleId,
			})
		}
		console.log(`many-styles: ${wb.styles.size} unique styles`)
		writeFixture('many-styles.xlsx', wb)
	}

	// 3. many-strings.xlsx - 10,000 rows with 5,000+ unique string values
	{
		const wb = createWorkbook()
		wb.addSheet('Strings')
		const sheet = wb.sheets[0] as NonNullable<(typeof wb.sheets)[0]>
		const rows = 10_000
		const uniqueCount = 5_500
		for (let r = 0; r < rows; r++) {
			const idx = r % uniqueCount
			sheet.cells.set(r, 0, {
				value: stringValue(`unique-string-${idx}-${'x'.repeat((idx % 20) + 1)}`),
				formula: null,
				styleId: DEFAULT_STYLE_ID,
			})
			sheet.cells.set(r, 1, {
				value: stringValue(`another-${idx}-${String(r).padStart(5, '0')}`),
				formula: null,
				styleId: DEFAULT_STYLE_ID,
			})
		}
		console.log(`many-strings: ${rows} rows, ${uniqueCount}+ unique strings`)
		writeFixture('many-strings.xlsx', wb)
	}

	// 4. formula-dense.xlsx - 5,000 rows where every cell has a formula
	{
		const wb = createWorkbook()
		wb.addSheet('Formulas')
		const sheet = wb.sheets[0] as NonNullable<(typeof wb.sheets)[0]>
		const rows = 5_000
		for (let r = 0; r < rows; r++) {
			const row1 = r + 1
			sheet.cells.set(r, 0, {
				value: numberValue(row1),
				formula: null,
				styleId: DEFAULT_STYLE_ID,
			})
			setFormulaCell(wb, 'Formulas', r, 1, `=A${row1}*2`)
			setFormulaCell(wb, 'Formulas', r, 2, `=SUM(A$1:A${row1})`)
			setFormulaCell(wb, 'Formulas', r, 3, `=B${row1}+C${row1}`)
			setFormulaCell(wb, 'Formulas', r, 4, `=IF(A${row1}>100,"big","small")`)
		}
		console.log(`formula-dense: ${rows} rows with formulas`)
		writeFixture('formula-dense.xlsx', wb)
	}

	// 5. merged-complex.xlsx - complex merge patterns
	{
		const wb = createWorkbook()
		wb.addSheet('Merged')
		const sheet = wb.sheets[0] as NonNullable<(typeof wb.sheets)[0]>
		// Populate some cells
		for (let r = 0; r < 50; r++) {
			for (let c = 0; c < 10; c++) {
				sheet.cells.set(r, c, {
					value: stringValue(`${indexToColumn(c)}${r + 1}`),
					formula: null,
					styleId: DEFAULT_STYLE_ID,
				})
			}
		}
		// Nested/overlapping-ish merge pattern: multiple non-overlapping merges
		sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 2, col: 3 } }) // A1:D3
		sheet.merges.push({ start: { row: 0, col: 5 }, end: { row: 4, col: 7 } }) // F1:H5
		sheet.merges.push({ start: { row: 5, col: 0 }, end: { row: 5, col: 2 } }) // A6:C6
		sheet.merges.push({ start: { row: 6, col: 0 }, end: { row: 10, col: 0 } }) // A7:A11 (tall narrow)
		sheet.merges.push({ start: { row: 12, col: 1 }, end: { row: 15, col: 4 } }) // B13:E16
		sheet.merges.push({ start: { row: 17, col: 2 }, end: { row: 20, col: 5 } }) // C18:F21
		sheet.merges.push({ start: { row: 22, col: 0 }, end: { row: 25, col: 3 } }) // A23:D26
		sheet.merges.push({ start: { row: 27, col: 4 }, end: { row: 30, col: 9 } }) // E28:J31 (large irregular)
		sheet.merges.push({ start: { row: 32, col: 0 }, end: { row: 32, col: 9 } }) // A33:J33 (full row)
		sheet.merges.push({ start: { row: 34, col: 0 }, end: { row: 39, col: 0 } }) // A35:A40
		console.log(`merged-complex: ${sheet.merges.length} merge regions`)
		writeFixture('merged-complex.xlsx', wb)
	}

	// 6. multi-sheet-10.xlsx - 10 sheets, each 1000 rows, cross-sheet references
	{
		const wb = createWorkbook()
		const sheetCount = 10
		const rowsPerSheet = 1000
		for (let s = 0; s < sheetCount; s++) {
			const sheet = wb.addSheet(`Sheet${s + 1}`)
			for (let r = 0; r < rowsPerSheet; r++) {
				sheet.cells.set(r, 0, {
					value: numberValue((s + 1) * (r + 1)),
					formula: null,
					styleId: DEFAULT_STYLE_ID,
				})
			}
		}
		// Add cross-sheet formulas on first sheet
		const first = wb.getSheet('Sheet1') as NonNullable<ReturnType<typeof wb.getSheet>>
		for (let r = 0; r < 100; r++) {
			const row1 = r + 1
			first.cells.set(r, 1, {
				value: EMPTY,
				formula: `=Sheet2!A${row1}+Sheet3!A${row1}`,
				styleId: DEFAULT_STYLE_ID,
			})
		}
		for (let r = 100; r < 200; r++) {
			const row1 = r + 1
			first.cells.set(r, 1, {
				value: EMPTY,
				formula: `=SUM(Sheet1!A1:A${row1})+Sheet5!A${row1}`,
				styleId: DEFAULT_STYLE_ID,
			})
		}
		console.log(`multi-sheet-10: ${sheetCount} sheets, ${rowsPerSheet} rows each`)
		writeFixture('multi-sheet-10.xlsx', wb)
	}

	console.log('Done. Generated 6 stress fixtures in', STRESS_DIR)
}

main()
