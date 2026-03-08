import { describe, expect, it } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { Workbook } from '@ascend/core'
import { booleanValue, numberValue, stringValue } from '@ascend/schema'
import { unzipSync } from 'fflate'
import type { PreservationCapsule } from '../preserve.ts'
import { readXlsx } from '../reader/index.ts'
import { writeXlsx } from './index.ts'

const S0 = 0 as StyleId

function roundTrip(wb: Workbook, capsules?: PreservationCapsule[]) {
	const written = writeXlsx(wb, capsules)
	if (!written.ok) throw new Error(`write failed: ${written.error.message}`)
	const read = readXlsx(written.value)
	if (!read.ok) throw new Error(`read failed: ${read.error.message}`)
	return { bytes: written.value, result: read.value }
}

describe('writeXlsx', () => {
	it('round-trips cell values correctly', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, { value: stringValue('Hello'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(42), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: booleanValue(true), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(3.14), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: booleanValue(false), formula: null, styleId: S0 })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s).toBeDefined()
		expect(s?.name).toBe('Test')
		expect(s?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'Hello' })
		expect(s?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 42 })
		expect(s?.cells.get(0, 2)?.value).toEqual({ kind: 'boolean', value: true })
		expect(s?.cells.get(1, 0)?.value).toEqual({ kind: 'number', value: 3.14 })
		expect(s?.cells.get(1, 1)?.value).toEqual({ kind: 'boolean', value: false })
	})

	it('round-trips multiple sheets', () => {
		const wb = new Workbook()
		const s1 = wb.addSheet('First')
		s1.cells.set(0, 0, { value: stringValue('A'), formula: null, styleId: S0 })
		const s2 = wb.addSheet('Second')
		s2.cells.set(0, 0, { value: numberValue(99), formula: null, styleId: S0 })

		const { result } = roundTrip(wb)

		expect(result.workbook.sheets).toHaveLength(2)
		expect(result.workbook.sheets[0]?.name).toBe('First')
		expect(result.workbook.sheets[1]?.name).toBe('Second')
		expect(result.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'A',
		})
		expect(result.workbook.sheets[1]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 99,
		})
	})

	it('preserves formula text on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: 'A1*2', styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(30), formula: 'SUM(A1,B1)', styleId: S0 })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.cells.get(0, 1)?.formula).toBe('A1*2')
		expect(s?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 20 })
		expect(s?.cells.get(1, 0)?.formula).toBe('SUM(A1,B1)')
	})

	it('preserves bold style on round-trip', () => {
		const wb = new Workbook()
		const boldId = wb.styles.register({ font: { bold: true } })
		const sheet = wb.addSheet('Styled')
		sheet.cells.set(0, 0, { value: stringValue('Bold'), formula: null, styleId: boldId })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		const cell = s?.cells.get(0, 0)
		expect(cell).toBeDefined()
		const style = result.workbook.styles.get(cell?.styleId ?? (0 as StyleId))
		expect(style?.font?.bold).toBe(true)
	})

	it('preserves number format on round-trip', () => {
		const wb = new Workbook()
		const pctId = wb.styles.register({ numberFormat: '0.00%' })
		const sheet = wb.addSheet('Fmt')
		sheet.cells.set(0, 0, { value: numberValue(0.75), formula: null, styleId: pctId })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		const cell = s?.cells.get(0, 0)
		expect(cell).toBeDefined()
		const style = result.workbook.styles.get(cell?.styleId ?? (0 as StyleId))
		expect(style?.numberFormat).toBe('0.00%')
	})

	it('preserves merges on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Merges')
		sheet.cells.set(0, 0, { value: stringValue('Merged'), formula: null, styleId: S0 })
		sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 1, col: 2 } })
		sheet.merges.push({ start: { row: 3, col: 0 }, end: { row: 3, col: 1 } })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.merges).toHaveLength(2)
		expect(s?.merges[0]).toEqual({ start: { row: 0, col: 0 }, end: { row: 1, col: 2 } })
		expect(s?.merges[1]).toEqual({ start: { row: 3, col: 0 }, end: { row: 3, col: 1 } })
	})

	it('preserves defined names on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.definedNames.set('Total', 'Data!$A$1')

		const { result } = roundTrip(wb)
		expect(result.workbook.definedNames.get('Total')).toBe('Data!$A$1')
	})

	it('preserves sheet-scoped defined names on round-trip', () => {
		const wb = new Workbook()
		const data = wb.addSheet('Data')
		const summary = wb.addSheet('Summary')
		data.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		summary.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })
		wb.definedNames.set('Budget', 'Summary!$A$1', { kind: 'sheet', sheetId: summary.id })

		const { result } = roundTrip(wb)
		const resolved = result.workbook.definedNames.resolve(
			'Budget',
			result.workbook.getSheet('Summary')?.id,
		)
		expect(resolved?.scope.kind).toBe('sheet')
		expect(resolved?.formula).toBe('Summary!$A$1')
	})

	it('preserves capsule parts through write-read cycle', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })

		const capsuleContent = new TextEncoder().encode('<chart>test chart data</chart>')
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/charts/chart1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
				relationships: [],
				content: capsuleContent,
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
			},
		]

		const { bytes, result } = roundTrip(wb, capsules)

		const entries = unzipSync(bytes)
		expect(entries['xl/charts/chart1.xml']).toBeDefined()

		const decoded = new TextDecoder().decode(entries['xl/charts/chart1.xml'])
		expect(decoded).toBe('<chart>test chart data</chart>')

		const readCapsules = result.capsules
		const chart = readCapsules.find((c) => c.partPath === 'xl/charts/chart1.xml')
		expect(chart).toBeDefined()
		expect(chart?.contentType).toContain('chart')
		expect(new TextDecoder().decode(chart?.content)).toBe('<chart>test chart data</chart>')
	})

	it('produces a valid ZIP file', () => {
		const wb = new Workbook()
		wb.addSheet('Empty')

		const written = writeXlsx(wb)
		expect(written.ok).toBe(true)
		if (!written.ok) return

		const entries = unzipSync(written.value)
		expect(entries['[Content_Types].xml']).toBeDefined()
		expect(entries['_rels/.rels']).toBeDefined()
		expect(entries['xl/workbook.xml']).toBeDefined()
		expect(entries['xl/_rels/workbook.xml.rels']).toBeDefined()
		expect(entries['xl/styles.xml']).toBeDefined()
		expect(entries['xl/worksheets/sheet1.xml']).toBeDefined()
	})
})
