import { describe, expect, test } from 'bun:test'
import { scanWorksheetXmlBytes, scanWorksheetXmlStructure } from './xlsx-xml-tokenizer.ts'

const encoder = new TextEncoder()

describe('xlsx XML tokenizer benchmark scanner', () => {
	test('counts worksheet rows, cells, values, formulas, and cell types', () => {
		const scan = scanWorksheetXmlStructure(`<worksheet><sheetData>
			<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>
			<row r="2"><c r="A2" t="inlineStr"><is><t>x</t></is></c><c r="B2"><f>A1</f><v>1</v></c></row>
		</sheetData></worksheet>`)

		expect(scan.rows).toBe(2)
		expect(scan.cells).toBe(4)
		expect(scan.values).toBe(3)
		expect(scan.formulas).toBe(1)
		expect(scan.inlineStrings).toBe(1)
		expect(scan.sharedStringCells).toBe(1)
		expect(scan.numericCells).toBe(2)
		expect(scan.maxCol).toBe(2)
		expect(scan.lastRow).toBe(2)
		expect(scan.checksum).not.toBe(0)
		expect(scanWorksheetXmlBytes(encoder.encode(xmlSummaryFixture()))).toEqual(
			scanWorksheetXmlStructure(xmlSummaryFixture()),
		)
	})

	test('returns an empty summary when sheetData is absent', () => {
		const scan = scanWorksheetXmlStructure('<worksheet/>')

		expect(scan.rows).toBe(0)
		expect(scan.cells).toBe(0)
		expect(scan.values).toBe(0)
		expect(scan.checksum).toBe(0)
		expect(scanWorksheetXmlBytes(encoder.encode('<worksheet/>'))).toEqual(scan)
	})

	test('byte scanner matches string scanner for self-closing worksheet nodes', () => {
		const xml = `<worksheet><sheetData>
			<row r="1"><c r="A1"/><c r="C1" t="str"/></row>
			<row r="3"/>
			<row r="4"><c r="AA4"><v>10</v></c></row>
		</sheetData></worksheet>`

		expect(scanWorksheetXmlBytes(encoder.encode(xml))).toEqual(scanWorksheetXmlStructure(xml))
	})
})

function xmlSummaryFixture(): string {
	return `<worksheet><sheetData>
			<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>
			<row r="2"><c r="A2" t="inlineStr"><is><t>x</t></is></c><c r="B2"><f>A1</f><v>1</v></c></row>
		</sheetData></worksheet>`
}
