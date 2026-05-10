import { describe, expect, it } from 'bun:test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import {
	DEFAULT_STYLE_ID,
	type Sheet,
	type Table,
	type Workbook,
} from '../../packages/core/src/index.ts'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import { type PreservationCapsule, readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { EMPTY, numberValue } from '../../packages/schema/src/index.ts'

const FILTER_FIXTURE_ROOT = resolve(import.meta.dir, '../xlsx/filter')
const XLSX_FIXTURE_ROOT = resolve(import.meta.dir, '../xlsx')

interface ReadFixtureResult {
	readonly workbook: Workbook
	readonly capsules: readonly PreservationCapsule[]
}

function readFixture(root: string, path: string): ReadFixtureResult {
	const result = readXlsx(readFileSync(resolve(root, path)))
	expect(result.ok).toBe(true)
	if (!result.ok) throw result.error
	return result.value
}

function roundTrip(fixture: ReadFixtureResult): Workbook {
	const written = writeXlsx(fixture.workbook, fixture.capsules)
	expect(written.ok).toBe(true)
	if (!written.ok) throw written.error
	const reread = readXlsx(written.value)
	expect(reread.ok).toBe(true)
	if (!reread.ok) throw reread.error
	return reread.value.workbook
}

function sheetByName(workbook: Workbook, name: string): Sheet {
	const sheet = workbook.sheets.find((entry) => entry.name === name)
	expect(sheet).toBeDefined()
	if (!sheet) throw new Error(`Missing sheet ${name}`)
	return sheet
}

function tableByName(workbook: Workbook, name: string): Table {
	const table = workbook.sheets
		.flatMap((sheet) => sheet.tables)
		.find((entry) => entry.name === name)
	expect(table).toBeDefined()
	if (!table) throw new Error(`Missing table ${name}`)
	return table
}

describe('filter feature contract', () => {
	it('captures POI worksheet top10 and dynamic filter criteria', () => {
		const { workbook } = readFixture(FILTER_FIXTURE_ROOT, 'poi/AutoFilter.xlsx')
		const top10 = sheetByName(workbook, 'Top10').autoFilter
		const bottom10 = sheetByName(workbook, 'Bot10').autoFilter
		const average = sheetByName(workbook, 'Average').autoFilter

		expect(top10?.columns).toContainEqual({ colId: 0, kind: 'top10', val: 10, filterVal: 2 })
		expect(bottom10?.columns).toContainEqual({
			colId: 0,
			kind: 'top10',
			top: false,
			val: 10,
			filterVal: 2,
		})
		expect(average?.columns).toEqual([
			{
				colId: 0,
				kind: 'dynamicFilter',
				dynamicFilterType: 'belowAverage',
				dynamicFilterVal: 2.380952380952381,
			},
			{
				colId: 4,
				kind: 'dynamicFilter',
				dynamicFilterType: 'aboveAverage',
				dynamicFilterVal: 2.4285714285714284,
			},
		])
	})

	it('captures POI worksheet value and custom filter criteria after round-trip', () => {
		const workbook = roundTrip(readFixture(FILTER_FIXTURE_ROOT, 'poi/AutoFilter.xlsx'))

		expect(sheetByName(workbook, 'One Cond').autoFilter).toEqual({
			ref: 'A1:E22',
			columns: [{ colId: 0, kind: 'filters', values: ['1'] }],
		})
		expect(sheetByName(workbook, 'Two Cond').autoFilter).toEqual({
			ref: 'A1:E22',
			columns: [
				{ colId: 0, kind: 'filters', values: ['1'] },
				{ colId: 4, kind: 'filters', values: ['3'] },
			],
		})
		expect(sheetByName(workbook, 'NE').autoFilter).toEqual({
			ref: 'A1:E22',
			columns: [
				{
					colId: 0,
					kind: 'customFilters',
					customFilters: [{ val: '3', operator: 'notEqual' }],
				},
			],
		})
		expect(sheetByName(workbook, 'GT').autoFilter).toEqual({
			ref: 'A1:E22',
			columns: [
				{
					colId: 0,
					kind: 'customFilters',
					customFilters: [{ val: '2', operator: 'greaterThan' }],
				},
			],
		})
		expect(sheetByName(workbook, 'AND Bounding').autoFilter).toEqual({
			ref: 'A1:E22',
			columns: [
				{
					colId: 0,
					kind: 'customFilters',
					and: true,
					customFilters: [
						{ val: '2', operator: 'greaterThanOrEqual' },
						{ val: '4', operator: 'lessThanOrEqual' },
					],
				},
			],
		})
		expect(sheetByName(workbook, 'OR Range').autoFilter).toEqual({
			ref: 'A1:E22',
			columns: [
				{
					colId: 0,
					kind: 'customFilters',
					customFilters: [
						{ val: '4', operator: 'greaterThanOrEqual' },
						{ val: '1', operator: 'lessThanOrEqual' },
					],
				},
			],
		})
	})

	it('captures POI table filter and sort states after round-trip', () => {
		const workbook = roundTrip(
			readFixture(FILTER_FIXTURE_ROOT, 'poi/ConditionalFormattingSamples.xlsx'),
		)

		expect(tableByName(workbook, 'Table1')).toMatchObject({
			autoFilter: {
				ref: 'A3:E25',
				columns: [
					{
						colId: 2,
						kind: 'customFilters',
						customFilters: [{ val: '1000', operator: 'greaterThan' }],
					},
				],
			},
			sortState: { ref: 'A3:E24', conditions: [{ ref: 'A2:A24' }] },
		})
		expect(tableByName(workbook, 'Table3')).toMatchObject({
			autoFilter: { ref: 'A2:F21', columns: [] },
			sortState: { ref: 'A2:G22', conditions: [{ ref: 'A1:A22' }] },
		})
		expect(sheetByName(workbook, 'Mountains').sortState).toEqual({
			ref: 'A3:I24',
			conditions: [{ ref: 'A2' }],
		})
	})

	it('preserves real OSS color and icon filter columns on round-trip', () => {
		const colorWorkbook = roundTrip(
			readFixture(XLSX_FIXTURE_ROOT, 'libreoffice/autofilter-colors.xlsx'),
		)
		const colorTable = colorWorkbook.sheets[0]?.tables[0]
		expect(colorTable?.autoFilter?.columns).toEqual([{ colId: 0, kind: 'colorFilter', dxfId: 0 }])

		const iconWorkbook = roundTrip(
			readFixture(FILTER_FIXTURE_ROOT, 'poi/ConditionalFormattingSamples.xlsx'),
		)
		const iconTable = tableByName(iconWorkbook, 'Table69')
		expect(iconTable.autoFilter?.columns).toEqual([
			{ colId: 1, kind: 'iconFilter', iconSet: '3TrafficLights1', iconId: 1 },
		])
	})

	it('captures ClosedXML custom auto-filters and sort state after round-trip', () => {
		const workbook = roundTrip(
			readFixture(XLSX_FIXTURE_ROOT, 'closedxml/AutoFilter_CustomAutoFilter.xlsx'),
		)

		expect(sheetByName(workbook, 'Single Column Numbers').autoFilter).toEqual({
			ref: 'A1:A7',
			columns: [
				{
					colId: 0,
					kind: 'customFilters',
					customFilters: [{ val: '3' }, { val: '4', operator: 'greaterThan' }],
				},
			],
			sortState: { ref: 'A2:A7', conditions: [{ ref: 'A1:A7', descending: true }] },
		})
		expect(sheetByName(workbook, 'Single Column Strings').autoFilter).toEqual({
			ref: 'A1:A7',
			columns: [
				{
					colId: 0,
					kind: 'customFilters',
					and: true,
					customFilters: [
						{ val: 'B', operator: 'greaterThanOrEqual' },
						{ val: 'D', operator: 'lessThanOrEqual' },
					],
				},
			],
			sortState: { ref: 'A2:A7', conditions: [{ ref: 'A1:A7', descending: true }] },
		})
		expect(sheetByName(workbook, 'Single Column Mixed').autoFilter).toEqual({
			ref: 'A1:A7',
			columns: [
				{
					colId: 0,
					kind: 'customFilters',
					customFilters: [{ val: 'D', operator: 'greaterThanOrEqual' }],
				},
			],
			sortState: { ref: 'A2:A7', conditions: [{ ref: 'A1:A7', descending: true }] },
		})
		expect(sheetByName(workbook, 'Multi Column').autoFilter).toEqual({
			ref: 'A1:C7',
			columns: [
				{
					colId: 1,
					kind: 'customFilters',
					customFilters: [{ val: '3' }, { val: '4', operator: 'greaterThan' }],
				},
				{ colId: 2, kind: 'customFilters', customFilters: [{ val: 'E' }] },
			],
			sortState: { ref: 'A2:C7', conditions: [{ ref: 'C1:C7', descending: true }] },
		})
	})

	it('captures LibreOffice worksheet filters and sort state after round-trip', () => {
		const customFilterWorkbook = roundTrip(
			readFixture(XLSX_FIXTURE_ROOT, 'libreoffice/autofilter.xlsx'),
		)
		expect(sheetByName(customFilterWorkbook, 'Sheet1').autoFilter).toEqual({
			ref: 'A1:C5',
			columns: [
				{
					colId: 2,
					kind: 'customFilters',
					and: true,
					customFilters: [{ val: '4', operator: 'equal' }],
				},
			],
		})

		const top10Workbook = roundTrip(
			readFixture(XLSX_FIXTURE_ROOT, 'libreoffice/tdf143068_top10filter.xlsx'),
		)
		expect(sheetByName(top10Workbook, 'Munka1').autoFilter).toEqual({
			ref: 'A1:A11',
			columns: [{ colId: 0, kind: 'top10', val: 4, filterVal: 7 }],
		})

		const sortWorkbook = roundTrip(
			readFixture(XLSX_FIXTURE_ROOT, 'libreoffice/sortconditionref2.xlsx'),
		)
		expect(sheetByName(sortWorkbook, 'Tabelle1').autoFilter).toEqual({
			ref: 'A10:B300',
			columns: [],
			sortState: { ref: 'A11:B13', conditions: [{ ref: 'A10:A300' }] },
		})
	})

	it('captures ExcelJS table filter criteria after round-trip', () => {
		const workbook = roundTrip(readFixture(XLSX_FIXTURE_ROOT, 'exceljs/test-issue-1669.xlsx'))

		expect(tableByName(workbook, 'Table1').autoFilter).toEqual({
			ref: 'A1:B6',
			columns: [
				{
					colId: 0,
					kind: 'customFilters',
					customFilters: [{ val: '4', operator: 'notEqual' }],
				},
			],
		})
		expect(tableByName(workbook, 'Table2').autoFilter).toEqual({
			ref: 'A1:B6',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					values: ['T123456789', 'T123456791', 'T123456793'],
				},
			],
		})
	})

	it('evaluates real LibreOffice color filters from style criteria', () => {
		const { workbook } = readFixture(XLSX_FIXTURE_ROOT, 'libreoffice/autofilter-colors.xlsx')
		const sheet = workbook.sheets[0]
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.rowDefs.clear()
		sheet.cells.set(0, 3, {
			value: EMPTY,
			formula: 'SUBTOTAL(9,A2:A11)',
			styleId: DEFAULT_STYLE_ID,
		})

		const result = recalculate(workbook, defaultCalcContext())

		expect(result.errors).toEqual([])
		expect(sheet.cells.get(0, 3)?.value).toEqual(numberValue(5))
	})
})
