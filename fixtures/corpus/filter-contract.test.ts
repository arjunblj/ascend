import { describe, expect, it } from 'bun:test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import type { Workbook } from '@ascend/core'
import { type PreservationCapsule, readXlsx, writeXlsx } from '@ascend/io-xlsx'

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

describe('filter feature contract', () => {
	it('captures POI worksheet top10 and dynamic filter criteria', () => {
		const { workbook } = readFixture(FILTER_FIXTURE_ROOT, 'poi/AutoFilter.xlsx')
		const top10 = workbook.sheets.find((sheet) => sheet.name === 'Top10')?.autoFilter
		const bottom10 = workbook.sheets.find((sheet) => sheet.name === 'Bot10')?.autoFilter
		const average = workbook.sheets.find((sheet) => sheet.name === 'Average')?.autoFilter

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

	it('preserves real OSS color and icon filter columns on round-trip', () => {
		const colorWorkbook = roundTrip(
			readFixture(XLSX_FIXTURE_ROOT, 'libreoffice/autofilter-colors.xlsx'),
		)
		const colorTable = colorWorkbook.sheets[0]?.tables[0]
		expect(colorTable?.autoFilter?.columns).toEqual([{ colId: 0, kind: 'colorFilter', dxfId: 0 }])

		const iconWorkbook = roundTrip(
			readFixture(FILTER_FIXTURE_ROOT, 'poi/ConditionalFormattingSamples.xlsx'),
		)
		const iconTable = iconWorkbook.sheets
			.flatMap((sheet) => sheet.tables)
			.find((table) => table.name === 'Table69')
		expect(iconTable?.autoFilter?.columns).toEqual([
			{ colId: 1, kind: 'iconFilter', iconSet: '3TrafficLights1', iconId: 1 },
		])
	})
})
