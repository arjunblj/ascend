import { describe, expect, test } from 'bun:test'
import { advancedFilterSparklineWorkbook } from '../../io-xlsx/src/reader/advanced-filter-sparkline.test.ts'
import { AscendWorkbook } from './index.ts'

describe('advanced filter and sparkline SDK inventory', () => {
	test('inspectSheet exposes custom sheet view filters and sparkline groups', async () => {
		const wb = await AscendWorkbook.open(advancedFilterSparklineWorkbook())
		const info = wb.inspect()
		const sheet = wb.inspectSheet('Data')

		expect(info.advancedFilterCount).toBe(1)
		expect(info.sparklineGroupCount).toBe(1)
		expect(info.sheets[0]?.advancedFilterCount).toBe(1)
		expect(info.sheets[0]?.sparklineGroupCount).toBe(1)
		expect(sheet?.advancedFilters?.[0]).toMatchObject({
			viewName: 'WestOnly',
			ref: 'A1:C20',
			filterColumnCount: 1,
			sortConditionCount: 1,
		})
		expect(sheet?.sparklineGroups?.[0]).toMatchObject({
			type: 'line',
			range: 'Data!B2:B4',
			locationRange: 'D2:D4',
			count: 1,
			markers: true,
		})
	})

	test('setSparklineGroup edits source and display flags through SDK save and reopen', async () => {
		const wb = await AscendWorkbook.open(advancedFilterSparklineWorkbook())
		const applied = wb.apply([
			{
				op: 'setSparklineGroup',
				sheet: 'Data',
				groupIndex: 0,
				range: 'Data!C2:C4',
				locationRange: 'E2:E4',
				type: 'column',
				markers: false,
				highPoint: false,
				displayXAxis: false,
			},
		])
		expect(applied.errors).toEqual([])

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(reopened.inspectSheet('Data')?.sparklineGroups?.[0]).toMatchObject({
			type: 'column',
			range: 'Data!C2:C4',
			locationRange: 'E2:E4',
			count: 1,
			markers: false,
			highPoint: false,
			displayXAxis: false,
		})
	})
})
