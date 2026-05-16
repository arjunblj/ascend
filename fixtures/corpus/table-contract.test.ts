import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { auditXlsxPackageGraphSafeEditIntegrity, inspectXlsxPackageGraph } from '@ascend/io-xlsx'
import { AscendWorkbook, type TableInfo } from '@ascend/sdk'

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

function tableContract(table: TableInfo): unknown {
	return {
		name: table.name,
		ref: table.ref,
		rowCount: table.rowCount,
		hasHeaders: table.hasHeaders,
		hasTotals: table.hasTotals,
		headerRow: table.headerRow,
		totalsRow: table.totalsRow,
		columnDefs: table.columnDefs.map((column) => ({
			name: column.name,
			id: column.id,
			...(column.formula ? { formula: column.formula } : {}),
			...(column.totalsRowFunction ? { totalsRowFunction: column.totalsRowFunction } : {}),
			...(column.totalsRowFormula ? { totalsRowFormula: column.totalsRowFormula } : {}),
			...(column.totalsRowLabel ? { totalsRowLabel: column.totalsRowLabel } : {}),
			...(column.dataDxfId !== undefined ? { dataDxfId: column.dataDxfId } : {}),
			...(column.totalsRowDxfId !== undefined ? { totalsRowDxfId: column.totalsRowDxfId } : {}),
		})),
	}
}

async function openTableContract(
	bytes: Uint8Array,
	sheetName: string,
	tableName: string,
): Promise<unknown> {
	const workbook = await AscendWorkbook.open(bytes)
	const table = workbook.inspectSheet(sheetName)?.tables?.find((entry) => entry.name === tableName)
	expect(table).toBeDefined()
	if (!table) throw new Error(`Missing table ${tableName}`)
	return tableContract(table)
}

async function safeEdit(bytes: Uint8Array, sheet: string, ref: string): Promise<Uint8Array> {
	const workbook = await AscendWorkbook.open(bytes)
	const result = workbook.apply([{ op: 'setCells', sheet, updates: [{ ref, value: 'safe-edit' }] }])
	expect(result.errors).toEqual([])
	return workbook.toBytes()
}

describe('table corpus contract', () => {
	test('preserves LibreOffice total-row functions, labels, styles, and visible values', async () => {
		const cases = [
			{
				path: '../xlsx/libreoffice/totalsRowFunction.xlsx',
				sheet: 'Present planner',
				table: 'PresentPlanner',
				editRef: 'H10',
				expected: {
					name: 'PresentPlanner',
					rowCount: 4,
					hasHeaders: true,
					hasTotals: true,
					headerRow: [
						{ kind: 'string', value: 'WHEN' },
						{ kind: 'string', value: 'WHO' },
						{ kind: 'string', value: 'WHAT' },
						{ kind: 'string', value: 'WHERE' },
						{ kind: 'string', value: 'HOW MUCH' },
						{ kind: 'string', value: 'NOTES' },
					],
					totalsRow: [
						{ kind: 'string', value: 'Total' },
						{ kind: 'empty' },
						{ kind: 'empty' },
						{ kind: 'empty' },
						{ kind: 'number', value: 350 },
						{ kind: 'empty' },
					],
					columnDefs: expect.arrayContaining([
						expect.objectContaining({ name: 'WHEN', totalsRowLabel: 'Total' }),
						expect.objectContaining({
							name: 'HOW MUCH',
							totalsRowFunction: 'sum',
							dataDxfId: 1,
							totalsRowDxfId: 2,
						}),
					]),
				},
			},
			{
				path: '../xlsx/libreoffice/tdf162963_TableWithTotalsEnabled.xlsx',
				sheet: 'Sheet1',
				table: 'myData',
				editRef: 'D10',
				expected: {
					name: 'myData',
					rowCount: 4,
					hasHeaders: true,
					hasTotals: true,
					headerRow: [
						{ kind: 'string', value: 'Name' },
						{ kind: 'string', value: 'Sales' },
					],
					totalsRow: [
						{ kind: 'string', value: 'All' },
						{ kind: 'number', value: 115 },
					],
					columnDefs: expect.arrayContaining([
						expect.objectContaining({ name: 'Name', totalsRowLabel: 'All' }),
						expect.objectContaining({
							name: 'Sales',
							totalsRowFunction: 'custom',
							totalsRowFormula: 'SUM(myData[Sales])',
						}),
					]),
				},
			},
		] as const

		for (const entry of cases) {
			const source = loadFixture(entry.path)
			const before = await openTableContract(source, entry.sheet, entry.table)
			expect(before).toMatchObject(entry.expected)

			const edited = await safeEdit(source, entry.sheet, entry.editRef)
			expect(await openTableContract(edited, entry.sheet, entry.table)).toEqual(before)
			expect(
				auditXlsxPackageGraphSafeEditIntegrity(
					inspectXlsxPackageGraph(source),
					inspectXlsxPackageGraph(edited),
				),
			).toEqual([])
		}
	})

	test('preserves POI current-row structured references in calculated table columns', async () => {
		const source = loadFixture('../xlsx/poi/StructuredReferences.xlsx')
		const before = await openTableContract(source, 'Table', '\\_Prime.1')
		expect(before).toMatchObject({
			name: '\\_Prime.1',
			rowCount: 6,
			hasHeaders: true,
			hasTotals: false,
			headerRow: [
				{ kind: 'string', value: 'calc=#*#' },
				{ kind: 'string', value: 'Name' },
				{ kind: 'string', value: 'Number' },
			],
			columnDefs: expect.arrayContaining([
				expect.objectContaining({
					name: 'calc=#*#',
					formula: '\\_Prime.1[[#This Row],[Number]]*\\_Prime.1[[#This Row],[Number]]',
				}),
			]),
		})

		const edited = await safeEdit(source, 'Table', 'E20')
		expect(await openTableContract(edited, 'Table', '\\_Prime.1')).toEqual(before)
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(edited),
			),
		).toEqual([])
	})
})
