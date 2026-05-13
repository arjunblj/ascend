import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import type { SheetX14ConditionalFormatInfo, Workbook } from '@ascend/core'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphSafeEditIntegrity,
	inspectXlsxPackageGraph,
	readXlsx,
} from '@ascend/io-xlsx'
import { AscendWorkbook } from '@ascend/sdk'

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

function readWorkbook(bytes: Uint8Array): Workbook {
	const result = readXlsx(bytes)
	expectOk(result)
	return result.value.workbook
}

function colorScaleContract(workbook: Workbook): readonly unknown[] {
	return workbook.sheets.flatMap((sheet) =>
		sheet.conditionalFormats.flatMap((format) =>
			format.rules.flatMap((rule) =>
				rule.colorScale
					? [
							{
								sheet: sheet.name,
								sqref: format.sqref,
								cfvo: rule.colorScale.cfvo.map((entry) => ({ ...entry })),
								colors: rule.colorScale.colors.map((color) => ({ ...color })),
							},
						]
					: [],
			),
		),
	)
}

function x14ConditionalFormatContract(
	workbook: Workbook,
): readonly SheetX14ConditionalFormatInfo[] {
	return workbook.sheets.flatMap((sheet) =>
		sheet.x14ConditionalFormats.map((format) => ({
			...format,
			formulas: [...format.formulas],
			...(format.dataBar
				? {
						dataBar: {
							...format.dataBar,
							cfvo: format.dataBar.cfvo.map((entry) => ({ ...entry })),
						},
					}
				: {}),
			...(format.iconSet
				? {
						iconSet: {
							...format.iconSet,
							cfvo: format.iconSet.cfvo.map((entry) => ({ ...entry })),
							...(format.iconSet.icons
								? { icons: format.iconSet.icons.map((entry) => ({ ...entry })) }
								: {}),
						},
					}
				: {}),
		})),
	)
}

async function applySafeEdit(bytes: Uint8Array, sheet: string, ref: string): Promise<Uint8Array> {
	const workbook = await AscendWorkbook.open(bytes)
	const apply = workbook.apply([{ op: 'setCells', sheet, updates: [{ ref, value: 'safe-edit' }] }])
	expect(apply.errors).toEqual([])
	return workbook.toBytes()
}

describe('conditional-format corpus contract', () => {
	test('retains LibreOffice color-scale formulas and colors after a safe edit', async () => {
		const source = loadFixture('../xlsx/libreoffice/colorscale.xlsx')
		const before = colorScaleContract(readWorkbook(source))
		expect(before).toContainEqual(
			expect.objectContaining({
				sheet: 'Sheet1',
				sqref: 'F3:F6',
				cfvo: expect.arrayContaining([{ type: 'formula', value: '2*A1+2' }]),
			}),
		)
		expect(before).toContainEqual(
			expect.objectContaining({
				sheet: 'Sheet2',
				sqref: 'F2:F7',
				cfvo: expect.arrayContaining([{ type: 'formula', value: '2*A1+3' }]),
			}),
		)

		const edited = await applySafeEdit(source, 'Sheet1', 'H10')
		expect(colorScaleContract(readWorkbook(edited))).toEqual(before)
	})

	test('retains POI x14 conditional-format payloads after a safe edit', async () => {
		const source = loadFixture('../xlsx/poi/NewStyleConditionalFormattings.xlsx')
		const before = x14ConditionalFormatContract(readWorkbook(source))
		expect(before).toHaveLength(3)
		expect(before).toContainEqual(
			expect.objectContaining({
				sqref: 'E2:E17',
				type: 'dataBar',
				id: '{9B4F274F-F774-40EE-9C50-A8B810847010}',
				dataBar: expect.objectContaining({
					cfvo: [{ type: 'autoMin' }, { type: 'autoMax' }],
					negativeFillColor: { rgb: 'FFFF0000' },
				}),
			}),
		)
		expect(before).toContainEqual(
			expect.objectContaining({
				sqref: 'U2:U17',
				type: 'iconSet',
				iconSet: expect.objectContaining({
					custom: true,
					icons: [
						{ iconSet: '3Signs', iconId: 0 },
						{ iconSet: '3Flags', iconId: 1 },
						{ iconSet: '3Symbols2', iconId: 2 },
					],
				}),
			}),
		)

		const edited = await applySafeEdit(source, 'CF', 'A30')
		expect(x14ConditionalFormatContract(readWorkbook(edited))).toEqual(before)
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(edited),
			),
		).toEqual([])
		expect(
			auditXlsxPackageGraphBytePreservation(inspectXlsxPackageGraph(source), source, edited),
		).toEqual([])
	})
})
