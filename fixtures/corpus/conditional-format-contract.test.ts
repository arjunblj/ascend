import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphSafeEditIntegrity,
	inspectXlsxPackageGraph,
} from '@ascend/io-xlsx'
import { AscendWorkbook, type SheetInspectInfo } from '@ascend/sdk'

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

function conditionalFormatContract(sheet: SheetInspectInfo): unknown {
	return {
		conditionalFormats: sheet.conditionalFormats?.map((format) => ({
			sqref: format.sqref,
			rules: format.rules.map((rule) => ({
				type: rule.type,
				formulas: rule.formulas,
				priority: rule.priority,
				operator: rule.operator,
				text: rule.text,
				rank: rule.rank,
				bottom: rule.bottom,
				percent: rule.percent,
				stopIfTrue: rule.stopIfTrue,
				dxfId: rule.dxfId,
				style: rule.style,
				colorScale: rule.colorScale,
				dataBar: rule.dataBar,
				iconSet: rule.iconSet,
				preservedRuleAttributes: rule.preservedRuleAttributes,
				preservedRuleChildXml: rule.preservedRuleChildXml,
			})),
		})),
		x14ConditionalFormats: sheet.x14ConditionalFormats?.map((format) => ({
			index: format.index,
			sqref: format.sqref,
			type: format.type,
			priority: format.priority,
			id: format.id,
			formulas: format.formulas,
			colorScale: format.colorScale,
			dataBar: format.dataBar,
			iconSet: format.iconSet,
			preservedRuleAttributes: format.preservedRuleAttributes,
			preservedRuleChildXml: format.preservedRuleChildXml,
		})),
	}
}

async function openConditionalFormatContract(
	bytes: Uint8Array,
	sheetName: string,
): Promise<unknown> {
	const workbook = await AscendWorkbook.open(bytes)
	const sheet = workbook.inspectSheet(sheetName)
	expect(sheet).toBeDefined()
	if (!sheet) throw new Error(`Missing sheet ${sheetName}`)
	return conditionalFormatContract(sheet)
}

async function safeEdit(bytes: Uint8Array, sheet: string, ref: string): Promise<Uint8Array> {
	const workbook = await AscendWorkbook.open(bytes)
	const result = workbook.apply([{ op: 'setCells', sheet, updates: [{ ref, value: 'safe-edit' }] }])
	expect(result.errors).toEqual([])
	return workbook.toBytes()
}

describe('conditional-format corpus contract', () => {
	test('preserves LibreOffice color-scale formulas and colors after a safe edit', async () => {
		const source = loadFixture('../xlsx/libreoffice/colorscale.xlsx')
		const beforeSheet1 = await openConditionalFormatContract(source, 'Sheet1')
		const beforeSheet2 = await openConditionalFormatContract(source, 'Sheet2')
		expect(beforeSheet1).toMatchObject({
			conditionalFormats: expect.arrayContaining([
				expect.objectContaining({
					sqref: 'F3:F6',
					rules: [
						expect.objectContaining({
							type: 'colorScale',
							colorScale: expect.objectContaining({
								cfvo: expect.arrayContaining([{ type: 'formula', value: '2*A1+2' }]),
							}),
						}),
					],
				}),
			]),
		})
		expect(beforeSheet2).toMatchObject({
			conditionalFormats: expect.arrayContaining([
				expect.objectContaining({
					sqref: 'F2:F7',
					rules: [
						expect.objectContaining({
							type: 'colorScale',
							colorScale: expect.objectContaining({
								cfvo: expect.arrayContaining([{ type: 'formula', value: '2*A1+3' }]),
							}),
						}),
					],
				}),
			]),
		})

		const edited = await safeEdit(source, 'Sheet1', 'H10')
		expect(await openConditionalFormatContract(edited, 'Sheet1')).toEqual(beforeSheet1)
		expect(await openConditionalFormatContract(edited, 'Sheet2')).toEqual(beforeSheet2)
	})

	test('preserves POI x14 conditional-format payloads after a safe edit', async () => {
		const source = loadFixture('../xlsx/poi/NewStyleConditionalFormattings.xlsx')
		const before = await openConditionalFormatContract(source, 'CF')
		expect(before).toMatchObject({
			x14ConditionalFormats: expect.arrayContaining([
				expect.objectContaining({
					sqref: 'E2:E17',
					type: 'dataBar',
					id: '{9B4F274F-F774-40EE-9C50-A8B810847010}',
					dataBar: expect.objectContaining({
						cfvo: [{ type: 'autoMin' }, { type: 'autoMax' }],
						negativeFillColor: { rgb: 'FFFF0000' },
					}),
				}),
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
			]),
		})

		const edited = await safeEdit(source, 'CF', 'A30')
		expect(await openConditionalFormatContract(edited, 'CF')).toEqual(before)
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

	test('preserves LibreOffice x14 priority and cross-sheet formula rules', async () => {
		const source = loadFixture('../xlsx/libreoffice/conditional_fmt_checkpriority.xlsx')
		const before = await openConditionalFormatContract(source, 'Sheet1')
		expect(before).toMatchObject({
			conditionalFormats: [
				{
					sqref: 'A1',
					rules: [
						expect.objectContaining({
							type: 'containsText',
							formulas: ['NOT(ISERROR(SEARCH("ABC",A1)))'],
							priority: 4,
							operator: 'containsText',
							text: 'ABC',
						}),
					],
				},
				{
					sqref: 'A3',
					rules: [
						expect.objectContaining({
							type: 'containsText',
							formulas: ['NOT(ISERROR(SEARCH("BAC",A3)))'],
							priority: 1,
							operator: 'containsText',
							text: 'BAC',
						}),
					],
				},
			],
			x14ConditionalFormats: [
				expect.objectContaining({
					sqref: 'A1',
					type: 'cellIs',
					priority: 3,
					formulas: ['Sheet2!$A$1'],
					preservedRuleAttributes: { operator: 'equal' },
				}),
				expect.objectContaining({
					sqref: 'A3',
					type: 'cellIs',
					priority: 2,
					formulas: ['Sheet2!$A$2'],
					preservedRuleAttributes: { operator: 'equal' },
				}),
			],
		})

		const edited = await safeEdit(source, 'Sheet1', 'C10')
		expect(await openConditionalFormatContract(edited, 'Sheet1')).toEqual(before)
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(edited),
			),
		).toEqual([])
	})

	test('preserves ClosedXML x14 data-bar regions and extension payloads', async () => {
		const source = loadFixture('../xlsx/closedxml/ConditionalFormatting_CFDataBars.xlsx')
		const before = await openConditionalFormatContract(source, 'Sheet1')
		expect(before).toMatchObject({
			conditionalFormats: expect.arrayContaining([
				expect.objectContaining({
					sqref: 'A2:A6',
					rules: [
						expect.objectContaining({
							type: 'dataBar',
							priority: 1,
							dataBar: expect.objectContaining({
								cfvo: [{ type: 'min' }, { type: 'max' }],
								color: { rgb: 'FFFFBF00' },
							}),
						}),
					],
				}),
				expect.objectContaining({
					sqref: 'C2:C6',
					rules: [
						expect.objectContaining({
							type: 'dataBar',
							priority: 3,
							dataBar: expect.objectContaining({
								cfvo: [
									{ type: 'num', value: '0' },
									{ type: 'num', value: '10' },
								],
								color: { rgb: 'FF536872' },
							}),
						}),
					],
				}),
			]),
			x14ConditionalFormats: expect.arrayContaining([
				expect.objectContaining({
					sqref: 'A2:A6',
					type: 'dataBar',
					dataBar: expect.objectContaining({
						cfvo: [{ type: 'autoMin' }, { type: 'autoMax' }],
						gradient: true,
						negativeFillColor: { rgb: 'FFFFBF00' },
					}),
				}),
				expect.objectContaining({
					sqref: 'C2:C6',
					type: 'dataBar',
					dataBar: expect.objectContaining({
						cfvo: [
							{ type: 'num', value: '0' },
							{ type: 'num', value: '10' },
						],
						axisColor: { rgb: 'FF000000' },
					}),
				}),
			]),
		})

		const edited = await safeEdit(source, 'Sheet1', 'H8')
		expect(await openConditionalFormatContract(edited, 'Sheet1')).toEqual(before)
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(edited),
			),
		).toEqual([])
	})

	test('preserves POI multi-rule formula, icon-set, data-bar, and color-scale sheets', async () => {
		const source = loadFixture('../xlsx/filter/poi/ConditionalFormattingSamples.xlsx')
		const cases = [
			{
				sheet: 'Products1',
				editRef: 'G25',
				expected: {
					conditionalFormats: expect.arrayContaining([
						expect.objectContaining({
							sqref: 'A3:A23',
							rules: [
								expect.objectContaining({
									type: 'timePeriod',
									formulas: ['AND(MONTH(A3)=MONTH(TODAY()),YEAR(A3)=YEAR(TODAY()))'],
									priority: 3,
								}),
							],
						}),
						expect.objectContaining({
							sqref: 'B3:B23',
							rules: [
								expect.objectContaining({
									type: 'containsText',
									formulas: ['NOT(ISERROR(SEARCH("Grain",B3)))'],
									priority: 4,
								}),
							],
						}),
					]),
				},
			},
			{
				sheet: 'Quarters',
				editRef: 'O10',
				expected: {
					conditionalFormats: expect.arrayContaining([
						expect.objectContaining({
							sqref: 'M4:M8',
							rules: [
								expect.objectContaining({
									type: 'iconSet',
									priority: 8,
									iconSet: expect.objectContaining({
										iconSet: '3Symbols2',
										showValue: false,
									}),
								}),
							],
						}),
					]),
				},
			},
			{
				sheet: 'Mountains',
				editRef: 'H26',
				expected: {
					conditionalFormats: expect.arrayContaining([
						expect.objectContaining({
							sqref: 'D3:D24',
							rules: [
								expect.objectContaining({
									type: 'dataBar',
									priority: 8,
									dataBar: expect.objectContaining({ color: { rgb: 'FF63C384' } }),
								}),
							],
						}),
					]),
				},
			},
			{
				sheet: 'Category sales',
				editRef: 'I24',
				expected: {
					conditionalFormats: expect.arrayContaining([
						expect.objectContaining({
							sqref: 'G3:G22',
							rules: [
								expect.objectContaining({
									type: 'colorScale',
									priority: 4,
									colorScale: expect.objectContaining({
										cfvo: [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
									}),
								}),
							],
						}),
					]),
				},
			},
		] as const

		for (const entry of cases) {
			const before = await openConditionalFormatContract(source, entry.sheet)
			expect(before).toMatchObject(entry.expected)

			const edited = await safeEdit(source, entry.sheet, entry.editRef)
			expect(await openConditionalFormatContract(edited, entry.sheet)).toEqual(before)
			expect(
				auditXlsxPackageGraphSafeEditIntegrity(
					inspectXlsxPackageGraph(source),
					inspectXlsxPackageGraph(edited),
				),
			).toEqual([])
		}
	})
})
