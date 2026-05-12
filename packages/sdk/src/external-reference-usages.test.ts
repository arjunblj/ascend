import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'

describe('external reference usages', () => {
	test('inspect maps external workbook references back to formulas and names', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setFormula',
				sheet: 'Sheet1',
				ref: 'A1',
				formula: '=SUM([Budget.xlsx]FY26!B2:B10)',
			},
			{
				op: 'setDefinedName',
				name: 'BudgetSource',
				ref: '[Budget.xlsx]FY26!A1:D10',
			},
		])

		expect(wb.inspect().externalReferenceUsages).toEqual([
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'cellFormula',
				sourceRef: 'Sheet1!A1',
				formula: 'SUM([Budget.xlsx]FY26!B2:B10)',
				references: ['[Budget.xlsx]FY26!B2:B10'],
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'definedName',
				name: 'BudgetSource',
				formula: '[Budget.xlsx]FY26!A1:D10',
				references: ['[Budget.xlsx]FY26!A1:D10'],
			},
		])
		expect(wb.externalReferenceUsages()).toEqual(wb.inspect().externalReferenceUsages)
	})

	test('inspect maps external workbook references in sheet metadata formulas', () => {
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as {
			wb: {
				chartParts: unknown[]
				externalReferenceDetails: Array<Record<string, unknown>>
				sheets: Array<{
					tables: unknown[]
					conditionalFormats: unknown[]
					dataValidations: unknown[]
					sparklineGroups: unknown[]
					x14ConditionalFormats: unknown[]
					x14DataValidations: unknown[]
				}>
			}
		}
		internal.wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			target: '../sources/Budget.xlsx',
		})
		internal.wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			series: [
				{
					nameRef: '[Budget.xlsx]FY26!$G$1',
					categoryRef: '[Budget.xlsx]FY26!$A$2:$A$10',
					valueRef: '[Budget.xlsx]FY26!$G$2:$G$10',
				},
			],
		})
		const sheet = internal.wb.sheets[0]
		if (!sheet) throw new Error('missing sheet')
		sheet.tables.push({
			name: 'Sales',
			columns: [
				{
					name: 'Amount',
					formula: '[Budget.xlsx]FY26!F1',
					totalsRowFormula: 'SUM([Budget.xlsx]FY26!F:F)',
				},
			],
		})
		sheet.conditionalFormats.push({
			sqref: 'A1:A10',
			rules: [{ type: 'expression', formulas: ['[Budget.xlsx]FY26!A1>0'] }],
		})
		sheet.dataValidations.push({
			sqref: 'B1:B10',
			type: 'list',
			formula1: '[Budget.xlsx]Lists!$A$1:$A$5',
		})
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'C1:C10',
			formulas: ['[Budget.xlsx]FY26!C1>0'],
			dataBar: { cfvo: [{ type: 'formula', value: '[Budget.xlsx]FY26!D1' }] },
		})
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'E1:E10',
			formula1: '[Budget.xlsx]Lists!$B$1:$B$5',
		})
		sheet.sparklineGroups.push({
			groupIndex: 0,
			range: '[Budget.xlsx]FY26!$H$2:$H$10',
			dateAxisRange: '[Budget.xlsx]FY26!$A$2:$A$10',
			sparklines: [
				{
					range: '[Budget.xlsx]FY26!$I$2:$I$10',
					locationRange: 'F2:F10',
				},
			],
			count: 1,
		})

		expect(wb.externalReferenceUsages()).toEqual([
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'chartSeriesName',
				sourceRef: 'xl/charts/chart1.xml#series0',
				formula: '[Budget.xlsx]FY26!$G$1',
				references: ['[Budget.xlsx]FY26!$G$1'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'chartSeriesCategory',
				sourceRef: 'xl/charts/chart1.xml#series0',
				formula: '[Budget.xlsx]FY26!$A$2:$A$10',
				references: ['[Budget.xlsx]FY26!$A$2:$A$10'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'chartSeriesValue',
				sourceRef: 'xl/charts/chart1.xml#series0',
				formula: '[Budget.xlsx]FY26!$G$2:$G$10',
				references: ['[Budget.xlsx]FY26!$G$2:$G$10'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'tableColumnFormula',
				sourceRef: 'Sheet1!Sales[Amount]',
				formula: '[Budget.xlsx]FY26!F1',
				references: ['[Budget.xlsx]FY26!F1'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'tableTotalsRowFormula',
				sourceRef: 'Sheet1!Sales[Amount]',
				formula: 'SUM([Budget.xlsx]FY26!F:F)',
				references: ['[Budget.xlsx]FY26!F:F'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'conditionalFormat',
				sourceRef: 'Sheet1!A1:A10',
				formula: '[Budget.xlsx]FY26!A1>0',
				references: ['[Budget.xlsx]FY26!A1'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'Lists',
				sourceKind: 'dataValidation',
				sourceRef: 'Sheet1!B1:B10',
				formula: '[Budget.xlsx]Lists!$A$1:$A$5',
				references: ['[Budget.xlsx]Lists!$A$1:$A$5'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'x14ConditionalFormat',
				sourceRef: 'Sheet1!C1:C10',
				formula: '[Budget.xlsx]FY26!C1>0',
				references: ['[Budget.xlsx]FY26!C1'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'x14ConditionalFormat',
				sourceRef: 'Sheet1!C1:C10',
				formula: '[Budget.xlsx]FY26!D1',
				references: ['[Budget.xlsx]FY26!D1'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'Lists',
				sourceKind: 'x14DataValidation',
				sourceRef: 'Sheet1!E1:E10',
				formula: '[Budget.xlsx]Lists!$B$1:$B$5',
				references: ['[Budget.xlsx]Lists!$B$1:$B$5'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'sparklineGroupRange',
				sourceRef: 'Sheet1!sparklineGroup0',
				formula: '[Budget.xlsx]FY26!$H$2:$H$10',
				references: ['[Budget.xlsx]FY26!$H$2:$H$10'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'sparklineDateAxisRange',
				sourceRef: 'Sheet1!sparklineGroup0',
				formula: '[Budget.xlsx]FY26!$A$2:$A$10',
				references: ['[Budget.xlsx]FY26!$A$2:$A$10'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
			{
				workbook: 'Budget.xlsx',
				sheet: 'FY26',
				sourceKind: 'sparklineRange',
				sourceRef: 'Sheet1!sparklineGroup0#sparkline0',
				formula: '[Budget.xlsx]FY26!$I$2:$I$10',
				references: ['[Budget.xlsx]FY26!$I$2:$I$10'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					target: '../sources/Budget.xlsx',
				},
			},
		])
		expect(wb.inspect().externalReferenceUsages).toEqual(wb.externalReferenceUsages())
	})

	test('resolves numeric workbook tokens to external link package targets', () => {
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as {
			wb: {
				externalReferences: string[]
				externalReferenceDetails: Array<Record<string, unknown>>
			}
		}
		internal.wb.externalReferences.push('xl/externalLinks/externalLink1.xml')
		internal.wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rId4',
			linkRelId: 'rIdExt',
			linkRelationshipKind: 'externalLinkPath',
			linkBindingStatus: 'externalBookRelId',
			target: '../sources/book1.xlsx',
			targetMode: 'External',
		})
		wb.apply([
			{
				op: 'setFormula',
				sheet: 'Sheet1',
				ref: 'A1',
				formula: '=SUM([1]Sheet1!B2:B10)',
			},
			{
				op: 'setDefinedName',
				name: 'ExternalSource',
				ref: '[1]Sheet1!A1:D10',
			},
		])

		expect(wb.externalReferenceUsages()).toEqual([
			{
				workbook: '1',
				sheet: 'Sheet1',
				sourceKind: 'cellFormula',
				sourceRef: 'Sheet1!A1',
				formula: 'SUM([1]Sheet1!B2:B10)',
				references: ['[1]Sheet1!B2:B10'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					relId: 'rId4',
					linkRelId: 'rIdExt',
					linkRelationshipKind: 'externalLinkPath',
					linkBindingStatus: 'externalBookRelId',
					target: '../sources/book1.xlsx',
					targetMode: 'External',
				},
			},
			{
				workbook: '1',
				sheet: 'Sheet1',
				sourceKind: 'definedName',
				name: 'ExternalSource',
				formula: '[1]Sheet1!A1:D10',
				references: ['[1]Sheet1!A1:D10'],
				externalReference: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					relId: 'rId4',
					linkRelId: 'rIdExt',
					linkRelationshipKind: 'externalLinkPath',
					linkBindingStatus: 'externalBookRelId',
					target: '../sources/book1.xlsx',
					targetMode: 'External',
				},
			},
		])
	})

	test('resolves named workbook tokens only when the target match is unique', () => {
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as {
			wb: {
				externalReferences: string[]
				externalReferenceDetails: Array<Record<string, unknown>>
			}
		}
		internal.wb.externalReferences.push(
			'xl/externalLinks/externalLink1.xml',
			'xl/externalLinks/externalLink2.xml',
			'xl/externalLinks/externalLink3.xml',
		)
		internal.wb.externalReferenceDetails.push(
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				target: '../sources/Budget.xlsx',
			},
			{
				partPath: 'xl/externalLinks/externalLink2.xml',
				target: '../archive/Ambiguous.xlsx',
			},
			{
				partPath: 'xl/externalLinks/externalLink3.xml',
				target: '../other/Ambiguous.xlsx',
			},
		)
		wb.apply([
			{
				op: 'setFormula',
				sheet: 'Sheet1',
				ref: 'A1',
				formula: '=[Budget.xlsx]FY26!A1',
			},
			{
				op: 'setFormula',
				sheet: 'Sheet1',
				ref: 'A2',
				formula: '=[Ambiguous.xlsx]FY26!A1',
			},
		])

		const usages = wb.externalReferenceUsages()
		expect(usages[0]).toMatchObject({
			workbook: 'Budget.xlsx',
			externalReference: {
				partPath: 'xl/externalLinks/externalLink1.xml',
				target: '../sources/Budget.xlsx',
			},
		})
		expect(usages[1]).toEqual({
			workbook: 'Ambiguous.xlsx',
			sheet: 'FY26',
			sourceKind: 'cellFormula',
			sourceRef: 'Sheet1!A2',
			formula: '[Ambiguous.xlsx]FY26!A1',
			references: ['[Ambiguous.xlsx]FY26!A1'],
		})
	})
})
