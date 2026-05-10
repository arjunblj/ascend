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
