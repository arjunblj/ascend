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
})
