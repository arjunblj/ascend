import { describe, expect, test } from 'bun:test'
import { validateExcelDefinedName } from './excel-names.ts'

describe('excel name validation', () => {
	test('validates public defined names', () => {
		for (const name of ['Budget', '_Budget.Total', '\\PrintArea']) {
			expect(validateExcelDefinedName(name), name).toBeNull()
		}

		for (const name of ['', 'A1', 'R1C1', 'R', 'C', 'Bad Name', '1Budget']) {
			expect(validateExcelDefinedName(name), name)?.toMatchObject({
				suggestedFix: expect.stringContaining('defined name'),
			})
		}
	})
})
