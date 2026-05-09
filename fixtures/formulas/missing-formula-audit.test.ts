import { describe, expect, test } from 'bun:test'
import { runMissingFormulaAudit } from './missing-formula-audit.ts'

describe('common Excel function registry coverage', () => {
	test('has no gaps in the tracked public Excel function set', () => {
		const audit = runMissingFormulaAudit()

		expect(audit.total).toBeGreaterThan(400)
		expect(audit.missing).toEqual([])
		expect(audit.coverage).toBe(1)
	})
})
