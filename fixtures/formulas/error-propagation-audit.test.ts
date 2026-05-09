import { describe, expect, test } from 'bun:test'
import { runErrorPropagationAudit } from './error-propagation-audit.ts'

describe('error propagation audit', () => {
	test('no unexpected non-propagating #REF! slots', () => {
		const { unexpected, registered } = runErrorPropagationAudit()
		expect(registered).toBeGreaterThan(0)
		expect(unexpected).toBe(0)
	})
})
