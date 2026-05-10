import { describe, expect, test } from 'bun:test'
import { runMissingFormulaAudit } from './missing-formula-audit.ts'

describe('common Excel function registry coverage', () => {
	test('has no gaps in the tracked public Excel function set', () => {
		const audit = runMissingFormulaAudit()

		expect(audit.total).toBeGreaterThan(400)
		expect(audit.missing).toEqual([])
		expect(audit.coverage).toBe(1)
	})

	test('separates registry coverage from JSON semantic corpus coverage', () => {
		const audit = runMissingFormulaAudit()

		expect(audit.semanticCorpus.fixtureFiles).toBeGreaterThan(20)
		expect(audit.semanticCorpus.totalCases).toBeGreaterThanOrEqual(720)
		expect(audit.semanticCorpus.coveredFunctions).toContain('SUM')
		expect(audit.semanticCorpus.coveredFunctions).toContain('AGGREGATE')
		expect(audit.semanticCorpus.trackedCovered).toContain('BINOM.DIST')
		expect(audit.semanticCorpus.trackedCovered).toContain('CHISQ.INV.RT')
		expect(audit.semanticCorpus.trackedCovered).toContain('FORECAST.ETS.CONFINT')
		expect(audit.semanticCorpus.trackedCovered).toContain('FREQUENCY')
		expect(audit.semanticCorpus.trackedCovered).toContain('RANDARRAY')
		expect(audit.semanticCorpus.trackedCovered).toContain('ISFORMULA')
		expect(audit.semanticCorpus.presentButNotCorpusCovered).toEqual([])
		expect(audit.semanticCorpus.trackedCoverage).toBe(1)
	})
})
