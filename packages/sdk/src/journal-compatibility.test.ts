import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import type { Operation } from '@ascend/schema'
import {
	MUTATION_JOURNAL_ISSUE_CODES,
	MUTATION_JOURNAL_ISSUE_SCHEMA,
	MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
	MUTATION_JOURNAL_REASON_CODES,
	MUTATION_JOURNAL_SURFACES,
	type MutationJournal,
	type MutationJournalStructuredIssue,
} from './journal.ts'
import { AscendWorkbook } from './workbook.ts'

const FIXTURE = JSON.parse(
	readFileSync(
		new URL('../../../fixtures/journal/mutation-journal-v1.json', import.meta.url),
		'utf-8',
	),
) as {
	readonly schemaVersion: number
	readonly schemaId: string
	readonly surfaces: readonly string[]
	readonly reasons: readonly string[]
	readonly scenario: {
		readonly ops: readonly Operation[]
		readonly journal: {
			readonly schemaVersion: number
			readonly schemaId: string
			readonly supported: boolean
			readonly exact: boolean
			readonly inverseOpCount: number
			readonly issueCount: number
			readonly issues: readonly MutationJournalStructuredIssue[]
		}
	}
}

describe('mutation journal v1 compatibility', () => {
	test('freezes the public issue schema and vocabulary', () => {
		expect(MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION).toBe(FIXTURE.schemaVersion)
		expect(MUTATION_JOURNAL_ISSUE_SCHEMA.$id).toBe(FIXTURE.schemaId)
		expect(MUTATION_JOURNAL_ISSUE_SCHEMA.properties.code.enum).toEqual(MUTATION_JOURNAL_ISSUE_CODES)
		expect(MUTATION_JOURNAL_ISSUE_SCHEMA.properties.surface.enum).toEqual(FIXTURE.surfaces)
		expect(MUTATION_JOURNAL_ISSUE_SCHEMA.properties.reason.enum).toEqual(FIXTURE.reasons)
		expect(MUTATION_JOURNAL_SURFACES).toEqual(FIXTURE.surfaces)
		expect(MUTATION_JOURNAL_REASON_CODES).toEqual(FIXTURE.reasons)
	})

	test('SDK journals emit structured v1 issues for the golden lossy edit', () => {
		const wb = AscendWorkbook.create()
		const result = wb.apply(FIXTURE.scenario.ops, { journal: true })
		expect(result.errors).toEqual([])
		expect(result.journal).toBeDefined()
		if (!result.journal) throw new Error('missing journal')
		expect(journalSummary(result.journal)).toEqual(FIXTURE.scenario.journal)
		expectStructuredIssues(result.journal.issues)
		for (const entry of result.journal.entries) expectStructuredIssues(entry.issues)
	})
})

export function journalSummary(journal: MutationJournal): typeof FIXTURE.scenario.journal {
	return {
		schemaVersion: journal.schemaVersion,
		schemaId: journal.schemaId,
		supported: journal.supported,
		exact: journal.exact,
		inverseOpCount: journal.inverseOps.length,
		issueCount: journal.issues.length,
		issues: journal.issues,
	}
}

export function expectStructuredIssues(issues: readonly MutationJournalStructuredIssue[]): void {
	for (const issue of issues) {
		expect(MUTATION_JOURNAL_ISSUE_CODES).toContain(issue.code)
		expect(MUTATION_JOURNAL_SURFACES).toContain(issue.surface)
		expect(MUTATION_JOURNAL_REASON_CODES).toContain(issue.reason)
		expect(typeof issue.message).toBe('string')
		if (issue.refs) {
			expect(issue.refs.every((ref) => typeof ref === 'string')).toBe(true)
		}
	}
}
