import { describe, expect, test } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'

const REPO_ROOT = new URL('../../../', import.meta.url)
const TRUST_MATRIX_PATH = 'docs/RELEASE_TRUST_MATRIX.md'

const EXPECTED_FAILURE_CLASSES = [
	'Invalid refs/ranges in edit metadata',
	'Public metadata shape errors',
	'Formula/shared-formula binding risks',
	'Range operation journal exactness',
	'Table and copy-sheet binding risks',
	'Package graph and post-write drift',
	'Real workbook formula metadata preservation',
]

const EXPECTED_BOUNDARY_CLASSES = [
	'Broad formula function coverage',
	'Product/DX orchestration such as progressive open or viewport merge helpers',
	'Reader/writer performance and benchmark tuning',
	'More malformed-field enumeration',
	'New unknown Excel feature implementation',
]

describe('release trust matrix', () => {
	test('cites live proof tests and fixtures for every release trust class', () => {
		const markdown = readRepoFile(TRUST_MATRIX_PATH)
		const rows = trustMatrixRows(markdown)
		expect(rows.map((row) => row.failureClass)).toEqual(EXPECTED_FAILURE_CLASSES)

		for (const row of rows) {
			const proofRefs = proofReferences(row.proof)
			expect(proofRefs.length, row.failureClass).toBeGreaterThan(0)
			for (const ref of proofRefs) {
				expect(existsSync(new URL(ref.path, REPO_ROOT)), ref.path).toBe(true)
				if (ref.path.endsWith('.test.ts')) {
					const source = readRepoFile(ref.path)
					for (const testName of ref.tests) {
						expect(
							source.includes(`test('${testName}'`) || source.includes(`test("${testName}"`),
							`${row.failureClass}: ${ref.path} -> ${testName}`,
						).toBe(true)
					}
				}
			}
		}
	})

	test('pins explicit out-of-scope boundaries for release completeness', () => {
		const markdown = readRepoFile(TRUST_MATRIX_PATH)
		const rows = completenessBoundaryRows(markdown)
		expect(rows.map((row) => row.outOfScopeClass)).toEqual(EXPECTED_BOUNDARY_CLASSES)

		for (const row of rows) {
			expect(row.boundaryReason).toMatch(/unless|only when|not required/)
			expect(row.evidence).toMatch(
				/(Formula-binding rows|journal\/schema compatibility|Package graph|Invalid refs\/ranges)/,
			)
		}
	})
})

interface TrustMatrixRow {
	readonly failureClass: string
	readonly proof: string
}

interface CompletenessBoundaryRow {
	readonly outOfScopeClass: string
	readonly boundaryReason: string
	readonly evidence: string
}

interface ProofReference {
	readonly path: string
	readonly tests: readonly string[]
}

function trustMatrixRows(markdown: string): TrustMatrixRow[] {
	const section = markdown.split('## Matrix')[1]?.split('## Completeness Boundary')[0]
	if (!section) return []
	return section
		.split('\n')
		.filter((line) => line.startsWith('| ') && !line.includes(' --- '))
		.slice(1)
		.map((line) => {
			const cells = line.split('|').map((cell) => cell.trim())
			return {
				failureClass: cells[1],
				proof: cells[4],
			}
		})
}

function completenessBoundaryRows(markdown: string): CompletenessBoundaryRow[] {
	const section = markdown.split('## Completeness Boundary')[1]?.split('## Release Discipline')[0]
	if (!section) return []
	return section
		.split('\n')
		.filter((line) => line.startsWith('| ') && !line.includes(' --- '))
		.slice(1)
		.map((line) => {
			const cells = line.split('|').map((cell) => cell.trim())
			return {
				outOfScopeClass: cells[1],
				boundaryReason: cells[2],
				evidence: cells[3],
			}
		})
}

function proofReferences(proof: string): ProofReference[] {
	const tokens = [...proof.matchAll(/`([^`]+)`/g)].map((match) => ({
		value: match[1],
		index: match.index ?? 0,
	}))
	const pathTokens = tokens.filter(
		(token) => token.value.endsWith('.test.ts') || token.value.endsWith('.json'),
	)
	return pathTokens.map((token, index) => {
		const nextPath = pathTokens[index + 1]
		const tests = tokens
			.filter(
				(candidate) =>
					candidate.index > token.index &&
					(nextPath === undefined || candidate.index < nextPath.index) &&
					!candidate.value.endsWith('.test.ts') &&
					!candidate.value.endsWith('.json'),
			)
			.map((candidate) => candidate.value)
		return { path: token.value, tests }
	})
}

function readRepoFile(path: string): string {
	return readFileSync(new URL(path, REPO_ROOT), 'utf-8')
}
