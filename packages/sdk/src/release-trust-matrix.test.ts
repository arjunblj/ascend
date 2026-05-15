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
})

interface TrustMatrixRow {
	readonly failureClass: string
	readonly proof: string
}

interface ProofReference {
	readonly path: string
	readonly tests: readonly string[]
}

function trustMatrixRows(markdown: string): TrustMatrixRow[] {
	return markdown
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
