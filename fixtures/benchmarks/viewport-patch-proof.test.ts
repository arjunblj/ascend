import { describe, expect, test } from 'bun:test'
import { runViewportPatchProof, viewportPatchProofMarkdown } from './viewport-patch-proof.ts'

describe('viewport patch proof harness', () => {
	test('proves retained patches and explicit invalidation reasons', async () => {
		const proof = await runViewportPatchProof()
		const byName = new Map(proof.cases.map((entry) => [entry.name, entry]))

		expect(proof.passed).toBe(true)
		expect(byName.get('retained-patch')).toMatchObject({
			passed: true,
			changedRefs: ['A1'],
		})
		expect(byName.get('skipped-token-retained')?.observed).toContain('patch:')
		expect(byName.get('invalid-token')).toMatchObject({ invalidationReason: 'base-token-invalid' })
		expect(byName.get('cross-session-token')).toMatchObject({
			invalidationReason: 'base-snapshot-missing',
		})
		expect(byName.get('expired-history')).toMatchObject({
			invalidationReason: 'base-token-expired',
		})
		expect(byName.get('projection-change')).toMatchObject({
			invalidationReason: 'base-snapshot-missing',
		})
		expect(byName.get('metadata-invalidation')).toMatchObject({
			invalidationReason: 'viewport-invalidated',
		})
	})

	test('renders claim-safe markdown boundaries', async () => {
		const proof = await runViewportPatchProof()
		const markdown = viewportPatchProofMarkdown(proof)

		expect(markdown).toContain('Viewport Patch Proof')
		expect(markdown).toContain('bounded single-session viewport patch retention')
		expect(markdown).toContain('not CRDT collaboration')
		expect(markdown).toContain('All passed: true')
	})
})
