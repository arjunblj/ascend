import { describe, expect, test } from 'bun:test'
import { runFormulaAssistProof } from './formula-assist-proof.ts'

describe('formula assist corpus proof', () => {
	test('covers public formulas and rejection-first prepareRename boundaries', () => {
		const result = runFormulaAssistProof({
			includeTimings: false,
			publicFormulaLimit: 250,
		})

		expect(result.passed).toBe(true)
		expect(result.publicFormulaCount).toBeGreaterThan(100)
		expect(result.sampledFormulaCount).toBe(250)
		expect(result.staticEdgeCaseCount).toBeGreaterThanOrEqual(10)
		expect(result.referenceCount).toBeGreaterThan(100)
		expect(result.renameOkCount).toBeGreaterThan(0)
		expect(result.renameRefusalCounts['no-symbol-at-cursor']).toBeGreaterThan(0)
		expect(result.renameRefusalCounts['workbook-context-required']).toBeGreaterThan(0)
		expect(result.renameRefusalCounts['reference-target-not-renameable']).toBeGreaterThan(0)
		expect(result.edgeCases).toContainEqual(
			expect.objectContaining({
				name: 'let-shadowed-inner-binding',
				observed: 'ok',
				role: 'let-binding-use',
			}),
		)
		expect(result.edgeCases).toContainEqual(
			expect.objectContaining({
				name: 'table-column-refusal',
				observed: 'workbook-context-required',
				role: 'table-column-use',
			}),
		)
		expect(result.edgeCases).toContainEqual(
			expect.objectContaining({
				name: 'external-reference-refusal',
				observed: 'reference-target-not-renameable',
				referenceKind: 'sheet-cell',
			}),
		)
		expect(result.boundary).toContain('does not apply workbook edits')
	})
})
