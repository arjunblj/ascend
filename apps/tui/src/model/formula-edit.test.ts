import { describe, expect, test } from 'bun:test'
import {
	backspace,
	createFormulaEditState,
	cycleReferenceMode,
	deleteForward,
	insertText,
	moveEditCursor,
} from './formula-edit.ts'

describe('formula edit model', () => {
	test('inserts and deletes at the cursor', () => {
		let state = createFormulaEditState('=A1', 1)
		state = insertText(state, 'SUM(')
		expect(state.buffer).toBe('=SUM(A1')
		expect(state.cursor).toBe(5)
		state = moveEditCursor(state, 'end')
		state = insertText(state, ')')
		expect(state.buffer).toBe('=SUM(A1)')
		state = moveEditCursor(state, -1)
		state = backspace(state)
		expect(state.buffer).toBe('=SUM(A)')
		state = deleteForward(state)
		expect(state.buffer).toBe('=SUM(A')
	})

	test('cycles Excel A1 reference modes with F4 semantics', () => {
		let state = createFormulaEditState('=A1+B2', 3)
		state = cycleReferenceMode(state)
		expect(state.buffer).toBe('=$A$1+B2')
		state = cycleReferenceMode(state)
		expect(state.buffer).toBe('=A$1+B2')
		state = cycleReferenceMode(state)
		expect(state.buffer).toBe('=$A1+B2')
		state = cycleReferenceMode(state)
		expect(state.buffer).toBe('=A1+B2')
	})

	test('uses shared SDK formula reference cycling for sheet-qualified refs', () => {
		const buffer = "='My Sheet'!A1+B2"
		let state = createFormulaEditState(buffer, buffer.indexOf('A1') + 1)
		state = cycleReferenceMode(state)
		expect(state.buffer).toBe("='My Sheet'!$A$1+B2")
		expect(state.cursor).toBe("='My Sheet'!$A$1".length)
	})
})
