import { cycleFormulaReferenceMode } from '@ascend/sdk'

export interface FormulaEditState {
	readonly buffer: string
	readonly cursor: number
	readonly pointMode: boolean
}

export function createFormulaEditState(buffer = '', cursor = buffer.length): FormulaEditState {
	return { buffer, cursor: clampCursor(cursor, buffer), pointMode: false }
}

export function insertText(state: FormulaEditState, text: string): FormulaEditState {
	return {
		...state,
		buffer: state.buffer.slice(0, state.cursor) + text + state.buffer.slice(state.cursor),
		cursor: state.cursor + text.length,
	}
}

export function backspace(state: FormulaEditState): FormulaEditState {
	if (state.cursor <= 0) return state
	return {
		...state,
		buffer: state.buffer.slice(0, state.cursor - 1) + state.buffer.slice(state.cursor),
		cursor: state.cursor - 1,
	}
}

export function deleteForward(state: FormulaEditState): FormulaEditState {
	if (state.cursor >= state.buffer.length) return state
	return {
		...state,
		buffer: state.buffer.slice(0, state.cursor) + state.buffer.slice(state.cursor + 1),
	}
}

export function moveEditCursor(
	state: FormulaEditState,
	delta: number | 'home' | 'end',
): FormulaEditState {
	const cursor =
		delta === 'home'
			? 0
			: delta === 'end'
				? state.buffer.length
				: clampCursor(state.cursor + delta, state.buffer)
	return { ...state, cursor }
}

export function cycleReferenceMode(state: FormulaEditState): FormulaEditState {
	const next = cycleFormulaReferenceMode(state.buffer, state.cursor)
	if (!next.changed) return state
	return {
		...state,
		buffer: next.formula,
		cursor: next.cursor,
	}
}

function clampCursor(cursor: number, buffer: string): number {
	return Math.max(0, Math.min(buffer.length, cursor))
}
