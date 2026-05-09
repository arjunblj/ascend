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
	const match = findReferenceAtCursor(state.buffer, state.cursor)
	if (!match) return state
	const next = cycleA1Reference(match.text)
	return {
		...state,
		buffer: `${state.buffer.slice(0, match.start)}${next}${state.buffer.slice(match.end)}`,
		cursor: match.start + next.length,
	}
}

function clampCursor(cursor: number, buffer: string): number {
	return Math.max(0, Math.min(buffer.length, cursor))
}

function findReferenceAtCursor(
	buffer: string,
	cursor: number,
): { start: number; end: number; text: string } | null {
	const referencePattern = /\$?[A-Za-z]{1,3}\$?\d{1,7}/g
	let best: { start: number; end: number; text: string } | null = null
	for (const match of buffer.matchAll(referencePattern)) {
		if (match.index === undefined) continue
		const start = match.index
		const text = match[0]
		const end = start + text.length
		if (!hasReferenceBoundary(buffer, start - 1) || !hasReferenceBoundary(buffer, end)) continue
		if (cursor >= start && cursor <= end) return { start, end, text }
		if (end <= cursor && (!best || end > best.end)) best = { start, end, text }
	}
	return best
}

function hasReferenceBoundary(buffer: string, index: number): boolean {
	if (index < 0 || index >= buffer.length) return true
	const char = buffer[index]
	return char === undefined || !/[A-Za-z0-9_$]/.test(char)
}

function cycleA1Reference(ref: string): string {
	const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/.exec(ref)
	if (!match) return ref
	const [, colAbs = '', col = '', rowAbs = '', row = ''] = match
	if (!colAbs && !rowAbs) return `$${col.toUpperCase()}$${row}`
	if (colAbs && rowAbs) return `${col.toUpperCase()}$${row}`
	if (!colAbs && rowAbs) return `$${col.toUpperCase()}${row}`
	return `${col.toUpperCase()}${row}`
}
