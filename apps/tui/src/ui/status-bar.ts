import { indexToColumn } from '@ascend/sdk'
import { selectionDimensions, selectionRef } from '../model/selection.ts'
import { fitAnsi, sanitizeTerminalText, visibleLength } from '../render/ansi-text.ts'
import { THEME } from '../render/styles.ts'
import type { SelectionState, TuiMode } from '../runtime/types.ts'

export function renderStatusBar(input: {
	readonly width: number
	readonly mode: TuiMode
	readonly sheet: string
	readonly dirty: boolean
	readonly selection: SelectionState
	readonly message: string
	readonly summary?: string
}): string {
	const dims = selectionDimensions(input.selection)
	const mode = modeLabel(input.mode)
	const base = `${mode} | ${sanitizeTerminalText(input.sheet)} | ${selectionRef(input.selection, indexToColumn)} | ${dims.rows}R x ${dims.cols}C | ${input.dirty ? 'Unsaved' : 'Saved'}`
	const left = input.message ? `${base} | ${sanitizeTerminalText(input.message)}` : base
	const right = input.summary
		? ` ${sanitizeTerminalText(input.summary)}   F1 Help  Ctrl+P Commands `
		: ' F1 Help  Ctrl+P Commands '
	const gap = Math.max(1, input.width - visibleLength(left) - visibleLength(right) - 1)
	return `${THEME.status}${fitAnsi(` ${left}${' '.repeat(gap)}${right}`, input.width)}${THEME.reset}`
}

function modeLabel(mode: TuiMode): string {
	switch (mode) {
		case 'ready':
			return 'READY'
		case 'entering':
			return 'ENTERING'
		case 'editing':
			return 'EDITING'
		case 'point':
			return 'POINT'
		case 'command':
			return 'COMMAND'
	}
}
