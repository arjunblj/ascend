import { fitAnsi, sanitizeTerminalText } from '../render/ansi-text.ts'
import { THEME } from '../render/styles.ts'
import type { TuiMode } from '../runtime/types.ts'

export function renderFormulaBar(input: {
	readonly width: number
	readonly address: string
	readonly content: string
	readonly mode: TuiMode
	readonly cursor?: number
}): string {
	const mode =
		input.mode === 'editing' || input.mode === 'entering' || input.mode === 'point' ? 'fx*' : 'fx '
	const label = input.address.includes(':') ? input.address : sanitizeTerminalText(input.address)
	const left = ` ${sanitizeTerminalText(label).padEnd(10)} | ${mode} `
	const contentWidth = Math.max(0, input.width - left.length)
	const scroll = formulaBarScrollOffset({
		width: input.width,
		address: input.address,
		content: input.content,
		mode: input.mode,
		cursor: input.cursor ?? 0,
	})
	return `${THEME.formula}${left}${fitAnsi(sanitizeTerminalText(input.content).slice(scroll), contentWidth)}${THEME.reset}`
}

export function formulaBarScrollOffset(input: {
	readonly width: number
	readonly address: string
	readonly content: string
	readonly mode: TuiMode
	readonly cursor: number
}): number {
	const mode =
		input.mode === 'editing' || input.mode === 'entering' || input.mode === 'point' ? 'fx*' : 'fx '
	const label = input.address.includes(':') ? input.address : sanitizeTerminalText(input.address)
	const prefixWidth = ` ${sanitizeTerminalText(label).padEnd(10)} | ${mode} `.length
	const contentWidth = Math.max(0, input.width - prefixWidth)
	if (contentWidth <= 0) return input.cursor
	return Math.max(0, input.cursor - contentWidth + 1)
}
