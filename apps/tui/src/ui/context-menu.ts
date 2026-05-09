import { fitAnsi, sanitizeTerminalText } from '../render/ansi-text.ts'
import { THEME } from '../render/styles.ts'
import type { ContextMenuState } from '../runtime/types.ts'

export function renderContextMenu(
	state: ContextMenuState,
	width: number,
	height: number,
): readonly string[] {
	const panelWidth = Math.max(34, Math.min(width - 2, 76))
	const maxItems = Math.max(0, height - 4)
	const lines = [
		boxLine(panelWidth),
		boxText(`${THEME.bold}${targetTitle(state)}${THEME.reset}`, panelWidth),
		...state.items.slice(0, maxItems).map((item, index) => {
			const selected = index === clampSelection(state.selectedIndex, state.items.length)
			const marker = selected ? `${THEME.active}>${THEME.reset}` : ' '
			const shortcut = item.shortcut ? `  ${item.shortcut}` : ''
			const detail = item.detail ? `  ${sanitizeTerminalText(item.detail)}` : ''
			return boxText(
				`${marker} ${sanitizeTerminalText(item.title)}${shortcut}${detail}`,
				panelWidth,
			)
		}),
		boxText('Up/Down Select   Enter Run   Esc Close', panelWidth),
		boxLine(panelWidth),
	]
	const leftPad = ' '.repeat(Math.max(0, Math.floor((width - panelWidth) / 2)))
	return lines.slice(0, height).map((line) => fitAnsi(`${leftPad}${line}`, width))
}

function targetTitle(state: ContextMenuState): string {
	return `Context: ${state.target} ${sanitizeTerminalText(state.address)}`
}

function boxText(text: string, width: number): string {
	return `| ${fitAnsi(text, width - 4)} |`
}

function boxLine(width: number): string {
	return `+${'-'.repeat(Math.max(0, width - 2))}+`
}

function clampSelection(index: number, length: number): number {
	if (length <= 0) return 0
	return Math.max(0, Math.min(length - 1, index))
}
