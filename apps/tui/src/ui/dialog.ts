import { fitAnsi, sanitizeTerminalText } from '../render/ansi-text.ts'
import { THEME } from '../render/styles.ts'
import type { DialogViewState } from '../runtime/types.ts'

export function renderDialog(
	dialog: DialogViewState,
	width: number,
	height: number,
): readonly string[] {
	const panelWidth = Math.max(20, Math.min(Math.max(1, width - 2), 82))
	const lines: string[] = []
	lines.push(boxLine('top', panelWidth))
	lines.push(boxText(`${THEME.bold}${dialog.title}${THEME.reset}`, panelWidth))
	lines.push(boxText(dialogScope(dialog), panelWidth))
	lines.push(boxLine('mid', panelWidth))
	for (const [index, field] of dialog.fields.slice(0, Math.max(0, height - 6)).entries()) {
		const active = index === dialog.activeField
		const required = field.required ? ' *' : ''
		const options = field.kind === 'select' ? '  Left/Right changes' : ''
		const value = renderFieldValue(field)
		const marker = active ? `${THEME.active}>${THEME.reset}` : ' '
		const text = `${marker} ${safeText(field.label)}${required}: ${value}${options}`
		lines.push(boxText(text, panelWidth))
	}
	if (dialog.fields.length === 0) {
		lines.push(boxText('Preview-only in this foundation slice.', panelWidth))
	}
	lines.push(boxLine('mid', panelWidth))
	lines.push(boxText(dialogPreview(dialog), panelWidth))
	lines.push(boxText('Tab Move   Space Toggle/Select   Enter Apply   Esc Close', panelWidth))
	lines.push(boxLine('bottom', panelWidth))

	const leftPad = ' '.repeat(Math.max(0, Math.floor((width - panelWidth) / 2)))
	const topPad = Math.max(0, Math.floor((height - lines.length) / 2))
	return [
		...Array.from({ length: topPad }, () => ''.padEnd(width)),
		...lines.map((line) => fitAnsi(`${leftPad}${line}`, width)),
	]
}

function dialogScope(dialog: DialogViewState): string {
	const rangeField = dialog.fields.find(
		(field) => field.name === 'range' || field.name === 'ref' || field.name === 'printArea',
	)
	return rangeField?.value ? `Scope ${safeText(rangeField.value)}` : 'Scope current selection'
}

function dialogPreview(dialog: DialogViewState): string {
	switch (dialog.id) {
		case 'format-cells':
			return 'Preview: applies number, font, alignment, fill, and border options.'
		case 'paste-special':
			return 'Result: controls whether formulas, values, formats, or layout are pasted.'
		case 'sort':
			return 'Preview: sort levels reorder rows in the selected range.'
		case 'filter':
			return 'Preview: header filter buttons and hidden-row counts are shown in the grid.'
		case 'find-replace':
			return 'Result: Find selects matches; Replace All journals every changed cell.'
		case 'data-validation':
			return 'Preview: selected cells get rule, input message, and error behavior.'
		case 'conditional-formatting':
			return 'Preview: rule is stored symbolically and rendered as terminal-safe cues.'
		case 'print-preview':
			return 'Preview: page setup and print area metadata update before export/print.'
		default:
			return 'Preview: operation is journaled and can be inspected before applying.'
	}
}

function boxText(text: string, width: number): string {
	return `| ${fitAnsi(text, width - 4)} |`
}

function boxLine(_kind: 'top' | 'mid' | 'bottom', width: number): string {
	return `+${'-'.repeat(Math.max(0, width - 2))}+`
}

function renderFieldValue(field: DialogViewState['fields'][number]): string {
	switch (field.kind) {
		case 'boolean':
			return field.value === 'true' ? '[x]' : '[ ]'
		case 'select':
			return `<${safeText(field.value || 'choose')}>`
		default:
			return field.value ? safeText(field.value) : '<empty>'
	}
}

function safeText(text: string): string {
	return sanitizeTerminalText(text)
}
