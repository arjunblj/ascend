import { fitAnsi, sanitizeTerminalText } from '../render/ansi-text.ts'
import { THEME } from '../render/styles.ts'
import type { OpenWorkbook } from '../runtime/types.ts'

export function renderWorkbookTabs(
	documents: readonly OpenWorkbook[],
	activeWorkbookId: string | null,
	width: number,
): string {
	const text =
		documents.length === 0
			? ' Ascend '
			: documents
					.map((doc, index) => {
						const active = doc.id === activeWorkbookId
						const dirty = doc.dirty ? ' *' : ''
						const name = sanitizeTerminalText(doc.name)
						return active ? `[${index + 1} ${name}${dirty}]` : ` ${index + 1} ${name}${dirty} `
					})
					.join(' ')
	return `${THEME.canvas}${fitAnsi(text, width)}${THEME.reset}`
}
