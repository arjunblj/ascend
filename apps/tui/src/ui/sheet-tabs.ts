import { fitAnsi, sanitizeTerminalText } from '../render/ansi-text.ts'
import { THEME } from '../render/styles.ts'

export function renderSheetTabs(
	sheets: readonly string[],
	activeIndex: number,
	width: number,
): string {
	const text =
		sheets.length === 0
			? ' [Sheet1] [+] '
			: `${sheets
					.map((sheet, index) => {
						const name = sanitizeTerminalText(sheet)
						return index === activeIndex ? `[${name}]` : ` ${name} `
					})
					.join(' ')} [+] `
	return `${THEME.ribbon}${fitAnsi(text, width)}${THEME.reset}`
}
