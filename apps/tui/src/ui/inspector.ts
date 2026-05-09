import { fitAnsi } from '../render/ansi-text.ts'

export function renderInspector(width: number, lines?: readonly string[]): readonly string[] {
	const content =
		lines && lines.length > 0
			? lines
			: ['Inspector', 'Formula, format, validation, comments, and trace details appear here.']
	return content.map((line) => fitAnsi(line, width))
}
