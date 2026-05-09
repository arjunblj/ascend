import sliceAnsi from 'slice-ansi'
import stringWidth from 'string-width'

const ESC = String.fromCharCode(27)

export function visibleLength(text: string): number {
	if (isFastTerminalText(text)) return fastVisibleLength(text)
	return stringWidth(text)
}

export function fitAnsi(text: string, width: number): string {
	const clipped = clipAnsi(text, width)
	const padding = Math.max(0, width - visibleLength(clipped))
	return padding === 0 ? clipped : `${clipped}${' '.repeat(padding)}`
}

export function clipAnsi(text: string, width: number): string {
	if (width <= 0) return ''
	if (isFastTerminalText(text)) return fastClipAnsi(text, width)
	return sliceAnsi(text, 0, width)
}

export function sanitizeTerminalText(text: string): string {
	let out = ''
	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i)
		if (code === 0x1b) {
			i = skipEscape(text, i)
			continue
		}
		if (code === 0x09) {
			out += ' '
			continue
		}
		if (code === 0x0a || code === 0x0d) {
			out += ' '
			continue
		}
		if (code < 0x20 || (code >= 0x7f && code <= 0x9f)) continue
		out += text[i]
	}
	return out
}

function isFastTerminalText(text: string): boolean {
	for (let i = 0; i < text.length; i++) {
		if (!isSingleWidthFastChar(text.charCodeAt(i))) return false
	}
	return true
}

function isSingleWidthFastChar(code: number): boolean {
	return (
		code <= 0x7f ||
		code === 0x2026 ||
		(code >= 0x2500 && code <= 0x257f) ||
		(code >= 0x2580 && code <= 0x259f) ||
		(code >= 0x25a0 && code <= 0x25ff)
	)
}

function fastVisibleLength(text: string): number {
	let width = 0
	for (let i = 0; i < text.length; i++) {
		if (isAnsiStart(text, i)) {
			i = skipAnsi(text, i)
			continue
		}
		width += 1
	}
	return width
}

function fastClipAnsi(text: string, width: number): string {
	let visible = 0
	let out = ''
	for (let i = 0; i < text.length; i++) {
		if (isAnsiStart(text, i)) {
			const end = skipAnsi(text, i)
			out += text.slice(i, end + 1)
			i = end
			continue
		}
		if (visible >= width) continue
		out += text[i]
		visible += 1
	}
	return out
}

function isAnsiStart(text: string, index: number): boolean {
	return text[index] === ESC && text[index + 1] === '['
}

function skipAnsi(text: string, index: number): number {
	for (let i = index + 2; i < text.length; i++) {
		const code = text.charCodeAt(i)
		if (code >= 0x40 && code <= 0x7e) return i
	}
	return text.length - 1
}

function skipEscape(text: string, index: number): number {
	const next = text[index + 1]
	if (next === '[') return skipAnsi(text, index)
	if (next === ']') {
		for (let i = index + 2; i < text.length; i++) {
			if (text.charCodeAt(i) === 0x07) return i
			if (text[i] === ESC && text[i + 1] === '\\') return i + 1
		}
		return text.length - 1
	}
	return Math.min(index + 1, text.length - 1)
}
