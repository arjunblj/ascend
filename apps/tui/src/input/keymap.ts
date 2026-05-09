import { StringDecoder } from 'node:string_decoder'
import type { InputEvent } from '../runtime/types.ts'

export function parseInputBuffer(buffer: Buffer): InputEvent {
	const raw = buffer.toString('utf8')
	return parseInputEvents(buffer)[0] ?? parseSingleInput(raw)
}

export function parseInputEvents(buffer: Buffer): readonly InputEvent[] {
	return parseInputString(buffer.toString('utf8'), false).events
}

export class TerminalInputParser {
	private readonly decoder = new StringDecoder('utf8')
	private pending = ''

	hasPending(): boolean {
		return this.pending.length > 0
	}

	push(buffer: Buffer): readonly InputEvent[] {
		const decoded = this.pending + this.decoder.write(buffer)
		const result = parseInputString(decoded, true)
		this.pending = result.pending
		return result.events
	}

	flush(): readonly InputEvent[] {
		const decoded = this.pending + this.decoder.end()
		this.pending = ''
		return decoded ? parseInputString(decoded, false).events : []
	}
}

function parseSingleInput(raw: string): InputEvent {
	const bracketedPaste = parseBracketedPaste(raw)
	if (bracketedPaste) return bracketedPaste
	const mapped = KEY_SEQUENCES.get(raw)
	if (mapped) return { kind: 'key', key: mapped, raw }
	if (raw.startsWith('\x1b[<')) return parseMouse(raw)
	const csiU = parseCsiU(raw)
	if (csiU) return csiU
	if (raw.length > 0 && [...raw].every((char) => char.charCodeAt(0) >= 0x20)) {
		return { kind: 'key', key: 'text', text: raw, raw }
	}
	return { kind: 'key', key: raw, raw }
}

function parseInputString(
	raw: string,
	allowPending: boolean,
): { readonly events: readonly InputEvent[]; readonly pending: string } {
	const events: InputEvent[] = []
	let index = 0
	while (index < raw.length) {
		const current = raw[index] ?? ''
		if (current !== '\x1b') {
			const control = parseControlAt(raw, index)
			if (control) {
				events.push(control.event)
				index += control.length
				continue
			}
			const start = index
			while (index < raw.length && raw[index] !== '\x1b' && !parseControlAt(raw, index)) {
				index += (raw.codePointAt(index) ?? 0) > 0xffff ? 2 : 1
			}
			const text = raw.slice(start, index)
			if (text) events.push({ kind: 'key', key: 'text', text, raw: text })
			continue
		}

		const bracketed = parseBracketedPasteAt(raw, index, allowPending)
		if (bracketed?.pending) return { events, pending: raw.slice(index) }
		if (bracketed) {
			events.push(bracketed.event)
			index += bracketed.length
			continue
		}

		const mouse = parseMouseAt(raw, index, allowPending)
		if (mouse?.pending) return { events, pending: raw.slice(index) }
		if (mouse) {
			events.push(mouse.event)
			index += mouse.length
			continue
		}

		const csiU = parseCsiUAt(raw, index, allowPending)
		if (csiU?.pending) return { events, pending: raw.slice(index) }
		if (csiU) {
			events.push(csiU.event)
			index += csiU.length
			continue
		}

		const mapped = parseMappedSequenceAt(raw, index, allowPending)
		if (mapped?.pending) return { events, pending: raw.slice(index) }
		if (mapped) {
			events.push(mapped.event)
			index += mapped.length
			continue
		}

		events.push({ kind: 'key', key: current, raw: current })
		index += 1
	}
	return { events, pending: '' }
}

const KEY_SEQUENCES = new Map<string, string>([
	['\x1b[A', 'ArrowUp'],
	['\x1b[B', 'ArrowDown'],
	['\x1b[C', 'ArrowRight'],
	['\x1b[D', 'ArrowLeft'],
	['\x1b[1;2A', 'Shift+ArrowUp'],
	['\x1b[1;2B', 'Shift+ArrowDown'],
	['\x1b[1;2C', 'Shift+ArrowRight'],
	['\x1b[1;2D', 'Shift+ArrowLeft'],
	['\x1b[1;3A', 'Alt+ArrowUp'],
	['\x1b[1;3B', 'Alt+ArrowDown'],
	['\x1b[1;3C', 'Alt+ArrowRight'],
	['\x1b[1;3D', 'Alt+ArrowLeft'],
	['\x1b[1;5A', 'Ctrl+ArrowUp'],
	['\x1b[1;5B', 'Ctrl+ArrowDown'],
	['\x1b[1;5C', 'Ctrl+ArrowRight'],
	['\x1b[1;5D', 'Ctrl+ArrowLeft'],
	['\x1b[1;6A', 'Ctrl+Shift+ArrowUp'],
	['\x1b[1;6B', 'Ctrl+Shift+ArrowDown'],
	['\x1b[1;6C', 'Ctrl+Shift+ArrowRight'],
	['\x1b[1;6D', 'Ctrl+Shift+ArrowLeft'],
	['\x1b[1;9A', 'Cmd+ArrowUp'],
	['\x1b[1;9B', 'Cmd+ArrowDown'],
	['\x1b[1;9C', 'Cmd+ArrowRight'],
	['\x1b[1;9D', 'Cmd+ArrowLeft'],
	['\x1b[1;10A', 'Cmd+Shift+ArrowUp'],
	['\x1b[1;10B', 'Cmd+Shift+ArrowDown'],
	['\x1b[1;10C', 'Cmd+Shift+ArrowRight'],
	['\x1b[1;10D', 'Cmd+Shift+ArrowLeft'],
	['\x1b[5;5~', 'Ctrl+PageUp'],
	['\x1b[6;5~', 'Ctrl+PageDown'],
	['\x1b[5;3~', 'Alt+PageUp'],
	['\x1b[6;3~', 'Alt+PageDown'],
	['\x1b[5;2~', 'Shift+PageUp'],
	['\x1b[6;2~', 'Shift+PageDown'],
	['\x1b[5~', 'PageUp'],
	['\x1b[6~', 'PageDown'],
	['\x1b[H', 'Home'],
	['\x1b[F', 'End'],
	['\x1b[1;5H', 'Ctrl+Home'],
	['\x1b[1;5F', 'Ctrl+End'],
	['\x1b[1;6H', 'Ctrl+Shift+Home'],
	['\x1b[1;6F', 'Ctrl+Shift+End'],
	['\x1b[13;2u', 'Shift+Enter'],
	['\x1b[32;2u', 'Shift+Space'],
	['\x1b[32;5u', 'Ctrl+Space'],
	['\x1b[3~', 'Delete'],
	['\x00', 'Ctrl+Space'],
	['\x7f', 'Backspace'],
	['\b', 'Backspace'],
	['\r', 'Enter'],
	['\t', 'Tab'],
	['\x1b[Z', 'Shift+Tab'],
	['\x1b', 'Escape'],
	['\x03', 'Ctrl+C'],
	['\x01', 'Ctrl+A'],
	['\x04', 'Ctrl+D'],
	['\x07', 'Ctrl+G'],
	['\x12', 'Ctrl+R'],
	['\x13', 'Ctrl+S'],
	['\x18', 'Ctrl+X'],
	['\x0f', 'Ctrl+O'],
	['\x1a', 'Ctrl+Z'],
	['\x19', 'Ctrl+Y'],
	['\x16', 'Ctrl+V'],
	['\x06', 'Ctrl+F'],
	['\x10', 'Ctrl+P'],
	['\x14', 'Ctrl+T'],
	['\x0c', 'Ctrl+L'],
	['\x1b=', 'Alt+='],
	['\x1b[1;3P', 'Alt+F1'],
	['\x1bOP', 'F1'],
	['\x1bOQ', 'F2'],
	['\x1b[1;2Q', 'Shift+F2'],
	['\x1bOR', 'F3'],
	['\x1bOS', 'F4'],
	['\x1b[15~', 'F5'],
	['\x1b[15;2~', 'Shift+F5'],
	['\x1b[17~', 'F6'],
	['\x1b[18~', 'F7'],
	['\x1b[19~', 'F8'],
	['\x1b[20~', 'F9'],
	['\x1b[21~', 'F10'],
	['\x1b[21;2~', 'Shift+F10'],
	['\x1b[23~', 'F11'],
	['\x1b[24~', 'F12'],
	['\x1b[29~', 'Menu'],
])

function parseBracketedPaste(raw: string): InputEvent | undefined {
	const start = '\x1b[200~'
	const end = '\x1b[201~'
	if (!raw.startsWith(start) || !raw.endsWith(end)) return undefined
	return {
		kind: 'key',
		key: 'text',
		text: raw.slice(start.length, -end.length),
		raw,
	}
}

function parseBracketedPasteAt(
	raw: string,
	index: number,
	allowPending: boolean,
):
	| { readonly event: InputEvent; readonly length: number; readonly pending?: false }
	| { readonly pending: true }
	| undefined {
	const start = '\x1b[200~'
	const end = '\x1b[201~'
	if (!raw.startsWith(start, index)) {
		if (allowPending && start.startsWith(raw.slice(index))) return { pending: true }
		return undefined
	}
	const endIndex = raw.indexOf(end, index + start.length)
	if (endIndex < 0) return allowPending ? { pending: true } : undefined
	const sequence = raw.slice(index, endIndex + end.length)
	const event = parseBracketedPaste(sequence)
	return event ? { event, length: sequence.length } : undefined
}

function parseCsiU(raw: string): InputEvent | undefined {
	if (!raw.startsWith('\x1b[') || !raw.endsWith('u')) return undefined
	const match = raw.slice(2, -1).match(/^(\d+)(?:;(\d+))?$/)
	if (!match) return undefined
	const codepoint = Number.parseInt(match[1] ?? '', 10)
	if (!Number.isFinite(codepoint) || codepoint <= 0) return undefined
	const key = keyNameFromCodepoint(codepoint)
	const modifiers = Number.parseInt(match[2] ?? '1', 10)
	const prefix = modifierPrefix(modifiers)
	const normalizedKey =
		(prefix || ((modifiers - 1) & 1) !== 0) && /^[a-z]$/.test(key) ? key.toUpperCase() : key
	return {
		kind: 'key',
		key: `${prefix}${normalizedKey}`,
		...(prefix ? {} : { text: normalizedKey }),
		raw,
	}
}

function parseCsiUAt(
	raw: string,
	index: number,
	allowPending: boolean,
):
	| { readonly event: InputEvent; readonly length: number; readonly pending?: false }
	| { readonly pending: true }
	| undefined {
	const fragment = raw.slice(index)
	const prefix = '\x1b['
	if (!fragment.startsWith(prefix)) {
		if (allowPending && prefix.startsWith(fragment)) return { pending: true }
		return undefined
	}
	const terminator = raw.indexOf('u', index + prefix.length)
	if (terminator >= 0) {
		const sequence = raw.slice(index, terminator + 1)
		const event = parseCsiU(sequence)
		if (event) return { event, length: sequence.length }
	}
	if (allowPending && isPartialCsiU(fragment.slice(prefix.length))) return { pending: true }
	return undefined
}

function keyNameFromCodepoint(codepoint: number): string {
	switch (codepoint) {
		case 9:
			return 'Tab'
		case 13:
			return 'Enter'
		case 27:
			return 'Escape'
		default:
			return String.fromCodePoint(codepoint)
	}
}

function modifierPrefix(modifiers: number): string {
	const flags = modifiers - 1
	const parts: string[] = []
	if ((flags & 4) !== 0) parts.push('Ctrl')
	if ((flags & 8) !== 0) parts.push('Cmd')
	if ((flags & 2) !== 0) parts.push('Alt')
	if ((flags & 1) !== 0) parts.push('Shift')
	return parts.length > 0 ? `${parts.join('+')}+` : ''
}

function parseMouse(raw: string): InputEvent {
	const match = raw.slice(3).match(/^(\d+);(\d+);(\d+)([mM])$/)
	if (!match) return { kind: 'key', key: raw, raw }
	const button = Number.parseInt(match[1] ?? '0', 10)
	const col = Number.parseInt(match[2] ?? '1', 10)
	const row = Number.parseInt(match[3] ?? '1', 10)
	const suffix = match[4]
	const action = button === 64 || button === 65 ? 'wheel' : suffix === 'm' ? 'release' : 'press'
	return {
		kind: 'mouse',
		action,
		row,
		col,
		button,
	}
}

function parseMouseAt(
	raw: string,
	index: number,
	allowPending: boolean,
):
	| { readonly event: InputEvent; readonly length: number; readonly pending?: false }
	| { readonly pending: true }
	| undefined {
	const fragment = raw.slice(index)
	const prefix = '\x1b[<'
	if (!fragment.startsWith(prefix)) {
		if (allowPending && prefix.startsWith(fragment)) return { pending: true }
		return undefined
	}
	const bodyStart = index + prefix.length
	for (let cursor = bodyStart; cursor < raw.length; cursor++) {
		const char = raw[cursor]
		if (char === '\x1b') break
		if (char === 'm' || char === 'M') {
			const sequence = raw.slice(index, cursor + 1)
			return { event: parseMouse(sequence), length: sequence.length }
		}
	}
	if (allowPending && isPartialMouse(fragment.slice(prefix.length))) return { pending: true }
	return undefined
}

function isPartialCsiU(body: string): boolean {
	if (body === '') return true
	const parts = body.split(';')
	if (parts.length > 2) return false
	return parts.every((part) => /^\d*$/.test(part))
}

function isPartialMouse(body: string): boolean {
	if (body === '') return true
	const parts = body.split(';')
	if (parts.length > 3) return false
	return parts.every((part) => /^\d*$/.test(part))
}

function parseControlAt(
	raw: string,
	index: number,
): { readonly event: InputEvent; readonly length: number } | undefined {
	const char = raw[index] ?? ''
	if (char === '\x1b') return undefined
	const mapped = KEY_SEQUENCES.get(char)
	if (!mapped) return undefined
	return { event: { kind: 'key', key: mapped, raw: char }, length: 1 }
}

function parseMappedSequenceAt(
	raw: string,
	index: number,
	allowPending: boolean,
):
	| { readonly event: InputEvent; readonly length: number; readonly pending?: false }
	| { readonly pending: true }
	| undefined {
	const fragment = raw.slice(index)
	let best: { sequence: string; key: string } | undefined
	for (const [sequence, key] of KEY_SEQUENCES) {
		if (!fragment.startsWith(sequence)) continue
		if (!best || sequence.length > best.sequence.length) best = { sequence, key }
	}
	if (best) {
		return {
			event: { kind: 'key', key: best.key, raw: best.sequence },
			length: best.sequence.length,
		}
	}
	if (allowPending) {
		for (const sequence of KEY_SEQUENCES.keys()) {
			if (sequence.startsWith(fragment)) return { pending: true }
		}
	}
	return undefined
}
