import { describe, expect, test } from 'bun:test'
import { parseInputBuffer, parseInputEvents, TerminalInputParser } from './keymap.ts'

describe('parseInputBuffer', () => {
	test('maps common keys', () => {
		expect(parseInputBuffer(Buffer.from('\x1b[A')).key).toBe('ArrowUp')
		expect(parseInputBuffer(Buffer.from('\x1b[3~')).key).toBe('Delete')
		expect(parseInputBuffer(Buffer.from('\x13')).key).toBe('Ctrl+S')
		expect(parseInputBuffer(Buffer.from('\x18')).key).toBe('Ctrl+X')
		expect(parseInputBuffer(Buffer.from('\x16')).key).toBe('Ctrl+V')
		expect(parseInputBuffer(Buffer.from('\x14')).key).toBe('Ctrl+T')
		expect(parseInputBuffer(Buffer.from('\x1b[20~')).key).toBe('F9')
		expect(parseInputBuffer(Buffer.from('\x1b[1;2Q')).key).toBe('Shift+F2')
		expect(parseInputBuffer(Buffer.from('\x1b[6;5~')).key).toBe('Ctrl+PageDown')
		expect(parseInputBuffer(Buffer.from('\x1b[1;6C')).key).toBe('Ctrl+Shift+ArrowRight')
		expect(parseInputBuffer(Buffer.from('\x1b[1;5H')).key).toBe('Ctrl+Home')
		expect(parseInputBuffer(Buffer.from('\x1b[1;5F')).key).toBe('Ctrl+End')
		expect(parseInputBuffer(Buffer.from('\x1b[1;6H')).key).toBe('Ctrl+Shift+Home')
		expect(parseInputBuffer(Buffer.from('\x1b[1;6F')).key).toBe('Ctrl+Shift+End')
		expect(parseInputBuffer(Buffer.from('\x1b[1;3B')).key).toBe('Alt+ArrowDown')
		expect(parseInputBuffer(Buffer.from('\x1b[15;2~')).key).toBe('Shift+F5')
		expect(parseInputBuffer(Buffer.from('\x1b[21;2~')).key).toBe('Shift+F10')
		expect(parseInputBuffer(Buffer.from('\x1b[29~')).key).toBe('Menu')
		expect(parseInputBuffer(Buffer.from('\x1b[13;2u')).key).toBe('Shift+Enter')
	})

	test('maps text chunks to text events', () => {
		const event = parseInputBuffer(Buffer.from('abc'))
		expect(event).toEqual({ kind: 'key', key: 'text', text: 'abc', raw: 'abc' })
	})

	test('maps bracketed paste to a single text event', () => {
		const event = parseInputBuffer(Buffer.from('\x1b[200~=SUM(A1:A3)\n42\x1b[201~'))
		expect(event).toEqual({
			kind: 'key',
			key: 'text',
			text: '=SUM(A1:A3)\n42',
			raw: '\x1b[200~=SUM(A1:A3)\n42\x1b[201~',
		})
	})

	test('maps CSI-u modified keys', () => {
		expect(parseInputBuffer(Buffer.from('\x1b[49;5u')).key).toBe('Ctrl+1')
		expect(parseInputBuffer(Buffer.from('\x1b[61;3u')).key).toBe('Alt+=')
		expect(parseInputBuffer(Buffer.from('\x1b[108;6u')).key).toBe('Ctrl+Shift+L')
		expect(parseInputBuffer(Buffer.from('\x1b[102;5u')).key).toBe('Ctrl+F')
		expect(parseInputBuffer(Buffer.from('\x1b[115;5u')).key).toBe('Ctrl+S')
		expect(parseInputBuffer(Buffer.from('\x1b[115;6u')).key).toBe('Ctrl+Shift+S')
		expect(parseInputBuffer(Buffer.from('\x1b[118;7u')).key).toBe('Ctrl+Alt+V')
		expect(parseInputBuffer(Buffer.from('\x1b[96;5u')).key).toBe('Ctrl+`')
	})

	test('maps SGR mouse press, release, and wheel', () => {
		expect(parseInputBuffer(Buffer.from('\x1b[<0;12;5M'))).toMatchObject({
			kind: 'mouse',
			action: 'press',
			col: 12,
			row: 5,
		})
		expect(parseInputBuffer(Buffer.from('\x1b[<0;12;5m'))).toMatchObject({
			kind: 'mouse',
			action: 'release',
		})
		expect(parseInputBuffer(Buffer.from('\x1b[<64;12;5M'))).toMatchObject({
			kind: 'mouse',
			action: 'wheel',
			button: 64,
		})
	})

	test('maps multiple terminal events from one read', () => {
		expect(parseInputEvents(Buffer.from('\x1b[C\x1b[C\r')).map(eventKey)).toEqual([
			'ArrowRight',
			'ArrowRight',
			'Enter',
		])
		expect(parseInputEvents(Buffer.from('abc\r')).map(eventKey)).toEqual(['text', 'Enter'])
	})

	test('keeps split escape sequences pending in streaming mode', () => {
		const parser = new TerminalInputParser()
		expect(parser.push(Buffer.from('\x1b['))).toEqual([])
		expect(parser.push(Buffer.from('C'))).toEqual([
			{ kind: 'key', key: 'ArrowRight', raw: '\x1b[C' },
		])
	})

	test('flush emits a bare Escape after ambiguity window', () => {
		const parser = new TerminalInputParser()
		expect(parser.push(Buffer.from('\x1b'))).toEqual([])
		expect(parser.hasPending()).toBe(true)
		expect(parser.flush()).toEqual([{ kind: 'key', key: 'Escape', raw: '\x1b' }])
	})

	test('keeps split UTF-8 and bracketed paste intact in streaming mode', () => {
		const parser = new TerminalInputParser()
		const text = Buffer.from('é')
		expect(parser.push(text.subarray(0, 1))).toEqual([])
		expect(parser.push(text.subarray(1))).toEqual([
			{ kind: 'key', key: 'text', text: 'é', raw: 'é' },
		])

		const paste = '\x1b[200~=SUM(A1:A3)\n42\x1b[201~'
		expect(parser.push(Buffer.from(paste.slice(0, 12)))).toEqual([])
		expect(parser.push(Buffer.from(paste.slice(12)))).toEqual([
			{
				kind: 'key',
				key: 'text',
				text: '=SUM(A1:A3)\n42',
				raw: paste,
			},
		])
	})
})

function eventKey(event: ReturnType<typeof parseInputEvents>[number]): string {
	return event.kind === 'key' ? event.key : event.kind
}
