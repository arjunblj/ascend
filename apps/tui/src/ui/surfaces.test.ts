import { describe, expect, test } from 'bun:test'
import stripAnsi from 'strip-ansi'
import { renderCommandPalette } from './command-palette.ts'
import { renderContextMenu } from './context-menu.ts'
import { renderDialog } from './dialog.ts'
import { renderFileHub } from './file-hub.ts'
import { renderRibbon } from './ribbon.ts'

describe('Excel-core TUI surfaces', () => {
	test('keeps command search discoverable in narrow ribbon layouts', () => {
		const text = stripAnsi(renderRibbon(54))

		expect(text).toContain('Alt/F10')
		expect(text).toContain('/ Search')
		expect(text.length).toBe(54)
	})

	test('renders File backstage with keyboard selection and safe labels', () => {
		const lines = renderFileHub(
			{
				visible: true,
				section: 'recent',
				query: '',
				selectedIndex: 1,
				entries: [
					{ label: 'Quarterly plan', detail: 'today', pinned: true },
					{ label: 'Unsafe\x1b[2J workbook\nname', detail: 'missing path', missing: true },
				],
			},
			72,
			10,
		)
		const text = stripAnsi(lines.join('\n'))

		expect(text).toContain('File backstage')
		expect(text).toContain('>   Unsafe workbook name')
		expect(text).toContain('missing')
		expect(text).not.toContain('\x1b[2J')
		expect(text).toContain('Up/Down Select')
	})

	test('shows command palette selected index and run fallback', () => {
		const lines = renderCommandPalette({ query: 'format', selectedIndex: 0 }, 96, 6)
		const text = stripAnsi(lines.join('\n'))

		expect(text).toContain('Search commands: format')
		expect(text).toContain('> ')
		expect(text).toContain('Enter Run')
		expect(text).toContain('Esc Close')
	})

	test('renders context menu as a selected keyboard surface', () => {
		const lines = renderContextMenu(
			{
				target: 'cell',
				address: 'B2',
				selectedIndex: 1,
				items: [
					{ id: 'copy', title: 'Copy', command: 'copy', shortcut: 'Ctrl+C' },
					{ id: 'format', title: 'Format Cells', command: 'format', shortcut: 'Ctrl+1' },
				],
			},
			72,
			8,
		)
		const text = stripAnsi(lines.join('\n'))

		expect(text).toContain('Context: cell B2')
		expect(text).toContain('> Format Cells  Ctrl+1')
		expect(text).toContain('Enter Run')
		expect(text).toContain('Esc Close')
	})

	test('sanitizes dialog labels and values without losing Excel-like affordances', () => {
		const lines = renderDialog(
			{
				id: 'find-replace',
				title: 'Find and Replace',
				activeField: 0,
				fields: [
					{
						name: 'findText',
						label: 'Find\x1b[2J what\nnow',
						kind: 'text',
						required: true,
						options: [],
						value: 'unsafe\x1b[3J value\ncell',
					},
					{
						name: 'matchCase',
						label: 'Match case',
						kind: 'boolean',
						required: false,
						options: [],
						value: 'true',
					},
				],
			},
			88,
			12,
		)
		const text = stripAnsi(lines.join('\n'))

		expect(text).toContain('Find and Replace')
		expect(text).toContain('> Find what now *: unsafe value cell')
		expect(text).toContain('Match case: [x]')
		expect(text).toContain('Tab Move')
		expect(text).not.toContain('\x1b[2J')
		expect(text).not.toContain('\x1b[3J')
	})
})
