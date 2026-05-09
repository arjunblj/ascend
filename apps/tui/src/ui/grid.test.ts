import { describe, expect, test } from 'bun:test'
import stripAnsi from 'strip-ansi'
import { createSelection } from '../model/selection.ts'
import { visibleLength } from '../render/ansi-text.ts'
import type {
	GridSemanticCell,
	GridSemanticFlag,
	GridSemanticModel,
	ViewportState,
} from '../runtime/types.ts'
import { renderGrid } from './grid.ts'

describe('renderGrid', () => {
	test('keeps wide Unicode cell content inside the column width', () => {
		const lines = renderGrid({
			viewport: viewport(4),
			selection: createSelection(),
			data: {
				range: 'A1:B1',
				cells: [
					{
						ref: 'A1',
						row: 0,
						col: 0,
						value: { kind: 'string', value: '東京AB' },
					},
					{
						ref: 'B1',
						row: 0,
						col: 1,
						value: { kind: 'string', value: 'Z' },
					},
				],
			},
			width: 32,
			height: 3,
			editBuffer: '',
			editing: false,
			showFormulas: false,
		})
		const text = stripAnsi(lines[1] ?? '')
		const cells = text.slice(6).split('|')
		expect(visibleLength(cells[1] ?? '')).toBe(4)
		expect(cells[2]?.trim()).toBe('Z')
	})

	test('right-aligns wide numeric-looking content by display width', () => {
		const lines = renderGrid({
			viewport: viewport(5),
			selection: createSelection(1, 0),
			data: {
				range: 'A2:A2',
				cells: [
					{
						ref: 'A2',
						row: 1,
						col: 0,
						value: { kind: 'number', value: 42 },
					},
				],
			},
			width: 24,
			height: 4,
			editBuffer: '漢42',
			editing: true,
			showFormulas: false,
		})
		const text = stripAnsi(lines[2] ?? '')
		const cell = text.slice(6).split('|')[1] ?? ''
		expect(visibleLength(cell)).toBe(5)
		expect(cell.trim()).toBe('漢42')
	})

	test('sanitizes terminal control sequences in cell text', () => {
		const lines = renderGrid({
			viewport: viewport(16),
			selection: createSelection(),
			data: {
				range: 'A1:A1',
				cells: [
					{
						ref: 'A1',
						row: 0,
						col: 0,
						value: { kind: 'string', value: 'safe\x1b[2J\tname\nnext' },
					},
				],
			},
			width: 48,
			height: 3,
			editBuffer: '',
			editing: false,
			showFormulas: false,
		})
		const text = stripAnsi(lines.join('\n'))
		expect(text).not.toContain('\x1b[2J')
		expect(text).toContain('safe name next')
	})

	test('renders native ASCII markers for Excel semantic cell state', () => {
		const lines = renderGrid({
			viewport: viewport(14),
			selection: createSelection(0, 0),
			data: {
				range: 'A1:B1',
				cells: [
					{ ref: 'A1', row: 0, col: 0, value: { kind: 'string', value: 'Region' } },
					{ ref: 'B1', row: 0, col: 1, value: { kind: 'number', value: 42 } },
				],
			},
			width: 48,
			height: 3,
			editBuffer: '',
			editing: false,
			showFormulas: false,
			semantics: semanticModel([
				['A1', ['filterActive', 'sortAsc', 'comment', 'validationDropdown', 'protected']],
				['B1', ['validationInvalid', 'conditionalFormat', 'hyperlink']],
			]),
		})

		const text = stripAnsi(lines.join('\n'))
		expect(text).toContain('A F A1')
		expect(text).toContain('Reg~')
		expect(text).toContain('FA1cdRO')
		expect(text).toContain('42 !*@')
	})
})

function viewport(width: number): ViewportState {
	return {
		topRow: 0,
		leftCol: 0,
		visibleRows: 3,
		visibleCols: 2,
		columnWidths: [width, width],
		overscanRows: 0,
		overscanCols: 0,
	}
}

function semanticModel(
	entries: readonly [string, readonly GridSemanticFlag[]][],
): GridSemanticModel {
	const cells = new Map<string, GridSemanticCell>()
	for (const [ref, flags] of entries) {
		const match = /^([A-Z]+)(\d+)$/.exec(ref)
		if (!match) continue
		const col = columnIndex(match[1] ?? 'A')
		const row = Number(match[2] ?? '1') - 1
		cells.set(ref, { ref, row, col, flags })
	}
	return { cells, frozenRows: 1, frozenCols: 1, protected: true, activeFilterRanges: ['A1:B2'] }
}

function columnIndex(label: string): number {
	let index = 0
	for (const char of label) index = index * 26 + char.charCodeAt(0) - 64
	return index - 1
}
