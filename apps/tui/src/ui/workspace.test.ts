import { describe, expect, test } from 'bun:test'
import { EMPTY } from '@ascend/schema'
import stripAnsi from 'strip-ansi'
import { createWorkspace } from '../runtime/workspace.ts'
import { renderWorkspace } from './workspace.ts'

describe('renderWorkspace', () => {
	test('summarizes only hydrated selected cells for huge selections', () => {
		const frame = renderWorkspace({
			size: { rows: 14, cols: 140 },
			workspace: createWorkspace(),
			sheetNames: ['Sheet1'],
			activeSheetIndex: 0,
			sheetName: 'Sheet1',
			mode: 'ready',
			selection: {
				active: { row: 999_999, col: 19 },
				anchor: { row: 0, col: 0 },
				kind: 'range',
			},
			viewport: {
				topRow: 0,
				leftCol: 0,
				visibleRows: 3,
				visibleCols: 2,
				columnWidths: [10, 10],
				overscanRows: 0,
				overscanCols: 0,
			},
			data: {
				ref: {
					start: { row: 0, col: 0 },
					end: { row: 2, col: 1 },
				},
				rowCount: 3,
				colCount: 2,
				cells: [
					{ ref: 'A1', row: 0, col: 0, value: { kind: 'number', value: 2 } },
					{ ref: 'B1', row: 0, col: 1, value: { kind: 'number', value: 3 } },
					{ ref: 'A2', row: 1, col: 0, value: EMPTY },
				],
			},
			editBuffer: '',
			formulaBarContent: '',
			editCursor: 0,
			commandPalette: { query: '', selectedIndex: 0 },
			activeDialog: undefined,
			contextMenu: undefined,
			inspectorLines: [],
			showFormulas: false,
			message: '',
			dirty: false,
			perfSummary: '',
		})

		const text = stripAnsi(frame.lines.join('\n'))
		expect(text).toContain('Count 2')
		expect(text).toContain('Sum 5')
	})

	test('keeps long formula edit cursors visible in the formula bar', () => {
		const frame = renderWorkspace({
			...baseWorkspaceInput(),
			size: { rows: 12, cols: 42 },
			mode: 'editing',
			editBuffer: '=SUM(A1:A999999)+SUM(B1:B999999)',
			formulaBarContent: '=SUM(A1:A999999)+SUM(B1:B999999)',
			editCursor: 33,
		})

		const text = stripAnsi(frame.lines[2] ?? '')
		expect(text).toContain('B999999')
		expect(frame.cursor?.visible).toBe(true)
		expect(frame.cursor?.col).toBeLessThanOrEqual(42)
	})
})

function baseWorkspaceInput(): Parameters<typeof renderWorkspace>[0] {
	return {
		size: { rows: 14, cols: 100 },
		workspace: {
			...createWorkspace(),
			fileHub: { ...createWorkspace().fileHub, visible: false },
			focusedRegion: 'grid',
		},
		sheetNames: ['Sheet1'],
		activeSheetIndex: 0,
		sheetName: 'Sheet1',
		mode: 'ready',
		selection: {
			active: { row: 0, col: 0 },
			anchor: { row: 0, col: 0 },
			kind: 'cell',
		},
		viewport: {
			topRow: 0,
			leftCol: 0,
			visibleRows: 3,
			visibleCols: 2,
			columnWidths: [10, 10],
			overscanRows: 0,
			overscanCols: 0,
		},
		data: undefined,
		editBuffer: '',
		formulaBarContent: '',
		editCursor: 0,
		commandPalette: { query: '', selectedIndex: 0 },
		activeDialog: undefined,
		contextMenu: undefined,
		inspectorLines: [],
		showFormulas: false,
		message: '',
		dirty: false,
		perfSummary: '',
	}
}
