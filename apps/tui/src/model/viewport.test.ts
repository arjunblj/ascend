import { describe, expect, test } from 'bun:test'
import { indexToColumn } from '@ascend/sdk'
import { createSelection, moveSelection } from './selection.ts'
import {
	createViewport,
	ensureSelectionVisible,
	resizeViewport,
	viewportRange,
} from './viewport.ts'

describe('viewport model', () => {
	test('viewportRange emits A1 ranges for visible bounds', () => {
		const viewport = createViewport({ rows: 12, cols: 40 })
		expect(viewportRange(viewport, indexToColumn)).toBe('A1:C6')
	})

	test('ensureSelectionVisible scrolls to include active cell', () => {
		const viewport = createViewport({ rows: 12, cols: 40 })
		const selection = moveSelection(createSelection(), 20, 5)
		const next = ensureSelectionVisible(viewport, selection)
		expect(next.topRow).toBeGreaterThan(0)
		expect(next.leftCol).toBeGreaterThan(0)
	})

	test('resizeViewport preserves scroll origin and recomputes visible shape', () => {
		const viewport = { ...createViewport({ rows: 20, cols: 100 }), topRow: 10, leftCol: 4 }
		const next = resizeViewport(viewport, { rows: 10, cols: 50 })
		expect(next.topRow).toBe(10)
		expect(next.leftCol).toBe(4)
		expect(next.visibleRows).not.toBe(viewport.visibleRows)
	})
})
