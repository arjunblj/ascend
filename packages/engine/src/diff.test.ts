import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { EMPTY, numberValue, stringValue } from '@ascend/schema'
import { diffWorkbooks } from './diff.ts'
import { compareSnapshots, createSnapshot } from './snapshot.ts'

const sid = 0 as StyleId

function cell(value: ReturnType<typeof numberValue>, formula: string | null = null) {
	return { value, formula, styleId: sid }
}

describe('diffWorkbooks', () => {
	test('identical workbooks produce empty diff', () => {
		const a = createWorkbook()
		const s1 = a.addSheet('Sheet1')
		s1.cells.set(0, 0, cell(numberValue(1)))

		const b = createWorkbook()
		const s2 = b.addSheet('Sheet1')
		s2.cells.set(0, 0, cell(numberValue(1)))

		const diff = diffWorkbooks(a, b)
		expect(diff.sheets).toEqual([])
		expect(diff.namesAdded).toEqual([])
		expect(diff.namesRemoved).toEqual([])
		expect(diff.namesChanged).toEqual([])
	})

	test('detects added cells', () => {
		const a = createWorkbook()
		a.addSheet('Sheet1')

		const b = createWorkbook()
		const s = b.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(1)))

		const diff = diffWorkbooks(a, b)
		expect(diff.sheets).toHaveLength(1)
		expect(diff.sheets[0]?.cellsAdded).toEqual(['A1'])
		expect(diff.sheets[0]?.cellsRemoved).toEqual([])
	})

	test('detects removed cells', () => {
		const a = createWorkbook()
		const sa = a.addSheet('Sheet1')
		sa.cells.set(0, 0, cell(numberValue(1)))
		sa.cells.set(1, 0, cell(numberValue(2)))

		const b = createWorkbook()
		const sb = b.addSheet('Sheet1')
		sb.cells.set(0, 0, cell(numberValue(1)))

		const diff = diffWorkbooks(a, b)
		expect(diff.sheets).toHaveLength(1)
		expect(diff.sheets[0]?.cellsRemoved).toEqual(['A2'])
	})

	test('detects changed values', () => {
		const a = createWorkbook()
		const sa = a.addSheet('Sheet1')
		sa.cells.set(0, 0, cell(numberValue(10)))

		const b = createWorkbook()
		const sb = b.addSheet('Sheet1')
		sb.cells.set(0, 0, cell(numberValue(99)))

		const diff = diffWorkbooks(a, b)
		expect(diff.sheets).toHaveLength(1)
		const changes = diff.sheets[0]?.cellsChanged
		expect(changes).toHaveLength(1)
		expect(changes?.[0]?.before).toEqual(numberValue(10))
		expect(changes?.[0]?.after).toEqual(numberValue(99))
	})

	test('detects formula changes', () => {
		const a = createWorkbook()
		const sa = a.addSheet('Sheet1')
		sa.cells.set(0, 0, cell(EMPTY, 'SUM(B1:B5)'))

		const b = createWorkbook()
		const sb = b.addSheet('Sheet1')
		sb.cells.set(0, 0, cell(EMPTY, 'SUM(B1:B10)'))

		const diff = diffWorkbooks(a, b)
		const changes = diff.sheets[0]?.cellsChanged
		expect(changes).toHaveLength(1)
		expect(changes?.[0]?.formulaBefore).toBe('SUM(B1:B5)')
		expect(changes?.[0]?.formulaAfter).toBe('SUM(B1:B10)')
	})

	test('detects added sheets', () => {
		const a = createWorkbook()
		a.addSheet('Sheet1')

		const b = createWorkbook()
		b.addSheet('Sheet1')
		const newSheet = b.addSheet('Sheet2')
		newSheet.cells.set(0, 0, cell(stringValue('x')))

		const diff = diffWorkbooks(a, b)
		const added = diff.sheets.find((s) => s.name === 'Sheet2')
		expect(added).toBeDefined()
		expect(added?.cellsAdded).toEqual(['A1'])
	})

	test('detects removed sheets', () => {
		const a = createWorkbook()
		a.addSheet('Sheet1')
		const removedSheet = a.addSheet('Sheet2')
		removedSheet.cells.set(0, 0, cell(numberValue(5)))

		const b = createWorkbook()
		b.addSheet('Sheet1')

		const diff = diffWorkbooks(a, b)
		const removed = diff.sheets.find((s) => s.name === 'Sheet2')
		expect(removed).toBeDefined()
		expect(removed?.cellsRemoved).toEqual(['A1'])
	})

	test('detects named range changes', () => {
		const a = createWorkbook()
		a.addSheet('Sheet1')
		a.definedNames.set('OldName', 'Sheet1!A1')
		a.definedNames.set('Changed', 'Sheet1!A1')

		const b = createWorkbook()
		b.addSheet('Sheet1')
		b.definedNames.set('NewName', 'Sheet1!B1')
		b.definedNames.set('Changed', 'Sheet1!B2')

		const diff = diffWorkbooks(a, b)
		expect(diff.namesAdded).toEqual(['NewName'])
		expect(diff.namesRemoved).toEqual(['OldName'])
		expect(diff.namesChanged).toEqual(['Changed'])
	})
})

describe('snapshot round-trip', () => {
	test('createSnapshot captures workbook state', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(42)))
		s.cells.set(0, 1, cell(EMPTY, 'A1*2'))
		wb.definedNames.set('Total', 'Sheet1!A1')

		const snap = createSnapshot(wb)
		expect(snap.sheets).toHaveLength(1)
		expect(snap.sheets[0]?.name).toBe('Sheet1')
		expect(snap.sheets[0]?.cells.A1?.value).toEqual(numberValue(42))
		expect(snap.sheets[0]?.cells.B1?.formula).toBe('A1*2')
		expect(snap.names.Total).toBe('Sheet1!A1')
	})

	test('createSnapshot keys sheet-scoped names with the sheet name', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		wb.definedNames.set('LocalTotal', 'Sheet1!A1', { kind: 'sheet', sheetId: s.id })

		const snap = createSnapshot(wb)
		expect(snap.names['Sheet1!LocalTotal']).toBe('Sheet1!A1')
	})

	test('compareSnapshots detects changes', () => {
		const wb1 = createWorkbook()
		const s1 = wb1.addSheet('Sheet1')
		s1.cells.set(0, 0, cell(numberValue(1)))

		const snap1 = createSnapshot(wb1)

		const wb2 = createWorkbook()
		const s2 = wb2.addSheet('Sheet1')
		s2.cells.set(0, 0, cell(numberValue(2)))
		s2.cells.set(1, 0, cell(numberValue(3)))

		const snap2 = createSnapshot(wb2)

		const diff = compareSnapshots(snap1, snap2)
		expect(diff.sheets).toHaveLength(1)
		expect(diff.sheets[0]?.cellsAdded).toEqual(['A2'])
		expect(diff.sheets[0]?.cellsChanged).toHaveLength(1)
		expect(diff.sheets[0]?.cellsChanged[0]?.ref).toBe('A1')
	})

	test('compareSnapshots ignores timestamp-only differences', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, cell(numberValue(1)))

		const snap = createSnapshot(wb)
		const sameButLater = { ...snap, timestamp: snap.timestamp + 1000 }

		const diff = compareSnapshots(snap, sameButLater)
		expect(diff.sheets).toEqual([])
		expect(diff.namesAdded).toEqual([])
		expect(diff.namesRemoved).toEqual([])
		expect(diff.namesChanged).toEqual([])
	})
})
