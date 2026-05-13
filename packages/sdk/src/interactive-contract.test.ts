import { describe, expect, test } from 'bun:test'
import { unlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
	DEFAULT_STYLE_ID,
	parseA1,
	type RangeRef,
	rangeIntersects,
	sqrefIntersects,
} from '@ascend/core'
import { numberValue } from '@ascend/schema'
import type { InteractiveViewportCell, InteractiveViewportPatch } from './index.ts'
import { AscendSession, AscendWorkbook } from './index.ts'
import { buildMutationJournal } from './journal.ts'

describe('interactive client contract', () => {
	test('viewport reads carry generation and load snapshot tokens', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'label' },
					{ ref: 'B2', value: 1 },
				],
			},
		])
		const before = wb.readWindowCompact('Sheet1', 'A1:B2', { includeRefs: true })
		if (!before) throw new Error('missing viewport')
		expect(before.snapshot.load.isPartial).toBe(false)
		expect(before.snapshot.generations).toEqual({
			workbook: 1,
			sheetMetadata: 0,
			formulas: 1,
			styles: 0,
		})
		expect(wb.validateReadSnapshot(before.snapshot)).toEqual({
			ok: true,
			current: true,
			expected: before.snapshot,
			receivedToken: before.snapshot.token,
		})

		const apply = wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B2', value: 7 }] },
		])
		expect(apply.errors).toEqual([])
		const stale = wb.validateReadSnapshot(before.snapshot)
		expect(stale.ok).toBe(false)
		if (stale.ok) throw new Error('snapshot unexpectedly current')
		expect(stale.error.details).toMatchObject({
			rule: 'stale-read-snapshot',
			receivedToken: before.snapshot.token,
			expectedGenerations: apply.generations,
		})

		const fresh = wb.readWindowCompact('Sheet1', 'A1:B2', { includeRefs: true })
		if (!fresh) throw new Error('missing fresh viewport')
		expect(fresh.snapshot.generations).toEqual(apply.generations)
		const patchedCells = before.cells.map((cell) =>
			cell.ref === 'B2' ? { ...cell, value: { kind: 'number' as const, value: 7 } } : cell,
		)
		expect(patchedCells).toEqual(fresh.cells)
	})

	test('interactive session returns viewport-shaped semantic payloads and stable patch tokens', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'C1', value: 'Amount' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 2 },
					{ ref: 'C2', value: 20 },
					{ ref: 'A3', value: 'East' },
					{ ref: 'B3', value: 3 },
					{ ref: 'C3', value: 30 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'H2', formula: 'B2*C2' },
			{ op: 'setStyle', sheet: 'Sheet1', range: 'B2:B2', style: { numberFormat: '0.00' } },
			{ op: 'setComment', sheet: 'Sheet1', ref: 'D4', text: 'review', author: 'agent' },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'E5', url: 'https://example.com' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A2:A5',
				rule: { type: 'list', formula1: '"West,East"', allowBlank: true },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'B2:B5',
				rule: { type: 'cellIs', operator: 'greaterThan', formula: '1', priority: 1 },
			},
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'F1:G1' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C5', column: 1, values: ['Open'] },
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:C3', name: 'Sales', hasHeaders: true },
			{ op: 'freezePane', sheet: 'Sheet1', row: 1, col: 1 },
			{ op: 'setRowHeight', sheet: 'Sheet1', row: 2, height: 24 },
		])

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const viewport = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 5,
			colCount: 8,
		})

		expect(viewport.load.mode).toBe('formula')
		expect(viewport.load.richSheetMetadataHydrated).toBe(true)
		expect(viewport.rowCount).toBe(5)
		expect(viewport.colCount).toBe(8)
		expect(viewport.flatValues[0]).toBe('Region')
		expect(viewport.displayText[0]).toBe('Region')
		expect(viewport.frozen).toEqual({ rows: 1, cols: 1 })
		expect(viewport.merges).toEqual([{ start: { row: 0, col: 5 }, end: { row: 0, col: 6 } }])
		expect(viewport.comments).toEqual([
			expect.objectContaining({ ref: 'D4', text: 'review', author: 'agent' }),
		])
		expect(viewport.hyperlinks[0]).toMatchObject({ ref: 'E5', target: 'https://example.com' })
		expect(viewport.dataValidations).toHaveLength(1)
		expect(viewport.conditionalFormats).toHaveLength(1)
		expect(viewport.tables.map((table) => table.name)).toEqual(['Sales'])
		expect(viewport.autoFilter?.ref).toBe('A1:C3')
		expect(viewport.rowLayout).toEqual([{ index: 2, size: 24 }])

		const formulaCell = viewport.cells.find((cell) => cell.ref === 'H2')
		expect(formulaCell?.formula).toBe('B2*C2')
		expect(formulaCell?.flags.formula).toBe(true)
		expect(viewport.cells.find((cell) => cell.ref === 'A2')?.flags.validation).toBe(true)
		expect(viewport.cells.find((cell) => cell.ref === 'B2')?.flags.conditionalFormat).toBe(true)
		expect(viewport.cells.find((cell) => cell.ref === 'A2')?.flags.table).toBe(true)

		const repeated = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 5,
			colCount: 8,
			changedSince: viewport.changeToken,
		})
		expect(repeated.patch).toMatchObject({
			baseToken: viewport.changeToken,
			changedCells: [],
			removedRefs: [],
			byteLength: 36,
		})
		expect(repeated.patch?.changeToken).toBe(repeated.changeToken)
		session.close()
	})

	test('interactive viewport overlay cache refreshes after metadata edits', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'old' },
					{ ref: 'B1', value: 'new' },
				],
			},
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1:A1',
				rule: { type: 'list', formula1: '"old,new"', allowBlank: true },
			},
		])
		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})
		expect(before.cells.find((cell) => cell.ref === 'A1')?.flags.validation).toBe(true)
		expect(before.cells.find((cell) => cell.ref === 'B1')?.flags.validation).toBe(false)

		const edit = await session.apply([
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'B1:B1',
				rule: { type: 'list', formula1: '"old,new"', allowBlank: true },
			},
		])
		expect(edit.apply.errors).toEqual([])
		const after = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})
		expect(after.cells.find((cell) => cell.ref === 'A1')?.flags.validation).toBe(true)
		expect(after.cells.find((cell) => cell.ref === 'B1')?.flags.validation).toBe(true)
		session.close()
	})

	test('interactive viewport overlay indexes match exhaustive metadata scans', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'C1', value: 'Amount' },
					{ ref: 'D1', value: 'Status' },
					{ ref: 'B2', value: 2 },
					{ ref: 'C3', value: 30 },
					{ ref: 'D4', value: 'Open' },
					{ ref: 'E5', value: 'Outside' },
				],
			},
			{ op: 'setComment', sheet: 'Sheet1', ref: 'B2', text: 'review' },
			{ op: 'setComment', sheet: 'Sheet1', ref: 'G7', text: 'outside' },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'C3', url: 'https://example.com' },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'H1', url: 'https://outside.example' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'B2:B4 D2:E2 H1:H2',
				rule: { type: 'list', formula1: '"Open,Closed"', allowBlank: true },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'C3:D5 A8:A9',
				rule: { type: 'expression', formula: 'C3>10', priority: 1 },
			},
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'B2:C3' },
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'F6:G7' },
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:D4', name: 'Sales', hasHeaders: true },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:D4', column: 3, values: ['Open'] },
		])
		const model = wb.getWorkbookModel()
		const sheet = model.getSheet('Sheet1')
		if (!sheet) throw new Error('Sheet1 missing')
		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		try {
			for (const request of [
				{ topRow: 1, leftCol: 1, rowCount: 3, colCount: 3 },
				{ topRow: 4, leftCol: 3, rowCount: 3, colCount: 2 },
				{ topRow: 0, leftCol: 7, rowCount: 2, colCount: 1 },
			]) {
				const viewport = session.readViewport({
					sheet: 'Sheet1',
					...request,
				})
				const range = viewport.viewport
				expect(viewport.comments.map((comment) => comment.ref).sort()).toEqual(
					[...sheet.comments.keys()].filter((ref) => cellRefOverlapsRange(ref, range)).sort(),
				)
				expect(viewport.hyperlinks.map((hyperlink) => hyperlink.ref).sort()).toEqual(
					[...sheet.hyperlinks.keys()].filter((ref) => cellRefOverlapsRange(ref, range)).sort(),
				)
				expect(viewport.merges.map(rangeKey).sort()).toEqual(
					sheet.merges
						.filter((merge) => rangeIntersects(merge, range))
						.map(rangeKey)
						.sort(),
				)
				expect(viewport.dataValidations.map((validation) => validation.sqref).sort()).toEqual(
					sheet.dataValidations
						.filter((validation) => sqrefIntersects(validation.sqref, range))
						.map((validation) => validation.sqref)
						.sort(),
				)
				expect(viewport.conditionalFormats.map((format) => format.sqref).sort()).toEqual(
					sheet.conditionalFormats
						.filter((format) => sqrefIntersects(format.sqref, range))
						.map((format) => format.sqref)
						.sort(),
				)
				expect(viewport.tables.map((table) => table.name).sort()).toEqual(
					sheet.tables
						.filter((table) => rangeIntersects(table.ref, range))
						.map((table) => table.name)
						.sort(),
				)
				expect(viewport.autoFilter?.ref ?? null).toBe(
					sheet.autoFilter && sqrefIntersects(sheet.autoFilter.ref, range)
						? sheet.autoFilter.ref
						: null,
				)
				for (const cell of viewport.cells) {
					expect(cell.flags).toMatchObject({
						comment: sheet.comments.has(cell.ref),
						hyperlink: sheet.hyperlinks.has(cell.ref),
						merged: sheet.merges.some((merge) => cellInRange(cell.row, cell.col, merge)),
						validation: sheet.dataValidations.some((validation) =>
							sqrefIntersects(validation.sqref, cellRange(cell.row, cell.col)),
						),
						conditionalFormat: sheet.conditionalFormats.some((format) =>
							sqrefIntersects(format.sqref, cellRange(cell.row, cell.col)),
						),
						table: sheet.tables.some((table) => cellInRange(cell.row, cell.col, table.ref)),
					})
				}
			}
		} finally {
			session.close()
		}
	})

	test('interactive viewport resolves shared formula member text', async () => {
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().getSheet('Sheet1')
		if (!sheet) throw new Error('missing sheet')
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: DEFAULT_STYLE_ID })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: DEFAULT_STYLE_ID })
		sheet.cells.set(0, 0, {
			value: numberValue(20),
			formula: 'B1*2',
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'A1' },
		})
		sheet.cells.set(1, 0, {
			value: numberValue(40),
			formula: null,
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
		})

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const viewport = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 2,
			colCount: 2,
		})
		expect(viewport.cells.find((cell) => cell.ref === 'A1')?.formula).toBe('B1*2')
		expect(viewport.cells.find((cell) => cell.ref === 'A2')?.formula).toBe('B2*2')
		session.close()
	})

	test('interactive session applies edits and returns pull viewport patches', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		expect(session.editReadiness()).toMatchObject({
			ready: false,
			preparing: false,
			write: null,
			promotedToFull: false,
		})
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})
		const edit = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
		])
		expect(edit.apply.errors).toEqual([])
		expect(edit.timings.applyMs).toBeNumber()
		expect(edit.timings.totalMs).toBeGreaterThanOrEqual(edit.timings.applyMs)
		expect(edit.load.read).toMatchObject({ mode: 'formula', isPartial: true })
		expect(edit.load.write).toMatchObject({ mode: 'full', isPartial: false })
		expect(edit.load.promotedToFull).toBe(true)
		expect(edit.recalc?.changed).toEqual(['Sheet1!B1'])
		expect(edit.generation.session).toBe(1)

		const after = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
			changedSince: before.changeToken,
		})
		expect(after.patch?.baseToken).toBe(before.changeToken)
		expect(after.patch?.changedCells.map((cell) => cell.ref).sort()).toEqual(['A1', 'B1'])
		expect(after.cells.find((cell) => cell.ref === 'A1')?.flatValue).toBe(5)
		expect(after.cells.find((cell) => cell.ref === 'B1')?.flatValue).toBe(10)
		session.close()
	})

	test('interactive pull patches materialize to the same cells as a fresh viewport read', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'C1', value: 3 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		const edit = await session.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 4 },
					{ ref: 'D1', value: 8 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*3' },
			{ op: 'clearRange', sheet: 'Sheet1', range: 'C1:C1', what: 'all' },
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'D1:D1', format: '0.00' },
		])
		expect(edit.apply.errors).toEqual([])

		const patch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: before.changeToken,
		})
		expect(patch?.baseToken).toBe(before.changeToken)
		expect(patch?.changeToken).toMatch(/^1:\d+$/)
		expect(patch?.changedCells.map((cell) => cell.ref).sort()).toEqual(['A1', 'B1', 'D1'])
		expect(patch?.removedRefs).toEqual(['C1'])
		if (!patch) throw new Error('expected viewport patch')
		expect(patch.byteLength).toBe(
			JSON.stringify({ changedCells: patch.changedCells, removedRefs: patch.removedRefs }).length,
		)

		const fresh = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(materializeViewportPatch(before.cells, patch)).toEqual(interactiveCellMap(fresh.cells))
		session.close()
	})

	test('interactive pull patches stay sheet-scoped for names with apostrophes and bangs', async () => {
		const sheetName = "Q1's Data!"
		const wb = AscendWorkbook.create()
		expect(wb.renameSheet('Sheet1', sheetName).errors).toEqual([])
		wb.apply([
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'B1', value: 2 },
				],
			},
			{ op: 'setFormula', sheet: sheetName, ref: 'C1', formula: 'A1+B1' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const before = session.readViewport({
			sheet: sheetName,
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		const edit = await session.apply([
			{ op: 'setCells', sheet: sheetName, updates: [{ ref: 'A1', value: 5 }] },
			{ op: 'setNumberFormat', sheet: sheetName, range: 'D1:D1', format: '0.00' },
		])
		expect(edit.apply.errors).toEqual([])
		expect(edit.recalc?.changed).toEqual([`${sheetName}!C1`])

		const patch = session.readViewportPatch({
			sheet: sheetName,
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: before.changeToken,
		})
		if (!patch) throw new Error('expected special sheet name patch')
		expect(patch.changedCells.map((cell) => [cell.ref, cell.flatValue]).sort()).toEqual([
			['A1', 5],
			['C1', 7],
			['D1', null],
		])
		const fresh = session.readViewport({
			sheet: sheetName,
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(materializeViewportPatch(before.cells, patch)).toEqual(interactiveCellMap(fresh.cells))
		session.close()
	})

	test('interactive pull patches materialize fillFormula outputs and recalculated values', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 2 },
					{ ref: 'A2', value: 4 },
				],
			},
		])

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 2,
			colCount: 3,
		})
		const edit = await session.apply([
			{ op: 'fillFormula', sheet: 'Sheet1', range: 'B1:B2', formula: 'A1*2' },
			{ op: 'setStyle', sheet: 'Sheet1', range: 'C2:C2', style: { numberFormat: '0.00' } },
		])
		expect(edit.apply.errors).toEqual([])
		expect(edit.recalc?.changed).toEqual(['Sheet1!B1', 'Sheet1!B2'])

		const patch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 2,
			colCount: 3,
			changedSince: before.changeToken,
		})
		if (!patch) throw new Error('expected fillFormula patch')
		expect(patch.changedCells.map((cell) => cell.ref).sort()).toEqual(['B1', 'B2', 'C2'])
		const changed = new Map(patch.changedCells.map((cell) => [cell.ref, cell]))
		expect(changed.get('B1')).toMatchObject({ formula: 'A1*2', flatValue: 4 })
		expect(changed.get('B2')).toMatchObject({ formula: 'A2*2', flatValue: 8 })
		expect(changed.get('C2')).toMatchObject({ formula: null, flatValue: null })
		expect(changed.get('C2')?.styleId).not.toBe(0)
		const fresh = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 2,
			colCount: 3,
		})
		expect(materializeViewportPatch(before.cells, patch)).toEqual(interactiveCellMap(fresh.cells))
		session.close()
	})

	test('interactive explicit recalc patches pending formula results after deferred recalc edits', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), {
			mode: 'full',
			prepareEdits: true,
		})
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})
		const edit = await session.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] }],
			{ recalc: false },
		)
		expect(edit.apply.errors).toEqual([])
		expect(edit.apply.recalcRequired).toBe(true)
		expect(edit.recalc).toBeNull()

		const afterEdit = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
			changedSince: before.changeToken,
		})
		expect(afterEdit.patch?.changedCells.map((cell) => cell.ref)).toEqual(['A1'])
		expect(afterEdit.cells.find((cell) => cell.ref === 'B1')?.flatValue).toBe(2)
		if (!afterEdit.patch) throw new Error('expected deferred-edit patch')
		expect(materializeViewportPatch(before.cells, afterEdit.patch)).toEqual(
			interactiveCellMap(afterEdit.cells),
		)

		const recalc = await session.apply([], { recalc: true })
		expect(recalc.apply.errors).toEqual([])
		expect(recalc.recalc?.errors).toEqual([])
		expect(recalc.recalc?.changed).toEqual(['Sheet1!B1'])
		expect(recalc.generation.session).toBe(edit.generation.session + 1)

		const recalcPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
			changedSince: afterEdit.changeToken,
		})
		if (!recalcPatch) throw new Error('expected explicit-recalc patch')
		expect(recalcPatch.changedCells.map((cell) => [cell.ref, cell.flatValue])).toEqual([['B1', 10]])
		const fresh = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})
		expect(materializeViewportPatch(afterEdit.cells, recalcPatch)).toEqual(
			interactiveCellMap(fresh.cells),
		)
		session.close()
	})

	test('interactive explicit recalc honors prior metadata invalidation before resuming patches', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'C1', value: 'review' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), {
			mode: 'full',
			prepareEdits: true,
		})
		const base = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 3,
		})
		const edit = await session.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] }],
			{ recalc: false },
		)
		expect(edit.apply.errors).toEqual([])
		const afterEdit = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 3,
			changedSince: base.changeToken,
		})
		expect(afterEdit.patch?.changedCells.map((cell) => cell.ref)).toEqual(['A1'])

		const metadata = await session.apply([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'C1', text: 'needs review' },
		])
		expect(metadata.apply.errors).toEqual([])
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 3,
				changedSince: afterEdit.changeToken,
			}),
		).toBeNull()
		const afterMetadata = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 3,
			changedSince: afterEdit.changeToken,
		})
		expect(afterMetadata.patch).toBeUndefined()
		expect(afterMetadata.cells.find((cell) => cell.ref === 'C1')?.flags.comment).toBe(true)

		const recalc = await session.apply([], { recalc: true })
		expect(recalc.apply.errors).toEqual([])
		expect(recalc.recalc?.changed).toEqual(['Sheet1!B1'])
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 3,
				changedSince: afterEdit.changeToken,
			}),
		).toBeNull()
		const recalcPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 3,
			changedSince: afterMetadata.changeToken,
		})
		if (!recalcPatch) throw new Error('expected resumed explicit-recalc patch')
		expect(recalcPatch.changedCells.map((cell) => [cell.ref, cell.flatValue])).toEqual([['B1', 10]])
		const fresh = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 3,
		})
		expect(materializeViewportPatch(afterMetadata.cells, recalcPatch)).toEqual(
			interactiveCellMap(fresh.cells),
		)
		session.close()
	})

	test('interactive explicit recalc advances generations for calc freshness metadata only', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		wb.recalc()
		const internal = wb as unknown as {
			wb: {
				calcSettings: {
					fullCalcOnLoad?: boolean
					calcCompleted?: boolean
					calcOnSave?: boolean
					forceFullCalc?: boolean
				}
			}
		}
		internal.wb.calcSettings = {
			...internal.wb.calcSettings,
			fullCalcOnLoad: true,
			calcCompleted: false,
			calcOnSave: false,
			forceFullCalc: true,
		}

		const session = await AscendSession.open(wb.toBytes(), {
			mode: 'full',
			prepareEdits: true,
		})
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})
		const recalc = await session.apply([], { recalc: true })
		expect(recalc.apply.errors).toEqual([])
		expect(recalc.recalc?.errors).toEqual([])
		expect(recalc.recalc?.changed).toEqual([])
		expect(recalc.generation.session).toBe(before.generation.session + 1)
		expect(recalc.generation.workbook).toBeGreaterThan(before.generation.workbook)
		expect(recalc.generation.formulas).toBeGreaterThan(before.generation.formulas)

		const patch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
			changedSince: before.changeToken,
		})
		if (!patch) throw new Error('expected empty calc-freshness patch')
		expect(patch.changedCells).toEqual([])
		expect(patch.removedRefs).toEqual([])
		const fresh = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})
		expect(materializeViewportPatch(before.cells, patch)).toEqual(interactiveCellMap(fresh.cells))
		session.close()
	})

	test('interactive sessions resume patching after an invalidating metadata refresh', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'B1', value: 2 },
				],
			},
		])

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const prepared = await session.prepareEdits()
		expect(prepared.load.promotedToFull).toBe(true)
		const base = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})

		const valueEdit = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 10 }] },
		])
		expect(valueEdit.apply.errors).toEqual([])
		const valuePatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
			changedSince: base.changeToken,
		})
		if (!valuePatch) throw new Error('expected value patch')
		const afterValue = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})
		expect(materializeViewportPatch(base.cells, valuePatch)).toEqual(
			interactiveCellMap(afterValue.cells),
		)

		const metadataEdit = await session.apply([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'A1', text: 'reviewed' },
		])
		expect(metadataEdit.apply.errors).toEqual([])
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 2,
				changedSince: afterValue.changeToken,
			}),
		).toBeNull()
		const afterMetadata = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
			changedSince: afterValue.changeToken,
		})
		expect(afterMetadata.patch).toBeUndefined()
		expect(afterMetadata.cells.find((cell) => cell.ref === 'A1')?.flags.comment).toBe(true)

		const resumedEdit = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 20 }] },
		])
		expect(resumedEdit.apply.errors).toEqual([])
		const resumedPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
			changedSince: afterMetadata.changeToken,
		})
		if (!resumedPatch) throw new Error('expected resumed patch')
		expect(resumedPatch.changedCells.map((cell) => [cell.ref, cell.flatValue])).toEqual([
			['B1', 20],
		])
		const afterResume = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})
		expect(materializeViewportPatch(afterMetadata.cells, resumedPatch)).toEqual(
			interactiveCellMap(afterResume.cells),
		)
		session.close()
	})

	test('interactive patches stay truthful across recalc style edits and metadata invalidation', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'C1', value: 10 },
					{ ref: 'D1', value: 'note target' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const base = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})

		const firstEdit = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C1:C1', format: '0.00' },
		])
		expect(firstEdit.apply.errors).toEqual([])
		expect(firstEdit.recalc?.changed).toEqual(['Sheet1!B1'])
		const firstPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: base.changeToken,
		})
		if (!firstPatch) throw new Error('expected first patch')
		expect(firstPatch.changedCells.map((cell) => cell.ref).sort()).toEqual(['A1', 'B1', 'C1'])
		const afterFirst = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(materializeViewportPatch(base.cells, firstPatch)).toEqual(
			interactiveCellMap(afterFirst.cells),
		)

		const metadataEdit = await session.apply([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'D1', text: 'metadata changed' },
		])
		expect(metadataEdit.apply.errors).toEqual([])
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 4,
				changedSince: afterFirst.changeToken,
			}),
		).toBeNull()
		const afterMetadata = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: afterFirst.changeToken,
		})
		expect(afterMetadata.patch).toBeUndefined()

		const secondEdit = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] },
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C1:C1', format: '$0.00' },
		])
		expect(secondEdit.apply.errors).toEqual([])
		expect(secondEdit.recalc?.changed).toEqual(['Sheet1!B1'])
		const secondPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: afterMetadata.changeToken,
		})
		if (!secondPatch) throw new Error('expected second patch')
		expect(secondPatch.changedCells.map((cell) => cell.ref).sort()).toEqual(['A1', 'B1', 'C1'])
		const afterSecond = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(materializeViewportPatch(afterMetadata.cells, secondPatch)).toEqual(
			interactiveCellMap(afterSecond.cells),
		)
		session.close()
	})

	test('interactive structural viewport shifts force fresh snapshots instead of ref patches', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'r1-a' },
					{ ref: 'B1', value: 'r1-b' },
					{ ref: 'C1', value: 'r1-c' },
					{ ref: 'A2', value: 'r2-a' },
					{ ref: 'B2', value: 'r2-b' },
					{ ref: 'C2', value: 'r2-c' },
					{ ref: 'A3', value: 'r3-a' },
					{ ref: 'B3', value: 'r3-b' },
					{ ref: 'C3', value: 'r3-c' },
				],
			},
		])

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const base = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 2,
			colCount: 2,
		})

		const rowDelete = await session.apply([{ op: 'deleteRows', sheet: 'Sheet1', at: 0, count: 1 }])
		expect(rowDelete.apply.errors).toEqual([])
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 2,
				colCount: 2,
				changedSince: base.changeToken,
			}),
		).toBeNull()
		const afterRowDelete = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 2,
			colCount: 2,
			changedSince: base.changeToken,
		})
		expect(afterRowDelete.patch).toBeUndefined()
		expect(new Map(afterRowDelete.cells.map((cell) => [cell.ref, cell.flatValue]))).toEqual(
			new Map([
				['A1', 'r2-a'],
				['B1', 'r2-b'],
				['A2', 'r3-a'],
				['B2', 'r3-b'],
			]),
		)

		const colDelete = await session.apply([{ op: 'deleteCols', sheet: 'Sheet1', at: 0, count: 1 }])
		expect(colDelete.apply.errors).toEqual([])
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 2,
				colCount: 2,
				changedSince: afterRowDelete.changeToken,
			}),
		).toBeNull()
		const afterColDelete = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 2,
			colCount: 2,
			changedSince: afterRowDelete.changeToken,
		})
		expect(afterColDelete.patch).toBeUndefined()
		expect(new Map(afterColDelete.cells.map((cell) => [cell.ref, cell.flatValue]))).toEqual(
			new Map([
				['A1', 'r2-b'],
				['B1', 'r2-c'],
				['A2', 'r3-b'],
				['B2', 'r3-c'],
			]),
		)
		session.close()
	})

	test('interactive edit journal undo patches materialize back to the base viewport', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'C1', value: 3 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		const edit = await session.apply(
			[
				{
					op: 'setCells',
					sheet: 'Sheet1',
					updates: [
						{ ref: 'A1', value: 4 },
						{ ref: 'D1', value: 8 },
					],
				},
				{ op: 'clearRange', sheet: 'Sheet1', range: 'C1:C1', what: 'all' },
				{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'D1:D1', format: '0.00' },
			],
			{ journal: true },
		)
		expect(edit.apply.errors).toEqual([])
		expect(edit.apply.journal?.exact).toBe(true)

		const editPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: before.changeToken,
		})
		if (!editPatch) throw new Error('expected edit patch')
		const edited = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(materializeViewportPatch(before.cells, editPatch)).toEqual(
			interactiveCellMap(edited.cells),
		)

		const undo = await session.apply(edit.apply.journal?.inverseOps ?? [])
		expect(undo.apply.errors).toEqual([])
		const undoPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: edited.changeToken,
		})
		if (!undoPatch) throw new Error('expected undo patch')
		const restored = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(materializeViewportPatch(edited.cells, undoPatch)).toEqual(
			interactiveCellMap(restored.cells),
		)
		expect(semanticInteractiveCellMap(restored.cells)).toEqual(
			semanticInteractiveCellMap(before.cells),
		)
		session.close()
	})

	test('interactive undo resumes patching after metadata invalidates the edited snapshot', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'C1', value: 3 },
					{ ref: 'D1', value: 'metadata target' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		const edit = await session.apply(
			[
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 4 }] },
				{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C1:C1', format: '0.00' },
			],
			{ journal: true },
		)
		expect(edit.apply.errors).toEqual([])
		expect(edit.apply.journal?.exact).toBe(true)
		const edited = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})

		const metadata = await session.apply([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'D1', text: 'metadata changed' },
		])
		expect(metadata.apply.errors).toEqual([])
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 4,
				changedSince: edited.changeToken,
			}),
		).toBeNull()
		const afterMetadata = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: edited.changeToken,
		})
		expect(afterMetadata.patch).toBeUndefined()
		expect(afterMetadata.comments).toHaveLength(1)

		const undo = await session.apply(edit.apply.journal?.inverseOps ?? [])
		expect(undo.apply.errors).toEqual([])
		const undoPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: afterMetadata.changeToken,
		})
		if (!undoPatch) throw new Error('expected undo patch after fresh metadata snapshot')
		const restored = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(materializeViewportPatch(afterMetadata.cells, undoPatch)).toEqual(
			interactiveCellMap(restored.cells),
		)
		expect(restored.comments).toHaveLength(1)
		const restoredSemantic = semanticInteractiveCellMap(restored.cells)
		const beforeSemantic = semanticInteractiveCellMap(before.cells)
		expect(restoredSemantic.get('A1')).toEqual(beforeSemantic.get('A1'))
		expect(restoredSemantic.get('B1')).toEqual(beforeSemantic.get('B1'))
		expect(restoredSemantic.get('C1')).toEqual(beforeSemantic.get('C1'))
		expect(restoredSemantic.get('D1')?.flags.comment).toBe(true)
		session.close()
	})

	test('interactive deferred recalc undo patches stay truthful after metadata invalidation', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'D1', value: 'metadata target' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const base = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		const edit = await session.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] }],
			{ recalc: false, journal: true },
		)
		expect(edit.apply.errors).toEqual([])
		expect(edit.apply.recalcRequired).toBe(true)
		expect(edit.recalc).toBeNull()
		expect(edit.apply.journal?.exact).toBe(true)
		const editPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: base.changeToken,
		})
		if (!editPatch) throw new Error('expected deferred edit patch')
		expect(editPatch.changedCells.map((cell) => [cell.ref, cell.flatValue])).toEqual([['A1', 5]])
		const afterEdit = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(afterEdit.cells.find((cell) => cell.ref === 'B1')?.flatValue).toBe(2)
		expect(materializeViewportPatch(base.cells, editPatch)).toEqual(
			interactiveCellMap(afterEdit.cells),
		)

		const recalc = await session.apply([], { recalc: true })
		expect(recalc.apply.errors).toEqual([])
		expect(recalc.recalc?.changed).toEqual(['Sheet1!B1'])
		const recalcPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: afterEdit.changeToken,
		})
		if (!recalcPatch) throw new Error('expected explicit recalc patch')
		expect(recalcPatch.changedCells.map((cell) => [cell.ref, cell.flatValue])).toEqual([['B1', 10]])
		const afterRecalc = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(materializeViewportPatch(afterEdit.cells, recalcPatch)).toEqual(
			interactiveCellMap(afterRecalc.cells),
		)

		const metadata = await session.apply([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'D1', text: 'reviewed' },
		])
		expect(metadata.apply.errors).toEqual([])
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 4,
				changedSince: afterRecalc.changeToken,
			}),
		).toBeNull()
		const afterMetadata = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: afterRecalc.changeToken,
		})
		expect(afterMetadata.patch).toBeUndefined()
		expect(afterMetadata.cells.find((cell) => cell.ref === 'D1')?.flags.comment).toBe(true)

		const undo = await session.apply(edit.apply.journal?.inverseOps ?? [])
		expect(undo.apply.errors).toEqual([])
		expect(undo.recalc?.changed).toEqual(['Sheet1!B1'])
		const undoPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: afterMetadata.changeToken,
		})
		if (!undoPatch) throw new Error('expected undo patch after metadata refresh')
		expect(undoPatch.changedCells.map((cell) => [cell.ref, cell.flatValue]).sort()).toEqual([
			['A1', 1],
			['B1', 2],
		])
		const restored = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(materializeViewportPatch(afterMetadata.cells, undoPatch)).toEqual(
			interactiveCellMap(restored.cells),
		)
		expect(restored.cells.find((cell) => cell.ref === 'D1')?.flags.comment).toBe(true)
		const restoredSemantic = semanticInteractiveCellMap(restored.cells)
		const baseSemantic = semanticInteractiveCellMap(base.cells)
		expect(restoredSemantic.get('A1')).toEqual(baseSemantic.get('A1'))
		expect(restoredSemantic.get('B1')).toEqual(baseSemantic.get('B1'))
		session.close()
	})

	test('interactive sessions can prepare mutable edit state before first edit', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 1,
		})
		const prepared = await session.prepareEdits()
		expect(prepared.load.promotedToFull).toBe(true)
		expect(prepared.timings.mutableWorkbookCached).toBe(false)
		expect(prepared.timings.mutableWorkbookOpenMs).toBeGreaterThanOrEqual(0)
		expect(prepared.timings.rebaseViewportSnapshotsMs).toBeGreaterThanOrEqual(0)
		expect(prepared.timings.totalMs).toBeGreaterThanOrEqual(
			prepared.timings.ensureMutableWorkbookMs,
		)
		expect(session.editReadiness()).toMatchObject({
			ready: true,
			preparing: false,
			generation: before.generation.session,
			promotedToFull: true,
			timings: {
				mutableWorkbookCached: false,
				mutableWorkbookReusedReadModel: false,
				mutableWorkbookOpenMs: prepared.timings.mutableWorkbookOpenMs,
				rebaseViewportSnapshotsMs: prepared.timings.rebaseViewportSnapshotsMs,
			},
		})

		const edit = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 10 }] },
		])
		expect(edit.apply.errors).toEqual([])
		expect(edit.load.promotedToFull).toBe(true)
		expect(edit.generation.session).toBe(before.generation.session + 1)
		expect(edit.timings.mutableWorkbookCached).toBe(true)
		expect(edit.timings.ensureMutableWorkbookMs).toBeLessThan(1)
		expect(
			session
				.readViewportPatch({
					sheet: 'Sheet1',
					topRow: 0,
					leftCol: 0,
					rowCount: 1,
					colCount: 1,
					changedSince: before.changeToken,
				})
				?.changedCells.map((cell) => [cell.ref, cell.flatValue]),
		).toEqual([['A1', 10]])
		session.close()

		const fullSession = await AscendSession.open(wb.toBytes(), {
			mode: 'full',
			richMetadata: true,
		})
		const fullReadDocument = fullSession.workbook()
		const fullPrepared = await fullSession.prepareEdits()
		expect(fullPrepared.load.promotedToFull).toBe(false)
		expect(fullPrepared.timings.mutableWorkbookReusedReadModel).toBe(true)
		const fullEdit = await fullSession.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 20 }] },
		])
		expect(fullEdit.apply.errors).toEqual([])
		expect(fullEdit.timings.mutableWorkbookCached).toBe(true)
		expect(fullReadDocument.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'number',
			value: 1,
		})
		fullSession.close()
	})

	test('interactive promotion invalidates pre-promotion tokens when full hydration changes viewport semantics', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'B1', value: 2 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: 'A1+B1' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), { mode: 'values' })
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(before.cells.find((cell) => cell.ref === 'C1')?.formula).toBeNull()

		const edit = await session.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'D1', value: 9 }] }],
			{ recalc: false },
		)
		expect(edit.apply.errors).toEqual([])
		expect(edit.load.promotedToFull).toBe(true)
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 4,
				changedSince: before.changeToken,
			}),
		).toBeNull()

		const refreshed = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: before.changeToken,
		})
		expect(refreshed.patch).toBeUndefined()
		expect(refreshed.cells.find((cell) => cell.ref === 'C1')?.formula).toBe('A1+B1')
		expect(refreshed.cells.find((cell) => cell.ref === 'D1')?.flatValue).toBe(9)

		const nextEdit = await session.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'D1', value: 10 }] }],
			{ recalc: false },
		)
		expect(nextEdit.apply.errors).toEqual([])
		const resumedPatch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: refreshed.changeToken,
		})
		expect(resumedPatch?.changedCells.map((cell) => [cell.ref, cell.flatValue])).toEqual([
			['D1', 10],
		])
		session.close()
	})

	test('prepareEdits advances session generation when edit-ready hydration invalidates viewport semantics', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'B1', value: 2 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: 'A1+B1' },
		])
		wb.recalc()

		const session = await AscendSession.open(wb.toBytes(), { mode: 'values' })
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		expect(before.generation.session).toBe(0)
		expect(before.cells.find((cell) => cell.ref === 'C1')?.flatValue).toBe(3)
		expect(before.cells.find((cell) => cell.ref === 'C1')?.formula).toBeNull()

		const prepared = await session.prepareEdits()
		expect(prepared.load.read).toMatchObject({
			mode: 'values',
			isPartial: true,
			partialReasons: ['only cell values are hydrated'],
		})
		expect(prepared.load.write).toMatchObject({ mode: 'full', isPartial: false })
		expect(session.editReadiness()).toMatchObject({
			ready: true,
			generation: 1,
			write: { mode: 'full', isPartial: false },
		})
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 4,
				changedSince: before.changeToken,
			}),
		).toBeNull()

		const refreshed = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: before.changeToken,
		})
		expect(refreshed.patch).toBeUndefined()
		expect(refreshed.generation.session).toBe(1)
		expect(refreshed.cells.find((cell) => cell.ref === 'C1')?.flatValue).toBe(3)
		expect(refreshed.cells.find((cell) => cell.ref === 'C1')?.formula).toBe('A1+B1')

		const edit = await session.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'D1', value: 9 }] }],
			{ recalc: false },
		)
		expect(edit.apply.errors).toEqual([])
		expect(edit.generation.session).toBe(2)
		const patch = session.readViewportPatch({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: refreshed.changeToken,
		})
		expect(patch?.changedCells.map((cell) => [cell.ref, cell.flatValue])).toEqual([['D1', 9]])
		const fresh = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		if (!patch) throw new Error('expected resumed patch')
		expect(materializeViewportPatch(refreshed.cells, patch)).toEqual(
			interactiveCellMap(fresh.cells),
		)
		session.close()
	})

	test('full read model reuse keeps retained read document and source bytes immutable', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])

		const session = await AscendSession.open(wb.toBytes(), {
			mode: 'full',
			richMetadata: true,
		})
		const readDocument = session.workbook()
		const beforeRaw = await readDocument.rawPackagePart({
			partPath: 'xl/worksheets/sheet1.xml',
			maxBytes: 2048,
		})
		expect(beforeRaw.origin).toBe('source')
		expect(beforeRaw.text).toContain('<v>1</v>')

		const prepared = await session.prepareEdits()
		expect(prepared.timings.mutableWorkbookReusedReadModel).toBe(true)
		const edit = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
		])
		expect(edit.apply.errors).toEqual([])
		expect(
			session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 1,
			}).cells[0]?.flatValue,
		).toBe(2)
		expect(readDocument.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'number',
			value: 1,
		})

		const afterRaw = await readDocument.rawPackagePart({
			partPath: 'xl/worksheets/sheet1.xml',
			maxBytes: 2048,
		})
		expect(afterRaw.origin).toBe('source')
		expect(afterRaw.sha256).toBe(beforeRaw.sha256)
		expect(afterRaw.text).toBe(beforeRaw.text)
		session.close()
	})

	test('interactive viewport tokens from other sessions refreshes and retained-log gaps force fresh reads', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
		const bytes = wb.toBytes()

		const session = await AscendSession.open(bytes, { mode: 'interactive' })
		const base = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 1,
		})
		const other = await AscendSession.open(bytes, { mode: 'interactive' })
		try {
			const crossSession = other.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 1,
				changedSince: base.changeToken,
			})
			expect(crossSession.cells[0]?.flatValue).toBe(1)
			expect(crossSession.patch).toBeUndefined()
			expect(
				other.readViewportPatch({
					sheet: 'Sheet1',
					topRow: 0,
					leftCol: 0,
					rowCount: 1,
					colCount: 1,
					changedSince: base.changeToken,
				}),
			).toBeNull()
		} finally {
			other.close()
		}

		const edit = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
		])
		expect(edit.apply.errors).toEqual([])
		const changed = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 1,
			changedSince: base.changeToken,
		})
		expect(changed.patch?.changedCells.map((cell) => [cell.ref, cell.flatValue])).toEqual([
			['A1', 2],
		])

		await session.refresh()
		const afterRefresh = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 1,
			changedSince: changed.changeToken,
		})
		expect(afterRefresh.patch).toBeUndefined()
		session.close()

		const expiring = await AscendSession.open(bytes, { mode: 'interactive' })
		try {
			const beforeGap = expiring.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 1,
			})
			for (let i = 0; i < 130; i++) {
				const gapEdit = await expiring.apply([
					{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: i + 2 }] },
				])
				expect(gapEdit.apply.errors).toEqual([])
			}
			expect(
				expiring.readViewportPatch({
					sheet: 'Sheet1',
					topRow: 0,
					leftCol: 0,
					rowCount: 1,
					colCount: 1,
					changedSince: beforeGap.changeToken,
				}),
			).toBeNull()
			const afterGap = expiring.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 1,
				changedSince: beforeGap.changeToken,
			})
			expect(afterGap.cells[0]?.flatValue).toBe(131)
			expect(afterGap.patch).toBeUndefined()
		} finally {
			expiring.close()
		}
	})

	test('interactive pull patches refuse metadata and layout edits that need fresh viewport state', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 2,
			colCount: 1,
		})
		const commentEdit = await session.apply([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'A1', text: 'needs review' },
		])
		expect(commentEdit.apply.errors).toEqual([])
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 2,
				colCount: 1,
				changedSince: before.changeToken,
			}),
		).toBeNull()

		const afterComment = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 2,
			colCount: 1,
			changedSince: before.changeToken,
		})
		expect(afterComment.comments).toHaveLength(1)
		expect(afterComment.cells.find((cell) => cell.ref === 'A1')?.flags.comment).toBe(true)
		expect(afterComment.patch).toBeUndefined()

		const layoutEdit = await session.apply([
			{ op: 'setRowHeight', sheet: 'Sheet1', row: 0, height: 28 },
		])
		expect(layoutEdit.apply.errors).toEqual([])
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 2,
				colCount: 1,
				changedSince: afterComment.changeToken,
			}),
		).toBeNull()
		const afterLayout = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 2,
			colCount: 1,
			changedSince: afterComment.changeToken,
		})
		expect(afterLayout.rowLayout).toEqual([{ index: 0, size: 28 }])
		expect(afterLayout.patch).toBeUndefined()
		session.close()
	})

	test('interactive pull patches reject tokens older than the retained change log', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 1,
		})
		for (let i = 0; i < 130; i++) {
			const edit = await session.apply([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: i + 2 }] },
			])
			expect(edit.apply.errors).toEqual([])
		}

		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 1,
				changedSince: before.changeToken,
			}),
		).toBeNull()
		session.close()
	})

	test('interactive session blocks edits against partial workbooks', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
				],
			},
		])

		const session = await AscendSession.open(wb.toBytes(), {
			mode: 'interactive',
			maxRows: 1,
		})
		const edit = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
		])
		expect(edit.apply.errors[0]?.message).toContain('partial workbook')
		expect(edit.load.read).toMatchObject({ mode: 'formula', isPartial: true, maxRows: 1 })
		expect(edit.load.write).toMatchObject({ mode: 'formula', isPartial: true, maxRows: 1 })
		expect(edit.load.promotedToFull).toBe(false)
		expect(edit.generation.session).toBe(0)
		session.close()
	})

	test('interactive explicit recalc stays read-only for capped partial sessions', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])

		const session = await AscendSession.open(wb.toBytes(), {
			mode: 'interactive',
			maxRows: 1,
		})
		const before = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 2,
		})
		const recalc = await session.apply([], { recalc: true })
		expect(recalc.apply.errors).toEqual([])
		expect(recalc.recalc?.errors[0]?.error.message).toContain('partial workbook')
		expect(recalc.load.write).toMatchObject({ mode: 'formula', isPartial: true, maxRows: 1 })
		expect(recalc.load.promotedToFull).toBe(false)
		expect(recalc.generation.session).toBe(before.generation.session)
		expect(
			session.readViewportPatch({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 2,
				changedSince: before.changeToken,
			}),
		)?.toMatchObject({
			changedCells: [],
			removedRefs: [],
		})
		session.close()
	})

	test('prepareEdits on a partial session does not report writable readiness', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
				],
			},
		])

		const session = await AscendSession.open(wb.toBytes(), {
			mode: 'interactive',
			maxRows: 1,
		})
		const viewport = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 1,
		})
		const prepared = await session.prepareEdits()
		expect(prepared.load.read).toMatchObject({ isPartial: true, maxRows: 1 })
		expect(prepared.load.write).toMatchObject({ isPartial: true, maxRows: 1 })
		expect(prepared.load.promotedToFull).toBe(false)
		expect(session.editReadiness()).toMatchObject({
			ready: false,
			preparing: false,
			promotedToFull: false,
			write: { isPartial: true, maxRows: 1 },
		})

		const edit = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
		])
		expect(edit.apply.errors[0]?.message).toContain('partial workbook')
		expect(edit.generation.session).toBe(viewport.generation.session)
		session.close()
	})

	test('interactive partial session refuses promotion after backing file changes', async () => {
		const input = join(tmpdir(), `ascend-stale-promotion-${Date.now()}-${process.pid}.xlsx`)
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'original' },
					{ ref: 'A2', value: 'not loaded' },
				],
			},
		])
		await wb.save(input)

		const session = await AscendSession.open(input, { mode: 'interactive', maxRows: 1 })
		try {
			const viewport = session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 1,
			})
			expect(viewport.cells[0]?.flatValue).toBe('original')

			const changed = AscendWorkbook.create()
			changed.apply([
				{
					op: 'setCells',
					sheet: 'Sheet1',
					updates: [
						{ ref: 'A1', value: 'changed elsewhere' },
						{ ref: 'A2', value: 'different tail value that changes the package size' },
					],
				},
			])
			await changed.save(input)
			expect(session.isStale()).toBe(true)

			let prepareError: unknown
			try {
				await session.prepareEdits()
			} catch (error) {
				prepareError = error
			}
			expect(prepareError).toBeInstanceOf(Error)
			expect((prepareError as Error).message).toContain(
				'Cannot promote a stale interactive session',
			)

			const edit = await session.apply([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'agent edit' }] },
			])
			expect(edit.apply.errors[0]).toMatchObject({
				code: 'VALIDATION_ERROR',
				details: {
					rule: 'stale-interactive-session',
					staleSession: true,
					requiredAction: 'refresh',
				},
			})
			expect(edit.load.promotedToFull).toBe(false)
			expect(edit.generation.session).toBe(viewport.generation.session)

			const reopened = await AscendWorkbook.open(input)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'changed elsewhere',
			})
		} finally {
			session.close()
			await unlink(input).catch(() => {})
		}
	})

	test('values-mode sessions refuse stale promotion before full writable hydration', async () => {
		const input = join(tmpdir(), `ascend-stale-values-promotion-${Date.now()}-${process.pid}.xlsx`)
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'B1', value: 2 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: 'A1+B1' },
		])
		wb.recalc()
		await wb.save(input)

		const session = await AscendSession.open(input, { mode: 'values' })
		try {
			const viewport = session.readViewport({
				sheet: 'Sheet1',
				topRow: 0,
				leftCol: 0,
				rowCount: 1,
				colCount: 3,
			})
			expect(viewport.cells.find((cell) => cell.ref === 'C1')?.flatValue).toBe(3)
			expect(viewport.cells.find((cell) => cell.ref === 'C1')?.formula).toBeNull()
			expect(session.editReadiness()).toMatchObject({
				ready: false,
				generation: viewport.generation.session,
				read: { mode: 'values', isPartial: true },
				write: null,
			})

			const changed = AscendWorkbook.create()
			changed.apply([
				{
					op: 'setCells',
					sheet: 'Sheet1',
					updates: [
						{ ref: 'A1', value: 'changed elsewhere' },
						{ ref: 'B1', value: 20 },
					],
				},
				{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: 'A1&B1' },
			])
			changed.recalc()
			await changed.save(input)
			expect(session.isStale()).toBe(true)

			let prepareError: unknown
			try {
				await session.prepareEdits()
			} catch (error) {
				prepareError = error
			}
			expect(prepareError).toBeInstanceOf(Error)
			expect((prepareError as Error).message).toContain(
				'Cannot promote a stale interactive session',
			)
			expect(session.editReadiness()).toMatchObject({
				ready: false,
				generation: viewport.generation.session,
				write: null,
			})

			const edit = await session.apply([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'agent edit' }] },
			])
			expect(edit.apply.errors[0]).toMatchObject({
				code: 'VALIDATION_ERROR',
				details: {
					rule: 'stale-interactive-session',
					staleSession: true,
					requiredAction: 'refresh',
				},
			})
			expect(edit.load).toMatchObject({
				read: { mode: 'values', isPartial: true },
				write: { mode: 'values', isPartial: true },
				promotedToFull: false,
			})
			expect(edit.generation.session).toBe(viewport.generation.session)

			const reopened = await AscendWorkbook.open(input)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'changed elsewhere',
			})
		} finally {
			session.close()
			await unlink(input).catch(() => {})
		}
	})

	test('apply returns dirty regions and monotonic generation tokens', () => {
		const wb = AscendWorkbook.create()

		const valueEdit = wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'B2', value: 10 },
					{ ref: 'D4', value: 20 },
				],
			},
		])
		expect(valueEdit.errors).toEqual([])
		expect(valueEdit.dirtyRegions).toEqual([
			{
				sheet: 'Sheet1',
				range: 'B2:D4',
				refs: ['Sheet1!B2', 'Sheet1!D4'],
			},
		])
		expect(valueEdit.generations).toEqual({
			workbook: 1,
			sheetMetadata: 0,
			formulas: 1,
			styles: 0,
		})

		const styleEdit = wb.apply([
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'B2:B2', format: '0.00' },
		])
		expect(styleEdit.errors).toEqual([])
		expect(styleEdit.dirtyRegions).toEqual([
			{ sheet: 'Sheet1', range: 'B2:B2', refs: ['Sheet1!B2'] },
		])
		expect(styleEdit.generations).toEqual({
			workbook: 2,
			sheetMetadata: 0,
			formulas: 1,
			styles: 1,
		})

		const structuralEdit = wb.apply([{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 }])
		expect(structuralEdit.errors).toEqual([])
		expect(structuralEdit.generations).toEqual({
			workbook: 3,
			sheetMetadata: 1,
			formulas: 2,
			styles: 1,
		})
	})

	test('no-op and failed applies do not advance workbook or session generations', async () => {
		const wb = AscendWorkbook.create()
		const initial = wb.readSnapshotInfo().generations
		const empty = wb.apply([])
		expect(empty).toMatchObject({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
			dirtyRegions: [],
			generations: initial,
			errors: [],
		})
		expect(wb.readSnapshotInfo().generations).toEqual(initial)

		const failed = wb.apply([
			{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 1 }] },
		])
		expect(failed.errors[0]?.code).toBe('SHEET_NOT_FOUND')
		expect(failed.generations).toEqual(initial)
		expect(wb.readSnapshotInfo().generations).toEqual(initial)

		const changed = wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 5 },
					{ ref: 'D1', value: null },
				],
			},
		])
		expect(changed.errors).toEqual([])
		expect(changed.affectedCells).toEqual(['A1'])
		const afterChanged = wb.readSnapshotInfo().generations
		const semanticNoOp = wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 5 },
					{ ref: 'D1', value: null },
				],
			},
		])
		expect(semanticNoOp).toMatchObject({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
			dirtyRegions: [],
			generations: afterChanged,
			errors: [],
		})
		expect(wb.readSnapshotInfo().generations).toEqual(afterChanged)

		const metadataOps = [
			{ op: 'setComment' as const, sheet: 'Sheet1', ref: 'A1', text: 'Review', author: 'Ada' },
			{
				op: 'setHyperlink' as const,
				sheet: 'Sheet1',
				ref: 'B1',
				url: 'https://example.com/report',
				display: 'Report',
				tooltip: 'Open report',
			},
		]
		const metadataChanged = wb.apply(metadataOps)
		expect(metadataChanged.errors).toEqual([])
		expect(metadataChanged.affectedCells.sort()).toEqual(['A1', 'B1'])
		const afterMetadataChanged = wb.readSnapshotInfo().generations
		const metadataNoOp = wb.apply(metadataOps)
		expect(metadataNoOp).toMatchObject({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
			dirtyRegions: [],
			generations: afterMetadataChanged,
			errors: [],
		})
		expect(wb.readSnapshotInfo().generations).toEqual(afterMetadataChanged)

		const styleOps = [
			{ op: 'setNumberFormat' as const, sheet: 'Sheet1', range: 'C1:C1', format: '0.00%' },
			{
				op: 'setStyle' as const,
				sheet: 'Sheet1',
				range: 'C1:C1',
				style: { font: { bold: true }, numberFormat: '0.00%' },
			},
		]
		const styleChanged = wb.apply(styleOps)
		expect(styleChanged.errors).toEqual([])
		expect(styleChanged.affectedCells).toEqual(['C1', 'C1'])
		const afterStyleChanged = wb.readSnapshotInfo().generations
		const styleNoOp = wb.apply(styleOps)
		expect(styleNoOp).toMatchObject({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
			dirtyRegions: [],
			generations: afterStyleChanged,
			errors: [],
		})
		expect(wb.readSnapshotInfo().generations).toEqual(afterStyleChanged)

		const fullSession = await AscendSession.open(wb.toBytes(), { mode: 'interactive' })
		const fullViewport = fullSession.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
		})
		const sessionNoOp = await fullSession.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 5 },
					{ ref: 'D1', value: null },
				],
			},
		])
		expect(sessionNoOp.apply).toMatchObject({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
			dirtyRegions: [],
			errors: [],
		})
		expect(sessionNoOp.generation.session).toBe(fullViewport.generation.session)
		const afterSessionNoOp = fullSession.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 4,
			changedSince: fullViewport.changeToken,
		})
		expect(afterSessionNoOp.patch).toMatchObject({
			baseToken: fullViewport.changeToken,
			changedCells: [],
			removedRefs: [],
			byteLength: 36,
		})
		expect(afterSessionNoOp.generation).toEqual(fullViewport.generation)
		expect(afterSessionNoOp.cells).toEqual(fullViewport.cells)

		const sessionMetadataNoOp = await fullSession.apply(metadataOps)
		expect(sessionMetadataNoOp.apply).toMatchObject({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
			dirtyRegions: [],
			errors: [],
		})
		expect(sessionMetadataNoOp.generation.session).toBe(fullViewport.generation.session)

		const sessionStyleChanged = await fullSession.apply(styleOps)
		expect(sessionStyleChanged.apply.errors).toEqual([])
		const sessionStyleNoOp = await fullSession.apply(styleOps)
		expect(sessionStyleNoOp.apply).toMatchObject({
			affectedCells: [],
			sheetsModified: [],
			recalcRequired: false,
			dirtyRegions: [],
			errors: [],
		})
		expect(sessionStyleNoOp.generation.session).toBe(sessionStyleChanged.generation.session)
		fullSession.close()

		const session = await AscendSession.open(wb.toBytes(), { mode: 'interactive', maxRows: 1 })
		const viewport = session.readViewport({
			sheet: 'Sheet1',
			topRow: 0,
			leftCol: 0,
			rowCount: 1,
			colCount: 1,
		})
		const sessionEmpty = await session.apply([])
		expect(sessionEmpty.apply.errors).toEqual([])
		const { session: viewportSessionGeneration, ...viewportWorkbookGenerations } =
			viewport.generation
		expect(sessionEmpty.apply.generations).toEqual(viewportWorkbookGenerations)
		expect(sessionEmpty.generation.session).toBe(viewportSessionGeneration)
		expect(sessionEmpty.load.promotedToFull).toBe(false)

		const sessionFailed = await session.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
		])
		expect(sessionFailed.apply.errors[0]?.message).toContain('partial workbook')
		expect(sessionFailed.generation.session).toBe(viewportSessionGeneration)
		session.close()
	})

	test('generation layers advance for formula recalc metadata style and workbook edits', () => {
		const wb = AscendWorkbook.create()

		const formulaEdit = wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '1+1' }])
		expect(formulaEdit.errors).toEqual([])
		expect(formulaEdit.generations).toEqual({
			workbook: 1,
			sheetMetadata: 0,
			formulas: 1,
			styles: 0,
		})

		const recalc = wb.recalc()
		expect(recalc.errors).toEqual([])
		expect(recalc.generations).toEqual({
			workbook: 2,
			sheetMetadata: 0,
			formulas: 2,
			styles: 0,
		})

		const commentEdit = wb.apply([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'B1', text: 'reviewed' },
		])
		expect(commentEdit.errors).toEqual([])
		expect(commentEdit.generations).toEqual({
			workbook: 3,
			sheetMetadata: 1,
			formulas: 2,
			styles: 0,
		})

		const styleEdit = wb.apply([
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C1:C1', format: '0.00' },
		])
		expect(styleEdit.errors).toEqual([])
		expect(styleEdit.generations).toEqual({
			workbook: 4,
			sheetMetadata: 1,
			formulas: 2,
			styles: 1,
		})

		const workbookMetadataEdit = wb.apply([
			{ op: 'setWorkbookProperties', properties: { date1904: true } },
		])
		expect(workbookMetadataEdit.errors).toEqual([])
		expect(workbookMetadataEdit.generations).toEqual({
			workbook: 5,
			sheetMetadata: 2,
			formulas: 3,
			styles: 1,
		})
	})

	test('recalc returns changed dirty regions and advances formula generation once', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'A1', value: 1 }],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		const before = wb.recalc()
		expect(before.errors).toEqual([])

		const apply = wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
		])
		const recalc = wb.recalc()

		expect(apply.dirtyRegions).toEqual([{ sheet: 'Sheet1', range: 'A1:A1', refs: ['Sheet1!A1'] }])
		expect(recalc.changed).toEqual(['Sheet1!B1'])
		expect(recalc.dirtyRegions).toEqual([{ sheet: 'Sheet1', range: 'B1:B1', refs: ['Sheet1!B1'] }])
		expect(recalc.generations.formulas).toBeGreaterThan(apply.generations.formulas)
		expect(recalc.generations.workbook).toBeGreaterThan(apply.generations.workbook)
	})

	test('apply can return a reversible journal for supported interactive edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])

		const apply = wb.apply(
			[
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
				{ op: 'setComment', sheet: 'Sheet1', ref: 'B2', text: 'review', author: 'agent' },
			],
			{ journal: true },
		)

		expect(apply.errors).toEqual([])
		expect(apply.journal?.supported).toBe(true)
		expect(apply.journal?.exact).toBe(true)
		expect(apply.journal?.inverseOps).toEqual([
			{ op: 'deleteComment', sheet: 'Sheet1', ref: 'B2' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
		])

		const undo = wb.apply(apply.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 1 })
		expect(wb.sheet('Sheet1')?.comment('B2')).toBeUndefined()
	})

	test('preview journals expose lossy formula preimages without mutating the workbook', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '1+1' }])

		const preview = wb.preview(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] }],
			{ journal: true },
		)

		expect(preview.wouldSucceed).toBe(true)
		expect(preview.journal?.supported).toBe(true)
		expect(preview.journal?.exact).toBe(false)
		expect(preview.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Formula cache for Sheet1!A1 cannot be restored with public operations',
				refs: ['Sheet1!A1'],
			},
		])
		expect(preview.journal?.inverseOps).toEqual([
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '1+1' },
		])
		expect(wb.sheet('Sheet1')?.cell('A1')?.formula).toBe('1+1')
	})

	test('journals do not claim exact rollback for supported mutations without inverse coverage', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 1 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])

		const changed = wb.apply(
			[
				{ op: 'appendRows', table: 'Sales', rows: [['East', 2]] },
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'D1', value: 'audit' }] },
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(false)
		expect(changed.journal?.exact).toBe(false)
		expect(changed.journal?.issues).toContainEqual({
			code: 'UNSUPPORTED_OPERATION',
			message: 'No reversible journal support for appendRows',
		})
		expect(changed.journal?.entries[0]).toMatchObject({
			opIndex: 0,
			supported: false,
			exact: false,
			inverseOps: [],
		})
		expect(changed.journal?.entries[1]).toMatchObject({
			opIndex: 1,
			supported: true,
			exact: true,
			inverseOps: [{ op: 'clearRange', sheet: 'Sheet1', range: 'D1', what: 'all' }],
		})
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'clearRange', sheet: 'Sheet1', range: 'D1', what: 'all' },
		])
	})

	test('journal inverse ops preserve internal order when restoring cleared cell state', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'C3', value: 12 }] }])
		wb.apply([{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C3:C3', format: '0.00' }])

		const cleared = wb.apply([{ op: 'clearRange', sheet: 'Sheet1', range: 'C3:C3', what: 'all' }], {
			journal: true,
		})
		expect(cleared.journal?.exact).toBe(true)

		const undo = wb.apply(cleared.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('C3')?.value).toEqual({ kind: 'number', value: 12 })
		expect(wb.cellStyle('Sheet1!C3')?.numberFormat).toBe('0.00')
	})

	test('journal inverse ops restore merge and unmerge metadata', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'mergeCells', sheet: 'Sheet1', range: 'A1:B2' }])

		const unmerged = wb.apply([{ op: 'unmergeCells', sheet: 'Sheet1', range: 'A1:B2' }], {
			journal: true,
		})

		expect(unmerged.journal?.exact).toBe(true)
		expect(unmerged.journal?.inverseOps).toEqual([
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'A1:B2' },
		])

		const undoUnmerge = wb.apply(unmerged.journal?.inverseOps ?? [], { transaction: true })
		expect(undoUnmerge.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.merges).toEqual([
			{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
		])

		const wb2 = AscendWorkbook.create()
		const merged = wb2.apply([{ op: 'mergeCells', sheet: 'Sheet1', range: 'C1:D1' }], {
			journal: true,
		})
		expect(merged.journal?.inverseOps).toEqual([
			{ op: 'unmergeCells', sheet: 'Sheet1', range: 'C1:D1' },
		])
	})

	test('journal inverse ops restore data validation changes', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1:A10',
				rule: {
					type: 'list',
					formula1: '"Open,Closed"',
					allowBlank: true,
					showErrorMessage: true,
				},
			},
		])

		const changed = wb.apply(
			[
				{
					op: 'setDataValidation',
					sheet: 'Sheet1',
					range: 'A1:A10',
					rule: { type: 'whole', operator: 'greaterThan', formula1: '0' },
				},
			],
			{ journal: true },
		)

		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1:A10',
				rule: {
					type: 'list',
					formula1: '"Open,Closed"',
					allowBlank: true,
					showErrorMessage: true,
				},
			},
		])

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.getDataValidations()).toEqual([
			{ type: 'list', formula: '"Open,Closed"', range: 'A1:A10' },
		])
	})

	test('journal inverse ops restore simple worksheet auto filters', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C10' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C10', column: 1, values: ['Open'] },
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:C10',
				sortRef: 'A2:C10',
				sortBy: 'B2:B10',
				descending: true,
			},
		])

		const cleared = wb.apply([{ op: 'clearAutoFilter', sheet: 'Sheet1' }], { journal: true })

		expect(cleared.journal?.exact).toBe(true)
		expect(cleared.journal?.inverseOps).toEqual([
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C10' },
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:C10',
				column: 1,
				values: ['Open'],
			},
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:C10',
				sortRef: 'A2:C10',
				sortBy: 'B2:B10',
				descending: true,
			},
		])

		const undo = wb.apply(cleared.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.autoFilter).toEqual({
			ref: 'A1:C10',
			columns: [{ colId: 1, kind: 'filters', values: ['Open'] }],
			sortState: {
				ref: 'A2:C10',
				conditions: [{ ref: 'B2:B10', descending: true }],
			},
		})

		const wb2 = AscendWorkbook.create()
		const created = wb2.apply([{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'D1:E5' }], {
			journal: true,
		})
		expect(created.journal?.inverseOps).toEqual([{ op: 'clearAutoFilter', sheet: 'Sheet1' }])
	})

	test('journal inverse ops restore deleted hyperlink metadata exactly', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'b2',
				url: 'https://example.com/report',
				display: 'Report',
				tooltip: 'Open report',
			},
		])

		const deleted = wb.apply([{ op: 'deleteHyperlink', sheet: 'Sheet1', ref: 'B2' }], {
			journal: true,
		})

		expect(deleted.errors).toEqual([])
		expect(deleted.journal?.exact).toBe(true)
		expect(deleted.journal?.inverseOps).toEqual([
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'B2',
				url: 'https://example.com/report',
				display: 'Report',
				tooltip: 'Open report',
			},
		])
		expect(wb.sheet('Sheet1')?.getHyperlinks()).toEqual([])

		const undo = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.getHyperlinks()).toEqual([
			{
				ref: 'B2',
				target: 'https://example.com/report',
				display: 'Report',
				tooltip: 'Open report',
			},
		])
	})

	test('journal inverse ops restore threaded comment text without losing metadata', () => {
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().sheets[0]
		sheet?.threadedComments.push(
			{
				ref: 'B2',
				text: 'original',
				id: 'tc-1',
				personId: 'person-1',
				author: 'Analyst',
				dateTime: '2026-05-13T10:00:00Z',
				done: false,
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
			{
				ref: 'B2',
				text: 'follow-up',
				id: 'tc-2',
				parentId: 'tc-1',
				personId: 'person-2',
				author: 'Reviewer',
				dateTime: '2026-05-13T10:05:00Z',
				done: true,
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
		)

		const changed = wb.apply(
			[
				{
					op: 'setThreadedComment',
					sheet: 'Sheet1',
					partPath: 'xl/threadedComments/threadedComment1.xml',
					threadedCommentId: 'tc-2',
					text: 'resolved by agent',
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setThreadedComment',
				sheet: 'Sheet1',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				threadedCommentId: 'tc-2',
				text: 'follow-up',
			},
		])
		const changedThread = wb.getWorkbookModel().sheets[0]?.threadedComments[1]
		expect(changedThread).toMatchObject({
			text: 'resolved by agent',
			id: 'tc-2',
			parentId: 'tc-1',
			personId: 'person-2',
			author: 'Reviewer',
			done: true,
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.getWorkbookModel().sheets[0]?.threadedComments[1]).toEqual({
			ref: 'B2',
			text: 'follow-up',
			id: 'tc-2',
			parentId: 'tc-1',
			personId: 'person-2',
			author: 'Reviewer',
			dateTime: '2026-05-13T10:05:00Z',
			done: true,
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})
	})

	test('threaded comment journals do not claim exact rollback for ambiguous selectors', () => {
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().sheets[0]
		sheet?.threadedComments.push(
			{ ref: 'B2', text: 'first', id: 'tc-1' },
			{ ref: 'B2', text: 'second', id: 'tc-2' },
		)

		const journal = buildMutationJournal(wb.getWorkbookModel(), [
			{ op: 'setThreadedComment', sheet: 'Sheet1', ref: 'B2', text: 'ambiguous' },
		])

		expect(journal.supported).toBe(true)
		expect(journal.exact).toBe(false)
		expect(journal.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Threaded comment selector on Sheet1 cannot be resolved exactly',
			},
		])
		expect(journal.inverseOps).toEqual([])
		expect(sheet?.threadedComments.map((comment) => comment.text)).toEqual(['first', 'second'])
	})

	test('journal inverse ops restore drawing text without losing relationship metadata', () => {
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().sheets[0]
		sheet?.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'graphicFrame',
			id: 7,
			name: 'Chart Callout',
			description: 'Revenue callout',
			text: 'Original',
			anchor: {
				kind: 'absolute',
				x: 1000,
				y: 2000,
				cx: 3000,
				cy: 4000,
			},
			relIds: ['rIdChart1'],
			relationshipRefs: [
				{
					id: 'rIdChart1',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
					target: '../charts/chart1.xml',
				},
			],
		})

		const changed = wb.apply(
			[
				{
					op: 'setDrawingText',
					sheet: 'Sheet1',
					drawingPartPath: 'xl/drawings/drawing1.xml',
					id: 7,
					text: 'Updated',
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setDrawingText',
				sheet: 'Sheet1',
				drawingPartPath: 'xl/drawings/drawing1.xml',
				id: 7,
				text: 'Original',
			},
		])
		expect(wb.getWorkbookModel().sheets[0]?.drawingObjectRefs[0]).toMatchObject({
			id: 7,
			name: 'Chart Callout',
			text: 'Updated',
			relIds: ['rIdChart1'],
		})

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.getWorkbookModel().sheets[0]?.drawingObjectRefs[0]).toEqual({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'graphicFrame',
			id: 7,
			name: 'Chart Callout',
			description: 'Revenue callout',
			text: 'Original',
			anchor: {
				kind: 'absolute',
				x: 1000,
				y: 2000,
				cx: 3000,
				cy: 4000,
			},
			relIds: ['rIdChart1'],
			relationshipRefs: [
				{
					id: 'rIdChart1',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
					target: '../charts/chart1.xml',
				},
			],
		})
	})

	test('drawing text journals do not claim exact rollback for ambiguous or non-text selectors', () => {
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().sheets[0]
		sheet?.drawingObjectRefs.push(
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				kind: 'textBox',
				id: 1,
				name: 'Duplicate',
				text: 'First',
			},
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				kind: 'shape',
				id: 2,
				name: 'Duplicate',
			},
		)

		const ambiguous = buildMutationJournal(wb.getWorkbookModel(), [
			{ op: 'setDrawingText', sheet: 'Sheet1', name: 'Duplicate', text: 'Updated' },
		])
		expect(ambiguous.supported).toBe(true)
		expect(ambiguous.exact).toBe(false)
		expect(ambiguous.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Drawing object selector on Sheet1 cannot be resolved to editable text exactly',
			},
		])
		expect(ambiguous.inverseOps).toEqual([])

		const noText = buildMutationJournal(wb.getWorkbookModel(), [
			{ op: 'setDrawingText', sheet: 'Sheet1', id: 2, text: 'Updated' },
		])
		expect(noText.supported).toBe(true)
		expect(noText.exact).toBe(false)
		expect(noText.inverseOps).toEqual([])
		expect(sheet?.drawingObjectRefs.map((object) => object.text)).toEqual(['First', undefined])
	})

	test('journal inverse ops restore chart series source refs without losing chart metadata', () => {
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			chartType: 'lineChart',
			title: 'Revenue',
			series: [
				{
					nameRef: 'Sheet1!$B$1',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$B$2:$B$4',
				},
				{
					nameText: 'Plan',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$C$2:$C$4',
				},
			],
		})

		const changed = wb.apply(
			[
				{
					op: 'setChartSeriesSource',
					partPath: 'xl/charts/chart1.xml',
					seriesIndex: 0,
					nameRef: 'Sheet1!$D$1',
					categoryRef: 'Sheet1!$A$2:$A$10',
					valueRef: 'Sheet1!$D$2:$D$10',
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setChartSeriesSource',
				partPath: 'xl/charts/chart1.xml',
				seriesIndex: 0,
				nameRef: 'Sheet1!$B$1',
				categoryRef: 'Sheet1!$A$2:$A$4',
				valueRef: 'Sheet1!$B$2:$B$4',
			},
		])
		expect(wb.getWorkbookModel().chartParts[0]).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			chartType: 'lineChart',
			title: 'Revenue',
			series: [
				{
					nameRef: 'Sheet1!$D$1',
					categoryRef: 'Sheet1!$A$2:$A$10',
					valueRef: 'Sheet1!$D$2:$D$10',
				},
				{
					nameText: 'Plan',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$C$2:$C$4',
				},
			],
		})

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.getWorkbookModel().chartParts[0]).toEqual({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			chartType: 'lineChart',
			title: 'Revenue',
			series: [
				{
					nameRef: 'Sheet1!$B$1',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$B$2:$B$4',
				},
				{
					nameText: 'Plan',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$C$2:$C$4',
				},
			],
		})
	})

	test('chart series journals mark rollback lossy when refs cannot be unset', () => {
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			series: [{ nameText: 'Actual', valueRef: 'Sheet1!$B$2:$B$4' }],
		})

		const journal = buildMutationJournal(wb.getWorkbookModel(), [
			{
				op: 'setChartSeriesSource',
				partPath: 'xl/charts/chart1.xml',
				seriesIndex: 0,
				nameRef: 'Sheet1!$B$1',
				valueRef: 'Sheet1!$C$2:$C$4',
			},
		])

		expect(journal.supported).toBe(true)
		expect(journal.exact).toBe(false)
		expect(journal.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Chart series selector cannot be restored exactly for series 0',
			},
		])
		expect(journal.inverseOps).toEqual([
			{
				op: 'setChartSeriesSource',
				partPath: 'xl/charts/chart1.xml',
				seriesIndex: 0,
				valueRef: 'Sheet1!$B$2:$B$4',
			},
		])
	})

	test('journal inverse ops restore pivot cache source and refresh metadata', () => {
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		wb.getWorkbookModel().pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			relId: 'rIdPivotCache1',
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
			recordsPartPath: 'xl/pivotCacheRecords/pivotCacheRecords1.xml',
			refreshOnLoad: false,
			enableRefresh: true,
			invalid: false,
			saveData: true,
			fields: [],
		})

		const changed = wb.apply(
			[
				{
					op: 'setPivotCache',
					pivotTable: 'PivotTable1',
					sourceSheet: 'RawData',
					sourceRef: 'A1:E20',
					refreshOnLoad: true,
					enableRefresh: false,
					invalid: true,
					saveData: false,
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setPivotCache',
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				sourceSheet: 'Raw',
				sourceRef: 'A1:D10',
				refreshOnLoad: false,
				enableRefresh: true,
				invalid: false,
				saveData: true,
			},
		])
		expect(wb.getWorkbookModel().pivotCaches[0]).toMatchObject({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			relId: 'rIdPivotCache1',
			recordsPartPath: 'xl/pivotCacheRecords/pivotCacheRecords1.xml',
			sourceSheet: 'RawData',
			sourceRef: 'A1:E20',
			refreshOnLoad: true,
			enableRefresh: false,
			invalid: true,
			saveData: false,
		})

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.getWorkbookModel().pivotCaches[0]).toEqual({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			relId: 'rIdPivotCache1',
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
			recordsPartPath: 'xl/pivotCacheRecords/pivotCacheRecords1.xml',
			refreshOnLoad: false,
			enableRefresh: true,
			invalid: false,
			saveData: true,
			fields: [],
		})
	})

	test('pivot cache journals mark rollback lossy when public ops cannot unset fields', () => {
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			sourceRef: 'A1:D10',
			fields: [],
		})

		const journal = buildMutationJournal(wb.getWorkbookModel(), [
			{
				op: 'setPivotCache',
				cacheId: 34,
				sourceSheet: 'RawData',
				sourceRef: 'A1:E20',
				refreshOnLoad: true,
			},
		])

		expect(journal.supported).toBe(true)
		expect(journal.exact).toBe(false)
		expect(journal.issues).toEqual([
			{ code: 'LOSSY_INVERSE', message: 'Pivot cache selector cannot be restored exactly' },
		])
		expect(journal.inverseOps).toEqual([
			{
				op: 'setPivotCache',
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				sourceRef: 'A1:D10',
			},
		])
	})

	test('journal inverse ops restore conditional format replacements', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'B1:B10',
				rule: {
					type: 'cellIs',
					operator: 'greaterThan',
					formula: '10',
					priority: 3,
					style: { font: { bold: true } },
				},
			},
		])

		const changed = wb.apply(
			[
				{
					op: 'setConditionalFormat',
					sheet: 'Sheet1',
					range: 'B1:B10',
					rule: { type: 'expression', formula: 'B1<0', priority: 1 },
				},
			],
			{ journal: true },
		)

		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'deleteConditionalFormat', sheet: 'Sheet1', range: 'B1:B10' },
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'B1:B10',
				rule: {
					type: 'cellIs',
					operator: 'greaterThan',
					formula: '10',
					priority: 3,
					style: { font: { bold: true } },
				},
				mode: 'replace',
			},
		])

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.getConditionalFormats()).toEqual([
			{ type: 'cellIs', priority: 3, range: 'B1:B10' },
		])
	})

	test('journal marks extended validation and autofilter metadata inverses lossy', () => {
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().getSheet('Sheet1')
		if (!sheet) throw new Error('Sheet1 missing')
		sheet.dataValidations.push({
			sqref: 'A1:A3',
			source: 'x14',
			uid: '{validation-uid}',
			type: 'list',
			formula1: '"Open,Closed"',
		})
		sheet.autoFilter = {
			ref: 'A1:B10',
			uid: '{filter-uid}',
			columns: [{ colId: 0, kind: 'filters', values: ['Open'], hiddenButton: true }],
			sortState: {
				ref: 'A1:B10',
				caseSensitive: true,
				conditions: [{ ref: 'A2:A10', descending: true }, { ref: 'B2:B10' }],
			},
		}

		const changed = wb.preview(
			[
				{
					op: 'setDataValidation',
					sheet: 'Sheet1',
					range: 'A1:A3',
					rule: { type: 'list', formula1: '"Yes,No"' },
				},
				{ op: 'clearAutoFilter', sheet: 'Sheet1' },
			],
			{ journal: true },
		)

		expect(changed.wouldSucceed).toBe(true)
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(false)
		expect(changed.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message:
					'Data validation extension metadata at Sheet1!A1:A3 cannot be restored with public operations',
				refs: ['Sheet1!A1:A3'],
			},
			{
				code: 'LOSSY_INVERSE',
				message:
					'AutoFilter column 0 on Sheet1!A1:B10 cannot be fully restored with public operations',
				refs: ['Sheet1!A1:B10'],
			},
			{
				code: 'LOSSY_INVERSE',
				message:
					'AutoFilter extension metadata on Sheet1!A1:B10 cannot be restored with public operations',
				refs: ['Sheet1!A1:B10'],
			},
			{
				code: 'LOSSY_INVERSE',
				message:
					'AutoFilter sort metadata on Sheet1!A1:B10 cannot be fully restored with public operations',
				refs: ['Sheet1!A1:B10'],
			},
		])
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:B10' },
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:B10',
				sortRef: 'A1:B10',
				sortBy: 'A2:A10',
				descending: true,
			},
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1:A3',
				rule: { type: 'list', formula1: '"Open,Closed"' },
			},
		])
	})

	test('journal inverse ops restore defined name edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A10' }])

		const changed = wb.apply([{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!B1:B10' }], {
			journal: true,
		})

		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A10' },
		])

		const undoChange = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undoChange.errors).toEqual([])
		expect(wb.definedName('Budget')?.formula).toBe('Sheet1!A1:A10')

		const deleted = wb.apply([{ op: 'deleteDefinedName', name: 'Budget' }], { journal: true })
		expect(deleted.journal?.inverseOps).toEqual([
			{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A10' },
		])

		const undoDelete = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undoDelete.errors).toEqual([])
		expect(wb.definedName('Budget')?.formula).toBe('Sheet1!A1:A10')

		const created = wb.apply([{ op: 'setDefinedName', name: 'Scratch', ref: 'Sheet1!C1:C2' }], {
			journal: true,
		})
		expect(created.journal?.inverseOps).toEqual([{ op: 'deleteDefinedName', name: 'Scratch' }])
	})

	test('journal inverse ops restore workbook and document properties', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setWorkbookProperties',
				properties: { codeName: 'OriginalBook', filterPrivacy: true, date1904: true },
			},
			{
				op: 'setDocumentProperties',
				properties: {
					core: { title: 'Original', creator: 'Analyst' },
					app: { Company: 'Ascend', TitlesOfParts: ['Sheet1'] },
					custom: [{ name: 'Desk', value: 'North', type: 'lpwstr', pid: 2 }],
				},
			},
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{
					op: 'setWorkbookProperties',
					mode: 'replace',
					properties: { codeName: 'AgentBook', date1904: false },
				},
				{
					op: 'setDocumentProperties',
					mode: 'replace',
					properties: {
						core: { title: 'Changed', lastModifiedBy: 'agent' },
						app: { Company: 'Agent' },
						custom: [{ name: 'Risk', value: true, type: 'bool', pid: 3 }],
					},
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setDocumentProperties',
				mode: 'replace',
				properties: {
					core: { title: 'Original', creator: 'Analyst' },
					app: { Company: 'Ascend', TitlesOfParts: ['Sheet1'] },
					custom: [{ name: 'Desk', value: 'North', type: 'lpwstr', pid: 2 }],
				},
			},
			{
				op: 'setWorkbookProperties',
				mode: 'replace',
				properties: { codeName: 'OriginalBook', filterPrivacy: true, date1904: true },
			},
		])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal inverse ops restore workbook views and calc settings', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setWorkbookView', index: 0, view: { activeTab: 0, firstSheet: 0, tabRatio: 600 } },
			{
				op: 'setCalcSettings',
				settings: {
					calcMode: 'auto',
					fullCalcOnLoad: false,
					calcCompleted: true,
					calcOnSave: true,
					forceFullCalc: false,
					calcId: 1,
					dateSystem: '1900',
					iterativeCalc: { enabled: false, maxIterations: 100, maxChange: 0.001 },
				},
			},
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{
					op: 'setWorkbookView',
					index: 0,
					mode: 'replace',
					view: { activeTab: 2, firstSheet: 1 },
				},
				{
					op: 'setWorkbookView',
					index: 1,
					view: { activeTab: 1, tabRatio: 720 },
				},
				{
					op: 'setCalcSettings',
					settings: {
						calcMode: 'manual',
						fullCalcOnLoad: true,
						calcCompleted: null,
						calcOnSave: false,
						forceFullCalc: true,
						calcId: 42,
						dateSystem: '1904',
						iterativeCalc: { enabled: true, maxIterations: 50, maxChange: 0.0001 },
					},
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setCalcSettings',
				settings: {
					calcMode: 'auto',
					fullCalcOnLoad: false,
					calcCompleted: true,
					calcOnSave: true,
					forceFullCalc: false,
					calcId: 1,
					dateSystem: '1900',
					iterativeCalc: { enabled: false, maxIterations: 100, maxChange: 0.001 },
				},
			},
			{ op: 'setWorkbookProperties', properties: { date1904: false }, mode: 'replace' },
			{ op: 'setWorkbookView', index: 1, view: null },
			{
				op: 'setWorkbookView',
				index: 0,
				view: { activeTab: 0, firstSheet: 0, tabRatio: 600 },
				mode: 'replace',
			},
		])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal inverse ops restore workbook protection when public ops can be exact', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setWorkbookProtection',
				protection: {
					lockStructure: true,
					workbookAlgorithmName: 'SHA-512',
					workbookSpinCount: 100000,
				},
			},
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{
					op: 'setWorkbookProtection',
					protection: { lockWindows: true, workbookPassword: 'ABCD' },
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setWorkbookProtection',
				protection: {
					lockStructure: true,
					workbookAlgorithmName: 'SHA-512',
					workbookSpinCount: 100000,
				},
			},
		])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)

		const unprotected = AscendWorkbook.create()
		const lossy = unprotected.preview(
			[{ op: 'setWorkbookProtection', protection: { lockStructure: true } }],
			{ journal: true },
		)
		expect(lossy.journal?.supported).toBe(true)
		expect(lossy.journal?.exact).toBe(false)
		expect(lossy.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Workbook protection absence cannot be restored exactly with public operations',
			},
		])
		expect(lossy.journal?.inverseOps).toEqual([{ op: 'setWorkbookProtection', protection: {} }])
	})

	test('theme journals restore representable edits and mark additions lossy', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setTheme',
				themeName: 'Office',
				colorSchemeName: 'Office Colors',
				majorFontLatin: 'Aptos Display',
				minorFontLatin: 'Aptos',
				themeColors: [
					{ slot: 'accent1', rgb: '4F81BD' },
					{ slot: 'lt1', systemColor: 'window', lastColor: 'FFFFFF' },
				],
			},
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{
					op: 'setTheme',
					themeName: 'Brand',
					colorSchemeName: 'Brand Colors',
					majorFontLatin: 'Inter Display',
					minorFontLatin: 'Inter',
					themeColors: [
						{ slot: 'accent1', rgb: '0F6CBD' },
						{ slot: 'lt1', systemColor: 'windowText', lastColor: '000000' },
					],
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setTheme',
				themeName: 'Office',
				colorSchemeName: 'Office Colors',
				majorFontLatin: 'Aptos Display',
				minorFontLatin: 'Aptos',
				themeColors: [
					{ slot: 'accent1', rgb: '4F81BD' },
					{ slot: 'lt1', systemColor: 'window', lastColor: 'FFFFFF' },
				],
			},
		])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)

		const sparse = AscendWorkbook.create()
		const lossy = sparse.preview(
			[
				{
					op: 'setTheme',
					themeName: 'New Brand',
					themeColors: [{ slot: 'accent1', rgb: '123456' }],
				},
			],
			{ journal: true },
		)
		expect(lossy.journal?.supported).toBe(true)
		expect(lossy.journal?.exact).toBe(false)
		expect(lossy.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Theme metadata field themeName cannot be removed with public operations',
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Theme color slot accent1 cannot be removed with public operations',
			},
		])
		expect(lossy.journal?.inverseOps).toEqual([])
	})

	test('journal exact inverse restores mixed semantic workbook state', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'C1', value: 'Amount' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 2 },
					{ ref: 'C2', value: 20 },
					{ ref: 'A3', value: 'East' },
					{ ref: 'B3', value: 3 },
					{ ref: 'C3', value: 30 },
				],
			},
			{ op: 'setStyle', sheet: 'Sheet1', range: 'B2:B2', style: { numberFormat: '0.00' } },
			{ op: 'setComment', sheet: 'Sheet1', ref: 'D4', text: 'initial', author: 'analyst' },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'E5', url: 'https://example.com' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A2:A5',
				rule: { type: 'list', formula1: '"West,East"', allowBlank: true },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'B2:B5',
				rule: { type: 'cellIs', operator: 'greaterThan', formula: '1', priority: 2 },
			},
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'F1:G1' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C5', column: 0, values: ['West'] },
			{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$B$2:$B$3' },
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:C3', name: 'Sales', hasHeaders: true },
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
			{ op: 'setTableColumn', table: 'Sales', column: 'Amount', formula: '=SUM([Qty])' },
			{ op: 'freezePane', sheet: 'Sheet1', row: 1, col: 1 },
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B2', value: 9 }] },
				{ op: 'setFormula', sheet: 'Sheet1', ref: 'H2', formula: 'B2*10' },
				{ op: 'setStyle', sheet: 'Sheet1', range: 'B2:B2', style: { numberFormat: '$0.00' } },
				{ op: 'setComment', sheet: 'Sheet1', ref: 'D4', text: 'changed', author: 'agent' },
				{
					op: 'setHyperlink',
					sheet: 'Sheet1',
					ref: 'E5',
					url: 'https://example.org',
					display: 'updated',
				},
				{
					op: 'setDataValidation',
					sheet: 'Sheet1',
					range: 'A2:A5',
					rule: { type: 'list', formula1: '"North,South"' },
				},
				{
					op: 'setConditionalFormat',
					sheet: 'Sheet1',
					range: 'B2:B5',
					rule: { type: 'expression', formula: 'B2>5', priority: 1 },
				},
				{ op: 'unmergeCells', sheet: 'Sheet1', range: 'F1:G1' },
				{ op: 'clearAutoFilter', sheet: 'Sheet1' },
				{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$C$2:$C$3' },
				{ op: 'renameTable', table: 'Sales', newName: 'Revenue' },
				{
					op: 'setTableColumn',
					table: 'Revenue',
					column: 'Amount',
					newName: 'RevenueAmount',
					formula: '=SUM([Qty])*2',
				},
				{ op: 'setTableStyle', table: 'Revenue', styleName: null },
				{ op: 'freezePane', sheet: 'Sheet1', row: 2, col: 0 },
				{ op: 'renameSheet', sheet: 'Sheet1', newName: 'Summary' },
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps.length).toBeGreaterThan(0)
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal exact inverse restores deleted metadata surfaces', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'B2', text: 'review', author: 'analyst' },
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'C3',
				url: 'https://example.com',
				display: 'Report',
			},
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'D2:D10',
				rule: { type: 'whole', operator: 'greaterThan', formula1: '0', allowBlank: true },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'E2:E10',
				rule: { type: 'cellIs', operator: 'greaterThan', formula: '5', priority: 3 },
			},
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:E10', column: 0, values: ['Open'] },
			{ op: 'setDefinedName', name: 'ReviewRange', ref: 'Sheet1!$B$2:$E$10' },
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{ op: 'deleteComment', sheet: 'Sheet1', ref: 'B2' },
				{ op: 'deleteHyperlink', sheet: 'Sheet1', ref: 'C3' },
				{ op: 'deleteDataValidation', sheet: 'Sheet1', range: 'D2:D10' },
				{ op: 'deleteConditionalFormat', sheet: 'Sheet1', range: 'E2:E10' },
				{ op: 'clearAutoFilter', sheet: 'Sheet1' },
				{ op: 'deleteDefinedName', name: 'ReviewRange' },
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal inverse ops restore sheet moves across adjacent renames', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'addSheet', name: 'Data' },
			{ op: 'addSheet', name: 'Report' },
			{ op: 'setCells', sheet: 'Report', updates: [{ ref: 'A1', value: 'report' }] },
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{ op: 'moveSheet', sheet: 'Report', position: 0 },
				{ op: 'renameSheet', sheet: 'Report', newName: 'Summary' },
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'renameSheet', sheet: 'Summary', newName: 'Report' },
			{ op: 'moveSheet', sheet: 'Report', position: 2 },
		])
		expect(wb.sheets).toEqual(['Summary', 'Sheet1', 'Data'])

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal inverse ops restore copied sheets and copied chart metadata', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Amount' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 20 },
				],
			},
		])
		wb.getWorkbookModel().chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			chartType: 'barChart',
			series: [{ nameRef: 'Sheet1!$B$1', categoryRef: 'Sheet1!$A$2', valueRef: 'Sheet1!$B$2' }],
		})
		const before = journalComparableState(wb)

		const copied = wb.apply([{ op: 'copySheet', sheet: 'Sheet1', newName: 'Copy', position: 0 }], {
			journal: true,
		})

		expect(copied.errors).toEqual([])
		expect(copied.journal?.supported).toBe(true)
		expect(copied.journal?.exact).toBe(true)
		expect(copied.journal?.inverseOps).toEqual([{ op: 'deleteSheet', sheet: 'Copy' }])
		expect(wb.sheets[0]).toBe('Copy')
		expect(wb.getWorkbookModel().chartParts).toHaveLength(2)
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(copied.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal inverse ops restore empty deleted sheets exactly', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Scratch', position: 0 }])
		const before = journalComparableState(wb)

		const deleted = wb.apply([{ op: 'deleteSheet', sheet: 'Scratch' }], { journal: true })

		expect(deleted.errors).toEqual([])
		expect(deleted.journal?.supported).toBe(true)
		expect(deleted.journal?.exact).toBe(true)
		expect(deleted.journal?.inverseOps).toEqual([{ op: 'addSheet', name: 'Scratch', position: 0 }])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('delete sheet journals surface lost sheet contents and dependent metadata', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'addSheet', name: 'Data' },
			{
				op: 'setCells',
				sheet: 'Data',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Amount' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 20 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: 'Data!B2' },
			{ op: 'setDefinedName', name: 'DataTotal', ref: 'Data!$B$2' },
			{ op: 'setComment', sheet: 'Data', ref: 'A1', text: 'source', author: 'analyst' },
		])
		wb.getWorkbookModel().chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			series: [{ valueRef: 'Data!$B$2' }],
		})

		const deleted = wb.apply([{ op: 'deleteSheet', sheet: 'Data' }], { journal: true })

		expect(deleted.errors).toEqual([])
		expect(deleted.journal?.supported).toBe(true)
		expect(deleted.journal?.exact).toBe(false)
		expect(deleted.journal?.inverseOps).toEqual([{ op: 'addSheet', name: 'Data', position: 1 }])
		expect(deleted.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Deleted sheet Data cannot be fully restored with public operations',
				refs: [
					'Data!cells:4',
					'Data!comments:1',
					'name:DataTotal',
					'Sheet1!A1',
					'chart:xl/charts/chart1.xml:series:0:valueRef',
				],
			},
		])
	})

	test('journal inverse ops restore existing row and column layout edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setRowHeight', sheet: 'Sheet1', row: 2, height: 24 },
			{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 14 },
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{ op: 'setRowHeight', sheet: 'Sheet1', row: 2, height: 32 },
				{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 20 },
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 14 },
			{ op: 'setRowHeight', sheet: 'Sheet1', row: 2, height: 24 },
		])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal marks newly-created row and column layout as lossy', () => {
		const wb = AscendWorkbook.create()
		const changed = wb.preview(
			[
				{ op: 'setRowHeight', sheet: 'Sheet1', row: 2, height: 32 },
				{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 20 },
			],
			{ journal: true },
		)

		expect(changed.wouldSucceed).toBe(true)
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(false)
		expect(changed.journal?.inverseOps).toEqual([])
		expect(changed.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Created row layout at Sheet1!3 cannot be cleared with public operations',
				refs: ['Sheet1!3'],
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Created col layout at Sheet1!B cannot be cleared with public operations',
				refs: ['Sheet1!B'],
			},
		])
	})

	test('journal inverse ops restore representable sheet tab color and protection metadata', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setTabColor', sheet: 'Sheet1', color: 'FF0000' },
			{
				op: 'setSheetProtection',
				sheet: 'Sheet1',
				password: 'ABCD',
				options: { formatCells: false, autoFilter: true },
			},
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{ op: 'setTabColor', sheet: 'Sheet1', color: '00FF00' },
				{
					op: 'setSheetProtection',
					sheet: 'Sheet1',
					password: 'DCBA',
					options: { insertRows: true, deleteRows: false },
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setSheetProtection',
				sheet: 'Sheet1',
				password: 'ABCD',
				options: { formatCells: false, autoFilter: true },
			},
			{ op: 'setTabColor', sheet: 'Sheet1', color: 'FF0000' },
		])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal classifies un-restorable sheet tab color and protection metadata as lossy', () => {
		const wb = AscendWorkbook.create()
		const modelSheet = wb.getWorkbookModel().getSheet('Sheet1')
		if (modelSheet) {
			modelSheet.tabColor = { theme: 2, tint: -0.25 }
			modelSheet.protection = { sheet: true, objects: true, algorithmName: 'SHA-512' }
		}

		const changed = wb.preview(
			[
				{ op: 'setTabColor', sheet: 'Sheet1', color: '00FF00' },
				{ op: 'setSheetProtection', sheet: 'Sheet1', options: { insertRows: true } },
			],
			{ journal: true },
		)

		expect(changed.wouldSucceed).toBe(true)
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(false)
		expect(changed.journal?.inverseOps).toEqual([{ op: 'setSheetProtection', sheet: 'Sheet1' }])
		expect(changed.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message:
					'Sheet tab color for Sheet1 uses unsupported color metadata and cannot be fully restored with public operations',
				refs: ['sheet:Sheet1:tabColor'],
			},
			{
				code: 'LOSSY_INVERSE',
				message:
					'Sheet protection for Sheet1 contains metadata that cannot be fully restored with public operations',
				refs: ['sheet:Sheet1:protection:objects', 'sheet:Sheet1:protection:algorithmName'],
			},
		])
	})

	test('journal marks newly-created tab color and sheet protection as lossy', () => {
		const wb = AscendWorkbook.create()
		const changed = wb.preview(
			[
				{ op: 'setTabColor', sheet: 'Sheet1', color: '00FF00' },
				{ op: 'setSheetProtection', sheet: 'Sheet1', options: { sort: true } },
			],
			{ journal: true },
		)

		expect(changed.wouldSucceed).toBe(true)
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(false)
		expect(changed.journal?.inverseOps).toEqual([])
		expect(changed.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Sheet tab color absence for Sheet1 cannot be restored with public operations',
				refs: ['sheet:Sheet1:tabColor'],
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Sheet protection absence for Sheet1 cannot be restored with public operations',
				refs: ['sheet:Sheet1:protection'],
			},
		])
	})

	test('journal inverse ops restore sheet and existing row/column visibility metadata', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'hideSheet', sheet: 'Sheet1', hidden: true },
			{ op: 'setRowHeight', sheet: 'Sheet1', row: 2, height: 24 },
			{ op: 'hideCols', sheet: 'Sheet1', at: 1, count: 1, hidden: true },
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{ op: 'hideSheet', sheet: 'Sheet1', hidden: false },
				{ op: 'hideRows', sheet: 'Sheet1', at: 2, count: 1, hidden: true },
				{ op: 'hideCols', sheet: 'Sheet1', at: 1, count: 1, hidden: false },
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'hideCols', sheet: 'Sheet1', at: 1, count: 1, hidden: true },
			{ op: 'setRowHeight', sheet: 'Sheet1', row: 2, height: 24 },
			{ op: 'hideSheet', sheet: 'Sheet1', hidden: true },
		])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal classifies un-restorable visibility metadata as lossy', () => {
		const wb = AscendWorkbook.create()
		const modelSheet = wb.getWorkbookModel().getSheet('Sheet1')
		if (modelSheet) modelSheet.state = 'veryHidden'

		const changed = wb.preview(
			[
				{ op: 'hideSheet', sheet: 'Sheet1', hidden: false },
				{ op: 'hideRows', sheet: 'Sheet1', at: 2, count: 1, hidden: true },
				{ op: 'hideCols', sheet: 'Sheet1', at: 1, count: 1, hidden: true },
			],
			{ journal: true },
		)

		expect(changed.wouldSucceed).toBe(true)
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(false)
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'hideSheet', sheet: 'Sheet1', hidden: true },
		])
		expect(changed.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message:
					'Sheet visibility for Sheet1 was veryHidden and cannot be restored with public operations',
				refs: ['sheet:Sheet1:state:veryHidden'],
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Created row hide metadata cannot be cleared with public operations',
				refs: ['Sheet1!3'],
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Created or unkeyed column hide metadata cannot be cleared with public operations',
				refs: ['Sheet1!B'],
			},
		])
	})

	test('journal inverse ops restore table renames and column metadata edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'C1', value: 'Amount' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 2 },
					{ ref: 'C2', value: 5 },
					{ ref: 'A3', value: 'Total' },
					{ ref: 'B3', value: 2 },
					{ ref: 'C3', value: 5 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:C3', name: 'Sales', hasHeaders: true },
			{
				op: 'setTableColumn',
				table: 'Sales',
				column: 'Amount',
				formula: '=[@Qty]*10',
				totalsRowFunction: 'sum',
			},
		])
		const internal = wb as unknown as {
			wb: {
				sheets: Array<{
					tables: Array<{
						hasTotals?: boolean
						ref: { start: { row: number; col: number }; end: { row: number; col: number } }
					}>
				}>
			}
		}
		const tableModel = internal.wb.sheets[0]?.tables[0]
		if (tableModel) {
			tableModel.hasTotals = true
			tableModel.ref = { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } }
		}

		const columnEdit = wb.apply(
			[
				{
					op: 'setTableColumn',
					table: 'Sales',
					column: 'Amount',
					newName: 'Revenue',
					formula: '=[@Qty]*12',
					totalsRowFunction: 'average',
				},
			],
			{ journal: true },
		)

		expect(columnEdit.journal?.exact).toBe(true)
		expect(columnEdit.journal?.inverseOps).toEqual([
			{
				op: 'setTableColumn',
				table: 'Sales',
				column: 'Revenue',
				newName: 'Amount',
				formula: '[@Qty]*10',
				totalsRowFunction: 'sum',
			},
		])

		const undoColumn = wb.apply(columnEdit.journal?.inverseOps ?? [], { transaction: true })
		expect(undoColumn.errors).toEqual([])
		expect(wb.table('Sales')?.columns).toEqual(['Region', 'Qty', 'Amount'])
		expect(wb.table('Sales')?.columnDefs[2]).toMatchObject({
			name: 'Amount',
			formula: '[@Qty]*10',
			totalsRowFunction: 'sum',
		})

		const renamed = wb.apply([{ op: 'renameTable', table: 'Sales', newName: 'RevenueTable' }], {
			journal: true,
		})
		expect(renamed.journal?.inverseOps).toEqual([
			{ op: 'renameTable', table: 'RevenueTable', newName: 'Sales' },
		])

		const undoRename = wb.apply(renamed.journal?.inverseOps ?? [], { transaction: true })
		expect(undoRename.errors).toEqual([])
		expect(wb.table('Sales')?.name).toBe('Sales')
	})

	test('journal inverse ops restore table lifecycle metadata edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Product' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'A2', value: 'Widget' },
					{ ref: 'B2', value: 2 },
					{ ref: 'A3', value: 'Gadget' },
					{ ref: 'B3', value: 3 },
				],
			},
		])

		const created = wb.apply(
			[{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B3', name: 'Sales', hasHeaders: true }],
			{ journal: true },
		)
		expect(created.journal?.exact).toBe(true)
		expect(created.journal?.inverseOps).toEqual([{ op: 'deleteTable', table: 'Sales' }])

		wb.apply([
			{ op: 'setTableColumn', table: 'Sales', column: 'Qty', formula: '=1+1' },
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
		])

		const resized = wb.apply([{ op: 'resizeTable', table: 'Sales', ref: 'A1:B2' }], {
			journal: true,
		})
		expect(resized.journal?.exact).toBe(true)
		expect(resized.journal?.inverseOps).toEqual([
			{ op: 'resizeTable', table: 'Sales', ref: 'A1:B3' },
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
			{ op: 'setTableColumn', table: 'Sales', column: 1, formula: '1+1' },
		])

		const undoResize = wb.apply(resized.journal?.inverseOps ?? [], { transaction: true })
		expect(undoResize.errors).toEqual([])
		expect(wb.table('Sales')?.ref).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 2, col: 1 },
		})
		expect(wb.table('Sales')?.columnDefs[1]?.formula).toBe('1+1')

		const deleted = wb.apply([{ op: 'deleteTable', table: 'Sales' }], { journal: true })
		expect(deleted.journal?.exact).toBe(true)
		expect(deleted.journal?.inverseOps).toEqual([
			{
				op: 'createTable',
				sheet: 'Sheet1',
				ref: 'A1:B3',
				name: 'Sales',
				hasHeaders: true,
			},
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
			{ op: 'setTableColumn', table: 'Sales', column: 0, newName: 'Product' },
			{ op: 'setTableColumn', table: 'Sales', column: 1, newName: 'Qty', formula: '1+1' },
		])

		const undoDelete = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undoDelete.errors).toEqual([])
		expect(wb.table('Sales')?.columns).toEqual(['Product', 'Qty'])
		expect(wb.table('Sales')?.columnDefs[1]?.formula).toBe('1+1')
		expect(wb.table('Sales')?.styleInfo).toEqual({ name: 'TableStyleMedium2' })
	})

	test('journal exact inverse restores table lifecycle mixed with structural row edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A4', value: 'Region' },
					{ ref: 'B4', value: 'Amount' },
					{ ref: 'A5', value: 'West' },
					{ ref: 'B5', value: 20 },
					{ ref: 'A6', value: 'East' },
					{ ref: 'B6', value: 30 },
				],
			},
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{
					op: 'createTable',
					sheet: 'Sheet1',
					ref: 'A4:B6',
					name: 'Sales',
					hasHeaders: true,
				},
				{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
				{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 },
				{ op: 'renameTable', table: 'Sales', newName: 'Revenue' },
				{
					op: 'setTableColumn',
					table: 'Revenue',
					column: 'Amount',
					newName: 'Net',
					formula: '=[@Net]*2',
					totalsRowFormula: 'SUM([Net])',
				},
				{ op: 'deleteRows', sheet: 'Sheet1', at: 2, count: 1 },
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.issues).toEqual([])
		expect(wb.table('Revenue')?.ref).toEqual({
			start: { row: 3, col: 0 },
			end: { row: 5, col: 1 },
		})
		expect(wb.table('Revenue')?.columns).toEqual(['Region', 'Net'])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal exact inverse restores table lifecycle mixed with structural column edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'D1', value: 'Region' },
					{ ref: 'E1', value: 'Amount' },
					{ ref: 'D2', value: 'West' },
					{ ref: 'E2', value: 20 },
					{ ref: 'D3', value: 'East' },
					{ ref: 'E3', value: 30 },
				],
			},
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{
					op: 'createTable',
					sheet: 'Sheet1',
					ref: 'D1:E3',
					name: 'Sales',
					hasHeaders: true,
				},
				{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
				{ op: 'insertCols', sheet: 'Sheet1', at: 1, count: 1 },
				{ op: 'renameTable', table: 'Sales', newName: 'Revenue' },
				{
					op: 'setTableColumn',
					table: 'Revenue',
					column: 'Amount',
					newName: 'Net',
					formula: '=[@Net]*2',
					totalsRowFormula: 'SUM([Net])',
				},
				{ op: 'deleteCols', sheet: 'Sheet1', at: 2, count: 1 },
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.issues).toEqual([])
		expect(wb.table('Revenue')?.ref).toEqual({
			start: { row: 0, col: 3 },
			end: { row: 2, col: 4 },
		})
		expect(wb.table('Revenue')?.columns).toEqual(['Region', 'Net'])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal inverse ops restore structural row and column edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'top' },
					{ ref: 'A2', value: 'deleted-row' },
					{ ref: 'B2', value: 4 },
					{ ref: 'A3', value: 'bottom' },
				],
			},
			{ op: 'setStyle', sheet: 'Sheet1', range: 'B2:B2', style: { numberFormat: '0.00' } },
		])

		const inserted = wb.apply([{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})
		expect(inserted.journal?.exact).toBe(true)
		expect(inserted.journal?.inverseOps).toEqual([
			{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 },
		])
		const undoInsert = wb.apply(inserted.journal?.inverseOps ?? [], { transaction: true })
		expect(undoInsert.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('A2')?.value).toEqual({
			kind: 'string',
			value: 'deleted-row',
		})

		const deletedRow = wb.apply([{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})
		expect(deletedRow.journal?.exact).toBe(true)
		expect(deletedRow.journal?.inverseOps).toEqual([
			{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 },
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A2', value: 'deleted-row' },
					{ ref: 'B2', value: 4 },
				],
			},
			{ op: 'clearRange', sheet: 'Sheet1', range: 'A2', what: 'styles' },
			{ op: 'clearRange', sheet: 'Sheet1', range: 'B2', what: 'styles' },
			{ op: 'setStyle', sheet: 'Sheet1', range: 'B2', style: { numberFormat: '0.00' } },
		])
		const undoDeleteRow = wb.apply(deletedRow.journal?.inverseOps ?? [], { transaction: true })
		expect(undoDeleteRow.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('A2')?.value).toEqual({
			kind: 'string',
			value: 'deleted-row',
		})
		expect(wb.cellStyle('Sheet1!B2')?.numberFormat).toBe('0.00')

		const deletedCol = wb.apply([{ op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})
		expect(deletedCol.journal?.exact).toBe(true)
		expect(deletedCol.journal?.inverseOps[0]).toEqual({
			op: 'insertCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
		})
		const undoDeleteCol = wb.apply(deletedCol.journal?.inverseOps ?? [], { transaction: true })
		expect(undoDeleteCol.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('B2')?.value).toEqual({ kind: 'number', value: 4 })
	})

	test('structural delete journals mark metadata intersections as lossy', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'header' }] },
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'C2:D2' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'F2:F2',
				rule: { type: 'list', formula1: '"yes,no"' },
			},
		])

		const deletedRow = wb.preview([{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})

		expect(deletedRow.wouldSucceed).toBe(true)
		expect(deletedRow.journal?.supported).toBe(true)
		expect(deletedRow.journal?.exact).toBe(false)
		expect(deletedRow.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Deleted row metadata on Sheet1 cannot be fully restored with public operations',
				refs: ['Sheet1!2'],
			},
		])
		expect(deletedRow.journal?.inverseOps[0]).toEqual({
			op: 'insertRows',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
		})
		expect(wb.sheet('Sheet1')?.merges).toEqual([
			{ start: { row: 1, col: 2 }, end: { row: 1, col: 3 } },
		])
	})

	test('column delete journals mark metadata intersections as lossy', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Status' },
					{ ref: 'C1', value: 'Amount' },
					{ ref: 'A2', value: 'A' },
					{ ref: 'B2', value: 'Open' },
					{ ref: 'C2', value: 10 },
				],
			},
			{ op: 'setComment', sheet: 'Sheet1', ref: 'B2', text: 'review' },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'B3', url: 'https://example.com' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'B2:B5',
				rule: { type: 'list', formula1: '"Open,Closed"' },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'B2:B5',
				rule: { type: 'expression', formula: 'B2="Open"', priority: 1 },
			},
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C5', column: 1, values: ['Open'] },
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:C2', name: 'Tasks', hasHeaders: true },
		])

		const deletedCol = wb.preview([{ op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})

		expect(deletedCol.wouldSucceed).toBe(true)
		expect(deletedCol.journal?.supported).toBe(true)
		expect(deletedCol.journal?.exact).toBe(false)
		expect(deletedCol.journal?.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message: 'Deleted column metadata on Sheet1 cannot be fully restored with public operations',
			refs: ['Sheet1!B'],
		})
		expect(deletedCol.journal?.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message:
				'Deleted column formula references on Sheet1 cannot be restored with public operations',
			refs: ['Sheet1!conditionalFormat:B2:B5:0:0:0'],
		})
		expect(deletedCol.journal?.inverseOps[0]).toEqual({
			op: 'insertCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
		})
		expect(wb.sheet('Sheet1')?.getComments()).toEqual([{ ref: 'B2', text: 'review' }])
	})

	test('structural delete journals mark represented package metadata as lossy', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Header' },
					{ ref: 'A2', value: 1 },
					{ ref: 'B2', value: 2 },
					{ ref: 'D2', value: 'Pivot' },
				],
			},
		])
		const model = wb.getWorkbookModel()
		const sheet = model.getSheet('Sheet1')
		if (!sheet) throw new Error('sheet missing')
		sheet.ignoredErrors.push({ sqref: 'A2:A2', formula: true })
		sheet.sortState = {
			ref: 'A1:A3',
			conditions: [{ ref: 'A2:A2' }],
		}
		sheet.advancedFilters.push({
			ref: 'B1:B3',
			autoFilter: {
				ref: 'B1:B3',
				columns: [],
				sortState: { ref: 'B1:B3', conditions: [{ ref: 'B2:B2' }] },
			},
			filterColumnCount: 0,
			sortConditionCount: 1,
		})
		sheet.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rId1',
			targetPath: '../media/image1.png',
			anchor: {
				kind: 'twoCell',
				from: { row: 1, col: 0 },
				to: { row: 2, col: 1 },
			},
		})
		sheet.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'shape',
			anchor: { kind: 'oneCell', from: { row: 1, col: 1 } },
		})
		model.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			series: [{ valueRef: 'Sheet1!$A$2:$A$2' }],
		})
		model.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 1,
			sourceSheet: 'Sheet1',
			sourceRef: 'A2:B4',
			fields: [],
		})
		model.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 1,
			locationRef: 'D2:E5',
			location: { ref: 'D2:E5' },
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})

		const deletedRow = wb.preview([{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})

		expect(deletedRow.wouldSucceed).toBe(true)
		expect(deletedRow.journal?.supported).toBe(true)
		expect(deletedRow.journal?.exact).toBe(false)
		expect(deletedRow.journal?.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message:
				'Deleted row represented metadata on Sheet1 cannot be fully restored with public operations',
			refs: [
				'Sheet1!ignoredError:A2:A2',
				'Sheet1!sortState:A1:A3',
				'Sheet1!sortState:condition:0:A2:A2',
				'Sheet1!advancedFilter:0:B1:B3',
				'Sheet1!advancedFilter:0:autoFilter:B1:B3',
				'Sheet1!advancedFilter:0:autoFilter:sortState:B1:B3',
				'Sheet1!advancedFilter:0:autoFilter:sortState:condition:0:B2:B2',
				'Sheet1!image:xl/drawings/drawing1.xml:0',
				'Sheet1!drawing:xl/drawings/drawing1.xml:0',
				'chart:xl/charts/chart1.xml:series:0:valueRef',
				'pivotCache:xl/pivotCache/pivotCacheDefinition1.xml:sourceRef',
				'pivotTable:xl/pivotTables/pivotTable1.xml:locationRef',
				'pivotTable:xl/pivotTables/pivotTable1.xml:location.ref',
			],
		})
	})

	test('column delete journals mark represented package metadata as lossy', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Header' },
					{ ref: 'B1', value: 'Sorted' },
					{ ref: 'B2', value: 1 },
					{ ref: 'C2', value: 2 },
					{ ref: 'D2', value: 'Pivot' },
				],
			},
		])
		const model = wb.getWorkbookModel()
		const sheet = model.getSheet('Sheet1')
		if (!sheet) throw new Error('sheet missing')
		sheet.ignoredErrors.push({ sqref: 'B2:B2', formula: true })
		sheet.sortState = {
			ref: 'B1:D1',
			conditions: [{ ref: 'B2:B2' }],
		}
		sheet.advancedFilters.push({
			ref: 'B1:B3',
			autoFilter: {
				ref: 'B1:B3',
				columns: [],
				sortState: { ref: 'B1:B3', conditions: [{ ref: 'B2:B2' }] },
			},
			filterColumnCount: 0,
			sortConditionCount: 1,
		})
		sheet.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rId1',
			targetPath: '../media/image1.png',
			anchor: {
				kind: 'twoCell',
				from: { row: 0, col: 1 },
				to: { row: 2, col: 2 },
			},
		})
		sheet.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'shape',
			anchor: { kind: 'oneCell', from: { row: 1, col: 1 } },
		})
		model.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			series: [{ valueRef: 'Sheet1!$B$2:$B$2' }],
		})
		model.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 1,
			sourceSheet: 'Sheet1',
			sourceRef: 'B2:C4',
			fields: [],
		})
		model.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 1,
			locationRef: 'B4:C8',
			location: { ref: 'B4:C8' },
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})

		const deletedCol = wb.preview([{ op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})

		expect(deletedCol.wouldSucceed).toBe(true)
		expect(deletedCol.journal?.supported).toBe(true)
		expect(deletedCol.journal?.exact).toBe(false)
		expect(deletedCol.journal?.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message:
				'Deleted column represented metadata on Sheet1 cannot be fully restored with public operations',
			refs: [
				'Sheet1!ignoredError:B2:B2',
				'Sheet1!sortState:B1:D1',
				'Sheet1!sortState:condition:0:B2:B2',
				'Sheet1!advancedFilter:0:B1:B3',
				'Sheet1!advancedFilter:0:autoFilter:B1:B3',
				'Sheet1!advancedFilter:0:autoFilter:sortState:B1:B3',
				'Sheet1!advancedFilter:0:autoFilter:sortState:condition:0:B2:B2',
				'Sheet1!image:xl/drawings/drawing1.xml:0',
				'Sheet1!drawing:xl/drawings/drawing1.xml:0',
				'chart:xl/charts/chart1.xml:series:0:valueRef',
				'pivotCache:xl/pivotCache/pivotCacheDefinition1.xml:sourceRef',
				'pivotTable:xl/pivotTables/pivotTable1.xml:locationRef',
				'pivotTable:xl/pivotTables/pivotTable1.xml:location.ref',
			],
		})
	})

	test('structural row delete exact journals restore shifted metadata outside the deleted band', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A4', value: 'Region' },
					{ ref: 'B4', value: 'Amount' },
					{ ref: 'A5', value: 'West' },
					{ ref: 'B5', value: 20 },
					{ ref: 'A6', value: 'East' },
					{ ref: 'B6', value: 30 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'H4', formula: 'B5*2' },
			{ op: 'setComment', sheet: 'Sheet1', ref: 'C4', text: 'below row', author: 'agent' },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'C5', url: 'https://example.com' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'D4:D6',
				rule: { type: 'whole', operator: 'greaterThan', formula1: '0' },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'E4:E6',
				rule: { type: 'expression', formula: 'B5>0' },
			},
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'F4:G4' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A4:B6', column: 0, values: ['West'] },
			{ op: 'setDefinedName', name: 'BelowBand', ref: 'Sheet1!$A$4:$B$6' },
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A4:B6', name: 'ShiftedRows', hasHeaders: true },
		])
		const before = journalComparableState(wb)

		const deleted = wb.apply([{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})

		expect(deleted.errors).toEqual([])
		expect(deleted.journal?.supported).toBe(true)
		expect(deleted.journal?.exact).toBe(true)
		expect(deleted.journal?.issues).toEqual([])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('structural column delete exact journals restore shifted metadata outside the deleted band', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'D1', value: 'Region' },
					{ ref: 'E1', value: 'Amount' },
					{ ref: 'D2', value: 'West' },
					{ ref: 'E2', value: 20 },
					{ ref: 'D3', value: 'East' },
					{ ref: 'E3', value: 30 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'K1', formula: 'E2*2' },
			{ op: 'setComment', sheet: 'Sheet1', ref: 'F1', text: 'right of column', author: 'agent' },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'F2', url: 'https://example.com' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'G1:G3',
				rule: { type: 'whole', operator: 'greaterThan', formula1: '0' },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'H1:H3',
				rule: { type: 'expression', formula: 'E2>0' },
			},
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'I1:J1' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'D1:E3', column: 0, values: ['West'] },
			{ op: 'setDefinedName', name: 'RightBand', ref: 'Sheet1!$D$1:$E$3' },
			{ op: 'createTable', sheet: 'Sheet1', ref: 'D1:E3', name: 'ShiftedCols', hasHeaders: true },
		])
		const before = journalComparableState(wb)

		const deleted = wb.apply([{ op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})

		expect(deleted.errors).toEqual([])
		expect(deleted.journal?.supported).toBe(true)
		expect(deleted.journal?.exact).toBe(true)
		expect(deleted.journal?.issues).toEqual([])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('structural delete journals mark broken external formula references as lossy', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 5 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A2' },
			{ op: 'setDefinedName', name: 'DeletedInput', ref: 'Sheet1!A2' },
		])

		const deletedRow = wb.apply([{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})

		expect(deletedRow.errors).toEqual([])
		expect(deletedRow.journal?.supported).toBe(true)
		expect(deletedRow.journal?.exact).toBe(false)
		expect(deletedRow.journal?.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message: 'Deleted row formula references on Sheet1 cannot be restored with public operations',
			refs: ['Sheet1!B1', 'name:DeletedInput'],
		})
		expect(wb.sheet('Sheet1')?.cell('B1')?.formula).toBe('#REF!')

		const undo = wb.apply(deletedRow.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 5 })
		expect(wb.sheet('Sheet1')?.cell('B1')?.formula).toBe('#REF!')
	})

	test('structural delete journals mark broken metadata formula references as lossy', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 5 }] },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'C1:C1',
				rule: { type: 'custom', formula1: 'A2>0' },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'D1:D1',
				rule: { type: 'expression', formula: 'A2>0' },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'E1:E1',
				rule: {
					type: 'colorScale',
					colorScale: {
						cfvo: [{ type: 'formula', value: 'A2' }, { type: 'max' }],
						colors: [{ rgb: 'FFFF0000' }, { rgb: 'FF00FF00' }],
					},
				},
			},
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'F10', value: 'Name' },
					{ ref: 'G10', value: 'Calc' },
					{ ref: 'F11', value: 'row' },
					{ ref: 'G11', value: 1 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'F10:G11', name: 'Audit', hasHeaders: true },
			{
				op: 'setTableColumn',
				table: 'Audit',
				column: 'Calc',
				formula: 'A2',
				totalsRowFormula: 'SUM(A2)',
			},
		])

		const deletedRow = wb.apply([{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})

		expect(deletedRow.errors).toEqual([])
		expect(deletedRow.journal?.supported).toBe(true)
		expect(deletedRow.journal?.exact).toBe(false)
		expect(deletedRow.journal?.issues).toContainEqual(
			expect.objectContaining({
				code: 'LOSSY_INVERSE',
				message:
					'Deleted row formula references on Sheet1 cannot be restored with public operations',
				refs: expect.arrayContaining([
					'Sheet1!validation:C1:C1:formula1',
					'Sheet1!conditionalFormat:D1:D1:0:0:0',
					'Sheet1!conditionalFormat:E1:E1:1:0:colorScale.cfvo:0',
					'Sheet1!table:Audit:Calc:formula',
					'Sheet1!table:Audit:Calc:totalsRowFormula',
				]),
			}),
		)
		expect(wb.sheet('Sheet1')?.dataValidations[0]?.formula1).toBe('#REF!>0')
		expect(wb.sheet('Sheet1')?.conditionalFormats[0]?.rules[0]?.formulas[0]).toBe('#REF!>0')
		expect(wb.sheet('Sheet1')?.conditionalFormats[1]?.rules[0]?.colorScale?.cfvo[0]?.value).toBe(
			'#REF!',
		)
		expect(wb.table('Audit')?.columnDefs[1]?.formula).toBe('#REF!')
		expect(wb.table('Audit')?.columnDefs[1]?.totalsRowFormula).toBe('SUM(#REF!)')

		const undo = wb.apply(deletedRow.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.dataValidations[0]?.formula1).toBe('#REF!>0')
		expect(wb.sheet('Sheet1')?.conditionalFormats[0]?.rules[0]?.formulas[0]).toBe('#REF!>0')
		expect(wb.sheet('Sheet1')?.conditionalFormats[1]?.rules[0]?.colorScale?.cfvo[0]?.value).toBe(
			'#REF!',
		)
		expect(wb.table('Audit')?.columnDefs[1]?.formula).toBe('#REF!')
		expect(wb.table('Audit')?.columnDefs[1]?.totalsRowFormula).toBe('SUM(#REF!)')
	})

	test('structural delete exact journals restore escaped-sheet formulas and defined names', () => {
		const wb = AscendWorkbook.create()
		const inputSheet = "Input.Data's Δ"
		wb.apply([
			{ op: 'renameSheet', sheet: 'Sheet1', newName: inputSheet },
			{ op: 'addSheet', name: 'Middle' },
			{ op: 'addSheet', name: 'Report' },
			{
				op: 'setCells',
				sheet: inputSheet,
				updates: [
					{ ref: 'A1', value: 'Header' },
					{ ref: 'A2', value: 'deleted' },
					{ ref: 'A4', value: 9 },
				],
			},
			{ op: 'setFormula', sheet: 'Report', ref: 'A1', formula: `${quoteSheet(inputSheet)}!A4` },
			{
				op: 'setFormula',
				sheet: 'Report',
				ref: 'A2',
				formula: `SUM(${quoteSheetSpan(inputSheet, 'Report')}!A4)`,
			},
			{ op: 'setDefinedName', name: 'GlobalLater', ref: `${quoteSheet(inputSheet)}!$A$4` },
			{ op: 'setDefinedName', name: 'LocalLater', scope: inputSheet, ref: '$A$4' },
		])
		const before = journalComparableState(wb)

		const deleted = wb.apply([{ op: 'deleteRows', sheet: inputSheet, at: 1, count: 1 }], {
			journal: true,
		})

		expect(deleted.errors).toEqual([])
		expect(deleted.journal?.supported).toBe(true)
		expect(deleted.journal?.exact).toBe(true)
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('structural column delete exact journals restore escaped-sheet formulas and defined names', () => {
		const wb = AscendWorkbook.create()
		const inputSheet = "Input.Data's Δ"
		wb.apply([
			{ op: 'renameSheet', sheet: 'Sheet1', newName: inputSheet },
			{ op: 'addSheet', name: 'Middle' },
			{ op: 'addSheet', name: 'Report' },
			{
				op: 'setCells',
				sheet: inputSheet,
				updates: [
					{ ref: 'A1', value: 'Header' },
					{ ref: 'B1', value: 'deleted' },
					{ ref: 'D1', value: 9 },
				],
			},
			{ op: 'setFormula', sheet: 'Report', ref: 'A1', formula: `${quoteSheet(inputSheet)}!D1` },
			{
				op: 'setFormula',
				sheet: 'Report',
				ref: 'A2',
				formula: `SUM(${quoteSheetSpan(inputSheet, 'Report')}!D1)`,
			},
			{ op: 'setDefinedName', name: 'GlobalLater', ref: `${quoteSheet(inputSheet)}!$D$1` },
			{ op: 'setDefinedName', name: 'LocalLater', scope: inputSheet, ref: '$D$1' },
		])
		const before = journalComparableState(wb)

		const deleted = wb.apply([{ op: 'deleteCols', sheet: inputSheet, at: 1, count: 1 }], {
			journal: true,
		})

		expect(deleted.errors).toEqual([])
		expect(deleted.journal?.supported).toBe(true)
		expect(deleted.journal?.exact).toBe(true)
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('structural delete journals surface escaped names 3D refs and x14 losses precisely', () => {
		const wb = AscendWorkbook.create()
		const inputSheet = "Input.Data's Δ"
		wb.apply([
			{ op: 'renameSheet', sheet: 'Sheet1', newName: inputSheet },
			{ op: 'addSheet', name: 'Middle' },
			{ op: 'addSheet', name: 'Report' },
			{ op: 'setCells', sheet: inputSheet, updates: [{ ref: 'A2', value: 5 }] },
			{
				op: 'setFormula',
				sheet: 'Report',
				ref: 'A1',
				formula: `SUM(${quoteSheetSpan(inputSheet, 'Report')}!A2)`,
			},
			{ op: 'setDefinedName', name: 'GlobalDeleted', ref: `${quoteSheet(inputSheet)}!$A$2` },
			{ op: 'setDefinedName', name: 'LocalDeleted', scope: inputSheet, ref: '$A$2' },
		])
		const sheet = wb.getWorkbookModel().getSheet(inputSheet)
		if (!sheet) throw new Error('input sheet missing')
		sheet.x14DataValidations.push({
			index: 7,
			sqref: 'C2:C5',
			type: 'list',
			formula1: '$A$1:$A$4',
			preservedChildXml: ['<x14ac:metadata flag="1"/>'],
		})
		sheet.x14ConditionalFormats.push({
			index: 8,
			sqref: 'D2:D5',
			type: 'dataBar',
			priority: 1,
			formulas: [],
			dataBar: { cfvo: [{ type: 'formula', value: '$A$2' }] },
			preservedRuleChildXml: ['<x14:extLst><x14:ext uri="{cf-extension}"/></x14:extLst>'],
		})

		const deleted = wb.preview([{ op: 'deleteRows', sheet: inputSheet, at: 1, count: 1 }], {
			journal: true,
		})

		expect(deleted.wouldSucceed).toBe(true)
		expect(deleted.journal?.supported).toBe(true)
		expect(deleted.journal?.exact).toBe(false)
		expect(deleted.journal?.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message: `Deleted row formula references on ${inputSheet} cannot be restored with public operations`,
			refs: expect.arrayContaining([
				`${inputSheet}!x14Validation:C2:C5:formula1`,
				`${inputSheet}!x14ConditionalFormat:D2:D5:8:dataBar.cfvo:0`,
				'Report!A1',
				'name:GlobalDeleted',
				`name:${inputSheet}!LocalDeleted`,
			]),
		})
		expect(deleted.journal?.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message: `Deleted row x14 metadata on ${inputSheet} cannot be fully restored with public operations`,
			refs: [`${inputSheet}!x14Validation:C2:C5:7`, `${inputSheet}!x14ConditionalFormat:D2:D5:8`],
		})
		expect(wb.sheet(inputSheet)?.cell('A2')?.value).toEqual({ kind: 'number', value: 5 })
	})

	test('column delete journals surface escaped names 3D refs and x14 losses precisely', () => {
		const wb = AscendWorkbook.create()
		const inputSheet = "Input.Data's Δ"
		wb.apply([
			{ op: 'renameSheet', sheet: 'Sheet1', newName: inputSheet },
			{ op: 'addSheet', name: 'Middle' },
			{ op: 'addSheet', name: 'Report' },
			{ op: 'setCells', sheet: inputSheet, updates: [{ ref: 'B1', value: 5 }] },
			{
				op: 'setFormula',
				sheet: 'Report',
				ref: 'A1',
				formula: `SUM(${quoteSheetSpan(inputSheet, 'Report')}!B1)`,
			},
			{ op: 'setDefinedName', name: 'GlobalDeleted', ref: `${quoteSheet(inputSheet)}!$B$1` },
			{ op: 'setDefinedName', name: 'LocalDeleted', scope: inputSheet, ref: '$B$1' },
		])
		const sheet = wb.getWorkbookModel().getSheet(inputSheet)
		if (!sheet) throw new Error('input sheet missing')
		sheet.x14DataValidations.push({
			index: 7,
			sqref: 'B2:E2',
			type: 'list',
			formula1: '$A$1:$B$1',
			preservedChildXml: ['<x14ac:metadata flag="1"/>'],
		})
		sheet.x14ConditionalFormats.push({
			index: 8,
			sqref: 'B3:E3',
			type: 'dataBar',
			priority: 1,
			formulas: [],
			dataBar: { cfvo: [{ type: 'formula', value: '$B$1' }] },
			preservedRuleChildXml: ['<x14:extLst><x14:ext uri="{cf-extension}"/></x14:extLst>'],
		})

		const deleted = wb.preview([{ op: 'deleteCols', sheet: inputSheet, at: 1, count: 1 }], {
			journal: true,
		})

		expect(deleted.wouldSucceed).toBe(true)
		expect(deleted.journal?.supported).toBe(true)
		expect(deleted.journal?.exact).toBe(false)
		expect(deleted.journal?.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message: `Deleted column formula references on ${inputSheet} cannot be restored with public operations`,
			refs: expect.arrayContaining([
				`${inputSheet}!x14Validation:B2:E2:formula1`,
				`${inputSheet}!x14ConditionalFormat:B3:E3:8:dataBar.cfvo:0`,
				'Report!A1',
				'name:GlobalDeleted',
				`name:${inputSheet}!LocalDeleted`,
			]),
		})
		expect(deleted.journal?.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message: `Deleted column x14 metadata on ${inputSheet} cannot be fully restored with public operations`,
			refs: [`${inputSheet}!x14Validation:B2:E2:7`, `${inputSheet}!x14ConditionalFormat:B3:E3:8`],
		})
		expect(wb.sheet(inputSheet)?.cell('B1')?.value).toEqual({ kind: 'number', value: 5 })
	})
})

function materializeViewportPatch(
	cells: readonly InteractiveViewportCell[],
	patch: InteractiveViewportPatch,
): Map<string, InteractiveViewportCell> {
	const next = interactiveCellMap(cells)
	for (const ref of patch.removedRefs) next.delete(ref)
	for (const cell of patch.changedCells) next.set(cell.ref, cell)
	return next
}

function interactiveCellMap(
	cells: readonly InteractiveViewportCell[],
): Map<string, InteractiveViewportCell> {
	return new Map(cells.map((cell) => [cell.ref, cell]))
}

function semanticInteractiveCellMap(
	cells: readonly InteractiveViewportCell[],
): Map<string, Omit<InteractiveViewportCell, 'styleId'>> {
	return new Map(
		cells.map((cell) => {
			const { styleId: _styleId, ...semanticCell } = cell
			return [cell.ref, semanticCell]
		}),
	)
}

function cellRefOverlapsRange(ref: string, range: RangeRef): boolean {
	const cell = parseA1(ref)
	return cellInRange(cell.row, cell.col, range)
}

function cellInRange(row: number, col: number, range: RangeRef): boolean {
	const startRow = Math.min(range.start.row, range.end.row)
	const endRow = Math.max(range.start.row, range.end.row)
	const startCol = Math.min(range.start.col, range.end.col)
	const endCol = Math.max(range.start.col, range.end.col)
	return row >= startRow && row <= endRow && col >= startCol && col <= endCol
}

function cellRange(row: number, col: number): RangeRef {
	return { start: { row, col }, end: { row, col } }
}

function rangeKey(range: RangeRef): string {
	return `${range.start.row}:${range.start.col}:${range.end.row}:${range.end.col}`
}

function quoteSheet(sheet: string): string {
	return `'${sheet.replace(/'/g, "''")}'`
}

function quoteSheetSpan(startSheet: string, endSheet: string): string {
	return `'${startSheet.replace(/'/g, "''")}:${endSheet.replace(/'/g, "''")}'`
}

function journalComparableState(wb: AscendWorkbook): object {
	const snapshot = wb.snapshot()
	return {
		sheets: snapshot.sheets,
		names: snapshot.names,
		workbookProtectionJson: snapshot.workbookProtectionJson,
		workbookProperties: wb.getWorkbookModel().workbookProperties,
		documentProperties: wb.inspect().documentProperties,
		workbookViews: wb.inspect().workbookViews,
		calcSettings: wb.getWorkbookModel().calcSettings,
		themeSummary: wb.inspect().themeSummary,
		sheetMetadata: wb.sheets.map((name) => {
			const sheet = wb.sheet(name)
			const modelSheet = wb.getWorkbookModel().getSheet(name)
			return {
				name,
				frozenRows: sheet?.frozenRows,
				frozenCols: sheet?.frozenCols,
				merges: sheet?.merges,
				dataValidations: sheet?.dataValidations,
				conditionalFormats: sheet?.conditionalFormats,
				x14DataValidations: modelSheet?.x14DataValidations,
				x14ConditionalFormats: modelSheet?.x14ConditionalFormats,
				comments: sheet?.getComments(),
				hyperlinks: sheet?.getHyperlinks(),
				autoFilter: sheet?.autoFilter,
				tables: modelSheet?.tables,
				tabColor: modelSheet?.tabColor,
				protection: modelSheet?.protection,
				state: modelSheet?.state,
				rowHeights: modelSheet ? [...modelSheet.rowHeights.entries()] : undefined,
				colWidths: modelSheet ? [...modelSheet.colWidths.entries()] : undefined,
				rowDefs: modelSheet ? [...modelSheet.rowDefs.entries()] : undefined,
				colDefs: modelSheet?.colDefs,
			}
		}),
		chartParts: wb.getWorkbookModel().chartParts,
		chartSheets: wb.getWorkbookModel().chartSheets,
		pivotCaches: wb.getWorkbookModel().pivotCaches,
		pivotTables: wb.getWorkbookModel().pivotTables,
		styles: {
			b2: wb.cellStyle(`${wb.sheets[0]}!B2`),
			c2: wb.cellStyle(`${wb.sheets[0]}!C2`),
		},
	}
}
