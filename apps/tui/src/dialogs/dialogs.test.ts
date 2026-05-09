import { describe, expect, test } from 'bun:test'
import { createSelection } from '../model/selection.ts'
import { buildDialogOperations, findDialog } from './index.ts'

const ctx = {
	sheet: 'Sheet1',
	selection: createSelection(),
}

describe('dialog operation contracts', () => {
	test('format-cells builds a style operation', () => {
		const ops = buildDialogOperations('format-cells', ctx, {
			numberFormat: '$#,##0.00',
			bold: true,
		})
		expect(ops).toEqual([
			{
				op: 'setStyle',
				sheet: 'Sheet1',
				range: 'A1',
				style: { numberFormat: '$#,##0.00', font: { bold: true } },
			},
		])
	})

	test('format-cells returns no operations for empty style input', () => {
		expect(buildDialogOperations('format-cells', ctx, {})).toEqual([])
	})

	test('sort builds a sortRange operation', () => {
		const ops = buildDialogOperations('sort', ctx, {
			range: 'A1:B5',
			column: 'A',
			descending: true,
		})
		expect(ops).toEqual([
			{ op: 'sortRange', sheet: 'Sheet1', range: 'A1:B5', by: [{ column: 'A', descending: true }] },
		])
	})

	test('sort rejects missing sort columns before engine dispatch', () => {
		expect(() => buildDialogOperations('sort', ctx, { range: 'A1:B5' })).toThrow(
			'Sort column is required',
		)
	})

	test('create-table builds a createTable operation', () => {
		const ops = buildDialogOperations('create-table', ctx, {
			ref: 'A1:B5',
			name: 'Revenue',
			hasHeaders: true,
		})
		expect(ops).toEqual([
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B5', name: 'Revenue', hasHeaders: true },
		])
	})

	test('comment builds a setComment operation', () => {
		const ops = buildDialogOperations('comment', ctx, {
			ref: 'B2',
			text: 'Review this',
			author: 'Ada',
		})
		expect(ops).toEqual([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'B2', text: 'Review this', author: 'Ada' },
		])
	})

	test('metadata dialogs reject invalid ranges and missing rule types', () => {
		expect(() =>
			buildDialogOperations('data-validation', ctx, {
				range: 'not-a-range',
				rule: { type: 'list', formula1: '"A,B"' },
			}),
		).toThrow('Validation range must be a valid A1 range')
		expect(() =>
			buildDialogOperations('data-validation', ctx, { rule: { formula1: '1' } }),
		).toThrow('Validation type must be one of')
		expect(() =>
			buildDialogOperations('conditional-formatting', ctx, { rule: { formula: 'TRUE' } }),
		).toThrow('Conditional format type must be one of')
		expect(() =>
			buildDialogOperations('comment', ctx, { ref: 'A1:B2', text: 'Review this' }),
		).toThrow('Comment cell must be a single A1 cell reference')
	})

	test('chart-wizard builds a setChartSeriesSource operation', () => {
		const ops = buildDialogOperations('chart-wizard', ctx, {
			seriesIndex: 1,
			sheet: 'Dashboard',
			chartIndex: 0,
			categoryRef: 'Sheet1!A2:A10',
			valueRef: 'Sheet1!B2:B10',
		})
		expect(ops).toEqual([
			{
				op: 'setChartSeriesSource',
				seriesIndex: 1,
				sheet: 'Dashboard',
				chartIndex: 0,
				categoryRef: 'Sheet1!A2:A10',
				valueRef: 'Sheet1!B2:B10',
			},
		])
	})

	test('pivot-fields builds a setPivotCache operation', () => {
		const ops = buildDialogOperations('pivot-fields', ctx, {
			pivotTable: 'PivotTable1',
			sourceSheet: 'Data',
			sourceRef: 'A1:D100',
			refreshOnLoad: true,
			invalid: true,
		})
		expect(ops).toEqual([
			{
				op: 'setPivotCache',
				pivotTable: 'PivotTable1',
				sourceSheet: 'Data',
				sourceRef: 'A1:D100',
				refreshOnLoad: true,
				invalid: true,
			},
		])
	})

	test('print-preview builds print area and page setup operations', () => {
		const ops = buildDialogOperations('print-preview', ctx, {
			range: 'A1:D20',
			orientation: 'landscape',
			fitToWidth: 1,
			fitToHeight: 0,
		})
		expect(ops).toEqual([
			{ op: 'setPrintArea', sheet: 'Sheet1', range: 'A1:D20' },
			{
				op: 'setPageSetup',
				sheet: 'Sheet1',
				setup: { orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 },
			},
		])
	})

	test('find-replace is a foundation dialog backed by the engine scanner', () => {
		expect(findDialog('find-replace')?.phase).toBe('foundation')
		expect(
			buildDialogOperations('find-replace', ctx, {
				range: 'A1:B10',
				findText: 'old',
				replaceText: 'new',
				action: 'replaceAll',
				lookIn: 'values',
			}),
		).toEqual([])
		expect(() =>
			buildDialogOperations('find-replace', ctx, {
				findText: '',
				action: 'replace',
				lookIn: 'values',
			}),
		).toThrow('Find text is required')
	})
})
