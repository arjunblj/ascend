import { describe, expect, test } from 'bun:test'
import {
	cycleFormulaReferenceMode,
	formulaTokenRanges,
	insertFormulaReference,
	referenceAtCursor,
} from './index.ts'

describe('formula editing utilities', () => {
	test('finds cell, range, sheet-qualified, structured, and spill references at the cursor', () => {
		expect(referenceAtCursor('SUM(A1:B2)', 6)).toEqual({
			text: 'A1:B2',
			start: 4,
			end: 9,
			kind: 'range',
		})
		expect(referenceAtCursor("'Q1 Plan'!$A$1+1", 12)).toEqual({
			text: "'Q1 Plan'!$A$1",
			start: 0,
			end: 14,
			kind: 'sheet-cell',
		})
		expect(referenceAtCursor('Sales[Amount]+A1', 8)).toEqual({
			text: 'Sales[Amount]',
			start: 0,
			end: 13,
			kind: 'structured',
		})
		expect(referenceAtCursor('A1#+B1', 2)).toEqual({
			text: 'A1#',
			start: 0,
			end: 3,
			kind: 'spill',
		})
		expect(referenceAtCursor('=SUM(A1)', 6)).toEqual({
			text: 'A1',
			start: 5,
			end: 7,
			kind: 'cell',
		})
		expect(referenceAtCursor('Table1[[#Totals],[Amount]]', 14)).toEqual({
			text: 'Table1[[#Totals],[Amount]]',
			start: 0,
			end: 26,
			kind: 'structured',
		})
	})

	test('cycles the cell reference under the cursor with Excel F4 semantics', () => {
		let result = cycleFormulaReferenceMode('A1+B2', 1)
		expect(result).toMatchObject({ formula: '$A$1+B2', cursor: 4, changed: true })
		result = cycleFormulaReferenceMode(result.formula, result.cursor)
		expect(result).toMatchObject({ formula: 'A$1+B2', cursor: 3, changed: true })
		result = cycleFormulaReferenceMode(result.formula, result.cursor)
		expect(result).toMatchObject({ formula: '$A1+B2', cursor: 3, changed: true })
		result = cycleFormulaReferenceMode(result.formula, result.cursor)
		expect(result).toMatchObject({ formula: 'A1+B2', cursor: 2, changed: true })
	})

	test('cycles sheet-qualified and range endpoint refs without rewriting sheet names', () => {
		expect(cycleFormulaReferenceMode("'My Sheet'!A1", 12)).toMatchObject({
			formula: "'My Sheet'!$A$1",
			changed: true,
		})
		expect(cycleFormulaReferenceMode('SUM(A1:B2)', 8)).toMatchObject({
			formula: 'SUM(A1:$B$2)',
			changed: true,
		})
	})

	test('inserts pointed references and can replace the reference under the cursor', () => {
		expect(insertFormulaReference('=SUM()', 5, "'Q1 Plan'!$A$1")).toEqual({
			formula: "=SUM('Q1 Plan'!$A$1)",
			cursor: 19,
			inserted: "'Q1 Plan'!$A$1",
		})

		expect(
			insertFormulaReference('=SUM(A1)+Table1[Amount]', 7, '$C$3', {
				replaceReferenceAtCursor: true,
			}),
		).toEqual({
			formula: '=SUM($C$3)+Table1[Amount]',
			cursor: 9,
			inserted: '$C$3',
			replaced: {
				text: 'A1',
				start: 5,
				end: 7,
				kind: 'cell',
			},
		})

		expect(
			insertFormulaReference('=SUM(A1)+Table1[Amount]', 15, 'Sales[Total]', {
				replaceReferenceAtCursor: true,
			}),
		).toEqual({
			formula: '=SUM(A1)+Sales[Total]',
			cursor: 21,
			inserted: 'Sales[Total]',
			replaced: {
				text: 'Table1[Amount]',
				start: 9,
				end: 23,
				kind: 'structured',
			},
		})
	})

	test('exposes token ranges with syntax classes for highlighting', () => {
		expect(formulaTokenRanges('SUM(A1, "x")').map((token) => token.className)).toEqual([
			'function',
			'punctuation',
			'reference',
			'punctuation',
			'whitespace',
			'literal',
			'punctuation',
		])
		expect(formulaTokenRanges('SUM(A1:').map((token) => token.text)).toEqual([
			'SUM',
			'(',
			'A1',
			':',
		])
	})
})
