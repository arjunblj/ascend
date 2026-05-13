import { describe, expect, test } from 'bun:test'
import {
	cycleFormulaReferenceMode,
	formulaDiagnostics,
	formulaFunctionCompletions,
	formulaFunctionSignature,
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
		expect(referenceAtCursor("='[Book.xlsx]Q1 Plan'!$A$1", 24)).toEqual({
			text: "'[Book.xlsx]Q1 Plan'!$A$1",
			start: 1,
			end: 26,
			kind: 'sheet-cell',
		})
		expect(referenceAtCursor('Table1[[#Totals],[Amount]]', 14)).toEqual({
			text: 'Table1[[#Totals],[Amount]]',
			start: 0,
			end: 26,
			kind: 'structured',
		})
	})

	test('does not treat earlier references as active in empty formula edit slots', () => {
		expect(referenceAtCursor('=A1 + ', 6)).toBeNull()
		expect(referenceAtCursor('=SUM(A1, )', 9)).toBeNull()
		expect(referenceAtCursor('=A1+ B2', 4)).toBeNull()
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

	test('does not cycle stale references from empty formula edit slots', () => {
		expect(cycleFormulaReferenceMode('=A1 + ', 6)).toEqual({
			formula: '=A1 + ',
			cursor: 6,
			changed: false,
		})
		expect(cycleFormulaReferenceMode('=SUM(A1, )', 9)).toEqual({
			formula: '=SUM(A1, )',
			cursor: 9,
			changed: false,
		})
		expect(cycleFormulaReferenceMode('=A1+ B2', 4)).toEqual({
			formula: '=A1+ B2',
			cursor: 4,
			changed: false,
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

	test('inserts pointed references in empty formula slots without replacing earlier refs', () => {
		expect(
			insertFormulaReference('=A1 + ', 6, 'B2', {
				replaceReferenceAtCursor: true,
			}),
		).toEqual({
			formula: '=A1 + B2',
			cursor: 8,
			inserted: 'B2',
		})

		expect(
			insertFormulaReference('=SUM(A1, )', 9, 'B2', {
				replaceReferenceAtCursor: true,
			}),
		).toEqual({
			formula: '=SUM(A1, B2)',
			cursor: 11,
			inserted: 'B2',
		})
	})

	test('returns formula parse diagnostics with stable source spans', () => {
		expect(formulaDiagnostics('=SUM(A1, 2)')).toEqual({ parseOk: true, diagnostics: [] })

		expect(formulaDiagnostics('=SUM(A1:')).toEqual({
			parseOk: false,
			diagnostics: [
				{
					code: 'formula-parse-error',
					severity: 'error',
					message: 'Unexpected token EOF "" at position 7',
					start: 8,
					end: 8,
				},
			],
		})

		expect(formulaDiagnostics('=)')).toEqual({
			parseOk: false,
			diagnostics: [
				{
					code: 'formula-parse-error',
					severity: 'error',
					message: 'Unexpected token CloseParen ")" at position 0',
					start: 1,
					end: 2,
				},
			],
		})
	})

	test('reports malformed structured references even when the parser accepts names', () => {
		expect(formulaDiagnostics('=Table1[Amount]')).toEqual({
			parseOk: true,
			diagnostics: [],
		})
		expect(formulaDiagnostics('=Table1[[#Totals],[Amount]]')).toEqual({
			parseOk: true,
			diagnostics: [],
		})
		expect(formulaDiagnostics("=Table1[Q1']Total]")).toEqual({
			parseOk: true,
			diagnostics: [],
		})

		expect(formulaDiagnostics('=Table1[')).toEqual({
			parseOk: false,
			diagnostics: [
				{
					code: 'formula-structured-reference-error',
					severity: 'error',
					message: 'Unterminated structured reference',
					start: 7,
					end: 8,
				},
			],
		})
		expect(formulaDiagnostics('=SUM(Table1[')).toEqual({
			parseOk: false,
			diagnostics: [
				{
					code: 'formula-structured-reference-error',
					severity: 'error',
					message: 'Unterminated structured reference',
					start: 11,
					end: 12,
				},
				{
					code: 'formula-parse-error',
					severity: 'error',
					message: 'Expected CloseParen, got EOF "" at position 11',
					start: 12,
					end: 12,
				},
			],
		})
		expect(formulaDiagnostics('=Table1[Amount]]')).toEqual({
			parseOk: false,
			diagnostics: [
				{
					code: 'formula-structured-reference-error',
					severity: 'error',
					message: 'Malformed structured reference',
					start: 7,
					end: 16,
				},
			],
		})
	})

	test('exposes function completions and generic signature help from the registry', () => {
		expect(formulaFunctionCompletions('xlo', { limit: 5 }).map((entry) => entry.name)).toEqual([
			'XLOOKUP',
		])
		expect(formulaFunctionSignature('if')).toEqual({
			name: 'IF',
			minArgs: 2,
			maxArgs: 3,
			volatile: false,
			label: 'IF(arg1, arg2, [arg3])',
			parameters: [
				{ label: 'arg1', index: 0, required: true },
				{ label: 'arg2', index: 1, required: true },
				{ label: 'arg3', index: 2, required: false },
			],
			variadic: false,
		})
		expect(formulaFunctionSignature('RAND')).toMatchObject({
			label: 'RAND()',
			volatile: true,
			parameters: [],
			variadic: false,
		})
		expect(formulaFunctionSignature('SUM')).toMatchObject({
			label: 'SUM(arg1, [arg2], ...)',
			variadic: true,
		})
		expect(formulaFunctionCompletions('sum', { limit: 2 })).toEqual([
			expect.objectContaining({ name: 'SUM' }),
			expect.objectContaining({ name: 'SUMIF' }),
		])
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
