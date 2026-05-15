import { describe, expect, test } from 'bun:test'
import {
	cycleFormulaReferenceMode,
	formulaAssist,
	formulaCodeActions,
	formulaDiagnostics,
	formulaFunctionCompletions,
	formulaFunctionSignature,
	formulaFunctionSignatureHelp,
	formulaHover,
	formulaReferenceRanges,
	formulaTokenRanges,
	insertFormulaReference,
	referenceAtCursor,
} from './index.ts'

describe('formula editing utilities', () => {
	test('bundles formula IDE assistance for API and MCP surfaces', () => {
		const result = formulaAssist('=SUM(A1:B2', {
			cursor: 8,
			prefix: 'SU',
			completionLimit: 3,
			functionName: 'SUM',
			reference: 'C1',
			replaceReferenceAtCursor: true,
			cycleReference: true,
		})

		expect(result.diagnostics.parseOk).toBe(false)
		expect(result.references).toEqual([{ text: 'A1:B2', start: 5, end: 10, kind: 'range' }])
		expect(result.activeReference).toMatchObject({ text: 'A1:B2', kind: 'range' })
		expect(result.hover).toMatchObject({ kind: 'reference', label: 'A1:B2' })
		expect(result.completions.some((completion) => completion.name === 'SUM')).toBe(true)
		expect(result.signature?.name).toBe('SUM')
		expect(result.signatureHelp?.signature.name).toBe('SUM')
		expect(result.codeActions).toContainEqual(
			expect.objectContaining({
				title: 'Cycle reference absolute/relative mode',
				kind: 'refactor.rewrite',
			}),
		)
		expect(result.codeActions).toContainEqual(
			expect.objectContaining({
				title: 'Replace reference with C1',
				kind: 'quickfix',
			}),
		)
		expect(result.cycle).toMatchObject({ formula: '=SUM(A1:$B$2', changed: true })
		expect(result.insertion).toMatchObject({ formula: '=SUM(C1', replaced: { text: 'A1:B2' } })
	})

	test('returns formula hover and code actions for LSP-style consumers', () => {
		expect(formulaHover('=SUM(A1:B2)', 2)).toMatchObject({
			kind: 'function',
			label: 'SUM',
			signature: { name: 'SUM' },
			start: 1,
			end: 4,
		})
		expect(formulaHover('=SUM(A1:B2)', 7)).toMatchObject({
			kind: 'reference',
			label: 'A1:B2',
			reference: { kind: 'range' },
		})
		expect(formulaHover('=SUM(Table1[[#Totals],[Amount])', 11)).toMatchObject({
			kind: 'diagnostic',
			label: 'formula-structured-reference-error',
		})

		const actions = formulaCodeActions('=SUM(A1:B2)', 9, {
			reference: 'C1:D2',
			replaceReferenceAtCursor: true,
			cycleReference: true,
		})
		expect(actions).toContainEqual(
			expect.objectContaining({
				title: 'Cycle reference absolute/relative mode',
				edit: { formula: '=SUM(A1:$B$2)', cursor: 12 },
			}),
		)
		expect(actions).toContainEqual(
			expect.objectContaining({
				title: 'Replace reference with C1:D2',
				edit: { formula: '=SUM(C1:D2)', cursor: 10 },
			}),
		)
	})

	test('finds cell, range, sheet-qualified, structured, and spill references at the cursor', () => {
		expect(formulaReferenceRanges("=SUM(A1,'Q2'!B2,Sales[Amount],A1#)")).toEqual([
			{ text: 'A1', start: 5, end: 7, kind: 'cell' },
			{ text: "'Q2'!B2", start: 8, end: 15, kind: 'sheet-cell' },
			{ text: 'Sales[Amount]', start: 16, end: 29, kind: 'structured' },
			{ text: 'A1#', start: 30, end: 33, kind: 'spill' },
		])
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
		expect(referenceAtCursor('=[Book.xlsx]Sheet1!A1', 5)).toEqual({
			text: '[Book.xlsx]Sheet1!A1',
			start: 1,
			end: 21,
			kind: 'sheet-cell',
		})
		expect(referenceAtCursor('=[Book.xlsx]Sheet1!A1:B2', 17)).toEqual({
			text: '[Book.xlsx]Sheet1!A1:B2',
			start: 1,
			end: 24,
			kind: 'sheet-range',
		})
		expect(referenceAtCursor('Table1[[#Totals],[Amount]]', 14)).toEqual({
			text: 'Table1[[#Totals],[Amount]]',
			start: 0,
			end: 26,
			kind: 'structured',
		})
		expect(referenceAtCursor('=SUM(Sheet1:Sheet3!A1)', 12)).toEqual({
			text: 'Sheet1:Sheet3!A1',
			start: 5,
			end: 21,
			kind: 'sheet-3d-cell',
		})
		expect(referenceAtCursor("=SUM('Q1 Plan':'Q3 Plan'!A1:B2)", 18)).toEqual({
			text: "'Q1 Plan':'Q3 Plan'!A1:B2",
			start: 5,
			end: 30,
			kind: 'sheet-3d-range',
		})
		expect(referenceAtCursor('=SUM([Book.xlsx]Sheet1:Sheet3!A1:B2)', 18)).toEqual({
			text: '[Book.xlsx]Sheet1:Sheet3!A1:B2',
			start: 5,
			end: 35,
			kind: 'sheet-3d-range',
		})
		expect(referenceAtCursor("=SUM('[Book.xlsx]Q1:Q3'!A1)", 18)).toEqual({
			text: "'[Book.xlsx]Q1:Q3'!A1",
			start: 5,
			end: 26,
			kind: 'sheet-3d-cell',
		})
		expect(referenceAtCursor("=SUM('C:/tmp/[Book.xlsx]Sheet1:Sheet3'!A1:B2)", 24)).toEqual({
			text: "'C:/tmp/[Book.xlsx]Sheet1:Sheet3'!A1:B2",
			start: 5,
			end: 44,
			kind: 'sheet-3d-range',
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

	test('cycles sheet-qualified refs when the cursor is in the qualifier', () => {
		expect(cycleFormulaReferenceMode("'My Sheet'!A1", 3)).toMatchObject({
			formula: "'My Sheet'!$A$1",
			cursor: 15,
			changed: true,
		})
		expect(cycleFormulaReferenceMode("='[Book.xlsx]Q1 Plan'!A1", 8)).toMatchObject({
			formula: "='[Book.xlsx]Q1 Plan'!$A$1",
			cursor: 26,
			changed: true,
		})
		expect(cycleFormulaReferenceMode('=[Book.xlsx]Sheet1!A1', 5)).toMatchObject({
			formula: '=[Book.xlsx]Sheet1!$A$1',
			cursor: 23,
			changed: true,
		})
		expect(cycleFormulaReferenceMode("'Q1 Plan'!A1:B2", 4)).toMatchObject({
			formula: "'Q1 Plan'!$A$1:B2",
			cursor: 14,
			changed: true,
		})
		expect(cycleFormulaReferenceMode('Sheet1:Sheet3!A1', 4)).toMatchObject({
			formula: 'Sheet1:Sheet3!$A$1',
			cursor: 18,
			changed: true,
			reference: {
				text: 'Sheet1:Sheet3!$A$1',
				start: 0,
				end: 18,
				kind: 'sheet-3d-cell',
			},
		})
		expect(cycleFormulaReferenceMode("'Q1 Plan':'Q3 Plan'!A1:B2", 10)).toMatchObject({
			formula: "'Q1 Plan':'Q3 Plan'!$A$1:B2",
			cursor: 24,
			changed: true,
			reference: {
				text: "'Q1 Plan':'Q3 Plan'!$A$1:B2",
				start: 0,
				end: 27,
				kind: 'sheet-3d-range',
			},
		})
		expect(cycleFormulaReferenceMode('=[Book.xlsx]Sheet1:Sheet3!A1', 9)).toMatchObject({
			formula: '=[Book.xlsx]Sheet1:Sheet3!$A$1',
			cursor: 30,
			changed: true,
			reference: {
				text: '[Book.xlsx]Sheet1:Sheet3!$A$1',
				start: 1,
				end: 30,
				kind: 'sheet-3d-cell',
			},
		})
		expect(cycleFormulaReferenceMode("='[Book.xlsx]Q1:Q3'!A1", 8)).toMatchObject({
			formula: "='[Book.xlsx]Q1:Q3'!$A$1",
			cursor: 24,
			changed: true,
			reference: {
				text: "'[Book.xlsx]Q1:Q3'!$A$1",
				start: 1,
				end: 24,
				kind: 'sheet-3d-cell',
			},
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

		expect(
			insertFormulaReference('=[Book.xlsx]Sheet1!A1+1', 6, 'B2', {
				replaceReferenceAtCursor: true,
			}),
		).toEqual({
			formula: '=B2+1',
			cursor: 3,
			inserted: 'B2',
			replaced: {
				text: '[Book.xlsx]Sheet1!A1',
				start: 1,
				end: 21,
				kind: 'sheet-cell',
			},
		})
		expect(
			insertFormulaReference('=SUM(Sheet1:Sheet3!A1)', 14, 'Sheet2!C3', {
				replaceReferenceAtCursor: true,
			}),
		).toEqual({
			formula: '=SUM(Sheet2!C3)',
			cursor: 14,
			inserted: 'Sheet2!C3',
			replaced: {
				text: 'Sheet1:Sheet3!A1',
				start: 5,
				end: 21,
				kind: 'sheet-3d-cell',
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

	test('reports malformed quoted sheet and workbook reference qualifiers', () => {
		expect(formulaDiagnostics("='Q1''s Plan'!A1")).toEqual({
			parseOk: true,
			diagnostics: [],
		})
		expect(formulaDiagnostics('=[Book.xlsx')).toEqual({
			parseOk: false,
			diagnostics: [
				{
					code: 'formula-reference-qualifier-error',
					severity: 'error',
					message: 'Unterminated external workbook or bracketed reference',
					start: 1,
					end: 11,
				},
			],
		})
		expect(formulaDiagnostics("='Q1 Plan!A1")).toEqual({
			parseOk: false,
			diagnostics: [
				{
					code: 'formula-reference-qualifier-error',
					severity: 'error',
					message: 'Unterminated quoted sheet or workbook reference',
					start: 1,
					end: 12,
				},
			],
		})
		expect(formulaDiagnostics("='[Book.xlsx]Q1 Plan!A1")).toEqual({
			parseOk: false,
			diagnostics: [
				{
					code: 'formula-reference-qualifier-error',
					severity: 'error',
					message: 'Unterminated quoted sheet or workbook reference',
					start: 1,
					end: 23,
				},
			],
		})
		expect(formulaDiagnostics("=SUM('Q1 Plan!A1)")).toEqual({
			parseOk: false,
			diagnostics: [
				{
					code: 'formula-reference-qualifier-error',
					severity: 'error',
					message: 'Unterminated quoted sheet or workbook reference',
					start: 5,
					end: 17,
				},
				{
					code: 'formula-parse-error',
					severity: 'error',
					message: 'Expected CloseParen, got EOF "" at position 16',
					start: 17,
					end: 17,
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

	test('resolves cursor-aware function signature help without client-side reparsing', () => {
		expect(formulaFunctionSignatureHelp('=SUM(A1, B2)', 5)).toMatchObject({
			signature: { name: 'SUM' },
			activeParameter: 0,
			callStart: 1,
			callEnd: 12,
			argumentListStart: 4,
			argumentListEnd: 12,
		})
		expect(formulaFunctionSignatureHelp('=SUM(A1, B2)', 9)).toMatchObject({
			signature: { name: 'SUM' },
			activeParameter: 1,
		})
		expect(formulaFunctionSignatureHelp('=SUM(A1)+B1', 10)).toBeNull()
		expect(formulaFunctionSignatureHelp('=NOT_REGISTERED(A1)', 16)).toBeNull()
	})

	test('keeps signature help scoped through nested, incomplete, and string arguments', () => {
		expect(formulaFunctionSignatureHelp('=IF(A1, SUM(B1, C1), 0)', 17)).toMatchObject({
			signature: { name: 'SUM' },
			activeParameter: 1,
		})
		expect(formulaFunctionSignatureHelp('=IF(A1, SUM(B1, C1), 0)', 21)).toMatchObject({
			signature: { name: 'IF' },
			activeParameter: 2,
		})
		expect(formulaFunctionSignatureHelp('=SUM(A1, ', 9)).toMatchObject({
			signature: { name: 'SUM' },
			activeParameter: 1,
			callStart: 1,
			callEnd: 9,
			argumentListStart: 4,
			argumentListEnd: 9,
		})
		expect(formulaFunctionSignatureHelp('=SUM("a,b", A1)', 13)).toMatchObject({
			signature: { name: 'SUM' },
			activeParameter: 1,
		})
		expect(formulaFunctionSignatureHelp('=SUM((A1,B1), C1)', 10)).toMatchObject({
			signature: { name: 'SUM' },
			activeParameter: 0,
		})
		expect(formulaFunctionSignatureHelp('=SUM((A1,B1), C1)', 15)).toMatchObject({
			signature: { name: 'SUM' },
			activeParameter: 1,
		})
		expect(formulaFunctionSignatureHelp('=SUM({1,2;3,4}, A1)', 18)).toMatchObject({
			signature: { name: 'SUM' },
			activeParameter: 1,
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
