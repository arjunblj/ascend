import type { Cell, Sheet, Workbook } from '@ascend/core'
import type { FormulaCellRef, FormulaNode } from '@ascend/formulas'
import { parseFormula, printFormula } from '@ascend/formulas'
import { shiftIndex } from './ref-shift.ts'

export function rewriteWorkbookFormulasForShift(
	workbook: Workbook,
	targetSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	for (const sheet of workbook.sheets) {
		const updates: [number, number, Cell][] = []
		for (const [row, col, existing] of sheet.cells.iterate()) {
			if (existing.formula === null) continue
			const parsed = parseFormula(existing.formula)
			if (!parsed.ok) continue
			const rewritten = rewriteNodeForShift(parsed.value, targetSheet, sheet.name, axis, at, delta)
			const nextFormula = printFormula(rewritten)
			if (nextFormula === existing.formula) continue
			updates.push([
				row,
				col,
				{
					value: existing.value,
					formula: nextFormula,
					styleId: existing.styleId,
				},
			])
		}
		for (const [row, col, updated] of updates) {
			sheet.cells.set(row, col, updated)
		}
	}
}

export function rewriteDefinedNameFormulasForShift(
	workbook: Workbook,
	targetSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	const entries = [...workbook.definedNames.list()]
	for (const entry of entries) {
		const scope = entry.scope.kind === 'sheet' ? entry.scope : undefined
		const scopeSheet = scope
			? workbook.sheets.find((sheet) => sheet.id === scope.sheetId)?.name
			: undefined
		const parsed = parseFormula(entry.formula)
		if (!parsed.ok) continue
		const rewritten = rewriteNodeForShift(
			parsed.value,
			targetSheet,
			scopeSheet ?? targetSheet,
			axis,
			at,
			delta,
		)
		const formula = printFormula(rewritten)
		if (formula === entry.formula) continue
		workbook.definedNames.set(entry.name, formula, entry.scope)
	}
}

export function rewriteFormulaTextForShift(
	formula: string | undefined,
	targetSheet: string,
	formulaSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): string | undefined {
	if (!formula) return formula
	const parsed = parseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(rewriteNodeForShift(parsed.value, targetSheet, formulaSheet, axis, at, delta))
}

export function rewriteSheetNameInFormulas(
	workbook: Workbook,
	oldName: string,
	newName: string,
): void {
	for (const sheet of workbook.sheets) {
		const updates: [number, number, Cell][] = []
		for (const [row, col, existing] of sheet.cells.iterate()) {
			if (existing.formula === null) continue
			const parsed = parseFormula(existing.formula)
			if (!parsed.ok) continue
			const rewritten = rewriteSheetName(parsed.value, oldName, newName)
			const nextFormula = printFormula(rewritten)
			if (nextFormula === existing.formula) continue
			updates.push([
				row,
				col,
				{
					value: existing.value,
					formula: nextFormula,
					styleId: existing.styleId,
				},
			])
		}
		for (const [row, col, updated] of updates) {
			sheet.cells.set(row, col, updated)
		}
	}
}

export function rewriteSheetNameInDefinedNames(
	workbook: Workbook,
	oldName: string,
	newName: string,
): void {
	const entries = [...workbook.definedNames.list()]
	for (const entry of entries) {
		const parsed = parseFormula(entry.formula)
		if (!parsed.ok) continue
		const rewritten = rewriteSheetName(parsed.value, oldName, newName)
		const formula = printFormula(rewritten)
		if (formula === entry.formula) continue
		workbook.definedNames.set(entry.name, formula, entry.scope)
	}
}

export function rewriteFormulaTextForRename(
	formula: string | undefined,
	oldName: string,
	newName: string,
): string | undefined {
	if (!formula) return formula
	const parsed = parseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(rewriteSheetName(parsed.value, oldName, newName))
}

export function rewriteSheetMetadataFormulasForShift(
	sheet: Sheet,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	for (let i = 0; i < sheet.dataValidations.length; i++) {
		const validation = sheet.dataValidations[i]
		if (!validation) continue
		const formula1 = rewriteFormulaTextForShift(
			validation.formula1,
			sheet.name,
			sheet.name,
			axis,
			at,
			delta,
		)
		const formula2 = rewriteFormulaTextForShift(
			validation.formula2,
			sheet.name,
			sheet.name,
			axis,
			at,
			delta,
		)
		sheet.dataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.conditionalFormats.length; i++) {
		const format = sheet.conditionalFormats[i]
		if (!format) continue
		sheet.conditionalFormats[i] = {
			...format,
			rules: format.rules.map((rule) => ({
				...rule,
				formulas: rule.formulas.map(
					(formula) =>
						rewriteFormulaTextForShift(formula, sheet.name, sheet.name, axis, at, delta) ?? formula,
				),
			})),
		}
	}
	for (let i = 0; i < sheet.tables.length; i++) {
		const table = sheet.tables[i]
		if (!table) continue
		const columns = table.columns.map((column) => {
			const formula = rewriteFormulaTextForShift(
				column.formula,
				sheet.name,
				sheet.name,
				axis,
				at,
				delta,
			)
			const totalsRowFormula = rewriteFormulaTextForShift(
				column.totalsRowFormula,
				sheet.name,
				sheet.name,
				axis,
				at,
				delta,
			)
			return {
				...column,
				...(formula !== undefined ? { formula } : {}),
				...(totalsRowFormula !== undefined ? { totalsRowFormula } : {}),
			}
		})
		sheet.tables[i] = { ...table, columns }
	}
}

export function rewriteSheetMetadataFormulasForRename(
	sheet: Sheet,
	oldName: string,
	newName: string,
): void {
	for (let i = 0; i < sheet.dataValidations.length; i++) {
		const validation = sheet.dataValidations[i]
		if (!validation) continue
		const formula1 = rewriteFormulaTextForRename(validation.formula1, oldName, newName)
		const formula2 = rewriteFormulaTextForRename(validation.formula2, oldName, newName)
		sheet.dataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.conditionalFormats.length; i++) {
		const format = sheet.conditionalFormats[i]
		if (!format) continue
		sheet.conditionalFormats[i] = {
			...format,
			rules: format.rules.map((rule) => ({
				...rule,
				formulas: rule.formulas.map(
					(formula) => rewriteFormulaTextForRename(formula, oldName, newName) ?? formula,
				),
			})),
		}
	}
	for (let i = 0; i < sheet.tables.length; i++) {
		const table = sheet.tables[i]
		if (!table) continue
		const columns = table.columns.map((column) => {
			const formula = rewriteFormulaTextForRename(column.formula, oldName, newName)
			const totalsRowFormula = rewriteFormulaTextForRename(
				column.totalsRowFormula,
				oldName,
				newName,
			)
			return {
				...column,
				...(formula !== undefined ? { formula } : {}),
				...(totalsRowFormula !== undefined ? { totalsRowFormula } : {}),
			}
		})
		sheet.tables[i] = { ...table, columns }
	}
}

function rewriteNodeForShift(
	node: FormulaNode,
	targetSheet: string,
	formulaSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): FormulaNode {
	const onTarget = (refSheet: string | undefined) => (refSheet ?? formulaSheet) === targetSheet

	const shiftRef = (ref: FormulaCellRef, isOnTarget: boolean): FormulaCellRef => {
		if (!isOnTarget) return ref
		if (axis === 'row') {
			if (delta > 0) {
				if (ref.row >= at) return { ...ref, row: ref.row + delta }
			} else {
				const deleteEnd = at - delta
				if (ref.row >= deleteEnd) return { ...ref, row: ref.row + delta }
			}
		} else {
			if (delta > 0) {
				if (ref.col >= at) return { ...ref, col: ref.col + delta }
			} else {
				const deleteEnd = at - delta
				if (ref.col >= deleteEnd) return { ...ref, col: ref.col + delta }
			}
		}
		return ref
	}

	switch (node.type) {
		case 'cellRef': {
			const ref = shiftRef(node.ref, onTarget(node.sheet))
			return node.sheet !== undefined
				? { type: 'cellRef', ref, sheet: node.sheet }
				: { type: 'cellRef', ref }
		}
		case 'rangeRef': {
			const hit = onTarget(node.sheet)
			const start = shiftRef(node.start, hit)
			const end = shiftRef(node.end, hit)
			return node.sheet !== undefined
				? { type: 'rangeRef', start, end, sheet: node.sheet }
				: { type: 'rangeRef', start, end }
		}
		case 'wholeRowRange': {
			if (!onTarget(node.sheet) || axis !== 'row') return node
			const startRow = shiftIndex(node.startRow, at, delta)
			const endRow = shiftIndex(node.endRow, at, delta)
			if (startRow === null || endRow === null) return node
			return node.sheet !== undefined
				? { type: 'wholeRowRange', startRow, endRow, sheet: node.sheet }
				: { type: 'wholeRowRange', startRow, endRow }
		}
		case 'wholeColumnRange': {
			if (!onTarget(node.sheet) || axis !== 'col') return node
			const startCol = shiftIndex(node.startCol, at, delta)
			const endCol = shiftIndex(node.endCol, at, delta)
			if (startCol === null || endCol === null) return node
			return node.sheet !== undefined
				? { type: 'wholeColumnRange', startCol, endCol, sheet: node.sheet }
				: { type: 'wholeColumnRange', startCol, endCol }
		}
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteNodeForShift(node.left, targetSheet, formulaSheet, axis, at, delta),
				right: rewriteNodeForShift(node.right, targetSheet, formulaSheet, axis, at, delta),
			}
		case 'unary':
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteNodeForShift(node.operand, targetSheet, formulaSheet, axis, at, delta),
			}
		case 'function':
			return {
				type: 'function',
				name: node.name,
				args: node.args.map((arg) =>
					rewriteNodeForShift(arg, targetSheet, formulaSheet, axis, at, delta),
				),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((row) =>
					row.map((cell) => rewriteNodeForShift(cell, targetSheet, formulaSheet, axis, at, delta)),
				),
			}
		default:
			return node
	}
}

function rewriteSheetName(node: FormulaNode, oldName: string, newName: string): FormulaNode {
	switch (node.type) {
		case 'cellRef':
			return node.sheet === oldName ? { type: 'cellRef', ref: node.ref, sheet: newName } : node
		case 'rangeRef':
			return node.sheet === oldName
				? { type: 'rangeRef', start: node.start, end: node.end, sheet: newName }
				: node
		case 'wholeRowRange':
			return node.sheet === oldName
				? { type: 'wholeRowRange', startRow: node.startRow, endRow: node.endRow, sheet: newName }
				: node
		case 'wholeColumnRange':
			return node.sheet === oldName
				? { type: 'wholeColumnRange', startCol: node.startCol, endCol: node.endCol, sheet: newName }
				: node
		case 'sheetSpanRef':
			return {
				type: 'sheetSpanRef',
				startSheet: node.startSheet === oldName ? newName : node.startSheet,
				endSheet: node.endSheet === oldName ? newName : node.endSheet,
				target: rewriteSheetName(node.target, oldName, newName),
			}
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteSheetName(node.left, oldName, newName),
				right: rewriteSheetName(node.right, oldName, newName),
			}
		case 'unary':
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteSheetName(node.operand, oldName, newName),
			}
		case 'function':
			return {
				type: 'function',
				name: node.name,
				args: node.args.map((arg) => rewriteSheetName(arg, oldName, newName)),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((row) => row.map((cell) => rewriteSheetName(cell, oldName, newName))),
			}
		default:
			return node
	}
}
