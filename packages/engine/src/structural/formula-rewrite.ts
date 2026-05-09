import type { Cell, Sheet, Workbook } from '@ascend/core'
import type { FormulaCellRef, FormulaNode } from '@ascend/formulas'
import { cachedParseFormula, printFormula } from '@ascend/formulas'
import { shiftIndex } from './ref-shift.ts'

export function rewriteWorkbookFormulasForShift(
	workbook: Workbook,
	targetSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	for (const sheet of workbook.sheets) {
		const isTarget = sheet.name === targetSheet
		const updates: [number, number, Cell][] = []
		for (const [row, col, existing] of sheet.cells.iterate()) {
			if (existing.formula === null) continue
			const parsed = cachedParseFormula(existing.formula)
			if (!parsed.ok) continue
			if (!isTarget && !formulaAstReferencesSheet(parsed.value, targetSheet)) continue
			const rewritten = rewriteFormulaAstForShift(
				parsed.value,
				targetSheet,
				sheet.name,
				axis,
				at,
				delta,
			)
			if (rewritten === parsed.value) continue
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
		const parsed = cachedParseFormula(entry.formula)
		if (!parsed.ok) continue
		const formulaSheet = scopeSheet ?? targetSheet
		if (formulaSheet !== targetSheet && !formulaAstReferencesSheet(parsed.value, targetSheet))
			continue
		const rewritten = rewriteFormulaAstForShift(
			parsed.value,
			targetSheet,
			formulaSheet,
			axis,
			at,
			delta,
		)
		if (rewritten === parsed.value) continue
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
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(
		rewriteFormulaAstForShift(parsed.value, targetSheet, formulaSheet, axis, at, delta),
	)
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
			const parsed = cachedParseFormula(existing.formula)
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
		const parsed = cachedParseFormula(entry.formula)
		if (!parsed.ok) continue
		const rewritten = rewriteSheetName(parsed.value, oldName, newName)
		const formula = printFormula(rewritten)
		if (formula === entry.formula) continue
		workbook.definedNames.set(entry.name, formula, entry.scope)
	}
}

export function rewriteTableNameInFormulas(
	workbook: Workbook,
	oldName: string,
	newName: string,
): void {
	for (const sheet of workbook.sheets) {
		const updates: [number, number, Cell][] = []
		for (const [row, col, existing] of sheet.cells.iterate()) {
			if (existing.formula === null) continue
			const rewrittenFormula = rewriteFormulaTextForTableRename(existing.formula, oldName, newName)
			if (rewrittenFormula === undefined) continue
			if (rewrittenFormula === existing.formula) continue
			updates.push([
				row,
				col,
				{
					value: existing.value,
					formula: rewrittenFormula,
					styleId: existing.styleId,
				},
			])
		}
		for (const [row, col, updated] of updates) {
			sheet.cells.set(row, col, updated)
		}
		rewriteSheetMetadataFormulasForTableRename(sheet, oldName, newName)
	}
}

export function rewriteTableNameInDefinedNames(
	workbook: Workbook,
	oldName: string,
	newName: string,
): void {
	const entries = [...workbook.definedNames.list()]
	for (const entry of entries) {
		const formula = rewriteFormulaTextForTableRename(entry.formula, oldName, newName)
		if (formula === undefined) continue
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
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(rewriteSheetName(parsed.value, oldName, newName))
}

export function rewriteFormulaTextForTableRename(
	formula: string | undefined,
	oldName: string,
	newName: string,
): string | undefined {
	if (!formula) return formula
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(rewriteTableName(parsed.value, oldName, newName))
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

export function rewriteSheetMetadataFormulasForTableRename(
	sheet: Sheet,
	oldName: string,
	newName: string,
): void {
	for (let i = 0; i < sheet.dataValidations.length; i++) {
		const validation = sheet.dataValidations[i]
		if (!validation) continue
		const formula1 = rewriteFormulaTextForTableRename(validation.formula1, oldName, newName)
		const formula2 = rewriteFormulaTextForTableRename(validation.formula2, oldName, newName)
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
					(formula) => rewriteFormulaTextForTableRename(formula, oldName, newName) ?? formula,
				),
			})),
		}
	}
	for (let i = 0; i < sheet.tables.length; i++) {
		const table = sheet.tables[i]
		if (!table) continue
		const columns = table.columns.map((column) => {
			const formula = rewriteFormulaTextForTableRename(column.formula, oldName, newName)
			const totalsRowFormula = rewriteFormulaTextForTableRename(
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

export function rewriteFormulaAstForShift(
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
			return ref === node.ref
				? node
				: node.sheet !== undefined
					? { type: 'cellRef', ref, sheet: node.sheet }
					: { type: 'cellRef', ref }
		}
		case 'rangeRef': {
			const hit = onTarget(node.sheet)
			const start = shiftRef(node.start, hit)
			const end = shiftRef(node.end, hit)
			return start === node.start && end === node.end
				? node
				: node.sheet !== undefined
					? { type: 'rangeRef', start, end, sheet: node.sheet }
					: { type: 'rangeRef', start, end }
		}
		case 'wholeRowRange': {
			if (!onTarget(node.sheet) || axis !== 'row') return node
			const startRow = shiftIndex(node.startRow, at, delta)
			const endRow = shiftIndex(node.endRow, at, delta)
			if (startRow === null || endRow === null) return node
			return startRow === node.startRow && endRow === node.endRow
				? node
				: node.sheet !== undefined
					? { type: 'wholeRowRange', startRow, endRow, sheet: node.sheet }
					: { type: 'wholeRowRange', startRow, endRow }
		}
		case 'wholeColumnRange': {
			if (!onTarget(node.sheet) || axis !== 'col') return node
			const startCol = shiftIndex(node.startCol, at, delta)
			const endCol = shiftIndex(node.endCol, at, delta)
			if (startCol === null || endCol === null) return node
			return startCol === node.startCol && endCol === node.endCol
				? node
				: node.sheet !== undefined
					? { type: 'wholeColumnRange', startCol, endCol, sheet: node.sheet }
					: { type: 'wholeColumnRange', startCol, endCol }
		}
		case 'binary': {
			const left = rewriteFormulaAstForShift(node.left, targetSheet, formulaSheet, axis, at, delta)
			const right = rewriteFormulaAstForShift(
				node.right,
				targetSheet,
				formulaSheet,
				axis,
				at,
				delta,
			)
			return left === node.left && right === node.right
				? node
				: { type: 'binary', op: node.op, left, right }
		}
		case 'dynamicRangeRef': {
			const start = rewriteFormulaAstForShift(
				node.start,
				targetSheet,
				formulaSheet,
				axis,
				at,
				delta,
			)
			const end = rewriteFormulaAstForShift(node.end, targetSheet, formulaSheet, axis, at, delta)
			return start === node.start && end === node.end
				? node
				: { type: 'dynamicRangeRef', start, end }
		}
		case 'unary': {
			const operand = rewriteFormulaAstForShift(
				node.operand,
				targetSheet,
				formulaSheet,
				axis,
				at,
				delta,
			)
			return operand === node.operand ? node : { type: 'unary', op: node.op, operand }
		}
		case 'spillRef': {
			const target = rewriteFormulaAstForShift(
				node.target,
				targetSheet,
				formulaSheet,
				axis,
				at,
				delta,
			)
			return target === node.target ? node : { type: 'spillRef', target }
		}
		case 'function': {
			const args = node.args.map((arg) =>
				rewriteFormulaAstForShift(arg, targetSheet, formulaSheet, axis, at, delta),
			)
			return args.every((arg, i) => arg === node.args[i])
				? node
				: { type: 'function', name: node.name, args }
		}
		case 'array': {
			const rows = node.rows.map((row) =>
				row.map((cell) =>
					rewriteFormulaAstForShift(cell, targetSheet, formulaSheet, axis, at, delta),
				),
			)
			return rows.every((row, ri) => row.every((cell, ci) => cell === node.rows[ri]?.[ci]))
				? node
				: { type: 'array', rows }
		}
		default:
			return node
	}
}

export function formulaAstReferencesSheet(node: FormulaNode, sheetName: string): boolean {
	switch (node.type) {
		case 'cellRef':
		case 'rangeRef':
		case 'wholeRowRange':
		case 'wholeColumnRange':
		case 'name':
			return node.sheet === sheetName
		case 'sheetSpanRef':
			return node.startSheet === sheetName || node.endSheet === sheetName
		case 'binary':
			return (
				formulaAstReferencesSheet(node.left, sheetName) ||
				formulaAstReferencesSheet(node.right, sheetName)
			)
		case 'dynamicRangeRef':
			return (
				formulaAstReferencesSheet(node.start, sheetName) ||
				formulaAstReferencesSheet(node.end, sheetName)
			)
		case 'unary':
			return formulaAstReferencesSheet(node.operand, sheetName)
		case 'spillRef':
			return formulaAstReferencesSheet(node.target, sheetName)
		case 'function':
			return node.args.some((arg) => formulaAstReferencesSheet(arg, sheetName))
		case 'array':
			return node.rows.some((row) => row.some((cell) => formulaAstReferencesSheet(cell, sheetName)))
		default:
			return false
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
		case 'dynamicRangeRef':
			return {
				type: 'dynamicRangeRef',
				start: rewriteSheetName(node.start, oldName, newName),
				end: rewriteSheetName(node.end, oldName, newName),
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

function rewriteTableName(node: FormulaNode, oldName: string, newName: string): FormulaNode {
	switch (node.type) {
		case 'structuredRef':
			return node.table === oldName ? { ...node, table: newName } : node
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteTableName(node.left, oldName, newName),
				right: rewriteTableName(node.right, oldName, newName),
			}
		case 'dynamicRangeRef':
			return {
				type: 'dynamicRangeRef',
				start: rewriteTableName(node.start, oldName, newName),
				end: rewriteTableName(node.end, oldName, newName),
			}
		case 'unary':
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteTableName(node.operand, oldName, newName),
			}
		case 'function':
			return {
				type: 'function',
				name: node.name,
				args: node.args.map((arg) => rewriteTableName(arg, oldName, newName)),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((row) => row.map((cell) => rewriteTableName(cell, oldName, newName))),
			}
		case 'sheetSpanRef':
			return {
				...node,
				target: rewriteTableName(node.target, oldName, newName),
			}
		case 'spillRef':
			return {
				type: 'spillRef',
				target: rewriteTableName(node.target, oldName, newName),
			}
		default:
			return node
	}
}
