import type { Workbook } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'
import { type EvalArg, functionRegistry, parseFormula, toNumber } from '@ascend/formulas'
import type { CellValue } from '@ascend/schema'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import type { CalcContext } from './calc-context.ts'
import { resolveStructuredRefRange } from './structured-refs.ts'

export interface EvalContext {
	readonly workbook: Workbook
	readonly calcContext: CalcContext
	readonly sheetIndex: number
	readonly row: number
	readonly col: number
	readonly definedNameStack?: readonly string[]
}

function resolveSheetIndex(
	wb: Workbook,
	sheetName: string | undefined,
	currentSheet: number,
): number {
	if (sheetName === undefined) return currentSheet
	const idx = wb.sheets.findIndex((s) => s.name.toLowerCase() === sheetName.toLowerCase())
	return idx
}

function getCellValue(wb: Workbook, sheetIndex: number, row: number, col: number): CellValue {
	const sheet = wb.sheets[sheetIndex]
	if (!sheet) return errorValue('#REF!')
	const cell = sheet.cells.get(row, col)
	return cell ? cell.value : EMPTY
}

function getRangeValues(
	wb: Workbook,
	sheetIndex: number,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): readonly (readonly CellValue[])[] {
	const sheet = wb.sheets[sheetIndex]
	if (!sheet) return [[errorValue('#REF!')]]
	const rows: CellValue[][] = []
	for (let r = startRow; r <= endRow; r++) {
		const row: CellValue[] = []
		for (let c = startCol; c <= endCol; c++) {
			const cell = sheet.cells.get(r, c)
			row.push(cell ? cell.value : EMPTY)
		}
		rows.push(row)
	}
	return rows
}

function coerceToNumber(v: CellValue): number | null {
	return toNumber(v)
}

function coerceToString(v: CellValue): string {
	switch (v.kind) {
		case 'number':
			return String(v.value)
		case 'string':
			return v.value
		case 'boolean':
			return v.value ? 'TRUE' : 'FALSE'
		case 'empty':
			return ''
		case 'date':
			return String(v.serial)
		case 'error':
			return v.value
		case 'richText':
			return v.runs.map((r) => r.text).join('')
	}
}

function evalBinary(op: string, left: CellValue, right: CellValue): CellValue {
	if (left.kind === 'error') return left
	if (right.kind === 'error') return right

	if (op === '&') {
		return stringValue(coerceToString(left) + coerceToString(right))
	}

	if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
		return evalComparison(op, left, right)
	}

	const ln = coerceToNumber(left)
	const rn = coerceToNumber(right)
	if (ln === null || rn === null) return errorValue('#VALUE!')

	switch (op) {
		case '+':
			return numberValue(ln + rn)
		case '-':
			return numberValue(ln - rn)
		case '*':
			return numberValue(ln * rn)
		case '/':
			return rn === 0 ? errorValue('#DIV/0!') : numberValue(ln / rn)
		case '^':
			return numberValue(ln ** rn)
		default:
			return errorValue('#VALUE!')
	}
}

function evalComparison(op: string, left: CellValue, right: CellValue): CellValue {
	const ln = coerceToNumber(left)
	const rn = coerceToNumber(right)

	if (ln !== null && rn !== null) {
		return booleanValue(comparePrimitive(op, ln, rn))
	}

	if (left.kind === 'string' || right.kind === 'string') {
		const ls = coerceToString(left).toLowerCase()
		const rs = coerceToString(right).toLowerCase()
		return booleanValue(comparePrimitive(op, ls, rs))
	}

	if (left.kind === 'boolean' && right.kind === 'boolean') {
		const lb = left.value ? 1 : 0
		const rb = right.value ? 1 : 0
		return booleanValue(comparePrimitive(op, lb, rb))
	}

	return booleanValue(comparePrimitive(op, coerceToString(left), coerceToString(right)))
}

function comparePrimitive<T extends number | string>(op: string, a: T, b: T): boolean {
	switch (op) {
		case '=':
			return a === b
		case '<>':
			return a !== b
		case '<':
			return a < b
		case '>':
			return a > b
		case '<=':
			return a <= b
		case '>=':
			return a >= b
		default:
			return false
	}
}

export function evaluate(node: FormulaNode, ctx: EvalContext): CellValue {
	switch (node.type) {
		case 'number':
			return numberValue(node.value)
		case 'string':
			return stringValue(node.value)
		case 'boolean':
			return booleanValue(node.value)
		case 'error':
			return errorValue(node.value)
		case 'missing':
			return EMPTY

		case 'cellRef': {
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return errorValue('#REF!')
			return getCellValue(ctx.workbook, si, node.ref.row, node.ref.col)
		}

		case 'rangeRef': {
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return errorValue('#REF!')
			const values = getRangeValues(
				ctx.workbook,
				si,
				node.start.row,
				node.start.col,
				node.end.row,
				node.end.col,
			)
			const firstRow = values[0]
			if (firstRow?.[0]) return firstRow[0]
			return EMPTY
		}

		case 'name': {
			return evaluateDefinedName(node.name, node.sheet, ctx)
		}

		case 'binary': {
			const left = evaluate(node.left, ctx)
			const right = evaluate(node.right, ctx)
			return evalBinary(node.op, left, right)
		}

		case 'unary': {
			const operand = evaluate(node.operand, ctx)
			if (operand.kind === 'error') return operand
			const n = coerceToNumber(operand)
			if (n === null) return errorValue('#VALUE!')
			switch (node.op) {
				case '+':
					return numberValue(n)
				case '-':
					return numberValue(-n)
				case '%':
					return numberValue(n / 100)
			}
			break
		}

		case 'function':
			return evalFunction(node.name, node.args, ctx)

		case 'array': {
			const firstRow = node.rows[0]
			if (firstRow) {
				const firstEl = firstRow[0]
				if (firstEl) return evaluate(firstEl, ctx)
			}
			return EMPTY
		}

		case 'structuredRef': {
			const resolved = resolveStructuredRefRange(
				ctx.workbook,
				node,
				ctx.sheetIndex,
				ctx.row,
				ctx.col,
			)
			if (!resolved) return errorValue('#REF!')
			const values = getRangeValues(
				ctx.workbook,
				resolved.sheetIndex,
				resolved.startRow,
				resolved.startCol,
				resolved.endRow,
				resolved.endCol,
			)
			const firstRow = values[0]
			return firstRow?.[0] ?? EMPTY
		}
	}
	return EMPTY
}

function evalFunction(name: string, argNodes: readonly FormulaNode[], ctx: EvalContext): CellValue {
	const upperName = name.toUpperCase()
	if (upperName === 'ROW' && argNodes.length === 0) {
		return numberValue(ctx.row + 1)
	}
	if (upperName === 'COLUMN' && argNodes.length === 0) {
		return numberValue(ctx.col + 1)
	}

	const def = functionRegistry.get(upperName)
	if (!def) return errorValue('#NAME?')

	const args = argNodes.map((argNode) => resolveArg(argNode, ctx))
	return def.evaluate(args, ctx.calcContext)
}

function resolveArg(node: FormulaNode, ctx: EvalContext): EvalArg {
	if (node.type === 'rangeRef') {
		const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
		if (si < 0) {
			return { value: errorValue('#REF!') }
		}
		const values = getRangeValues(
			ctx.workbook,
			si,
			node.start.row,
			node.start.col,
			node.end.row,
			node.end.col,
		)
		const firstRow = values[0]
		const firstVal = firstRow ? (firstRow[0] ?? EMPTY) : EMPTY
		return {
			value: firstVal,
			kind: 'range',
			values,
			ref: {
				kind: 'range',
				sheetIndex: si,
				row: node.start.row,
				col: node.start.col,
				endRow: node.end.row,
				endCol: node.end.col,
			},
		}
	}

	if (node.type === 'cellRef') {
		const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
		if (si < 0) {
			return { value: errorValue('#REF!') }
		}
		return {
			value: getCellValue(ctx.workbook, si, node.ref.row, node.ref.col),
			ref: {
				kind: 'cell',
				sheetIndex: si,
				row: node.ref.row,
				col: node.ref.col,
			},
		}
	}

	if (node.type === 'name') {
		const resolved = resolveDefinedName(node.name, node.sheet, ctx)
		if (!resolved) return { value: errorValue('#NAME?') }
		return resolveArg(resolved.ast, resolved.ctx)
	}

	if (node.type === 'structuredRef') {
		const resolved = resolveStructuredRefRange(ctx.workbook, node, ctx.sheetIndex, ctx.row, ctx.col)
		if (!resolved) return { value: errorValue('#REF!') }
		const values = getRangeValues(
			ctx.workbook,
			resolved.sheetIndex,
			resolved.startRow,
			resolved.startCol,
			resolved.endRow,
			resolved.endCol,
		)
		const firstRow = values[0]
		const firstVal = firstRow ? (firstRow[0] ?? EMPTY) : EMPTY
		return {
			value: firstVal,
			kind: 'range',
			values,
			ref: {
				kind: 'range',
				sheetIndex: resolved.sheetIndex,
				row: resolved.startRow,
				col: resolved.startCol,
				endRow: resolved.endRow,
				endCol: resolved.endCol,
			},
		}
	}

	const value = evaluate(node, ctx)
	return { value }
}

function evaluateDefinedName(name: string, sheet: string | undefined, ctx: EvalContext): CellValue {
	const resolved = resolveDefinedName(name, sheet, ctx)
	if (!resolved) return errorValue('#NAME?')
	return evaluate(resolved.ast, resolved.ctx)
}

function resolveDefinedName(
	name: string,
	sheet: string | undefined,
	ctx: EvalContext,
): { ast: FormulaNode; ctx: EvalContext } | null {
	const explicitSheet = sheet ? ctx.workbook.getSheet(sheet) : undefined
	const currentSheet = ctx.workbook.sheets[ctx.sheetIndex]
	const entry = ctx.workbook.definedNames.resolve(name, currentSheet?.id, explicitSheet?.id)
	if (!entry) return null

	const entryKey =
		entry.scope.kind === 'workbook'
			? `workbook:${entry.name.toLowerCase()}`
			: `sheet:${entry.scope.sheetId}:${entry.name.toLowerCase()}`
	if (ctx.definedNameStack?.includes(entryKey)) return null

	const parsed = parseFormula(entry.formula)
	if (!parsed.ok) return null

	let sheetIndex = ctx.sheetIndex
	if (entry.scope.kind === 'sheet') {
		const scope = entry.scope
		const localSheetIndex = ctx.workbook.sheets.findIndex(
			(workbookSheet) => workbookSheet.id === scope.sheetId,
		)
		if (localSheetIndex >= 0) sheetIndex = localSheetIndex
	}

	return {
		ast: parsed.value,
		ctx: {
			...ctx,
			sheetIndex,
			definedNameStack: [...(ctx.definedNameStack ?? []), entryKey],
		},
	}
}
