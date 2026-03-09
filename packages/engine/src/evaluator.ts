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
	if (upperName === 'INDIRECT' || upperName === 'OFFSET') {
		return resolveReferenceFunction(upperName, argNodes, ctx).value
	}

	const def = functionRegistry.get(upperName)
	if (!def) return errorValue('#NAME?')
	if (argNodes.length < def.minArgs || argNodes.length > def.maxArgs) {
		return errorValue('#VALUE!')
	}

	const args = argNodes.map((argNode) => resolveArg(argNode, ctx))
	return def.evaluate(args, {
		...ctx.calcContext,
		sheetIndex: ctx.sheetIndex,
		row: ctx.row,
		col: ctx.col,
	})
}

function resolveArg(node: FormulaNode, ctx: EvalContext): EvalArg {
	if (node.type === 'function') {
		const upperName = node.name.toUpperCase()
		if (upperName === 'INDIRECT' || upperName === 'OFFSET') {
			return resolveReferenceFunction(upperName, node.args, ctx)
		}
	}

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

function resolveReferenceFunction(
	name: 'INDIRECT' | 'OFFSET',
	argNodes: readonly FormulaNode[],
	ctx: EvalContext,
): EvalArg {
	return name === 'INDIRECT'
		? resolveIndirectReference(argNodes, ctx)
		: resolveOffsetReference(argNodes, ctx)
}

function resolveIndirectReference(argNodes: readonly FormulaNode[], ctx: EvalContext): EvalArg {
	const textValue = evaluate(argNodes[0] ?? { type: 'missing' }, ctx)
	if (textValue.kind === 'error') return { value: textValue }
	const refText = coerceToString(textValue).trim()
	if (refText.length === 0) return { value: errorValue('#REF!') }

	const useA1Value =
		argNodes.length > 1 ? evaluate(argNodes[1] as FormulaNode, ctx) : booleanValue(true)
	if (useA1Value.kind === 'error') return { value: useA1Value }
	const useA1 =
		useA1Value.kind === 'boolean'
			? useA1Value.value
			: useA1Value.kind === 'number'
				? useA1Value.value !== 0
				: true

	if (useA1) {
		const parsed = parseFormula(refText)
		if (!parsed.ok) return { value: errorValue('#REF!') }
		const resolved = resolveReferenceNode(parsed.value, ctx)
		return resolved ?? { value: errorValue('#REF!') }
	}

	return resolveR1C1Reference(refText, ctx) ?? { value: errorValue('#REF!') }
}

function resolveOffsetReference(argNodes: readonly FormulaNode[], ctx: EvalContext): EvalArg {
	const base = resolveArg(argNodes[0] ?? { type: 'missing' }, ctx)
	if (base.value.kind === 'error') return base
	if (!base.ref) return { value: errorValue('#VALUE!') }

	const rowOffset = coerceToNumber(evaluate(argNodes[1] ?? { type: 'missing' }, ctx))
	if (rowOffset === null) return { value: errorValue('#VALUE!') }
	const colOffset = coerceToNumber(evaluate(argNodes[2] ?? { type: 'missing' }, ctx))
	if (colOffset === null) return { value: errorValue('#VALUE!') }

	const baseHeight =
		base.ref.kind === 'range' ? (base.ref.endRow ?? base.ref.row) - base.ref.row + 1 : 1
	const baseWidth =
		base.ref.kind === 'range' ? (base.ref.endCol ?? base.ref.col) - base.ref.col + 1 : 1

	const height =
		argNodes.length > 3 ? coerceToNumber(evaluate(argNodes[3] as FormulaNode, ctx)) : baseHeight
	if (height === null) return { value: errorValue('#VALUE!') }
	const width =
		argNodes.length > 4 ? coerceToNumber(evaluate(argNodes[4] as FormulaNode, ctx)) : baseWidth
	if (width === null) return { value: errorValue('#VALUE!') }

	const targetHeight = Math.trunc(height)
	const targetWidth = Math.trunc(width)
	if (targetHeight <= 0 || targetWidth <= 0) return { value: errorValue('#REF!') }

	const startRow = base.ref.row + Math.trunc(rowOffset)
	const startCol = base.ref.col + Math.trunc(colOffset)
	const endRow = startRow + targetHeight - 1
	const endCol = startCol + targetWidth - 1
	if (startRow < 0 || startCol < 0 || endRow < 0 || endCol < 0) {
		return { value: errorValue('#REF!') }
	}

	return makeRangeArg(ctx.workbook, base.ref.sheetIndex, startRow, startCol, endRow, endCol)
}

function resolveReferenceNode(node: FormulaNode, ctx: EvalContext): EvalArg | null {
	switch (node.type) {
		case 'cellRef': {
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return { value: errorValue('#REF!') }
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
		case 'rangeRef': {
			const si = resolveSheetIndex(ctx.workbook, node.sheet, ctx.sheetIndex)
			if (si < 0) return { value: errorValue('#REF!') }
			return makeRangeArg(
				ctx.workbook,
				si,
				node.start.row,
				node.start.col,
				node.end.row,
				node.end.col,
			)
		}
		case 'name': {
			const resolved = resolveDefinedName(node.name, node.sheet, ctx)
			if (!resolved) return { value: errorValue('#NAME?') }
			return (
				resolveReferenceNode(resolved.ast, resolved.ctx) ?? resolveArg(resolved.ast, resolved.ctx)
			)
		}
		case 'structuredRef': {
			const resolved = resolveStructuredRefRange(
				ctx.workbook,
				node,
				ctx.sheetIndex,
				ctx.row,
				ctx.col,
			)
			if (!resolved) return { value: errorValue('#REF!') }
			return makeRangeArg(
				ctx.workbook,
				resolved.sheetIndex,
				resolved.startRow,
				resolved.startCol,
				resolved.endRow,
				resolved.endCol,
			)
		}
		default:
			return null
	}
}

function makeRangeArg(
	workbook: Workbook,
	sheetIndex: number,
	startRow: number,
	startCol: number,
	endRow: number,
	endCol: number,
): EvalArg {
	const values = getRangeValues(workbook, sheetIndex, startRow, startCol, endRow, endCol)
	const firstRow = values[0]
	return {
		value: firstRow?.[0] ?? EMPTY,
		kind: 'range',
		values,
		ref: {
			kind: 'range',
			sheetIndex,
			row: startRow,
			col: startCol,
			endRow,
			endCol,
		},
	}
}

function resolveR1C1Reference(refText: string, ctx: EvalContext): EvalArg | null {
	const sheetSplit = refText.lastIndexOf('!')
	const sheetName =
		sheetSplit === -1 ? undefined : refText.slice(0, sheetSplit).replace(/^'|'$/g, '')
	const body = sheetSplit === -1 ? refText : refText.slice(sheetSplit + 1)
	const [startToken, endToken = body] = body.split(':')
	if (!startToken) return null
	const start = parseR1C1CellRef(startToken, ctx.row, ctx.col)
	const end = parseR1C1CellRef(endToken, ctx.row, ctx.col)
	if (!start || !end) return null
	const sheetIndex = resolveSheetIndex(ctx.workbook, sheetName, ctx.sheetIndex)
	if (sheetIndex < 0) return { value: errorValue('#REF!') }
	if (start.row < 0 || start.col < 0 || end.row < 0 || end.col < 0) {
		return { value: errorValue('#REF!') }
	}

	return makeRangeArg(ctx.workbook, sheetIndex, start.row, start.col, end.row, end.col)
}

function parseR1C1CellRef(
	token: string,
	currentRow: number,
	currentCol: number,
): { row: number; col: number } | null {
	const match = /^(R(?:\[-?\d+\]|\d+)?)(C(?:\[-?\d+\]|\d+)?)$/i.exec(token.trim())
	if (!match) return null
	const rowToken = match[1]
	const colToken = match[2]
	if (!rowToken || !colToken) return null
	return {
		row: parseR1C1Coordinate(rowToken, currentRow),
		col: parseR1C1Coordinate(colToken, currentCol),
	}
}

function parseR1C1Coordinate(token: string, currentIndex: number): number {
	const value = token.slice(1)
	if (value.startsWith('[') && value.endsWith(']')) {
		return currentIndex + Number.parseInt(value.slice(1, -1) || '0', 10)
	}
	if (value === '') return currentIndex
	return Number.parseInt(value, 10) - 1
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
