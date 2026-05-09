import type { CellFormulaBinding } from '@ascend/core'
import type { FormulaNode, StructuredRefNode, Token } from '@ascend/formulas'
import { functionRegistry, printFormula, tokenize } from '@ascend/formulas'
import type { CellValue } from '@ascend/schema'
import type { FormulaInfo, FormulaReferenceInfo, FormulaReferenceScope } from './types.ts'

interface BuildFormulaInfoInput {
	readonly ref: string
	readonly formula: string
	readonly value: CellValue
	readonly binding?: CellFormulaBinding | undefined
	readonly tokens: readonly Token[]
	readonly ast?: FormulaNode
	readonly normalizedFormula?: string
	readonly functions?: readonly string[]
	readonly references?: readonly FormulaReferenceInfo[]
	readonly parseError?: string
	readonly volatile?: boolean
}

export function buildFormulaInfo(input: BuildFormulaInfoInput): FormulaInfo {
	const references = input.references ?? (input.ast ? collectFormulaReferences(input.ast) : [])
	return {
		ref: input.ref,
		formula: input.formula,
		normalizedFormula:
			input.normalizedFormula ?? (input.ast ? printFormula(input.ast) : input.formula),
		value: input.value,
		...(input.binding ? { binding: input.binding } : {}),
		references,
		refs: flattenLegacyReferenceTexts(references),
		functions: input.functions
			? [...input.functions]
			: input.ast
				? [...collectFunctionNames(input.ast)]
				: [],
		volatile: input.ast ? hasVolatileFunction(input.ast) : (input.volatile ?? false),
		tokens: input.tokens,
		...(input.ast ? { ast: input.ast } : {}),
		...(input.parseError ? { parseError: input.parseError } : {}),
	}
}

export function tokenizeFormulaInput(formula: string): readonly Token[] {
	return tokenize(formula).filter((token) => token.type !== 'Whitespace' && token.type !== 'EOF')
}

export function collectFormulaReferences(node: FormulaNode): readonly FormulaReferenceInfo[] {
	const refs: FormulaReferenceInfo[] = []
	appendFormulaReferences(node, refs)
	return refs
}

export function flattenLegacyReferenceTexts(
	references: readonly FormulaReferenceInfo[],
): readonly string[] {
	const texts: string[] = []
	for (const reference of references) {
		if (reference.kind === 'union' || reference.kind === 'intersection') {
			texts.push(...flattenLegacyReferenceTexts(reference.members))
			continue
		}
		texts.push(reference.text)
	}
	return texts
}

export function hasVolatileFunction(node: FormulaNode): boolean {
	switch (node.type) {
		case 'function':
			if (functionRegistry.get(node.name.toUpperCase())?.volatile) return true
			return node.args.some(hasVolatileFunction)
		case 'binary':
			return hasVolatileFunction(node.left) || hasVolatileFunction(node.right)
		case 'dynamicRangeRef':
			return hasVolatileFunction(node.start) || hasVolatileFunction(node.end)
		case 'unary':
			return hasVolatileFunction(node.operand)
		case 'array':
			return node.rows.some((row) => row.some(hasVolatileFunction))
		case 'spillRef':
			return hasVolatileFunction(node.target)
		case 'sheetSpanRef':
			return hasVolatileFunction(node.target)
		case 'wholeRowRange':
		case 'wholeColumnRange':
			return false
		default:
			return false
	}
}

export function collectFunctionNames(node: FormulaNode, out = new Set<string>()): Set<string> {
	switch (node.type) {
		case 'function':
			out.add(node.name)
			for (const arg of node.args) collectFunctionNames(arg, out)
			break
		case 'binary':
			collectFunctionNames(node.left, out)
			collectFunctionNames(node.right, out)
			break
		case 'dynamicRangeRef':
			collectFunctionNames(node.start, out)
			collectFunctionNames(node.end, out)
			break
		case 'unary':
			collectFunctionNames(node.operand, out)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) collectFunctionNames(cell, out)
			}
			break
		case 'spillRef':
			collectFunctionNames(node.target, out)
			break
		case 'sheetSpanRef':
			collectFunctionNames(node.target, out)
			break
		case 'wholeRowRange':
		case 'wholeColumnRange':
			break
	}
	return out
}

function appendFormulaReferences(node: FormulaNode, out: FormulaReferenceInfo[]): void {
	const direct = toFormulaReferenceInfo(node)
	if (direct) {
		out.push(direct)
		return
	}
	switch (node.type) {
		case 'binary':
			appendFormulaReferences(node.left, out)
			appendFormulaReferences(node.right, out)
			break
		case 'dynamicRangeRef':
			appendFormulaReferences(node.start, out)
			appendFormulaReferences(node.end, out)
			break
		case 'unary':
			appendFormulaReferences(node.operand, out)
			break
		case 'function':
			for (const arg of node.args) appendFormulaReferences(arg, out)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) appendFormulaReferences(cell, out)
			}
			break
		case 'sheetSpanRef':
			appendFormulaReferences(node.target, out)
			break
		default:
			break
	}
}

function toFormulaReferenceInfo(node: FormulaNode): FormulaReferenceInfo | undefined {
	switch (node.type) {
		case 'cellRef':
			return { kind: 'cell', text: printFormula(node), scope: formulaReferenceScope(node.sheet) }
		case 'rangeRef':
			return { kind: 'range', text: printFormula(node), scope: formulaReferenceScope(node.sheet) }
		case 'dynamicRangeRef':
			return { kind: 'range', text: printFormula(node), scope: formulaReferenceScope(undefined) }
		case 'wholeRowRange':
			return {
				kind: 'wholeRow',
				text: printFormula(node),
				scope: formulaReferenceScope(node.sheet),
			}
		case 'wholeColumnRange':
			return {
				kind: 'wholeColumn',
				text: printFormula(node),
				scope: formulaReferenceScope(node.sheet),
			}
		case 'name':
			return { kind: 'name', text: printFormula(node), scope: formulaReferenceScope(node.sheet) }
		case 'structuredRef':
			return structuredFormulaReferenceInfo(node)
		case 'spillRef': {
			const target = toFormulaReferenceInfo(node.target)
			return {
				kind: 'spill',
				text: printFormula(node),
				targetText: printFormula(node.target),
				...(target ? { target } : {}),
			}
		}
		case 'sheetSpanRef': {
			const target = toFormulaReferenceInfo(node.target)
			if (!target) return undefined
			return {
				...target,
				text: printFormula(node),
				scope: { kind: 'sheetSpan', startSheet: node.startSheet, endSheet: node.endSheet },
			}
		}
		case 'unary':
			if (node.op !== '@') return undefined
			return {
				kind: 'implicitIntersection',
				text: printFormula(node),
				targetText: printFormula(node.operand),
				...(toFormulaReferenceInfo(node.operand)
					? { target: toFormulaReferenceInfo(node.operand) as FormulaReferenceInfo }
					: {}),
			}
		case 'binary':
			if (node.op !== ',' && node.op !== ' ') return undefined
			return buildCompoundReferenceInfo(node)
		default:
			return undefined
	}
}

function buildCompoundReferenceInfo(
	node: Extract<FormulaNode, { type: 'binary' }>,
): FormulaReferenceInfo | undefined {
	const left = toFormulaReferenceInfo(node.left)
	const right = toFormulaReferenceInfo(node.right)
	if (!left || !right) return undefined
	return {
		kind: node.op === ',' ? 'union' : 'intersection',
		text: printFormula(node),
		members: [left, right],
	}
}

function structuredFormulaReferenceInfo(node: StructuredRefNode): FormulaReferenceInfo {
	return {
		kind: 'structured',
		text: printFormula(node),
		scope: { kind: 'local' },
		table: node.table,
		specifiers: [...node.specifiers],
		...(node.column ? { column: node.column } : {}),
		...(node.endColumn ? { endColumn: node.endColumn } : {}),
	}
}

function formulaReferenceScope(sheet: string | undefined): FormulaReferenceScope {
	if (sheet) {
		const external = /^\[([^\]]+)\](.+)$/.exec(sheet)
		if (external) {
			const workbook = external[1]
			const externalSheet = external[2]
			if (workbook && externalSheet) {
				return { kind: 'external', workbook, sheet: externalSheet }
			}
		}
	}
	return sheet ? { kind: 'sheet', sheet } : { kind: 'local' }
}
