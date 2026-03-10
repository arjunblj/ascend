import type { FormulaNode } from '@ascend/formulas'
import { parseFormula, printFormula } from '@ascend/formulas'

const FUTURE_FUNCTIONS = new Set([
	'ANCHORARRAY',
	'CHOOSECOLS',
	'DROP',
	'FILTER',
	'HSTACK',
	'LET',
	'RANDARRAY',
	'SEQUENCE',
	'SINGLE',
	'SORT',
	'SORTBY',
	'TAKE',
	'TEXTSPLIT',
	'TOCOL',
	'TOROW',
	'UNIQUE',
	'XMATCH',
	'XLOOKUP',
])

export function normalizeStoredFormulaText(formula: string): string {
	if (!needsStoredFormulaNormalization(formula)) return formula
	const stripped = stripStoragePrefixes(formula)
	if (!requiresStoredAstNormalization(stripped)) return stripped
	const parsed = parseFormula(stripped)
	if (!parsed.ok) return stripped
	return printFormula(rewriteFromStoredAst(parsed.value))
}

export function toStoredFormulaText(formula: string): string {
	if (!needsStoredFormulaDenormalization(formula)) return formula
	const parsed = parseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(rewriteToStoredAst(parsed.value))
}

function rewriteFromStoredAst(node: FormulaNode): FormulaNode {
	switch (node.type) {
		case 'function': {
			const name = stripStoragePrefixes(node.name)
			const upper = name.toUpperCase()
			if (upper === 'ANCHORARRAY' && node.args[0]) {
				return { type: 'spillRef', target: rewriteFromStoredAst(node.args[0]) }
			}
			if (upper === 'SINGLE' && node.args[0]) {
				return { type: 'unary', op: '@', operand: rewriteFromStoredAst(node.args[0]) }
			}
			return {
				type: 'function',
				name,
				args: node.args.map(rewriteFromStoredAst),
			}
		}
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteFromStoredAst(node.left),
				right: rewriteFromStoredAst(node.right),
			}
		case 'unary':
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteFromStoredAst(node.operand),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((row) => row.map(rewriteFromStoredAst)),
			}
		case 'spillRef':
			return {
				type: 'spillRef',
				target: rewriteFromStoredAst(node.target),
			}
		case 'sheetSpanRef':
			return {
				type: 'sheetSpanRef',
				startSheet: node.startSheet,
				endSheet: node.endSheet,
				target: rewriteFromStoredAst(node.target),
			}
		default:
			return node
	}
}

function rewriteToStoredAst(node: FormulaNode): FormulaNode {
	switch (node.type) {
		case 'function': {
			const name = shouldPrefixFutureFunction(node.name)
				? `_xlfn.${stripStoragePrefixes(node.name)}`
				: stripStoragePrefixes(node.name)
			return {
				type: 'function',
				name,
				args: node.args.map(rewriteToStoredAst),
			}
		}
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteToStoredAst(node.left),
				right: rewriteToStoredAst(node.right),
			}
		case 'unary':
			if (node.op === '@') {
				return {
					type: 'function',
					name: '_xlfn.SINGLE',
					args: [rewriteToStoredAst(node.operand)],
				}
			}
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteToStoredAst(node.operand),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((row) => row.map(rewriteToStoredAst)),
			}
		case 'spillRef':
			return {
				type: 'function',
				name: '_xlfn.ANCHORARRAY',
				args: [rewriteToStoredAst(node.target)],
			}
		case 'sheetSpanRef':
			return {
				type: 'sheetSpanRef',
				startSheet: node.startSheet,
				endSheet: node.endSheet,
				target: rewriteToStoredAst(node.target),
			}
		default:
			return node
	}
}

function stripStoragePrefixes(text: string): string {
	return text
		.replace(/_xlfn\._xlws\./gi, '')
		.replace(/_xlfn\./gi, '')
		.replace(/_xlws\./gi, '')
}

function shouldPrefixFutureFunction(name: string): boolean {
	return FUTURE_FUNCTIONS.has(stripStoragePrefixes(name).toUpperCase())
}

function needsStoredFormulaNormalization(formula: string): boolean {
	return formula.includes('_xlfn.') || formula.includes('_xlws.')
}

function requiresStoredAstNormalization(formula: string): boolean {
	const upper = formula.toUpperCase()
	return upper.includes('ANCHORARRAY(') || upper.includes('SINGLE(')
}

function needsStoredFormulaDenormalization(formula: string): boolean {
	if (formula.includes('#') || formula.includes('@')) return true
	const upper = formula.toUpperCase()
	for (const name of FUTURE_FUNCTIONS) {
		if (name === 'ANCHORARRAY' || name === 'SINGLE') continue
		if (upper.includes(`${name}(`)) return true
	}
	return false
}
