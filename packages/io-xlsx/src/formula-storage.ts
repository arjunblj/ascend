import type { FormulaNode } from '@ascend/formulas'
import { parseFormula, printFormula } from '@ascend/formulas'

const FUTURE_FUNCTIONS = new Set([
	'AGGREGATE',
	'ANCHORARRAY',
	'BASE',
	'BETA.DIST',
	'BETA.INV',
	'BINOM.DIST',
	'BINOM.DIST.RANGE',
	'BINOM.INV',
	'BITAND',
	'BITLSHIFT',
	'BITOR',
	'BITRSHIFT',
	'BITXOR',
	'BYCOL',
	'BYROW',
	'CEILING.MATH',
	'CEILING.PRECISE',
	'CHISQ.DIST',
	'CHISQ.DIST.RT',
	'CHISQ.INV',
	'CHISQ.INV.RT',
	'CHISQ.TEST',
	'CHOOSECOLS',
	'CHOOSEROWS',
	'COMBINA',
	'CONCAT',
	'CONFIDENCE.NORM',
	'CONFIDENCE.T',
	'COVARIANCE.P',
	'COVARIANCE.S',
	'DAYS',
	'DECIMAL',
	'DROP',
	'ERF.PRECISE',
	'ERFC.PRECISE',
	'EXPAND',
	'EXPON.DIST',
	'F.DIST',
	'F.DIST.RT',
	'F.INV',
	'F.INV.RT',
	'F.TEST',
	'FIELDVALUE',
	'FILTER',
	'FLOOR.MATH',
	'FLOOR.PRECISE',
	'FORECAST.ETS',
	'FORECAST.ETS.CONFINT',
	'FORECAST.ETS.SEASONALITY',
	'FORECAST.ETS.STAT',
	'FORMULATEXT',
	'GAMMA',
	'GAMMA.DIST',
	'GAMMA.INV',
	'GAUSS',
	'HSTACK',
	'HYPGEOM.DIST',
	'IFNA',
	'IFS',
	'IMAGE',
	'ISFORMULA',
	'ISO.CEILING',
	'ISOWEEKNUM',
	'LAMBDA',
	'LET',
	'LOGNORM.DIST',
	'LOGNORM.INV',
	'MAKEARRAY',
	'MAP',
	'MAXIFS',
	'MINIFS',
	'MODE.MULT',
	'MODE.SNGL',
	'MUNIT',
	'NEGBINOM.DIST',
	'NETWORKDAYS.INTL',
	'NORM.DIST',
	'NORM.INV',
	'NORM.S.DIST',
	'NORM.S.INV',
	'NUMBERVALUE',
	'PDURATION',
	'PERCENTILE.EXC',
	'PERCENTILE.INC',
	'PERCENTRANK.EXC',
	'PERCENTRANK.INC',
	'PERMUTATIONA',
	'PHI',
	'POISSON.DIST',
	'QUARTILE.EXC',
	'QUARTILE.INC',
	'RANDARRAY',
	'RANK.AVG',
	'RANK.EQ',
	'REDUCE',
	'REGEXEXTRACT',
	'REGEXREPLACE',
	'REGEXTEST',
	'RRI',
	'SCAN',
	'SEQUENCE',
	'SHEET',
	'SHEETS',
	'SINGLE',
	'SORT',
	'SORTBY',
	'STDEV.P',
	'STDEV.S',
	'STOCKHISTORY',
	'SWITCH',
	'T.DIST',
	'T.DIST.2T',
	'T.DIST.RT',
	'T.INV',
	'T.INV.2T',
	'T.TEST',
	'TAKE',
	'TEXTAFTER',
	'TEXTBEFORE',
	'TEXTJOIN',
	'TEXTSPLIT',
	'TOCOL',
	'TOROW',
	'UNICHAR',
	'UNICODE',
	'UNIQUE',
	'VAR.P',
	'VAR.S',
	'VSTACK',
	'WEIBULL.DIST',
	'WORKDAY.INTL',
	'WRAPCOLS',
	'WRAPROWS',
	'XLOOKUP',
	'XMATCH',
	'XOR',
	'Z.TEST',
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
	const rewritten = rewriteToStoredAst(parsed.value)
	const originalPrinted = printFormula(parsed.value)
	const rewrittenPrinted = printFormula(rewritten)
	return rewrittenPrinted === originalPrinted ? formula : rewrittenPrinted
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

const STORAGE_PREFIX_RE = /_xlfn\.(?:_xlws\.)?|_xlws\./gi

function stripStoragePrefixes(text: string): string {
	return text.replace(STORAGE_PREFIX_RE, '')
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
	if (!formula.includes('(')) return false
	const upper = formula.toUpperCase()
	for (const name of FUTURE_FUNCTIONS) {
		if (name === 'ANCHORARRAY' || name === 'SINGLE') continue
		if (upper.includes(`${name}(`)) return true
	}
	return false
}
