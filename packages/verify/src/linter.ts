import type { Workbook } from '@ascend/core'
import { toA1 } from '@ascend/core'
import { analyzeWorkbookFormulas, type WorkbookFormulaAnalysis } from '@ascend/engine'
import type { FormulaNode } from '@ascend/formulas'
import { functionRegistry } from '@ascend/formulas'

export interface LintOptions {
	readonly volatileThreshold?: number
	readonly fragileRefThreshold?: number
	readonly complexityDepthWarning?: number
	readonly complexityDepthError?: number
}

export interface LintResult {
	readonly violations: readonly LintViolation[]
}

export interface LintViolation {
	readonly rule: string
	readonly severity: 'warning' | 'info' | 'error'
	readonly message: string
	readonly ref: string
	readonly formula: string
}

const VOLATILE_FNS = new Set<string>()
for (const [name, def] of functionRegistry) {
	if (def.volatile) VOLATILE_FNS.add(name)
}

const SAFE_NUMBERS = new Set([0, 1, 100])
const COMMON_PERCENTAGES = new Set([
	0.01, 0.05, 0.1, 0.15, 0.2, 0.25, 0.5, 0.75, 2, 10, 12, 24, 30, 52, 60, 90, 100, 360, 365,
])

function countVolatileCalls(node: FormulaNode): number {
	if (node.type === 'function') {
		const isSelf = VOLATILE_FNS.has(node.name.toUpperCase()) ? 1 : 0
		return isSelf + node.args.reduce((n, a) => n + countVolatileCalls(a), 0)
	}
	if (node.type === 'binary') return countVolatileCalls(node.left) + countVolatileCalls(node.right)
	if (node.type === 'dynamicRangeRef')
		return countVolatileCalls(node.start) + countVolatileCalls(node.end)
	if (node.type === 'unary') return countVolatileCalls(node.operand)
	if (node.type === 'array') {
		return node.rows.reduce((n, row) => n + row.reduce((m, c) => m + countVolatileCalls(c), 0), 0)
	}
	if (node.type === 'spillRef') return countVolatileCalls(node.target)
	return 0
}

function findMagicNumbers(node: FormulaNode): number[] {
	const magic: number[] = []
	walkForMagic(node, magic)
	return magic
}

function walkForMagic(node: FormulaNode, out: number[]): void {
	switch (node.type) {
		case 'number':
			if (!SAFE_NUMBERS.has(node.value) && !COMMON_PERCENTAGES.has(node.value)) {
				out.push(node.value)
			}
			break
		case 'binary':
			walkForMagic(node.left, out)
			walkForMagic(node.right, out)
			break
		case 'dynamicRangeRef':
			walkForMagic(node.start, out)
			walkForMagic(node.end, out)
			break
		case 'unary':
			walkForMagic(node.operand, out)
			break
		case 'function':
			for (const arg of node.args) walkForMagic(arg, out)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) walkForMagic(cell, out)
			}
			break
		case 'spillRef':
			walkForMagic(node.target, out)
			break
		default:
			break
	}
}

function findFragileRefs(node: FormulaNode, threshold: number): boolean {
	switch (node.type) {
		case 'rangeRef': {
			const rowSpan = Math.abs(node.end.row - node.start.row) + 1
			const colSpan = Math.abs(node.end.col - node.start.col) + 1
			if (rowSpan * colSpan >= threshold) {
				const allAbsolute =
					node.start.rowAbsolute &&
					node.start.colAbsolute &&
					node.end.rowAbsolute &&
					node.end.colAbsolute
				return !allAbsolute
			}
			return false
		}
		case 'binary':
			return findFragileRefs(node.left, threshold) || findFragileRefs(node.right, threshold)
		case 'dynamicRangeRef':
			return findFragileRefs(node.start, threshold) || findFragileRefs(node.end, threshold)
		case 'unary':
			return findFragileRefs(node.operand, threshold)
		case 'function':
			return node.args.some((a) => findFragileRefs(a, threshold))
		case 'array':
			return node.rows.some((row) => row.some((c) => findFragileRefs(c, threshold)))
		case 'spillRef':
			return findFragileRefs(node.target, threshold)
		default:
			return false
	}
}

function astDepth(node: FormulaNode): number {
	switch (node.type) {
		case 'binary':
			return 1 + Math.max(astDepth(node.left), astDepth(node.right))
		case 'dynamicRangeRef':
			return 1 + Math.max(astDepth(node.start), astDepth(node.end))
		case 'unary':
			return 1 + astDepth(node.operand)
		case 'function':
			return 1 + (node.args.length > 0 ? Math.max(...node.args.map(astDepth)) : 0)
		case 'array':
			return 1 + node.rows.reduce((max, row) => Math.max(max, ...row.map(astDepth)), 0)
		case 'spillRef':
			return 1 + astDepth(node.target)
		case 'sheetSpanRef':
			return 1 + astDepth(node.target)
		default:
			return 1
	}
}

function collectNameReferences(node: FormulaNode, names: Set<string>): void {
	switch (node.type) {
		case 'name':
			names.add(node.name.toLowerCase())
			break
		case 'binary':
			collectNameReferences(node.left, names)
			collectNameReferences(node.right, names)
			break
		case 'dynamicRangeRef':
			collectNameReferences(node.start, names)
			collectNameReferences(node.end, names)
			break
		case 'unary':
			collectNameReferences(node.operand, names)
			break
		case 'function':
			for (const arg of node.args) collectNameReferences(arg, names)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) collectNameReferences(cell, names)
			}
			break
		case 'spillRef':
			collectNameReferences(node.target, names)
			break
		case 'sheetSpanRef':
			collectNameReferences(node.target, names)
			break
		default:
			break
	}
}

export function lint(
	workbook: Workbook,
	analysis?: WorkbookFormulaAnalysis,
	options?: LintOptions,
): LintResult {
	const violations: LintViolation[] = []
	const compiled = analysis ?? analyzeWorkbookFormulas(workbook)

	const volatileThreshold = options?.volatileThreshold ?? 10
	const fragileRefThreshold = options?.fragileRefThreshold ?? 100
	const complexityWarning = options?.complexityDepthWarning ?? 10
	const complexityError = options?.complexityDepthError ?? 20

	const sheetStats = new Map<string, { volatileCount: number; volatileCells: number }>()
	const referencedNames = new Set<string>()

	for (const formula of compiled.formulas.values()) {
		const ref = `${formula.sheetName}!${toA1({ row: formula.row, col: formula.col })}`
		if (!formula.ast) {
			violations.push({
				rule: 'parse-error',
				severity: 'warning',
				message: `Unparseable formula: ${formula.formula}`,
				ref,
				formula: formula.formula,
			})
			continue
		}

		const ast = formula.ast
		const volatileCalls = countVolatileCalls(ast)
		if (volatileCalls > 0) {
			const stats = sheetStats.get(formula.sheetName) ?? { volatileCount: 0, volatileCells: 0 }
			stats.volatileCount += volatileCalls
			stats.volatileCells += 1
			sheetStats.set(formula.sheetName, stats)
		}

		const magic = findMagicNumbers(ast)
		if (magic.length > 0) {
			violations.push({
				rule: 'hardcoded-in-formula',
				severity: 'info',
				message: `Formula contains magic number(s): ${magic.join(', ')}`,
				ref,
				formula: formula.formula,
			})
		}

		if (findFragileRefs(ast, fragileRefThreshold)) {
			violations.push({
				rule: 'fragile-refs',
				severity: 'warning',
				message: 'Non-absolute reference spanning a large range',
				ref,
				formula: formula.formula,
			})
		}

		const depth = astDepth(ast)
		if (depth > complexityError) {
			violations.push({
				rule: 'complex-formula',
				severity: 'error',
				message: `Formula has depth ${depth} (exceeds error threshold of ${complexityError})`,
				ref,
				formula: formula.formula,
			})
		} else if (depth > complexityWarning) {
			violations.push({
				rule: 'complex-formula',
				severity: 'warning',
				message: `Formula has depth ${depth} (exceeds warning threshold of ${complexityWarning})`,
				ref,
				formula: formula.formula,
			})
		}

		collectNameReferences(ast, referencedNames)
	}

	for (const [sheetName, stats] of sheetStats) {
		if (stats.volatileCount > volatileThreshold) {
			violations.push({
				rule: 'volatile-overuse',
				severity: 'warning',
				message: `Sheet "${sheetName}" has ${stats.volatileCount} volatile function calls across ${stats.volatileCells} cell(s)`,
				ref: `${sheetName}!A1`,
				formula: '',
			})
		}
	}

	for (const entry of workbook.definedNames.list()) {
		if (!referencedNames.has(entry.name.toLowerCase())) {
			violations.push({
				rule: 'unused-name',
				severity: 'info',
				message: `Defined name "${entry.name}" is never referenced by any formula`,
				ref: entry.formula,
				formula: '',
			})
		}
	}

	return { violations }
}
