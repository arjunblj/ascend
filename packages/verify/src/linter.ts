import type { Workbook } from '@ascend/core'
import { toA1 } from '@ascend/core'
import { analyzeWorkbook, type WorkbookAnalysis } from '@ascend/engine'
import type { FormulaNode } from '@ascend/formulas'
import { functionRegistry } from '@ascend/formulas'

export interface LintResult {
	readonly violations: readonly LintViolation[]
}

export interface LintViolation {
	readonly rule: string
	readonly severity: 'warning' | 'info'
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
	if (node.type === 'unary') return countVolatileCalls(node.operand)
	if (node.type === 'array') {
		return node.rows.reduce((n, row) => n + row.reduce((m, c) => m + countVolatileCalls(c), 0), 0)
	}
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
		default:
			break
	}
}

const LARGE_RANGE_THRESHOLD = 100

function findFragileRefs(node: FormulaNode): boolean {
	switch (node.type) {
		case 'rangeRef': {
			const rowSpan = Math.abs(node.end.row - node.start.row) + 1
			const colSpan = Math.abs(node.end.col - node.start.col) + 1
			if (rowSpan * colSpan >= LARGE_RANGE_THRESHOLD) {
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
			return findFragileRefs(node.left) || findFragileRefs(node.right)
		case 'unary':
			return findFragileRefs(node.operand)
		case 'function':
			return node.args.some((a) => findFragileRefs(a))
		case 'array':
			return node.rows.some((row) => row.some((c) => findFragileRefs(c)))
		default:
			return false
	}
}

export function lint(workbook: Workbook, analysis?: WorkbookAnalysis): LintResult {
	const violations: LintViolation[] = []
	const compiled = analysis ?? analyzeWorkbook(workbook)

	for (const sheet of workbook.sheets) {
		let sheetVolatileCount = 0
		const volatileCells: string[] = []

		for (const formula of compiled.formulas.values()) {
			const ref = `${sheet.name}!${toA1({ row: formula.row, col: formula.col })}`
			if (formula.sheetName !== sheet.name) continue
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

			const vc = countVolatileCalls(ast)
			if (vc > 0) {
				sheetVolatileCount += vc
				volatileCells.push(ref)
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

			if (findFragileRefs(ast)) {
				violations.push({
					rule: 'fragile-refs',
					severity: 'warning',
					message: 'Non-absolute reference spanning a large range',
					ref,
					formula: formula.formula,
				})
			}
		}

		if (sheetVolatileCount > 10) {
			violations.push({
				rule: 'volatile-overuse',
				severity: 'warning',
				message: `Sheet "${sheet.name}" has ${sheetVolatileCount} volatile function calls across ${volatileCells.length} cell(s)`,
				ref: `${sheet.name}!A1`,
				formula: '',
			})
		}
	}

	return { violations }
}
