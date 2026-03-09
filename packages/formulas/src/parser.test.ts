import { describe, expect, it } from 'bun:test'
import type { FormulaNode } from './ast.ts'
import { parseFormula } from './parser.ts'
import { printFormula } from './printer.ts'
import { extractRefs, rewriteRefs } from './refs.ts'

function p(formula: string): FormulaNode {
	const result = parseFormula(formula)
	if (!result.ok) throw new Error(result.error.message)
	return result.value
}

describe('parse', () => {
	it('parses number literals', () => {
		expect(p('42')).toEqual({ type: 'number', value: 42 })
		expect(p('3.14')).toEqual({ type: 'number', value: 3.14 })
		expect(p('1.5E+3')).toEqual({ type: 'number', value: 1500 })
	})

	it('parses string literals', () => {
		expect(p('"hello"')).toEqual({ type: 'string', value: 'hello' })
		expect(p('"say ""hi"""')).toEqual({ type: 'string', value: 'say "hi"' })
	})

	it('parses boolean literals', () => {
		expect(p('TRUE')).toEqual({ type: 'boolean', value: true })
		expect(p('FALSE')).toEqual({ type: 'boolean', value: false })
	})

	it('parses error literals', () => {
		expect(p('#N/A')).toEqual({ type: 'error', value: '#N/A' })
		expect(p('#DIV/0!')).toEqual({ type: 'error', value: '#DIV/0!' })
		expect(p('#VALUE!')).toEqual({ type: 'error', value: '#VALUE!' })
	})

	it('parses addition and subtraction', () => {
		const node = p('1+2')
		expect(node).toEqual({
			type: 'binary',
			op: '+',
			left: { type: 'number', value: 1 },
			right: { type: 'number', value: 2 },
		})
	})

	it('respects multiplication before addition', () => {
		const node = p('1+2*3')
		expect(node.type).toBe('binary')
		if (node.type !== 'binary') return
		expect(node.op).toBe('+')
		expect(node.left).toEqual({ type: 'number', value: 1 })
		expect(node.right).toEqual({
			type: 'binary',
			op: '*',
			left: { type: 'number', value: 2 },
			right: { type: 'number', value: 3 },
		})
	})

	it('respects exponentiation before multiplication', () => {
		const node = p('2*3^4')
		expect(node.type).toBe('binary')
		if (node.type !== 'binary') return
		expect(node.op).toBe('*')
		expect(node.right).toEqual({
			type: 'binary',
			op: '^',
			left: { type: 'number', value: 3 },
			right: { type: 'number', value: 4 },
		})
	})

	it('parses exponentiation as right associative', () => {
		const node = p('2^3^2')
		expect(node).toEqual({
			type: 'binary',
			op: '^',
			left: { type: 'number', value: 2 },
			right: {
				type: 'binary',
				op: '^',
				left: { type: 'number', value: 3 },
				right: { type: 'number', value: 2 },
			},
		})
	})

	it('parses unary minus after exponentiation precedence', () => {
		const node = p('-2^2')
		expect(node).toEqual({
			type: 'unary',
			op: '-',
			operand: {
				type: 'binary',
				op: '^',
				left: { type: 'number', value: 2 },
				right: { type: 'number', value: 2 },
			},
		})
	})

	it('parses unary minus', () => {
		const node = p('-A1')
		expect(node.type).toBe('unary')
		if (node.type !== 'unary') return
		expect(node.op).toBe('-')
		expect(node.operand.type).toBe('cellRef')
	})

	it('parses unary plus', () => {
		const node = p('+5')
		expect(node).toEqual({
			type: 'unary',
			op: '+',
			operand: { type: 'number', value: 5 },
		})
	})

	it('parses postfix percent', () => {
		const node = p('50%')
		expect(node).toEqual({
			type: 'unary',
			op: '%',
			operand: { type: 'number', value: 50 },
		})
	})

	it('parses spill references', () => {
		const node = p('A1#')
		expect(node).toEqual({
			type: 'spillRef',
			target: {
				type: 'cellRef',
				ref: { row: 0, col: 0, rowAbsolute: false, colAbsolute: false },
			},
		})
	})

	it('parses implicit intersection prefix', () => {
		const node = p('@A1')
		expect(node).toEqual({
			type: 'unary',
			op: '@',
			operand: {
				type: 'cellRef',
				ref: { row: 0, col: 0, rowAbsolute: false, colAbsolute: false },
			},
		})
	})

	it('parses concatenation', () => {
		const node = p('"a"&"b"')
		expect(node).toEqual({
			type: 'binary',
			op: '&',
			left: { type: 'string', value: 'a' },
			right: { type: 'string', value: 'b' },
		})
	})

	it('parses comparison operators', () => {
		for (const op of ['=', '<>', '<', '>', '<=', '>='] as const) {
			const node = p(`1${op}2`)
			expect(node.type).toBe('binary')
			if (node.type === 'binary') expect(node.op).toBe(op)
		}
	})

	it('parses function calls', () => {
		const node = p('SUM(1,2,3)')
		expect(node.type).toBe('function')
		if (node.type !== 'function') return
		expect(node.name).toBe('SUM')
		expect(node.args).toHaveLength(3)
	})

	it('parses function with missing arguments', () => {
		const node = p('IF(A1,,0)')
		expect(node.type).toBe('function')
		if (node.type !== 'function') return
		expect(node.name).toBe('IF')
		expect(node.args).toHaveLength(3)
		expect(node.args[1]).toEqual({ type: 'missing' })
	})

	it('parses cell references', () => {
		const node = p('A1')
		expect(node.type).toBe('cellRef')
		if (node.type !== 'cellRef') return
		expect(node.ref).toEqual({ row: 0, col: 0, rowAbsolute: false, colAbsolute: false })
	})

	it('parses absolute cell references', () => {
		const node = p('$A$1')
		expect(node.type).toBe('cellRef')
		if (node.type !== 'cellRef') return
		expect(node.ref.rowAbsolute).toBe(true)
		expect(node.ref.colAbsolute).toBe(true)
	})

	it('parses range references', () => {
		const node = p('A1:B10')
		expect(node.type).toBe('rangeRef')
		if (node.type !== 'rangeRef') return
		expect(node.start).toEqual({ row: 0, col: 0, rowAbsolute: false, colAbsolute: false })
		expect(node.end).toEqual({ row: 9, col: 1, rowAbsolute: false, colAbsolute: false })
	})

	it('parses sheet-qualified cell references', () => {
		const node = p('Sheet1!A1')
		expect(node.type).toBe('cellRef')
		if (node.type !== 'cellRef') return
		expect(node.sheet).toBe('Sheet1')
		expect(node.ref.row).toBe(0)
	})

	it('parses quoted sheet name references', () => {
		const node = p("'My Sheet'!A1")
		expect(node.type).toBe('cellRef')
		if (node.type !== 'cellRef') return
		expect(node.sheet).toBe('My Sheet')
	})

	it('parses sheet-qualified range references', () => {
		const node = p('Sheet2!A1:B10')
		expect(node.type).toBe('rangeRef')
		if (node.type !== 'rangeRef') return
		expect(node.sheet).toBe('Sheet2')
	})

	it('parses parenthesized expressions', () => {
		const node = p('(1+2)*3')
		expect(node.type).toBe('binary')
		if (node.type !== 'binary') return
		expect(node.op).toBe('*')
		expect(node.left).toEqual({
			type: 'binary',
			op: '+',
			left: { type: 'number', value: 1 },
			right: { type: 'number', value: 2 },
		})
	})

	it('parses array literals', () => {
		const node = p('{1,2;3,4}')
		expect(node.type).toBe('array')
		if (node.type !== 'array') return
		expect(node.rows).toHaveLength(2)
		expect(node.rows[0]).toHaveLength(2)
	})

	it('parses named ranges', () => {
		const node = p('MyRange')
		expect(node).toEqual({ type: 'name', name: 'MyRange' })
	})

	it('parses sheet-qualified defined names', () => {
		const node = p('Sheet1!Budget')
		expect(node).toEqual({ type: 'name', name: 'Budget', sheet: 'Sheet1' })
	})

	it('parses SUM(A1:B10)', () => {
		const node = p('SUM(A1:B10)')
		expect(node.type).toBe('function')
		if (node.type !== 'function') return
		expect(node.name).toBe('SUM')
		expect(node.args).toHaveLength(1)
		expect(node.args[0]?.type).toBe('rangeRef')
	})

	it('parses IF(A1>0,A1*2,"negative")', () => {
		const node = p('IF(A1>0,A1*2,"negative")')
		expect(node.type).toBe('function')
		if (node.type !== 'function') return
		expect(node.name).toBe('IF')
		expect(node.args).toHaveLength(3)

		const cond = node.args[0]
		expect(cond?.type).toBe('binary')
		if (cond?.type !== 'binary') return
		expect(cond.op).toBe('>')
	})

	it('parses VLOOKUP(A1,Sheet2!A1:B10,2,FALSE)', () => {
		const node = p('VLOOKUP(A1,Sheet2!A1:B10,2,FALSE)')
		expect(node.type).toBe('function')
		if (node.type !== 'function') return
		expect(node.name).toBe('VLOOKUP')
		expect(node.args).toHaveLength(4)
		expect(node.args[1]?.type).toBe('rangeRef')
		if (node.args[1]?.type === 'rangeRef') {
			expect(node.args[1].sheet).toBe('Sheet2')
		}
		expect(node.args[3]).toEqual({ type: 'boolean', value: false })
	})

	it('returns error result for invalid input', () => {
		const result = parseFormula(')')
		expect(result.ok).toBe(false)
	})
})

describe('printFormula', () => {
	it('roundtrips simple expressions', () => {
		const cases = ['1+2', '1+2*3', '(1+2)*3', '"hello"', 'TRUE', '#N/A', 'A1', '$A$1']
		for (const c of cases) {
			expect(printFormula(p(c))).toBe(c)
		}
	})

	it('roundtrips function calls', () => {
		expect(printFormula(p('SUM(A1:B10)'))).toBe('SUM(A1:B10)')
		expect(printFormula(p('IF(A1>0,A1*2,"neg")'))).toBe('IF(A1>0,A1*2,"neg")')
	})

	it('roundtrips sheet-qualified references', () => {
		expect(printFormula(p('Sheet1!A1'))).toBe('Sheet1!A1')
		expect(printFormula(p("'My Sheet'!A1"))).toBe("'My Sheet'!A1")
		expect(printFormula(p('Sheet1!Budget'))).toBe('Sheet1!Budget')
	})

	it('preserves parentheses where needed', () => {
		expect(printFormula(p('1-(2+3)'))).toBe('1-(2+3)')
		expect(printFormula(p('(1+2)*3'))).toBe('(1+2)*3')
	})

	it('roundtrips unary operators', () => {
		expect(printFormula(p('-A1'))).toBe('-A1')
		expect(printFormula(p('50%'))).toBe('50%')
		expect(printFormula(p('@A1'))).toBe('@A1')
		expect(printFormula(p('A1#'))).toBe('A1#')
	})

	it('roundtrips array literals', () => {
		expect(printFormula(p('{1,2;3,4}'))).toBe('{1,2;3,4}')
	})
})

describe('extractRefs', () => {
	it('extracts cell refs from simple formula', () => {
		const refs = extractRefs(p('A1+B2'))
		expect(refs).toHaveLength(2)
		expect(refs[0]?.kind).toBe('cell')
		expect(refs[1]?.kind).toBe('cell')
	})

	it('extracts range refs', () => {
		const refs = extractRefs(p('SUM(A1:B10)'))
		expect(refs).toHaveLength(1)
		expect(refs[0]?.kind).toBe('range')
	})

	it('extracts refs from nested expressions', () => {
		const refs = extractRefs(p('IF(A1>0,B1*C1,D1)'))
		expect(refs).toHaveLength(4)
	})
})

describe('rewriteRefs', () => {
	it('shifts row references', () => {
		const original = p('A1+B2')
		const shifted = rewriteRefs(original, (ref) => ({
			...ref,
			row: ref.row + 1,
		}))
		expect(printFormula(shifted)).toBe('A2+B3')
	})

	it('shifts column references', () => {
		const original = p('A1')
		const shifted = rewriteRefs(original, (ref) => ({
			...ref,
			col: ref.col + 1,
		}))
		expect(printFormula(shifted)).toBe('B1')
	})

	it('preserves non-ref nodes', () => {
		const original = p('1+2')
		const same = rewriteRefs(original, (ref) => ref)
		expect(same).toEqual(original)
	})
})
