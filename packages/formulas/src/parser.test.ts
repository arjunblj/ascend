import { describe, expect, it, test } from 'bun:test'
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

	it('parses whole-column ranges', () => {
		expect(p('A:C')).toEqual({ type: 'wholeColumnRange', startCol: 0, endCol: 2 })
	})

	it('parses whole-row ranges', () => {
		expect(p('1:3')).toEqual({ type: 'wholeRowRange', startRow: 0, endRow: 2 })
	})

	it('parses union references at top level', () => {
		expect(p('A1,B2')).toEqual({
			type: 'binary',
			op: ',',
			left: {
				type: 'cellRef',
				ref: { row: 0, col: 0, rowAbsolute: false, colAbsolute: false },
			},
			right: {
				type: 'cellRef',
				ref: { row: 1, col: 1, rowAbsolute: false, colAbsolute: false },
			},
		})
	})

	it('parses intersection references with whitespace', () => {
		expect(p('A:A 2:2')).toEqual({
			type: 'binary',
			op: ' ',
			left: { type: 'wholeColumnRange', startCol: 0, endCol: 0 },
			right: { type: 'wholeRowRange', startRow: 1, endRow: 1 },
		})
	})

	it('parses sheet-qualified whole-column and whole-row ranges', () => {
		expect(p('Sheet1!A:C')).toEqual({
			type: 'wholeColumnRange',
			startCol: 0,
			endCol: 2,
			sheet: 'Sheet1',
		})
		expect(p('Sheet1!1:3')).toEqual({
			type: 'wholeRowRange',
			startRow: 0,
			endRow: 2,
			sheet: 'Sheet1',
		})
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

	it('parses workbook-qualified external references', () => {
		const node = p('[Book.xlsx]Sheet1!A1')
		expect(node).toEqual({
			type: 'cellRef',
			sheet: '[Book.xlsx]Sheet1',
			ref: { row: 0, col: 0, rowAbsolute: false, colAbsolute: false },
		})
	})

	it('parses 3D sheet-span references', () => {
		expect(p('Sheet1:Sheet3!A1')).toEqual({
			type: 'sheetSpanRef',
			startSheet: 'Sheet1',
			endSheet: 'Sheet3',
			target: {
				type: 'cellRef',
				ref: { row: 0, col: 0, rowAbsolute: false, colAbsolute: false },
			},
		})
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
		expect(printFormula(p('[Book.xlsx]Sheet1!A1'))).toBe('[Book.xlsx]Sheet1!A1')
		expect(printFormula(p('Sheet1:Sheet3!A1'))).toBe('Sheet1:Sheet3!A1')
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

	it('roundtrips whole-row and whole-column ranges', () => {
		expect(printFormula(p('A:C'))).toBe('A:C')
		expect(printFormula(p('1:3'))).toBe('1:3')
		expect(printFormula(p('Sheet1!A:C'))).toBe('Sheet1!A:C')
	})

	it('roundtrips union and intersection references', () => {
		expect(printFormula(p('A1,B2'))).toBe('A1,B2')
		expect(printFormula(p('A:A 2:2'))).toBe('A:A 2:2')
		expect(printFormula(p('SUM((A1,B2))'))).toBe('SUM((A1,B2))')
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

	it('extracts refs from union and intersection expressions', () => {
		expect(extractRefs(p('A1,B2'))).toHaveLength(2)
		expect(extractRefs(p('A:A 2:2'))).toHaveLength(2)
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

describe('fuzz: parseFormula never throws', () => {
	function xorshift32(state: { s: number }): number {
		let s = state.s
		s ^= s << 13
		s ^= s >> 17
		s ^= s << 5
		state.s = s >>> 0
		return (s >>> 0) / 0xffffffff
	}

	function pick<T>(arr: readonly T[], rng: { s: number }): T {
		return arr[Math.floor(xorshift32(rng) * arr.length)] as T
	}

	function randInt(rng: { s: number }, min: number, max: number): number {
		return min + Math.floor(xorshift32(rng) * (max - min + 1))
	}

	function randCol(rng: { s: number }): string {
		const cols = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
		return cols[randInt(rng, 0, 25)] as string
	}

	function randCellRef(rng: { s: number }): string {
		const abs = xorshift32(rng) < 0.3 ? '$' : ''
		const rowAbs = xorshift32(rng) < 0.3 ? '$' : ''
		return `${abs}${randCol(rng)}${rowAbs}${randInt(rng, 1, 999)}`
	}

	function randRange(rng: { s: number }): string {
		return `${randCellRef(rng)}:${randCellRef(rng)}`
	}

	const FN_NAMES = [
		'SUM',
		'AVERAGE',
		'COUNT',
		'MAX',
		'MIN',
		'IF',
		'VLOOKUP',
		'INDEX',
		'MATCH',
		'CONCATENATE',
		'LEFT',
		'RIGHT',
		'MID',
		'LEN',
		'TRIM',
		'ROUND',
		'ABS',
		'AND',
		'OR',
		'NOT',
		'IFERROR',
		'SUMIF',
		'COUNTIF',
		'COUNTA',
	] as const

	const BINARY_OPS = ['+', '-', '*', '/', '^', '&', '=', '<>', '<', '>', '<=', '>='] as const

	function randAtom(rng: { s: number }): string {
		const kind = randInt(rng, 0, 5)
		switch (kind) {
			case 0:
				return String(randInt(rng, 0, 9999))
			case 1:
				return `"${String.fromCharCode(...Array.from({ length: randInt(rng, 0, 8) }, () => randInt(rng, 65, 90)))}"` // random quoted string
			case 2:
				return randCellRef(rng)
			case 3:
				return xorshift32(rng) < 0.5 ? 'TRUE' : 'FALSE'
			case 4:
				return pick(['#N/A', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#NULL!'], rng)
			default:
				return randRange(rng)
		}
	}

	function randExpr(rng: { s: number }, depth: number): string {
		if (depth <= 0) return randAtom(rng)
		const kind = randInt(rng, 0, 5)
		switch (kind) {
			case 0: {
				const fn = pick(FN_NAMES, rng)
				const argc = randInt(rng, 1, 4)
				const args = Array.from({ length: argc }, () => randExpr(rng, depth - 1))
				return `${fn}(${args.join(',')})`
			}
			case 1:
				return `${randExpr(rng, depth - 1)}${pick(BINARY_OPS, rng)}${randExpr(rng, depth - 1)}`
			case 2:
				return `(${randExpr(rng, depth - 1)})`
			case 3:
				return `-${randExpr(rng, depth - 1)}`
			case 4:
				return `${randExpr(rng, depth - 1)}%`
			default:
				return randAtom(rng)
		}
	}

	function randGarbage(rng: { s: number }): string {
		const chars =
			'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+-*/^&=<>(),:;!@#$%{}[] "'
		const len = randInt(rng, 1, 40)
		return Array.from({ length: len }, () => chars[randInt(rng, 0, chars.length - 1)]).join('')
	}

	it('valid formula patterns parse without throwing', () => {
		const formulas = [
			'SUM(A1:B10)',
			'IF(A1>0,B1,C1)',
			'VLOOKUP(D1,A:B,2,0)',
			'AVERAGE(A1:A100)',
			'COUNT(A:A)',
			'MAX(1,2,3)',
			'MIN(A1,B1)',
			'CONCATENATE("a","b")',
			'INDEX(A1:C10,2,3)',
			'MATCH(A1,B1:B10,0)',
			'SUMIF(A:A,">0")',
			'COUNTIF(B:B,"test")',
			'AND(TRUE,FALSE)',
			'OR(A1>0,B1<10)',
			'NOT(TRUE)',
			'ROUND(3.14,1)',
			'LEFT("hello",3)',
			'RIGHT("world",3)',
			'MID("test",2,2)',
			'LEN("abc")',
			'TRIM("  hi  ")',
			'1+2*3',
			'(1+2)*3',
			'-A1',
			'50%',
			'A1&B1',
			'A1=B1',
			'A1<>B1',
			'{1,2;3,4}',
			'Sheet1!A1',
			'SUM(Sheet1!A:A)',
			'IF(AND(A1>0,B1<10),A1*B1,0)',
			'IFERROR(A1/B1,0)',
			'A1+B1*C1-D1/E1^2',
		]
		for (const f of formulas) {
			const result = parseFormula(f)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	it('invalid formula patterns parse without throwing', () => {
		const formulas = [
			')',
			'(()',
			'SUM(',
			'+',
			',,',
			'===',
			'!!!',
			'{',
			'}',
			'(',
			'SUM(,)',
			'""',
			'"""',
			'1++2',
			'--A1',
			'SUM(1,2,',
			'IF(A1>)',
			')))',
			'SUM(A1:)',
			':B1',
			'A1:',
		]
		for (const f of formulas) {
			const result = parseFormula(f)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('empty string', () => {
		const result = parseFormula('')
		expect(result.ok === true || result.ok === false).toBe(true)
	})

	test('very long formula (10K chars)', () => {
		const long = `SUM(${Array.from({ length: 500 }, (_, i) => `A${i + 1}`).join(',')})` // ~3K chars
		const result = parseFormula(long)
		expect(result.ok === true || result.ok === false).toBe(true)

		const longer = `${'A1+'.repeat(3333)}A1`
		const result2 = parseFormula(longer)
		expect(result2.ok === true || result2.ok === false).toBe(true)
	})

	test('deeply nested parens (100 levels)', () => {
		const deep = `${'('.repeat(100)}1${')'.repeat(100)}`
		const result = parseFormula(deep)
		expect(result.ok === true || result.ok === false).toBe(true)
	})

	test('max args (255)', () => {
		const args = Array.from({ length: 255 }, (_, i) => `A${i + 1}`).join(',')
		const result = parseFormula(`SUM(${args})`)
		expect(result.ok === true || result.ok === false).toBe(true)
	})

	test('special characters: unicode, newlines, tabs, null bytes', () => {
		const specials = [
			'"日本語"',
			'"hello\\nworld"',
			'A1+\t+B1',
			'\0',
			'"emoji🎉"',
			'"null\\0byte"',
			'SUM(\t)',
			'"café"',
			'"naïve"',
			'"中文"',
		]
		for (const f of specials) {
			const result = parseFormula(f)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('random generated formulas never throw (100 cases)', () => {
		const rng = { s: 42 }
		for (let i = 0; i < 100; i++) {
			const depth = randInt(rng, 0, 4)
			const formula = xorshift32(rng) < 0.7 ? randExpr(rng, depth) : randGarbage(rng)
			const result = parseFormula(formula)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('random garbage strings never throw (50 cases)', () => {
		const rng = { s: 99999 }
		for (let i = 0; i < 50; i++) {
			const formula = randGarbage(rng)
			const result = parseFormula(formula)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('corrupted valid formulas never throw (50 cases)', () => {
		const rng = { s: 77777 }
		const bases = [
			'SUM(A1:B10)',
			'IF(A1>0,B1,C1)',
			'VLOOKUP(D1,A:B,2,0)',
			'INDEX(A1:C10,2,3)',
			'A1+B1*C1',
		]
		for (let i = 0; i < 50; i++) {
			const base = pick(bases, rng)
			const chars = [...base]
			const mutations = randInt(rng, 1, 3)
			for (let m = 0; m < mutations; m++) {
				const action = randInt(rng, 0, 2)
				const pos = randInt(rng, 0, chars.length - 1)
				if (action === 0 && chars.length > 0) {
					chars.splice(pos, 1)
				} else if (action === 1) {
					chars.splice(pos, 0, String.fromCharCode(randInt(rng, 32, 126)))
				} else if (chars.length > 0) {
					chars[pos] = String.fromCharCode(randInt(rng, 32, 126))
				}
			}
			const result = parseFormula(chars.join(''))
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})
})
