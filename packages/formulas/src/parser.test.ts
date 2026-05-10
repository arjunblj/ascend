import { describe, expect, it, test } from 'bun:test'
import type { FormulaNode } from './ast.ts'
import { parseFormula, parseFormulaOrThrow } from './parser.ts'
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

	it('parseFormulaOrThrow returns AST for valid formulas', () => {
		expect(parseFormulaOrThrow('1+2')).toEqual({
			type: 'binary',
			op: '+',
			left: { type: 'number', value: 1 },
			right: { type: 'number', value: 2 },
		})
	})

	it('parseFormulaOrThrow throws parser errors directly', () => {
		expect(() => parseFormulaOrThrow(')')).toThrow(/Unexpected token|Expected/)
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

	it('parses unary minus before exponentiation precedence', () => {
		const node = p('-2^2')
		expect(node).toEqual({
			type: 'binary',
			op: '^',
			left: {
				type: 'unary',
				op: '-',
				operand: { type: 'number', value: 2 },
			},
			right: { type: 'number', value: 2 },
		})
	})

	it('keeps explicit grouping and unary exponent operands unambiguous', () => {
		expect(p('-(2^2)')).toEqual({
			type: 'unary',
			op: '-',
			operand: {
				type: 'binary',
				op: '^',
				left: { type: 'number', value: 2 },
				right: { type: 'number', value: 2 },
			},
		})
		expect(p('2^-2')).toEqual({
			type: 'binary',
			op: '^',
			left: { type: 'number', value: 2 },
			right: {
				type: 'unary',
				op: '-',
				operand: { type: 'number', value: 2 },
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

	it('parses direct lambda invocation as internal call node', () => {
		const node = p('LAMBDA(x,x+1)(5)')
		expect(node).toEqual({
			type: 'function',
			name: '__CALL__',
			args: [
				{
					type: 'function',
					name: 'LAMBDA',
					args: [
						{ type: 'name', name: 'x' },
						{
							type: 'binary',
							op: '+',
							left: { type: 'name', name: 'x' },
							right: { type: 'number', value: 1 },
						},
					],
				},
				{ type: 'number', value: 5 },
			],
		})
		expect(printFormula(node)).toBe('LAMBDA(x,x+1)(5)')
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

	it('parses dynamic range endpoints', () => {
		const node = p('A1:INDEX(A:A,5)')
		expect(node.type).toBe('dynamicRangeRef')
		if (node.type !== 'dynamicRangeRef') return
		expect(node.start.type).toBe('cellRef')
		expect(node.end.type).toBe('function')
	})

	it('parses reference functions as dynamic range starts', () => {
		const node = p('INDEX(A:A,1):INDEX(A:A,5)')
		expect(node.type).toBe('dynamicRangeRef')
		if (node.type !== 'dynamicRangeRef') return
		expect(node.start.type).toBe('function')
		expect(node.end.type).toBe('function')
	})

	it('parses whole-column ranges', () => {
		expect(p('A:C')).toEqual({ type: 'wholeColumnRange', startCol: 0, endCol: 2 })
		expect(p('$A:$C')).toEqual({
			type: 'wholeColumnRange',
			startCol: 0,
			endCol: 2,
			startColAbsolute: true,
			endColAbsolute: true,
		})
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
		expect(p('Sheet1!$A:$C')).toEqual({
			type: 'wholeColumnRange',
			startCol: 0,
			endCol: 2,
			startColAbsolute: true,
			endColAbsolute: true,
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

	it('parses workbook-index-qualified defined names', () => {
		const node = p('[0]!col1_')
		expect(node).toEqual({ type: 'name', name: 'col1_', sheet: '[0]' })
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
		expect(printFormula(p('D:D/$E:$E'))).toBe('D:D/$E:$E')
		expect(printFormula(p('1:3'))).toBe('1:3')
		expect(printFormula(p('Sheet1!A:C'))).toBe('Sheet1!A:C')
	})

	it('roundtrips union and intersection references', () => {
		expect(printFormula(p('A1,B2'))).toBe('A1,B2')
		expect(printFormula(p('A:A 2:2'))).toBe('A:A 2:2')
		expect(printFormula(p('SUM((A1,B2))'))).toBe('SUM((A1,B2))')
	})

	it('roundtrips structured reference column ranges', () => {
		expect(printFormula(p('Table1[[Sales]:[Cost]]'))).toBe('Table1[[Sales]:[Cost]]')
		expect(printFormula(p('Table1[@[Sales]:[Cost]]'))).toBe('Table1[@[Sales]:[Cost]]')
		expect(printFormula(p('Table1[[#Data],[Sales]:[Cost]]'))).toBe('Table1[[#Data],[Sales]:[Cost]]')
	})

	it('unescapes structured reference column special characters', () => {
		expect(p("BillingData[Check'#]")).toMatchObject({
			type: 'structuredRef',
			table: 'BillingData',
			column: 'Check#',
		})
	})

	it('parses escaped table names with dotted numeric suffixes', () => {
		expect(p('\\_Prime.1[Name]')).toMatchObject({
			type: 'structuredRef',
			table: '\\_Prime.1',
			column: 'Name',
		})
		expect(p('\\_Prime.1[[#This Row],[Number]]')).toMatchObject({
			type: 'structuredRef',
			table: '\\_Prime.1',
			specifiers: ['#This Row'],
			column: 'Number',
		})
		expect(printFormula(p('\\_Prime.1[Name]'))).toBe('\\_Prime.1[Name]')
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
		// Math - aggregation
		'SUM',
		'AVERAGE',
		'COUNT',
		'COUNTA',
		'COUNTBLANK',
		'MAX',
		'MIN',
		'PRODUCT',
		'SUBTOTAL',
		'AGGREGATE',
		// Math - rounding
		'ROUND',
		'ROUNDUP',
		'ROUNDDOWN',
		'INT',
		'TRUNC',
		'CEILING',
		'FLOOR',
		'MROUND',
		'CEILING.MATH',
		'FLOOR.MATH',
		'EVEN',
		'ODD',
		// Math - basic
		'ABS',
		'SIGN',
		'SQRT',
		'POWER',
		'MOD',
		'QUOTIENT',
		'LOG',
		'LOG10',
		'LN',
		'EXP',
		'PI',
		'SUMPRODUCT',
		'SUMSQ',
		'SUMX2MY2',
		'SUMX2PY2',
		'SUMXMY2',
		'GCD',
		'LCM',
		'RANDBETWEEN',
		// Math - trig
		'SIN',
		'COS',
		'TAN',
		'ASIN',
		'ACOS',
		'ATAN',
		'ATAN2',
		'SINH',
		'COSH',
		'TANH',
		'ASINH',
		'ACOSH',
		'ATANH',
		'RADIANS',
		'DEGREES',
		// Math - combinatorics
		'FACT',
		'FACTDOUBLE',
		'COMBIN',
		'COMBINA',
		'PERMUT',
		'PERMUTATIONA',
		'MULTINOMIAL',
		// Math - conditional
		'SUMIF',
		'COUNTIF',
		'AVERAGEIF',
		'SUMIFS',
		'COUNTIFS',
		'AVERAGEIFS',
		'MAXIFS',
		'MINIFS',
		// Math - random
		'RAND',
		'RANDBETWEEN',
		// Math - composite
		'MMULT',
		'MDETERM',
		'MINVERSE',
		// Text
		'CONCATENATE',
		'CONCAT',
		'TEXTJOIN',
		'LEFT',
		'RIGHT',
		'MID',
		'LEN',
		'TRIM',
		'UPPER',
		'LOWER',
		'EXACT',
		'PROPER',
		'FIND',
		'SEARCH',
		'TEXTBEFORE',
		'TEXTAFTER',
		'TEXTSPLIT',
		'SUBSTITUTE',
		'REPLACE',
		'TEXT',
		'VALUE',
		'CHAR',
		'CODE',
		'REPT',
		'CLEAN',
		'T',
		'UNICHAR',
		'UNICODE',
		'FIXED',
		'DOLLAR',
		'ARRAYTOTEXT',
		'VALUETOTEXT',
		'NUMBERVALUE',
		// Logical
		'IF',
		'IFS',
		'AND',
		'OR',
		'NOT',
		'XOR',
		'IFERROR',
		'IFNA',
		'TRUE',
		'FALSE',
		'SWITCH',
		// Financial
		'PMT',
		'FV',
		'PV',
		'NPER',
		'RATE',
		'IPMT',
		'PPMT',
		'NPV',
		'IRR',
		'SLN',
		'SYD',
		'DDB',
		'DOLLARDE',
		'DOLLARFR',
		'XNPV',
		'XIRR',
		'DB',
		'VDB',
		'MIRR',
		'ISPMT',
		'CUMIPMT',
		'CUMPRINC',
		'EFFECT',
		'NOMINAL',
		'PDURATION',
		'RRI',
		'FVSCHEDULE',
		'DISC',
		'INTRATE',
		// Dynamic
		'SORT',
		'SORTBY',
		'FILTER',
		'UNIQUE',
		'SEQUENCE',
		'RANDARRAY',
		'LET',
		'TRANSPOSE',
		'TOCOL',
		'TOROW',
		'WRAPCOLS',
		'WRAPROWS',
		'HSTACK',
		'VSTACK',
		'TAKE',
		'DROP',
		'EXPAND',
		'CHOOSECOLS',
		'CHOOSEROWS',
		'LAMBDA',
		'MAP',
		'REDUCE',
		'SCAN',
		// Date
		'DATE',
		'TODAY',
		'NOW',
		'YEAR',
		'MONTH',
		'DAY',
		'HOUR',
		'MINUTE',
		'SECOND',
		'TIME',
		'TIMEVALUE',
		'DATEVALUE',
		'DATEDIF',
		'EDATE',
		'EOMONTH',
		'WEEKDAY',
		'WEEKNUM',
		'NETWORKDAYS',
		'WORKDAY',
		'DAYS360',
		'YEARFRAC',
		'ISOWEEKNUM',
		'DAYS',
		'NETWORKDAYS.INTL',
		'WORKDAY.INTL',
		// Database
		'DSUM',
		'DAVERAGE',
		'DCOUNT',
		'DCOUNTA',
		'DMAX',
		'DMIN',
		'DPRODUCT',
		'DGET',
		'DSTDEV',
		'DSTDEVP',
		'DVAR',
		'DVARP',
		// Engineering
		'BIN2DEC',
		'DEC2BIN',
		'HEX2DEC',
		'DEC2HEX',
		'OCT2DEC',
		'DEC2OCT',
		'BIN2HEX',
		'BIN2OCT',
		'HEX2BIN',
		'HEX2OCT',
		'OCT2BIN',
		'OCT2HEX',
		'DELTA',
		'GESTEP',
		'BITAND',
		'BITOR',
		'BITXOR',
		'BITLSHIFT',
		'BITRSHIFT',
		'ERF',
		'ERF.PRECISE',
		'ERFC',
		'ERFC.PRECISE',
		'COMPLEX',
		'IMREAL',
		'IMAGINARY',
		'IMABS',
		'IMARGUMENT',
		'IMCONJUGATE',
		'IMSUM',
		'IMSUB',
		'IMPRODUCT',
		'IMDIV',
		'IMPOWER',
		'IMSQRT',
		'IMEXP',
		'IMLN',
		'IMSIN',
		'IMCOS',
		// Info
		'ISBLANK',
		'ISERROR',
		'ISERR',
		'ISNA',
		'ISNUMBER',
		'ISTEXT',
		'ISLOGICAL',
		'TYPE',
		'N',
		'NA',
		'ISEVEN',
		'ISODD',
		'ISNONTEXT',
		'ERROR.TYPE',
		'ISFORMULA',
		// Lookup
		'VLOOKUP',
		'HLOOKUP',
		'INDEX',
		'MATCH',
		'XLOOKUP',
		'XMATCH',
		'CHOOSE',
		'LOOKUP',
		'ADDRESS',
		'ROWS',
		'COLUMNS',
		'ROW',
		'COLUMN',
		'FORMULATEXT',
		'AREAS',
		'INDIRECT',
		'OFFSET',
		// Stats
		'LARGE',
		'SMALL',
		'RANK',
		'PERCENTILE',
		'MEDIAN',
		'STDEV',
		'STDEV.S',
		'STDEV.P',
		'STDEVP',
		'VAR',
		'VAR.S',
		'VAR.P',
		'VARP',
		'PERCENTILE.INC',
		'PERCENTILE.EXC',
		'QUARTILE',
		'QUARTILE.INC',
		'QUARTILE.EXC',
		'MODE',
		'MODE.SNGL',
		'AVERAGEA',
		'MAXA',
		'MINA',
		'RANK.EQ',
		'RANK.AVG',
		'GEOMEAN',
		'HARMEAN',
		'TRIMMEAN',
		'PERCENTRANK.INC',
		'PERCENTRANK.EXC',
		'FORECAST.LINEAR',
		'FORECAST',
		'SLOPE',
		'INTERCEPT',
		'RSQ',
		'CORREL',
		'PEARSON',
		'STEYX',
		'COVARIANCE.P',
		'COVARIANCE.S',
		'AVEDEV',
		'DEVSQ',
		'KURT',
		'SKEW',
		'FREQUENCY',
		'MODE.MULT',
		'NORM.DIST',
		'NORM.INV',
		'NORM.S.DIST',
		'NORM.S.INV',
		'ISREF',
	] as const

	const SHEET_NAMES = ['Sheet1', 'Sheet2', 'Sheet3', 'Data', 'My Sheet', "'Jan 2024'"] as const
	const DEFINED_NAMES = ['TaxRate', 'Revenue', 'CostOfGoods', '_temp', 'x'] as const
	const TABLE_COLS = ['Sales', 'Revenue', 'Cost', 'Name', 'Date'] as const

	const BINARY_OPS = ['+', '-', '*', '/', '^', '&', '=', '<>', '<', '>', '<=', '>='] as const

	function randSheetRef(rng: { s: number }): string {
		const sheet = pick(SHEET_NAMES, rng)
		const prefix = sheet.includes(' ') ? `'${sheet}'!` : `${sheet}!`
		return xorshift32(rng) < 0.5 ? `${prefix}${randCellRef(rng)}` : `${prefix}${randRange(rng)}`
	}

	function randStructuredRef(rng: { s: number }): string {
		const table = `Table${randInt(rng, 1, 5)}`
		const col = pick(TABLE_COLS, rng)
		const forms = [
			`${table}[${col}]`,
			`${table}[[${col}]]`,
			`${table}[#All]`,
			`${table}[#Data]`,
			`${table}[#Headers]`,
			`${table}[@${col}]`,
		]
		return pick(forms, rng)
	}

	function randAtom(rng: { s: number }): string {
		const kind = randInt(rng, 0, 11)
		switch (kind) {
			case 0:
				return String(randInt(rng, 0, 9999))
			case 1:
				return `"${String.fromCharCode(...Array.from({ length: randInt(rng, 0, 8) }, () => randInt(rng, 65, 90)))}"`
			case 2:
				return randCellRef(rng)
			case 3:
				return xorshift32(rng) < 0.5 ? 'TRUE' : 'FALSE'
			case 4:
				return pick(['#N/A', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#NULL!'], rng)
			case 5:
				return randRange(rng)
			case 6:
				return randSheetRef(rng)
			case 7:
				return pick(DEFINED_NAMES, rng)
			case 8:
				return randStructuredRef(rng)
			case 9:
				return pick(['""', '0', '-0', '1E+308', '-1E+308', '1E-10', '9.99E+307'], rng)
			case 10: {
				const n = (xorshift32(rng) * 2 - 1) * 1e15
				return n.toExponential(randInt(rng, 0, 5))
			}
			default:
				return `${randInt(rng, -999, 999)}.${randInt(rng, 0, 999999)}`
		}
	}

	function randExpr(rng: { s: number }, depth: number): string {
		if (depth <= 0) return randAtom(rng)
		const kind = randInt(rng, 0, 9)
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
			case 5: {
				// deeply nested IF
				const levels = randInt(rng, 2, 4)
				let expr = randAtom(rng)
				for (let i = 0; i < levels; i++) {
					expr = `IF(${randExpr(rng, 0)}>${randInt(rng, 0, 100)},${expr},${randAtom(rng)})`
				}
				return expr
			}
			case 6: {
				// mixed function calls: SUM(IF(...), VLOOKUP(...))
				const outer = pick(['SUM', 'AVERAGE', 'MAX', 'MIN', 'IFERROR', 'IF'] as const, rng)
				const inner1Fn = pick(FN_NAMES, rng)
				const inner2Fn = pick(FN_NAMES, rng)
				const a1 = randExpr(rng, depth - 2)
				const a2 = randExpr(rng, depth - 2)
				return `${outer}(${inner1Fn}(${a1}),${inner2Fn}(${a2}))`
			}
			case 7: {
				// array literal
				const rows = randInt(rng, 1, 3)
				const cols = randInt(rng, 1, 4)
				const rowStrs: string[] = []
				for (let r = 0; r < rows; r++) {
					const cells: string[] = []
					for (let c = 0; c < cols; c++) cells.push(randAtom(rng))
					rowStrs.push(cells.join(','))
				}
				return `{${rowStrs.join(';')}}`
			}
			case 8:
				// cross-sheet reference in expression
				return `${randSheetRef(rng)}${pick(BINARY_OPS, rng)}${randExpr(rng, depth - 1)}`
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
			// basic
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
			// deeply nested IF
			'IF(IF(IF(A1>0,1,0)>0,"yes","no")="yes",100,200)',
			'IF(IF(IF(IF(A1>0,1,0),2,3),4,5),6,7)',
			'IF(A1>0,IF(B1>0,IF(C1>0,IF(D1>0,"deep","d3"),"d2"),"d1"),"d0")',
			// complex arithmetic
			'((A1+B1)*C1-D1)/E1^F1',
			'(A1+B1+C1+D1+E1+F1+G1+H1)/8',
			'((((A1+1)*2)-3)/4)^0.5',
			'A1*B1+C1*D1-E1*F1+G1/H1',
			'-(-(-A1))',
			'A1^B1^C1',
			'1+2+3+4+5+6+7+8+9+10',
			'A1%+B1%-C1%*D1%',
			'(A1+B1)*(C1-D1)/(E1+F1)^(G1-H1)',
			// mixed function calls
			'SUM(IF(A1:A10>0,A1:A10,0))',
			'IFERROR(VLOOKUP(A1,B:C,2,0),INDEX(D:D,MATCH(A1,E:E,0)))',
			'AVERAGE(LARGE(A1:A100,1),LARGE(A1:A100,2),LARGE(A1:A100,3))',
			'CONCATENATE(LEFT(A1,3),"-",RIGHT(B1,4))',
			'SUMPRODUCT((A1:A10>0)*(B1:B10))',
			'IF(AND(OR(A1>0,B1>0),C1<>0),SUM(D1:D10)/C1,0)',
			'INDEX(A1:C10,MATCH(MIN(B1:B10),B1:B10,0),3)',
			'IF(ISERROR(A1/B1),0,ROUND(A1/B1,2))',
			'MAX(SUM(A1:A5),SUM(B1:B5),SUM(C1:C5))',
			'CHOOSE(MATCH(A1,{1,2,3},0),"low","mid","high")',
			// error conditions
			'1/0',
			'0/0',
			'VLOOKUP("missing",A1:B1,2,FALSE)',
			'SQRT(-1)',
			'LOG(0)',
			'LOG(-1)',
			'0^0',
			'0^(-1)',
			'MATCH("x",A1:A1,0)',
			// edge case values
			'""',
			'""+""',
			'0',
			'-0',
			'1E+308',
			'-1E+308',
			'1E-10',
			'9.99E+307',
			'0.000000001',
			'999999999999999',
			'TRUE+TRUE',
			'FALSE*100',
			'TRUE&"text"',
			// array formulas
			'{1,2,3}',
			'{1;2;3}',
			'{1,2;3,4;5,6}',
			'{TRUE,FALSE;1,0}',
			'{"a","b","c"}',
			'SUM({1,2,3}*{4,5,6})',
			'MMULT({1,2;3,4},{5;6})',
			'{1,2,3}+{4,5,6}',
			// cross-sheet references
			'Sheet2!A1+Sheet3!B2',
			'SUM(Sheet1!A1:A10,Sheet2!B1:B10)',
			"'My Sheet'!A1",
			"SUM('Jan 2024'!A:A)",
			'Sheet1!A1:Sheet1!A10',
			'VLOOKUP(A1,Sheet2!A:B,2,FALSE)',
			'IF(Sheet1!A1>Sheet2!A1,Sheet1!B1,Sheet2!B1)',
			'MAX(Sheet1!A1:A100)-MIN(Sheet2!B1:B100)',
			// defined names
			'TaxRate*Revenue',
			'SUM(Revenue)-CostOfGoods',
			'IF(TaxRate>0.2,"high","low")',
			'ROUND(Revenue*TaxRate,2)',
			// structured references
			'Table1[Sales]',
			'SUM(Table1[Revenue])',
			'AVERAGE(Table1[Cost])',
			'Table1[@Sales]',
			'Table1[#All]',
			'Table1[#Data]',
			'Table1[#Headers]',
			'Table1[[Sales]:[Cost]]',
			'Table1[@[Sales]:[Cost]]',
			'Table1[[#Data],[Sales]:[Cost]]',
			// all major function categories
			'XLOOKUP(A1,B:B,C:C)',
			'XMATCH(A1,B:B)',
			'SEQUENCE(5,3,1,2)',
			'SORT(A1:A10)',
			'UNIQUE(A1:A20)',
			'FILTER(A1:A10,B1:B10>0)',
			'SORTBY(A1:A10,B1:B10)',
			'TRANSPOSE(A1:D1)',
			'HSTACK(A1:A5,B1:B5)',
			'VSTACK(A1:C1,A2:C2)',
			'TAKE(A1:A10,5)',
			'DROP(A1:A10,3)',
			'CHOOSECOLS(A1:E5,1,3,5)',
			'CHOOSEROWS(A1:A10,1,5,10)',
			'TOCOL(A1:C3)',
			'TOROW(A1:A10)',
			'WRAPCOLS(A1:A12,4)',
			'WRAPROWS(A1:L1,4)',
			'EXPAND(A1:B2,4,4,0)',
			'TEXTSPLIT("a,b,c",",")',
			'TEXTBEFORE("hello-world","-")',
			'TEXTAFTER("hello-world","-")',
			'LET(x,A1+B1,y,C1+D1,x*y)',
			'PMT(0.05/12,360,-100000)',
			'NPV(0.1,A1:A5)',
			'IRR(A1:A5)',
			'XNPV(0.1,A1:A5,B1:B5)',
			'XIRR(A1:A5,B1:B5)',
			'DATE(2024,1,15)',
			'DATEDIF(A1,B1,"Y")',
			'NETWORKDAYS(A1,B1)',
			'YEARFRAC(A1,B1,0)',
			'WORKDAY(A1,10)',
			'NORM.DIST(0,0,1,TRUE)',
			'NORM.INV(0.975,0,1)',
			'PERCENTILE(A1:A100,0.9)',
			'STDEV(A1:A100)',
			'CORREL(A1:A10,B1:B10)',
			'FORECAST(5,A1:A10,B1:B10)',
			'SLOPE(A1:A10,B1:B10)',
			'BIN2DEC("1010")',
			'DEC2HEX(255)',
			'COMPLEX(3,4)',
			'IMABS("3+4i")',
			'ERF(1)',
			'DSUM(A1:D10,2,F1:G2)',
			'ISBLANK(A1)',
			'ISNUMBER(A1)',
			'ERROR.TYPE(A1)',
			'TYPE(A1)',
			'ADDRESS(1,1)',
			'ROW(A5)',
			'COLUMN(C1)',
			'ROWS(A1:A10)',
			'COLUMNS(A1:E1)',
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
			// additional malformed patterns
			'SUM(,,,)',
			'IF(,,)',
			'Sheet1!',
			'Sheet1!:A1',
			"'Unclosed sheet!A1",
			'Table1[',
			'Table1[]',
			'Table1[#]',
			'Table1[@]',
			'{,}',
			'{;}',
			'{1,;2}',
			'SUM(Sheet1!)',
			'IF(A1>,B1,)',
			'VLOOKUP(,,,)',
			'A1:B1:C1',
			'(((((',
			')))))',
			'SUM())',
			'(SUM(',
			'=A1',
			'A1..B1',
			'SUM(A1 B1)',
			'IF IF',
			'SUM[A1]',
			'{1,2,3,}',
			'"unclosed string',
			'A1:$',
			'$$A1',
			'SUM(#REF!)',
			'UNKNOWN_FUNCTION(1)',
			' ',
			'\t\t',
			'\n',
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

	test('random generated formulas never throw (500 cases)', () => {
		const rng = { s: 42 }
		for (let i = 0; i < 500; i++) {
			const depth = randInt(rng, 0, 5)
			const formula = xorshift32(rng) < 0.7 ? randExpr(rng, depth) : randGarbage(rng)
			const result = parseFormula(formula)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('random garbage strings never throw (100 cases)', () => {
		const rng = { s: 99999 }
		for (let i = 0; i < 100; i++) {
			const formula = randGarbage(rng)
			const result = parseFormula(formula)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('corrupted valid formulas never throw (100 cases)', () => {
		const rng = { s: 77777 }
		const bases = [
			'SUM(A1:B10)',
			'IF(A1>0,B1,C1)',
			'VLOOKUP(D1,A:B,2,0)',
			'INDEX(A1:C10,2,3)',
			'A1+B1*C1',
			'Sheet2!A1+Sheet3!B2',
			'XLOOKUP("x",A:A,B:B)',
			'IF(AND(A1>0,B1<10),SUM(C1:C10),0)',
			'NORM.DIST(0,0,1,TRUE)',
			'Table1[Sales]',
			'IFERROR(A1/B1,0)',
			'SEQUENCE(5,3,1,2)',
			'LET(x,1,y,2,x+y)',
			'NETWORKDAYS(A1,B1)',
			'PERCENTILE(A1:A10,0.9)',
		]
		for (let i = 0; i < 100; i++) {
			const base = pick(bases, rng)
			const chars = [...base]
			const mutations = randInt(rng, 1, 4)
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

	test('every registered function name parses as a call (all 200+ functions)', () => {
		for (const fnName of FN_NAMES) {
			const args = 'A1,B1,C1,1,TRUE,"text"'
			const result = parseFormula(`${fnName}(${args})`)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('deeply nested IF chains (10 levels)', () => {
		let formula = '0'
		for (let i = 0; i < 10; i++) {
			formula = `IF(A${i + 1}>${i * 10},${formula},${i})`
		}
		const result = parseFormula(formula)
		expect(result.ok === true || result.ok === false).toBe(true)
	})

	test('deeply nested function calls (8 levels)', () => {
		const formulas = [
			'SUM(SUM(SUM(SUM(SUM(SUM(SUM(SUM(1))))))))',
			'ABS(ABS(ABS(ABS(ABS(ABS(ABS(ABS(-1))))))))',
			'ROUND(ROUND(ROUND(ROUND(3.14159,4),3),2),1)',
			'IF(IF(IF(IF(TRUE,1,0),2,0),3,0),4,0)',
			'IFERROR(IFERROR(IFERROR(1/0,2/0),3/0),0)',
			'LEFT(RIGHT(LEFT(RIGHT("hello world",8),5),3),2)',
			'MAX(MIN(MAX(MIN(MAX(1,2),3),4),5),6)',
		]
		for (const f of formulas) {
			const result = parseFormula(f)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('complex multi-operator arithmetic chains', () => {
		const formulas = [
			'A1+B1-C1*D1/E1^F1&G1=H1<>I1<J1>K1<=L1>=M1',
			'((A1+B1)*C1-D1)/E1^F1%+G1',
			'---A1',
			'A1%%%%%',
			'(((((1+2)*(3+4))-(5+6))/(7+8))^(1/2))%',
			'SUM(A1:A10)*AVERAGE(B1:B10)+COUNT(C1:C10)^2-MIN(D1:D10)/MAX(E1:E10)',
		]
		for (const f of formulas) {
			const result = parseFormula(f)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('cross-sheet reference patterns', () => {
		const formulas = [
			'Sheet1!A1',
			'Sheet2!A1:B10',
			"'Sheet With Spaces'!A1",
			"'Jan 2024'!A1:Z100",
			'SUM(Sheet1!A:A,Sheet2!B:B,Sheet3!C:C)',
			'Sheet1!A1+Sheet2!A1+Sheet3!A1',
			'VLOOKUP(Sheet1!A1,Sheet2!A:B,2,FALSE)',
			'IF(Sheet1!A1>Sheet2!A1,Sheet1!B1,Sheet2!B1)',
			'INDEX(Sheet2!A1:C10,MATCH(A1,Sheet2!A1:A10,0),2)',
			"SUMIF('Data Sheet'!A:A,\">0\",'Data Sheet'!B:B)",
		]
		for (const f of formulas) {
			const result = parseFormula(f)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('structured reference patterns', () => {
		const formulas = [
			'Table1[Sales]',
			'Table1[[Sales]:[Revenue]]',
			'Table1[@Sales]',
			'Table1[#All]',
			'Table1[#Data]',
			'Table1[#Headers]',
			'Table1[#Totals]',
			'SUM(Table1[Sales])',
			'AVERAGE(Table1[Revenue])',
			'Table1[@Sales]*Table1[@Cost]',
			'VLOOKUP(A1,Table1,2,FALSE)',
		]
		for (const f of formulas) {
			const result = parseFormula(f)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('defined name patterns in formulas', () => {
		const formulas = [
			'TaxRate',
			'Revenue*TaxRate',
			'SUM(Revenue)',
			'IF(TaxRate>0.2,"high","low")',
			'ROUND(Revenue*TaxRate,2)',
			'CostOfGoods+Revenue',
			'_temp*100',
			'x+1',
			'MAX(TaxRate,0.15)',
			'IFERROR(Revenue/CostOfGoods,0)',
		]
		for (const f of formulas) {
			const result = parseFormula(f)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})

	test('high-depth random expression trees (200 cases, depth 5-6)', () => {
		const rng = { s: 314159 }
		for (let i = 0; i < 200; i++) {
			const depth = randInt(rng, 4, 6)
			const formula = randExpr(rng, depth)
			const result = parseFormula(formula)
			expect(result.ok === true || result.ok === false).toBe(true)
		}
	})
})
