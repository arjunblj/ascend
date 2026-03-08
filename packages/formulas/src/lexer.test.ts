import { describe, expect, it } from 'bun:test'
import { tokenize } from './lexer.ts'
import { TokenType } from './tokens.ts'

function values(formula: string) {
	return tokenize(formula)
		.filter((t) => t.type !== TokenType.Whitespace && t.type !== TokenType.EOF)
		.map((t) => t.value)
}

describe('tokenize', () => {
	it('tokenizes integers', () => {
		const tokens = tokenize('42')
		expect(tokens[0]?.type).toBe(TokenType.Number)
		expect(tokens[0]?.value).toBe('42')
	})

	it('tokenizes decimals and scientific notation', () => {
		expect(tokenize('3.14')[0]?.value).toBe('3.14')
		expect(tokenize('.5')[0]?.value).toBe('.5')
		expect(tokenize('1.5E+3')[0]?.value).toBe('1.5E+3')
		expect(tokenize('2e-10')[0]?.value).toBe('2e-10')
	})

	it('tokenizes string literals with escaped quotes', () => {
		const tokens = tokenize('"hello"')
		expect(tokens[0]?.type).toBe(TokenType.String)
		expect(tokens[0]?.value).toBe('hello')

		const escaped = tokenize('"say ""hi"""')
		expect(escaped[0]?.value).toBe('say "hi"')
	})

	it('tokenizes boolean values', () => {
		expect(tokenize('TRUE')[0]?.type).toBe(TokenType.Boolean)
		expect(tokenize('TRUE')[0]?.value).toBe('TRUE')
		expect(tokenize('FALSE')[0]?.type).toBe(TokenType.Boolean)
		expect(tokenize('false')[0]?.value).toBe('FALSE')
	})

	it('tokenizes simple cell references', () => {
		const tokens = tokenize('A1')
		expect(tokens[0]?.type).toBe(TokenType.CellRef)
		expect(tokens[0]?.value).toBe('A1')
	})

	it('tokenizes absolute cell references', () => {
		for (const ref of ['$A$1', '$A1', 'A$1', '$AB$100']) {
			expect(tokenize(ref)[0]?.type).toBe(TokenType.CellRef)
			expect(tokenize(ref)[0]?.value).toBe(ref)
		}
	})

	it('tokenizes sheet names followed by bang as Name', () => {
		const tokens = tokenize('Sheet1!A1')
		expect(tokens[0]?.type).toBe(TokenType.Name)
		expect(tokens[0]?.value).toBe('Sheet1')
		expect(tokens[1]?.type).toBe(TokenType.Bang)
		expect(tokens[2]?.type).toBe(TokenType.CellRef)
		expect(tokens[2]?.value).toBe('A1')
	})

	it('tokenizes quoted sheet names', () => {
		const tokens = tokenize("'My Sheet'!A1")
		expect(tokens[0]?.type).toBe(TokenType.Name)
		expect(tokens[0]?.value).toBe('My Sheet')
		expect(tokens[1]?.type).toBe(TokenType.Bang)
		expect(tokens[2]?.type).toBe(TokenType.CellRef)
	})

	it('tokenizes single-char operators', () => {
		for (const op of ['+', '-', '*', '/', '^', '&', '=', '%']) {
			const tokens = tokenize(op)
			expect(tokens[0]?.type).toBe(TokenType.Operator)
			expect(tokens[0]?.value).toBe(op)
		}
	})

	it('tokenizes multi-char operators', () => {
		for (const op of ['<>', '<=', '>=']) {
			const tokens = tokenize(op)
			expect(tokens[0]?.type).toBe(TokenType.Operator)
			expect(tokens[0]?.value).toBe(op)
		}
		expect(tokenize('<')[0]?.value).toBe('<')
		expect(tokenize('>')[0]?.value).toBe('>')
	})

	it('detects function names when followed by open paren', () => {
		const tokens = tokenize('SUM(')
		expect(tokens[0]?.type).toBe(TokenType.Function)
		expect(tokens[0]?.value).toBe('SUM')
		expect(tokens[1]?.type).toBe(TokenType.OpenParen)
	})

	it('tokenizes error values', () => {
		const errors = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']
		for (const e of errors) {
			const tokens = tokenize(e)
			expect(tokens[0]?.type).toBe(TokenType.Error)
			expect(tokens[0]?.value).toBe(e)
		}
	})

	it('preserves whitespace tokens', () => {
		const tokens = tokenize('A1 B1')
		expect(tokens[0]?.type).toBe(TokenType.CellRef)
		expect(tokens[1]?.type).toBe(TokenType.Whitespace)
		expect(tokens[2]?.type).toBe(TokenType.CellRef)
	})

	it('tokenizes punctuation', () => {
		expect(tokenize('(')[0]?.type).toBe(TokenType.OpenParen)
		expect(tokenize(')')[0]?.type).toBe(TokenType.CloseParen)
		expect(tokenize('{')[0]?.type).toBe(TokenType.OpenBrace)
		expect(tokenize('}')[0]?.type).toBe(TokenType.CloseBrace)
		expect(tokenize(',')[0]?.type).toBe(TokenType.Comma)
		expect(tokenize(';')[0]?.type).toBe(TokenType.Semicolon)
		expect(tokenize(':')[0]?.type).toBe(TokenType.Colon)
		expect(tokenize('!')[0]?.type).toBe(TokenType.Bang)
	})

	it('tokenizes array literal brackets', () => {
		const v = values('{1,2;3,4}')
		expect(v).toEqual(['{', '1', ',', '2', ';', '3', ',', '4', '}'])
	})

	it('tokenizes a complex formula', () => {
		const v = values('IF(A1>0,A1*2,"neg")')
		expect(v).toEqual(['IF', '(', 'A1', '>', '0', ',', 'A1', '*', '2', ',', 'neg', ')'])
	})

	it('tokenizes named ranges', () => {
		const tokens = tokenize('MyRange')
		expect(tokens[0]?.type).toBe(TokenType.Name)
		expect(tokens[0]?.value).toBe('MyRange')
	})

	it('ends with EOF token', () => {
		const tokens = tokenize('')
		expect(tokens).toHaveLength(1)
		expect(tokens[0]?.type).toBe(TokenType.EOF)
	})
})
