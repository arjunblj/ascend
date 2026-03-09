import { columnToIndex } from '@ascend/core'
import { ascendError, type ExcelError, err, ok, type Result } from '@ascend/schema'
import type {
	BinaryOp,
	CellRefNode,
	FormulaCellRef,
	FormulaNode,
	RangeRefNode,
	StructuredRefNode,
	UnaryOp,
} from './ast.ts'
import { tokenize } from './lexer.ts'
import { type Token, TokenType } from './tokens.ts'

const EOF_TOKEN: Token = { type: TokenType.EOF, value: '', position: -1 }

function parseCellRefValue(raw: string): FormulaCellRef {
	let i = 0
	const colAbsolute = raw.charAt(i) === '$'
	if (colAbsolute) i++

	let colStr = ''
	while (i < raw.length && /[A-Za-z]/.test(raw.charAt(i))) {
		colStr += raw.charAt(i)
		i++
	}

	const rowAbsolute = raw.charAt(i) === '$'
	if (rowAbsolute) i++

	const rowStr = raw.slice(i)
	return {
		row: Number.parseInt(rowStr, 10) - 1,
		col: columnToIndex(colStr.toUpperCase()),
		rowAbsolute,
		colAbsolute,
	}
}

class FormulaParser {
	private readonly tokens: readonly Token[]
	private pos = 0

	constructor(tokens: readonly Token[]) {
		this.tokens = tokens.filter((t) => t.type !== TokenType.Whitespace)
	}

	private peek(): Token {
		return this.tokens[this.pos] ?? EOF_TOKEN
	}

	private advance(): Token {
		const token = this.peek()
		if (token.type !== TokenType.EOF) this.pos++
		return token
	}

	private expect(type: TokenType): Token {
		const token = this.peek()
		if (token.type !== type) {
			throw new Error(
				`Expected ${type}, got ${token.type} "${token.value}" at position ${token.position}`,
			)
		}
		return this.advance()
	}

	private isOp(...ops: string[]): boolean {
		const t = this.peek()
		return t.type === TokenType.Operator && ops.includes(t.value)
	}

	parse(): FormulaNode {
		const node = this.parseExpression()
		if (this.peek().type !== TokenType.EOF) {
			const leftover = this.peek()
			throw new Error(
				`Unexpected token ${leftover.type} "${leftover.value}" at position ${leftover.position}`,
			)
		}
		return node
	}

	private parseExpression(): FormulaNode {
		return this.parseComparison()
	}

	private parseComparison(): FormulaNode {
		let left = this.parseConcatenation()
		while (this.isOp('=', '<>', '<', '>', '<=', '>=')) {
			const op = this.advance().value as BinaryOp
			const right = this.parseConcatenation()
			left = { type: 'binary', op, left, right }
		}
		return left
	}

	private parseConcatenation(): FormulaNode {
		let left = this.parseAddition()
		while (this.isOp('&')) {
			this.advance()
			const right = this.parseAddition()
			left = { type: 'binary', op: '&', left, right }
		}
		return left
	}

	private parseAddition(): FormulaNode {
		let left = this.parseMultiplication()
		while (this.isOp('+', '-')) {
			const op = this.advance().value as BinaryOp
			const right = this.parseMultiplication()
			left = { type: 'binary', op, left, right }
		}
		return left
	}

	private parseMultiplication(): FormulaNode {
		let left = this.parseUnaryPrefix()
		while (this.isOp('*', '/')) {
			const op = this.advance().value as BinaryOp
			const right = this.parseUnaryPrefix()
			left = { type: 'binary', op, left, right }
		}
		return left
	}

	private parseExponentiation(): FormulaNode {
		const left = this.parsePostfix()
		if (this.isOp('^')) {
			this.advance()
			const right = this.parseUnaryPrefix()
			return { type: 'binary', op: '^', left, right }
		}
		return left
	}

	private parseUnaryPrefix(): FormulaNode {
		if (this.isOp('+', '-', '@')) {
			const op = this.advance().value as UnaryOp
			const operand = this.parseUnaryPrefix()
			return { type: 'unary', op, operand }
		}
		return this.parseExponentiation()
	}

	private parsePostfix(): FormulaNode {
		let node = this.parseAtom()
		while (this.isOp('%', '#')) {
			const op = this.advance().value
			if (op === '%') {
				node = { type: 'unary', op: '%' as const, operand: node }
			} else {
				node = { type: 'spillRef', target: node }
			}
		}
		return node
	}

	private parseAtom(): FormulaNode {
		const token = this.peek()

		if (token.type === TokenType.Number) {
			this.advance()
			return { type: 'number', value: Number.parseFloat(token.value) }
		}

		if (token.type === TokenType.String) {
			this.advance()
			return { type: 'string', value: token.value }
		}

		if (token.type === TokenType.Boolean) {
			this.advance()
			return { type: 'boolean', value: token.value === 'TRUE' }
		}

		if (token.type === TokenType.Error) {
			this.advance()
			return { type: 'error', value: token.value as ExcelError }
		}

		if (token.type === TokenType.Function) {
			return this.parseFunctionCall()
		}

		if (token.type === TokenType.CellRef) {
			return this.parseCellOrRange()
		}

		if (token.type === TokenType.Name) {
			if (token.value.startsWith('[')) {
				this.advance()
				return this.buildStructuredRef('', token.value)
			}
			return this.parseNameOrSheetRef()
		}

		if (token.type === TokenType.OpenParen) {
			this.advance()
			const expr = this.parseExpression()
			this.expect(TokenType.CloseParen)
			return expr
		}

		if (token.type === TokenType.OpenBrace) {
			return this.parseArrayLiteral()
		}

		throw new Error(`Unexpected token ${token.type} "${token.value}" at position ${token.position}`)
	}

	private parseFunctionCall(): FormulaNode {
		const name = this.advance().value
		this.expect(TokenType.OpenParen)
		const args: FormulaNode[] = []

		if (this.peek().type !== TokenType.CloseParen) {
			args.push(this.parseArgOrMissing())
			while (this.peek().type === TokenType.Comma) {
				this.advance()
				args.push(this.parseArgOrMissing())
			}
		}

		this.expect(TokenType.CloseParen)
		return { type: 'function', name, args }
	}

	private parseArgOrMissing(): FormulaNode {
		const next = this.peek()
		if (next.type === TokenType.Comma || next.type === TokenType.CloseParen) {
			return { type: 'missing' }
		}
		return this.parseExpression()
	}

	private parseCellOrRange(sheet?: string): CellRefNode | RangeRefNode {
		const token = this.advance()
		const ref = parseCellRefValue(token.value)

		if (this.peek().type === TokenType.Colon) {
			this.advance()
			const endToken = this.expect(TokenType.CellRef)
			const end = parseCellRefValue(endToken.value)
			if (sheet !== undefined) {
				return { type: 'rangeRef', start: ref, end, sheet }
			}
			return { type: 'rangeRef', start: ref, end }
		}

		if (sheet !== undefined) {
			return { type: 'cellRef', ref, sheet }
		}
		return { type: 'cellRef', ref }
	}

	private parseNameOrSheetRef(): FormulaNode {
		const token = this.advance()

		if (this.peek().type === TokenType.Bang) {
			this.advance()
			const sheet = token.value
			if (this.peek().type === TokenType.CellRef) {
				return this.parseCellOrRange(sheet)
			}
			if (this.peek().type === TokenType.Name) {
				const name = this.advance().value
				return { type: 'name', name, sheet }
			}
			throw new Error(
				`Expected cell reference or name after "${sheet}!" at position ${this.peek().position}`,
			)
		}

		if (this.peek().type === TokenType.Name && this.peek().value.startsWith('[')) {
			const bracketToken = this.advance()
			return this.buildStructuredRef(token.value, bracketToken.value)
		}

		return { type: 'name', name: token.value }
	}

	private buildStructuredRef(table: string, bracket: string): StructuredRefNode {
		const content = bracket.slice(1, -1)
		const specifiers: string[] = []
		let column: string | undefined

		if (content.startsWith('@')) {
			specifiers.push('@')
			const col = content.slice(1)
			if (col.length > 0) column = col
		} else if (content.startsWith('#')) {
			specifiers.push(content)
		} else if (content.startsWith('[')) {
			const parts = content.split('],[')
			for (const part of parts) {
				const cleaned = part.replace(/^\[|]$/g, '')
				if (cleaned.startsWith('#')) {
					specifiers.push(cleaned)
				} else {
					column = cleaned
				}
			}
		} else {
			column = content
		}

		if (column !== undefined) {
			return { type: 'structuredRef', table, specifiers, column }
		}
		return { type: 'structuredRef', table, specifiers }
	}

	private parseArrayLiteral(): FormulaNode {
		this.expect(TokenType.OpenBrace)
		const rows: FormulaNode[][] = []
		let currentRow: FormulaNode[] = [this.parseExpression()]

		while (this.peek().type !== TokenType.CloseBrace) {
			if (this.peek().type === TokenType.Comma) {
				this.advance()
				currentRow.push(this.parseExpression())
			} else if (this.peek().type === TokenType.Semicolon) {
				this.advance()
				rows.push(currentRow)
				currentRow = [this.parseExpression()]
			} else {
				throw new Error(`Expected , or ; in array literal at position ${this.peek().position}`)
			}
		}

		rows.push(currentRow)
		this.expect(TokenType.CloseBrace)
		return { type: 'array', rows }
	}
}

export function parse(tokens: readonly Token[]): Result<FormulaNode> {
	try {
		const parser = new FormulaParser(tokens)
		return ok(parser.parse())
	} catch (e) {
		return err(ascendError('FORMULA_PARSE_ERROR', e instanceof Error ? e.message : String(e)))
	}
}

export function parseFormula(formula: string): Result<FormulaNode> {
	return parse(tokenize(formula))
}
