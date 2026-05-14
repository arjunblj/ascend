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
	const colAbsolute = raw[i] === '$'
	if (colAbsolute) i++

	const colStart = i
	while (i < raw.length) {
		const ch = raw.charCodeAt(i)
		if ((ch >= 65 && ch <= 90) || (ch >= 97 && ch <= 122)) i++
		else break
	}

	const rowAbsolute = raw[i] === '$'
	if (rowAbsolute) i++

	return {
		row: Number.parseInt(raw.slice(i), 10) - 1,
		col: columnToIndex(raw.slice(colStart, i - (rowAbsolute ? 1 : 0))),
		rowAbsolute,
		colAbsolute,
	}
}

const COL_LABEL_RE = /^\$?[A-Za-z]{1,3}$/
function isColumnLabel(raw: string): boolean {
	return COL_LABEL_RE.test(raw)
}

function columnLabelToIndex(raw: string): number {
	return columnToIndex((raw.startsWith('$') ? raw.slice(1) : raw).toUpperCase())
}

function wholeColumnRangeNode(startRaw: string, endRaw: string): FormulaNode {
	return {
		type: 'wholeColumnRange',
		startCol: columnLabelToIndex(startRaw),
		endCol: columnLabelToIndex(endRaw),
		...(startRaw.startsWith('$') ? { startColAbsolute: true } : {}),
		...(endRaw.startsWith('$') ? { endColAbsolute: true } : {}),
	}
}

function isStructuredRefEscapedChar(ch: string | undefined): boolean {
	return ch === '[' || ch === ']' || ch === '#' || ch === "'" || ch === '@'
}

function stripOuterStructuredRefBrackets(text: string): string {
	return text.startsWith('[') && text.endsWith(']') ? text.slice(1, -1) : text
}

function hasOuterStructuredRefBrackets(text: string): boolean {
	return text.startsWith('[') && text.endsWith(']')
}

function splitStructuredRefParts(content: string): string[] {
	const parts: string[] = []
	let start = 0
	let depth = 0
	for (let i = 0; i < content.length; i++) {
		const ch = content[i]
		if (ch === "'" && isStructuredRefEscapedChar(content[i + 1])) {
			i++
		} else if (ch === '[') {
			depth++
		} else if (ch === ']') {
			depth--
		} else if (ch === ',' && depth === 0) {
			parts.push(content.slice(start, i))
			start = i + 1
		}
	}
	parts.push(content.slice(start))
	return parts
}

function findStructuredRefTopLevelColon(content: string): number {
	let depth = 0
	for (let i = 0; i < content.length; i++) {
		const ch = content[i]
		if (ch === "'" && isStructuredRefEscapedChar(content[i + 1])) {
			i++
		} else if (ch === '[') {
			depth++
		} else if (ch === ']') {
			depth--
		} else if (ch === ':' && depth === 0) {
			return i
		}
	}
	return -1
}

function unescapeStructuredRefColumn(name: string): string {
	return name.replace(/'([#@[\]'])/g, '$1')
}

class FormulaParser {
	private readonly tokens: readonly Token[]
	private pos = 0

	constructor(tokens: readonly Token[]) {
		this.tokens = tokens
	}

	private peek(skipWhitespace = false): Token {
		if (!skipWhitespace) return this.tokens[this.pos] ?? EOF_TOKEN
		return this.lookahead(0, true)
	}

	private advance(skipWhitespace = false): Token {
		if (skipWhitespace) this.skipWhitespace()
		const token = this.peek()
		if (token.type !== TokenType.EOF) this.pos++
		return token
	}

	private expect(type: TokenType, skipWhitespace = true): Token {
		const token = this.peek(skipWhitespace)
		if (token.type !== type) {
			throw new Error(
				`Expected ${type}, got ${token.type} "${token.value}" at position ${token.position}`,
			)
		}
		return this.advance(skipWhitespace)
	}

	private isOp(...ops: string[]): boolean {
		const t = this.peek(true)
		return t.type === TokenType.Operator && ops.includes(t.value)
	}

	private skipWhitespace(): void {
		while (this.peek().type === TokenType.Whitespace) this.pos++
	}

	private lookahead(offset: number, skipWhitespace = false): Token {
		let index = this.pos
		let remaining = offset
		while (true) {
			const token = this.tokens[index] ?? EOF_TOKEN
			if (!skipWhitespace || token.type !== TokenType.Whitespace) {
				if (remaining === 0) return token
				remaining--
			}
			if (token.type === TokenType.EOF) return EOF_TOKEN
			index++
		}
	}

	parseOrThrow(): FormulaNode {
		const node = this.parseExpression()
		this.skipWhitespace()
		if (this.peek().type !== TokenType.EOF) {
			const leftover = this.peek()
			throw new Error(
				`Unexpected token ${leftover.type} "${leftover.value}" at position ${leftover.position}`,
			)
		}
		return node
	}

	private parseExpression(allowUnionComma = true): FormulaNode {
		return this.parseComparison(allowUnionComma)
	}

	private parseComparison(allowUnionComma: boolean): FormulaNode {
		let left = this.parseConcatenation(allowUnionComma)
		while (this.isOp('=', '<>', '<', '>', '<=', '>=')) {
			const op = this.advance(true).value as BinaryOp
			const right = this.parseConcatenation(allowUnionComma)
			left = { type: 'binary', op, left, right }
		}
		return left
	}

	private parseConcatenation(allowUnionComma: boolean): FormulaNode {
		let left = this.parseAddition(allowUnionComma)
		while (this.isOp('&')) {
			this.advance(true)
			const right = this.parseAddition(allowUnionComma)
			left = { type: 'binary', op: '&', left, right }
		}
		return left
	}

	private parseAddition(allowUnionComma: boolean): FormulaNode {
		let left = this.parseMultiplication(allowUnionComma)
		while (this.isOp('+', '-')) {
			const op = this.advance(true).value as BinaryOp
			const right = this.parseMultiplication(allowUnionComma)
			left = { type: 'binary', op, left, right }
		}
		return left
	}

	private parseMultiplication(allowUnionComma: boolean): FormulaNode {
		let left = this.parseExponentiation(allowUnionComma)
		while (this.isOp('*', '/')) {
			const op = this.advance(true).value as BinaryOp
			const right = this.parseExponentiation(allowUnionComma)
			left = { type: 'binary', op, left, right }
		}
		return left
	}

	private parseExponentiation(allowUnionComma: boolean): FormulaNode {
		const left = this.parseUnaryPrefix(allowUnionComma)
		if (this.isOp('^')) {
			this.advance(true)
			const right = this.parseExponentiation(allowUnionComma)
			return { type: 'binary', op: '^', left, right }
		}
		return left
	}

	private parseUnaryPrefix(allowUnionComma: boolean): FormulaNode {
		if (this.isOp('+', '-', '@')) {
			const op = this.advance(true).value as UnaryOp
			const operand = this.parseUnaryPrefix(allowUnionComma)
			return { type: 'unary', op, operand }
		}
		return this.parseReferenceUnion(allowUnionComma)
	}

	private parseReferenceUnion(allowUnionComma: boolean): FormulaNode {
		let left = this.parseReferenceIntersection()
		while (allowUnionComma) {
			this.skipWhitespace()
			if (this.peek().type !== TokenType.Comma || !isReferenceLike(left)) break
			this.advance()
			const right = this.parseReferenceIntersection()
			if (!isReferenceLike(right)) {
				throw new Error(
					`Union operator requires references at position ${this.peek(true).position}`,
				)
			}
			left = { type: 'binary', op: ',', left, right }
		}
		return left
	}

	private parseReferenceIntersection(): FormulaNode {
		let left = this.parseReferenceRange()
		while (isReferenceLike(left) && this.peek().type === TokenType.Whitespace) {
			const savedPos = this.pos
			this.advance()
			if (!canStartReferenceExpression(this.peek(true))) {
				this.pos = savedPos
				break
			}
			const right = this.parseReferenceRange()
			if (!isReferenceLike(right)) {
				throw new Error(
					`Intersection operator requires references at position ${this.peek(true).position}`,
				)
			}
			left = { type: 'binary', op: ' ', left, right }
		}
		return left
	}

	private parseReferenceRange(): FormulaNode {
		let left = this.parsePostfix()
		while (isReferenceLike(left) && this.peek(true).type === TokenType.Colon) {
			this.expect(TokenType.Colon)
			const right = this.parsePostfix()
			if (!isReferenceLike(right)) {
				throw new Error(
					`Range operator requires references at position ${this.peek(true).position}`,
				)
			}
			left = makeRangeFromEndpoints(left, right)
		}
		return left
	}

	private parsePostfix(): FormulaNode {
		let node = this.parseAtom()
		while (true) {
			if (this.peek(true).type === TokenType.OpenParen) {
				node = this.parsePostfixCall(node)
				continue
			}
			if (!this.isOp('%', '#')) break
			const op = this.advance(true).value
			if (op === '%') node = { type: 'unary', op: '%' as const, operand: node }
			else node = { type: 'spillRef', target: node }
		}
		return node
	}

	private parseAtom(): FormulaNode {
		const token = this.peek(true)

		if (token.type === TokenType.Number) {
			if (
				this.lookahead(1, true).type === TokenType.Colon &&
				this.lookahead(2, true).type === TokenType.Number
			) {
				const start = Number.parseInt(this.advance(true).value, 10) - 1
				this.expect(TokenType.Colon)
				const end = Number.parseInt(this.expect(TokenType.Number).value, 10) - 1
				return { type: 'wholeRowRange', startRow: start, endRow: end }
			}
			this.advance(true)
			return { type: 'number', value: Number.parseFloat(token.value) }
		}

		if (token.type === TokenType.String) {
			this.advance(true)
			return { type: 'string', value: token.value }
		}

		if (token.type === TokenType.Boolean) {
			this.advance(true)
			return { type: 'boolean', value: token.value === 'TRUE' }
		}

		if (token.type === TokenType.Error) {
			this.advance(true)
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
				if (
					isSheetNameSegmentToken(this.lookahead(1, true)) &&
					this.lookahead(2, true).type === TokenType.Colon &&
					isSheetNameSegmentToken(this.lookahead(3, true)) &&
					this.lookahead(4, true).type === TokenType.Bang
				) {
					const workbookToken = this.advance(true).value
					const startSheet = this.advance(true).value
					this.expect(TokenType.Colon)
					const endSheet = this.advance(true).value
					this.expect(TokenType.Bang)
					return this.parseSheetSpanQualifiedReference(`${workbookToken}${startSheet}`, endSheet)
				}
				if (
					isSheetNameSegmentToken(this.lookahead(1, true)) &&
					this.lookahead(2, true).type === TokenType.Bang
				) {
					const workbookToken = this.advance(true).value
					const sheetToken = this.advance(true).value
					this.expect(TokenType.Bang)
					return this.parseSheetQualifiedReference(`${workbookToken}${sheetToken}`)
				}
				if (this.lookahead(1, true).type === TokenType.Bang) {
					const workbookToken = this.advance(true).value
					this.expect(TokenType.Bang)
					const spanToken = splitSheetSpanToken(workbookToken)
					if (spanToken) {
						if (!spanToken.startSheet || !spanToken.endSheet) {
							throw new Error(
								`Invalid 3D sheet span "${workbookToken}" at position ${token.position}`,
							)
						}
						return this.parseSheetSpanQualifiedReference(spanToken.startSheet, spanToken.endSheet)
					}
					return this.parseSheetQualifiedReference(workbookToken)
				}
				this.advance(true)
				return this.buildStructuredRef('', token.value)
			}
			if (
				isColumnLabel(token.value) &&
				this.lookahead(1, true).type === TokenType.Colon &&
				this.lookahead(2, true).type === TokenType.Name &&
				isColumnLabel(this.lookahead(2, true).value)
			) {
				const startCol = this.advance(true).value
				this.expect(TokenType.Colon)
				const endCol = this.expect(TokenType.Name).value
				return wholeColumnRangeNode(startCol, endCol)
			}
			return this.parseNameOrSheetRef()
		}

		if (token.type === TokenType.OpenParen) {
			this.advance(true)
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
		const name = this.advance(true).value
		const args = this.parseCallArgs()
		return { type: 'function', name, args }
	}

	private parsePostfixCall(callee: FormulaNode): FormulaNode {
		const args = this.parseCallArgs()
		return { type: 'function', name: '__CALL__', args: [callee, ...args] }
	}

	private parseCallArgs(): FormulaNode[] {
		this.expect(TokenType.OpenParen)
		const args: FormulaNode[] = []

		if (this.peek(true).type !== TokenType.CloseParen) {
			args.push(this.parseArgOrMissing())
			while (this.peek(true).type === TokenType.Comma) {
				this.advance(true)
				args.push(this.parseArgOrMissing())
			}
		}

		this.expect(TokenType.CloseParen)
		return args
	}

	private parseArgOrMissing(): FormulaNode {
		const next = this.peek(true)
		if (next.type === TokenType.Comma || next.type === TokenType.CloseParen) {
			return { type: 'missing' }
		}
		return this.parseExpression(false)
	}

	private parseCellOrRange(sheet?: string): CellRefNode | RangeRefNode {
		const token = this.advance(true)
		const ref = parseCellRefValue(token.value)

		if (
			this.peek(true).type === TokenType.Colon &&
			this.lookahead(1, true).type === TokenType.CellRef
		) {
			this.expect(TokenType.Colon)
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
		const token = this.advance(true)
		const spanToken = splitSheetSpanToken(token.value)
		if (spanToken && this.peek(true).type === TokenType.Bang) {
			if (!spanToken.startSheet || !spanToken.endSheet) {
				throw new Error(`Invalid 3D sheet span "${token.value}" at position ${token.position}`)
			}
			this.expect(TokenType.Bang)
			return this.parseSheetSpanQualifiedReference(spanToken.startSheet, spanToken.endSheet)
		}

		if (
			this.peek(true).type === TokenType.Colon &&
			this.lookahead(1, true).type === TokenType.Name &&
			this.lookahead(2, true).type === TokenType.Bang
		) {
			this.expect(TokenType.Colon)
			const endSheet = this.expect(TokenType.Name).value
			this.expect(TokenType.Bang)
			return this.parseSheetSpanQualifiedReference(token.value, endSheet)
		}

		if (this.peek(true).type === TokenType.Bang) {
			this.expect(TokenType.Bang)
			return this.parseSheetQualifiedReference(token.value)
		}

		if (this.peek(true).type === TokenType.Name && this.peek(true).value.startsWith('[')) {
			const bracketToken = this.advance(true)
			return this.buildStructuredRef(token.value, bracketToken.value)
		}

		return { type: 'name', name: token.value }
	}

	private parseSheetSpanQualifiedReference(startSheet: string, endSheet: string): FormulaNode {
		return {
			type: 'sheetSpanRef',
			startSheet,
			endSheet,
			target: this.parseQualifiedReferenceTarget(),
		}
	}

	private parseSheetQualifiedReference(sheet: string): FormulaNode {
		const target = this.parseQualifiedReferenceTarget()
		if (target.type === 'name' && this.peek(true).type === TokenType.OpenParen) {
			return { type: 'function', name: `${sheet}!${target.name}`, args: this.parseCallArgs() }
		}
		switch (target.type) {
			case 'error':
				return target
			case 'cellRef':
				return { type: 'cellRef', ref: target.ref, sheet }
			case 'rangeRef':
				return { type: 'rangeRef', start: target.start, end: target.end, sheet }
			case 'wholeRowRange':
				return { type: 'wholeRowRange', startRow: target.startRow, endRow: target.endRow, sheet }
			case 'wholeColumnRange':
				return {
					type: 'wholeColumnRange',
					startCol: target.startCol,
					endCol: target.endCol,
					...(target.startColAbsolute ? { startColAbsolute: true } : {}),
					...(target.endColAbsolute ? { endColAbsolute: true } : {}),
					sheet,
				}
			case 'name':
				return { type: 'name', name: target.name, sheet }
			default:
				throw new Error(`Unsupported sheet-qualified reference target "${target.type}"`)
		}
	}

	private parseQualifiedReferenceTarget(): FormulaNode {
		if (
			this.peek(true).type === TokenType.Number &&
			this.lookahead(1, true).type === TokenType.Colon &&
			this.lookahead(2, true).type === TokenType.Number
		) {
			const start = Number.parseInt(this.advance(true).value, 10) - 1
			this.expect(TokenType.Colon)
			const end = Number.parseInt(this.expect(TokenType.Number).value, 10) - 1
			return { type: 'wholeRowRange', startRow: start, endRow: end }
		}
		if (
			this.peek(true).type === TokenType.Name &&
			isColumnLabel(this.peek(true).value) &&
			this.lookahead(1, true).type === TokenType.Colon &&
			this.lookahead(2, true).type === TokenType.Name &&
			isColumnLabel(this.lookahead(2, true).value)
		) {
			const startCol = this.advance(true).value
			this.expect(TokenType.Colon)
			const endCol = this.expect(TokenType.Name).value
			return wholeColumnRangeNode(startCol, endCol)
		}
		if (this.peek(true).type === TokenType.CellRef) {
			return this.parseCellOrRange()
		}
		if (this.peek(true).type === TokenType.Error) {
			const token = this.advance(true)
			const error = token.value as ExcelError
			if (error === '#REF!') return { type: 'error', value: error }
			throw new Error(
				`Unsupported sheet-qualified error reference "${error}" at position ${token.position}`,
			)
		}
		if (this.peek(true).type === TokenType.Name) {
			const name = this.advance(true).value
			return { type: 'name', name }
		}
		if (this.peek(true).type === TokenType.Function) {
			const name = this.advance(true).value
			return { type: 'name', name }
		}
		throw new Error(
			`Expected cell reference or name after qualifier at position ${this.peek(true).position}`,
		)
	}

	private buildStructuredRef(table: string, bracket: string): StructuredRefNode {
		const content = bracket.slice(1, -1)
		const specifiers: string[] = []
		let column: string | undefined
		let endColumn: string | undefined

		const setColumnSpec = (text: string): void => {
			const rangeSeparator = findStructuredRefTopLevelColon(text)
			if (rangeSeparator >= 0) {
				const startColumn = text.slice(0, rangeSeparator).trim()
				const endColumnText = text.slice(rangeSeparator + 1).trim()
				if (
					hasOuterStructuredRefBrackets(startColumn) &&
					hasOuterStructuredRefBrackets(endColumnText)
				) {
					column = unescapeStructuredRefColumn(stripOuterStructuredRefBrackets(startColumn))
					endColumn = unescapeStructuredRefColumn(stripOuterStructuredRefBrackets(endColumnText))
					return
				}
			}
			column = unescapeStructuredRefColumn(stripOuterStructuredRefBrackets(text))
		}

		if (content.startsWith('@')) {
			specifiers.push('@')
			const col = content.slice(1)
			if (col.length > 0) setColumnSpec(col)
		} else if (content.startsWith('#')) {
			specifiers.push(content)
		} else if (content.startsWith('[')) {
			const parts = splitStructuredRefParts(content)
			for (const part of parts) {
				const cleaned = stripOuterStructuredRefBrackets(part)
				if (cleaned.startsWith('#')) {
					specifiers.push(cleaned)
				} else {
					setColumnSpec(part)
				}
			}
		} else {
			column = unescapeStructuredRefColumn(content)
		}

		if (column !== undefined) {
			return {
				type: 'structuredRef',
				table,
				specifiers,
				column,
				...(endColumn ? { endColumn } : {}),
			}
		}
		return { type: 'structuredRef', table, specifiers }
	}

	private parseArrayLiteral(): FormulaNode {
		this.expect(TokenType.OpenBrace)
		const rows: FormulaNode[][] = []
		let currentRow: FormulaNode[] = []
		let expectValue = true

		while (this.peek(true).type !== TokenType.CloseBrace) {
			if (this.peek(true).type === TokenType.Comma) {
				if (expectValue) currentRow.push({ type: 'missing' })
				this.advance(true)
				expectValue = true
			} else if (this.peek(true).type === TokenType.Semicolon) {
				if (expectValue) currentRow.push({ type: 'missing' })
				this.advance(true)
				rows.push(currentRow)
				currentRow = []
				expectValue = true
			} else {
				if (!expectValue) {
					throw new Error(
						`Expected , or ; in array literal at position ${this.peek(true).position}`,
					)
				}
				currentRow.push(this.parseExpression(false))
				expectValue = false
			}
		}

		if (expectValue && currentRow.length > 0) currentRow.push({ type: 'missing' })
		rows.push(currentRow)
		this.expect(TokenType.CloseBrace)
		return { type: 'array', rows }
	}
}

export function parse(tokens: readonly Token[]): Result<FormulaNode> {
	try {
		return ok(parseOrThrow(tokens))
	} catch (e) {
		return err(ascendError('FORMULA_PARSE_ERROR', e instanceof Error ? e.message : String(e)))
	}
}

export function parseFormula(formula: string): Result<FormulaNode> {
	return parse(tokenize(formula))
}

export function parseOrThrow(tokens: readonly Token[]): FormulaNode {
	return new FormulaParser(tokens).parseOrThrow()
}

export function parseFormulaOrThrow(formula: string): FormulaNode {
	return parseOrThrow(tokenize(formula))
}

const PARSE_CACHE_LIMIT = 8192
const PARSE_CACHE_EVICT = 2048
const globalParseCache = new Map<string, Result<FormulaNode>>()

export function cachedParseFormula(formula: string): Result<FormulaNode> {
	const hit = globalParseCache.get(formula)
	if (hit) {
		if (globalParseCache.size >= PARSE_CACHE_LIMIT) {
			globalParseCache.delete(formula)
			globalParseCache.set(formula, hit)
		}
		return hit
	}
	const result = parseFormula(formula)
	if (globalParseCache.size >= PARSE_CACHE_LIMIT) {
		const iter = globalParseCache.keys()
		for (let i = 0; i < PARSE_CACHE_EVICT; i++) {
			const k = iter.next().value
			if (k !== undefined) globalParseCache.delete(k)
		}
	}
	globalParseCache.set(formula, result)
	return result
}

export function clearGlobalParseCache(): void {
	globalParseCache.clear()
}

function isReferenceLike(node: FormulaNode): boolean {
	switch (node.type) {
		case 'cellRef':
		case 'rangeRef':
		case 'dynamicRangeRef':
		case 'wholeRowRange':
		case 'wholeColumnRange':
		case 'name':
		case 'structuredRef':
		case 'spillRef':
		case 'sheetSpanRef':
			return true
		case 'function':
			return isReferenceFunctionName(node.name)
		case 'binary':
			return node.op === ',' || node.op === ' '
		default:
			return false
	}
}

function isReferenceFunctionName(name: string): boolean {
	switch (name.toUpperCase()) {
		case 'INDEX':
		case 'INDIRECT':
		case 'OFFSET':
			return true
		default:
			return false
	}
}

function isSheetNameSegmentToken(token: Token): boolean {
	return token.type === TokenType.Name || token.type === TokenType.CellRef
}

function splitSheetSpanToken(
	token: string,
): { readonly startSheet: string; readonly endSheet: string } | null {
	const colon = sheetSpanColonIndex(token)
	if (colon < 0) return null
	return {
		startSheet: token.slice(0, colon),
		endSheet: token.slice(colon + 1),
	}
}

function sheetSpanColonIndex(token: string): number {
	const open = token.indexOf('[')
	if (open < 0) return token.indexOf(':')
	const close = token.indexOf(']', open + 1)
	if (close < 0) return -1
	return token.indexOf(':', close + 1)
}

function makeRangeFromEndpoints(left: FormulaNode, right: FormulaNode): FormulaNode {
	const normalizedRight = inheritEndpointSheet(left, right)
	if (left.type === 'cellRef' && normalizedRight.type === 'cellRef') {
		const sheet = left.sheet ?? normalizedRight.sheet
		if (sheet !== undefined) {
			return { type: 'rangeRef', start: left.ref, end: normalizedRight.ref, sheet }
		}
		return { type: 'rangeRef', start: left.ref, end: normalizedRight.ref }
	}
	return { type: 'dynamicRangeRef', start: left, end: normalizedRight }
}

function inheritEndpointSheet(left: FormulaNode, right: FormulaNode): FormulaNode {
	const sheet = left.type === 'cellRef' ? left.sheet : undefined
	if (sheet === undefined) return right
	switch (right.type) {
		case 'cellRef':
			return right.sheet === undefined ? { type: 'cellRef', ref: right.ref, sheet } : right
		case 'rangeRef':
			return right.sheet === undefined
				? { type: 'rangeRef', start: right.start, end: right.end, sheet }
				: right
		case 'wholeRowRange':
			return right.sheet === undefined
				? {
						type: 'wholeRowRange',
						startRow: right.startRow,
						endRow: right.endRow,
						sheet,
					}
				: right
		case 'wholeColumnRange':
			return right.sheet === undefined
				? {
						type: 'wholeColumnRange',
						startCol: right.startCol,
						endCol: right.endCol,
						...(right.startColAbsolute ? { startColAbsolute: true } : {}),
						...(right.endColAbsolute ? { endColAbsolute: true } : {}),
						sheet,
					}
				: right
		default:
			return right
	}
}

function canStartReferenceExpression(token: Token): boolean {
	switch (token.type) {
		case TokenType.CellRef:
		case TokenType.Name:
		case TokenType.Number:
		case TokenType.Function:
		case TokenType.OpenParen:
			return true
		default:
			return false
	}
}

export function normalizeFormulaInput(formula: string): string {
	return formula.startsWith('=') ? formula.slice(1) : formula
}
