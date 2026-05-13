import { type Token, TokenType, tokenize } from '@ascend/formulas'

export type FormulaTokenClass =
	| 'reference'
	| 'name'
	| 'function'
	| 'literal'
	| 'operator'
	| 'punctuation'
	| 'whitespace'

export interface FormulaTokenRange {
	readonly type: TokenType
	readonly text: string
	readonly start: number
	readonly end: number
	readonly className: FormulaTokenClass
}

export interface FormulaReferenceRange {
	readonly text: string
	readonly start: number
	readonly end: number
	readonly kind: 'cell' | 'range' | 'sheet-cell' | 'sheet-range' | 'structured' | 'spill'
}

export interface CycleReferenceResult {
	readonly formula: string
	readonly cursor: number
	readonly changed: boolean
	readonly reference?: FormulaReferenceRange
}

export interface InsertFormulaReferenceOptions {
	readonly replaceReferenceAtCursor?: boolean
}

export interface InsertFormulaReferenceResult {
	readonly formula: string
	readonly cursor: number
	readonly inserted: string
	readonly replaced?: FormulaReferenceRange
}

interface TokenSpan {
	readonly token: Token
	readonly start: number
	readonly end: number
	readonly text: string
}

export function formulaTokenRanges(formula: string): FormulaTokenRange[] {
	const spans = tokenSpans(formula)
	return spans
		.filter((span) => span.token.type !== TokenType.EOF)
		.map((span) => ({
			type: span.token.type,
			text: span.text,
			start: span.start,
			end: span.end,
			className: classifyFormulaToken(span.token),
		}))
}

export function referenceAtCursor(formula: string, cursor: number): FormulaReferenceRange | null {
	const spans = tokenSpans(formula)
	const references = collectFormulaReferenceRanges(formula, spans)
	let best: FormulaReferenceRange | null = null
	for (const reference of references) {
		if (cursor >= reference.start && cursor <= reference.end) return reference
		if (reference.end <= cursor && (!best || reference.end > best.end)) best = reference
	}
	return best
}

export function cycleFormulaReferenceMode(formula: string, cursor: number): CycleReferenceResult {
	const spans = tokenSpans(formula)
	let target: TokenSpan | undefined
	for (const span of spans) {
		if (span.token.type !== TokenType.CellRef) continue
		if (cursor >= span.start && cursor <= span.end) {
			target = span
			break
		}
		if (span.end <= cursor) target = span
	}
	if (!target) return { formula, cursor, changed: false }
	const cycled = cycleA1Reference(target.text)
	if (cycled === target.text) return { formula, cursor, changed: false }
	const nextFormula = `${formula.slice(0, target.start)}${cycled}${formula.slice(target.end)}`
	const reference = referenceAtCursor(nextFormula, target.start + cycled.length)
	const result: CycleReferenceResult = {
		formula: nextFormula,
		cursor: target.start + cycled.length,
		changed: true,
	}
	return reference ? { ...result, reference } : result
}

export function insertFormulaReference(
	formula: string,
	cursor: number,
	reference: string,
	options: InsertFormulaReferenceOptions = {},
): InsertFormulaReferenceResult {
	const clampedCursor = Math.max(0, Math.min(formula.length, cursor))
	const replaced = options.replaceReferenceAtCursor
		? referenceAtCursor(formula, clampedCursor)
		: null
	const start = replaced?.start ?? clampedCursor
	const end = replaced?.end ?? clampedCursor
	return {
		formula: `${formula.slice(0, start)}${reference}${formula.slice(end)}`,
		cursor: start + reference.length,
		inserted: reference,
		...(replaced ? { replaced } : {}),
	}
}

function tokenSpans(formula: string): TokenSpan[] {
	const tokens = tokenize(formula)
	return tokens.map((token, index) => {
		const next = tokens[index + 1]
		const end = next ? next.position : formula.length
		return {
			token,
			start: token.position,
			end,
			text: formula.slice(token.position, end),
		}
	})
}

function collectFormulaReferenceRanges(
	formula: string,
	spans: readonly TokenSpan[],
): FormulaReferenceRange[] {
	const references: FormulaReferenceRange[] = []
	for (let i = 0; i < spans.length; i++) {
		const span = spans[i]
		if (!span || span.token.type === TokenType.EOF) continue
		if (span.token.type === TokenType.CellRef) {
			const range = collectCellReferenceRange(formula, spans, i)
			references.push(range)
			continue
		}
		if (span.token.type !== TokenType.Name) continue
		const next = nextNonWhitespace(spans, i + 1)
		if (next?.span.token.type === TokenType.Name && next.span.text.startsWith('[')) {
			references.push({
				text: formula.slice(span.start, next.span.end),
				start: span.start,
				end: next.span.end,
				kind: 'structured',
			})
			i = next.index
			continue
		}
		if (next?.span.token.type === TokenType.Bang) {
			const cell = nextNonWhitespace(spans, next.index + 1)
			if (cell?.span.token.type === TokenType.CellRef) {
				const local = collectCellReferenceRange(formula, spans, cell.index)
				references.push({
					text: formula.slice(span.start, local.end),
					start: span.start,
					end: local.end,
					kind: local.kind === 'range' ? 'sheet-range' : 'sheet-cell',
				})
				i = cell.index
			}
		}
	}
	return references
}

function collectCellReferenceRange(
	formula: string,
	spans: readonly TokenSpan[],
	index: number,
): FormulaReferenceRange {
	const start = spans[index]
	if (!start) throw new Error('Missing cell reference token')
	let end = start.end
	let kind: FormulaReferenceRange['kind'] = 'cell'
	const afterCell = nextNonWhitespace(spans, index + 1)
	if (afterCell?.span.token.type === TokenType.Colon) {
		const rangeEnd = nextNonWhitespace(spans, afterCell.index + 1)
		if (rangeEnd?.span.token.type === TokenType.CellRef) {
			end = rangeEnd.span.end
			kind = 'range'
		}
	}
	const afterRange = nextNonWhitespace(spans, index + 1)
	if (
		kind === 'cell' &&
		afterRange?.span.token.type === TokenType.Operator &&
		afterRange.span.text === '#'
	) {
		end = afterRange.span.end
		kind = 'spill'
	}
	return {
		text: formula.slice(start.start, end),
		start: start.start,
		end,
		kind,
	}
}

function nextNonWhitespace(
	spans: readonly TokenSpan[],
	start: number,
): { readonly span: TokenSpan; readonly index: number } | undefined {
	for (let index = start; index < spans.length; index++) {
		const span = spans[index]
		if (!span || span.token.type === TokenType.EOF) return undefined
		if (span.token.type !== TokenType.Whitespace) return { span, index }
	}
	return undefined
}

function classifyFormulaToken(token: Token): FormulaTokenClass {
	switch (token.type) {
		case TokenType.CellRef:
			return 'reference'
		case TokenType.Name:
			return 'name'
		case TokenType.Function:
			return 'function'
		case TokenType.Number:
		case TokenType.String:
		case TokenType.Boolean:
		case TokenType.Error:
			return 'literal'
		case TokenType.Operator:
			return 'operator'
		case TokenType.Whitespace:
			return 'whitespace'
		default:
			return 'punctuation'
	}
}

function cycleA1Reference(ref: string): string {
	const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/.exec(ref)
	if (!match) return ref
	const [, colAbs = '', col = '', rowAbs = '', row = ''] = match
	if (!colAbs && !rowAbs) return `$${col.toUpperCase()}$${row}`
	if (colAbs && rowAbs) return `${col.toUpperCase()}$${row}`
	if (!colAbs && rowAbs) return `$${col.toUpperCase()}${row}`
	return `${col.toUpperCase()}${row}`
}
