import { type Token, TokenType } from './tokens.ts'

const CELL_REF_RE = /^\$?[A-Za-z]{1,3}\$?\d+$/

const ERROR_RE = /^#(?:GETTING_DATA|NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!)/

function isDigit(ch: string | undefined): ch is string {
	return ch !== undefined && ch >= '0' && ch <= '9'
}

function isAlpha(ch: string | undefined): ch is string {
	return ch !== undefined && ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z'))
}

function isIdStart(ch: string | undefined): ch is string {
	return isAlpha(ch) || ch === '_'
}

function isIdPart(ch: string | undefined): ch is string {
	return isAlpha(ch) || isDigit(ch) || ch === '_' || ch === '$'
}

export function tokenize(formula: string): Token[] {
	const tokens: Token[] = []
	const len = formula.length
	let pos = 0

	while (pos < len) {
		const ch = formula[pos]

		if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
			const start = pos
			pos++
			while (pos < len) {
				const ws = formula[pos]
				if (ws !== ' ' && ws !== '\t' && ws !== '\n' && ws !== '\r') break
				pos++
			}
			tokens.push({ type: TokenType.Whitespace, value: formula.slice(start, pos), position: start })
			continue
		}

		if (ch === '"') {
			const start = pos
			pos++
			const strStart = pos
			let hasEscape = false
			while (pos < len) {
				if (formula[pos] === '"') {
					if (formula[pos + 1] === '"') {
						hasEscape = true
						pos += 2
					} else {
						break
					}
				} else {
					pos++
				}
			}
			const raw = formula.slice(strStart, pos)
			if (pos < len) pos++
			tokens.push({
				type: TokenType.String,
				value: hasEscape ? raw.replaceAll('""', '"') : raw,
				position: start,
			})
			continue
		}

		if (ch === '#') {
			const start = pos
			ERROR_RE.lastIndex = 0
			const m = ERROR_RE.exec(formula.slice(pos))
			if (m) {
				tokens.push({ type: TokenType.Error, value: m[0], position: start })
				pos += m[0].length
			} else {
				tokens.push({ type: TokenType.Operator, value: '#', position: start })
				pos++
			}
			continue
		}

		if (isDigit(ch) || (ch === '.' && isDigit(formula[pos + 1]))) {
			const start = pos
			while (pos < len && isDigit(formula[pos])) pos++
			if (formula[pos] === '.') {
				pos++
				while (pos < len && isDigit(formula[pos])) pos++
			}
			const expCh = formula[pos]
			if (expCh === 'E' || expCh === 'e') {
				pos++
				const sign = formula[pos]
				if (sign === '+' || sign === '-') pos++
				while (pos < len && isDigit(formula[pos])) pos++
			}
			tokens.push({ type: TokenType.Number, value: formula.slice(start, pos), position: start })
			continue
		}

		if (ch === "'") {
			const start = pos
			pos++
			const qStart = pos
			let hasEscape = false
			while (pos < len) {
				if (formula[pos] === "'") {
					if (formula[pos + 1] === "'") {
						hasEscape = true
						pos += 2
					} else {
						break
					}
				} else {
					pos++
				}
			}
			const raw = formula.slice(qStart, pos)
			if (pos < len) pos++
			tokens.push({
				type: TokenType.Name,
				value: hasEscape ? raw.replaceAll("''", "'") : raw,
				position: start,
			})
			continue
		}

		if (isIdStart(ch) || ch === '$') {
			const start = pos
			while (pos < len && isIdPart(formula[pos])) pos++
			while (formula[pos] === '.' && pos + 1 < len && isAlpha(formula[pos + 1])) {
				pos++
				while (pos < len && isIdPart(formula[pos])) pos++
			}
			const raw = formula.slice(start, pos)

			if (formula[pos] === '!') {
				tokens.push({ type: TokenType.Name, value: raw, position: start })
				continue
			}

			const upper = raw.toUpperCase()
			if ((upper === 'TRUE' || upper === 'FALSE') && formula[pos] !== '(') {
				tokens.push({ type: TokenType.Boolean, value: upper, position: start })
				continue
			}

			if (formula[pos] === '(') {
				tokens.push({ type: TokenType.Function, value: raw, position: start })
				continue
			}

			if (CELL_REF_RE.test(raw)) {
				tokens.push({ type: TokenType.CellRef, value: raw, position: start })
				continue
			}

			tokens.push({ type: TokenType.Name, value: raw, position: start })
			continue
		}

		if (ch === '<') {
			const start = pos
			pos++
			const next = formula[pos]
			if (next === '>') {
				pos++
				tokens.push({ type: TokenType.Operator, value: '<>', position: start })
			} else if (next === '=') {
				pos++
				tokens.push({ type: TokenType.Operator, value: '<=', position: start })
			} else {
				tokens.push({ type: TokenType.Operator, value: '<', position: start })
			}
			continue
		}

		if (ch === '>') {
			const start = pos
			pos++
			if (formula[pos] === '=') {
				pos++
				tokens.push({ type: TokenType.Operator, value: '>=', position: start })
			} else {
				tokens.push({ type: TokenType.Operator, value: '>', position: start })
			}
			continue
		}

		if (
			ch === '+' ||
			ch === '-' ||
			ch === '*' ||
			ch === '/' ||
			ch === '^' ||
			ch === '&' ||
			ch === '=' ||
			ch === '%' ||
			ch === '@'
		) {
			tokens.push({ type: TokenType.Operator, value: ch, position: pos })
			pos++
			continue
		}

		if (ch === '(') {
			tokens.push({ type: TokenType.OpenParen, value: '(', position: pos })
			pos++
			continue
		}
		if (ch === ')') {
			tokens.push({ type: TokenType.CloseParen, value: ')', position: pos })
			pos++
			continue
		}
		if (ch === '{') {
			tokens.push({ type: TokenType.OpenBrace, value: '{', position: pos })
			pos++
			continue
		}
		if (ch === '}') {
			tokens.push({ type: TokenType.CloseBrace, value: '}', position: pos })
			pos++
			continue
		}
		if (ch === ',') {
			tokens.push({ type: TokenType.Comma, value: ',', position: pos })
			pos++
			continue
		}
		if (ch === ';') {
			tokens.push({ type: TokenType.Semicolon, value: ';', position: pos })
			pos++
			continue
		}
		if (ch === ':') {
			tokens.push({ type: TokenType.Colon, value: ':', position: pos })
			pos++
			continue
		}
		if (ch === '!') {
			tokens.push({ type: TokenType.Bang, value: '!', position: pos })
			pos++
			continue
		}

		if (ch === '[') {
			const start = pos
			pos++
			let depth = 1
			while (pos < len && depth > 0) {
				const bc = formula[pos]
				if (bc === '[') depth++
				else if (bc === ']') depth--
				if (depth > 0) pos++
			}
			if (pos < len) pos++
			tokens.push({ type: TokenType.Name, value: formula.slice(start, pos), position: start })
			continue
		}

		pos++
	}

	tokens.push({ type: TokenType.EOF, value: '', position: pos })
	return tokens
}
