import { type Token, TokenType } from './tokens.ts'

const CELL_REF_RE = /^\$?[A-Za-z]{1,3}\$?\d+$/

const ERROR_LITERALS = [
	'#GETTING_DATA',
	'#NULL!',
	'#DIV/0!',
	'#VALUE!',
	'#REF!',
	'#NAME?',
	'#NUM!',
	'#N/A',
	'#SPILL!',
	'#CALC!',
] as const

function isDigit(ch: string): boolean {
	return ch >= '0' && ch <= '9'
}

function isAlpha(ch: string): boolean {
	return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z')
}

function isIdStart(ch: string): boolean {
	return isAlpha(ch) || ch === '_'
}

function isIdPart(ch: string): boolean {
	return isAlpha(ch) || isDigit(ch) || ch === '_' || ch === '$'
}

export function tokenize(formula: string): Token[] {
	const tokens: Token[] = []
	let pos = 0

	while (pos < formula.length) {
		const ch = formula.charAt(pos)

		if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
			const start = pos
			pos++
			while (pos < formula.length) {
				const ws = formula.charAt(pos)
				if (ws !== ' ' && ws !== '\t' && ws !== '\n' && ws !== '\r') break
				pos++
			}
			tokens.push({ type: TokenType.Whitespace, value: formula.slice(start, pos), position: start })
			continue
		}

		if (ch === '"') {
			const start = pos
			pos++
			let value = ''
			while (pos < formula.length) {
				const sc = formula.charAt(pos)
				if (sc === '"') {
					if (formula.charAt(pos + 1) === '"') {
						value += '"'
						pos += 2
					} else {
						pos++
						break
					}
				} else {
					value += sc
					pos++
				}
			}
			tokens.push({ type: TokenType.String, value, position: start })
			continue
		}

		if (ch === '#') {
			const start = pos
			let matched = false
			for (const pattern of ERROR_LITERALS) {
				if (formula.startsWith(pattern, pos)) {
					tokens.push({ type: TokenType.Error, value: pattern, position: start })
					pos += pattern.length
					matched = true
					break
				}
			}
			if (!matched) {
				tokens.push({ type: TokenType.Operator, value: '#', position: start })
				pos++
			}
			continue
		}

		if (isDigit(ch) || (ch === '.' && isDigit(formula.charAt(pos + 1)))) {
			const start = pos
			while (pos < formula.length && isDigit(formula.charAt(pos))) pos++
			if (formula.charAt(pos) === '.') {
				pos++
				while (pos < formula.length && isDigit(formula.charAt(pos))) pos++
			}
			const expCh = formula.charAt(pos)
			if (expCh === 'E' || expCh === 'e') {
				pos++
				const sign = formula.charAt(pos)
				if (sign === '+' || sign === '-') pos++
				while (pos < formula.length && isDigit(formula.charAt(pos))) pos++
			}
			tokens.push({ type: TokenType.Number, value: formula.slice(start, pos), position: start })
			continue
		}

		if (ch === "'") {
			const start = pos
			pos++
			let value = ''
			while (pos < formula.length) {
				const qc = formula.charAt(pos)
				if (qc === "'") {
					if (formula.charAt(pos + 1) === "'") {
						value += "'"
						pos += 2
					} else {
						pos++
						break
					}
				} else {
					value += qc
					pos++
				}
			}
			tokens.push({ type: TokenType.Name, value, position: start })
			continue
		}

		if (isIdStart(ch) || ch === '$') {
			const start = pos
			while (pos < formula.length && isIdPart(formula.charAt(pos))) pos++
			while (
				formula.charAt(pos) === '.' &&
				pos + 1 < formula.length &&
				isAlpha(formula.charAt(pos + 1))
			) {
				pos++
				while (pos < formula.length && isIdPart(formula.charAt(pos))) pos++
			}
			const raw = formula.slice(start, pos)

			if (formula.charAt(pos) === '!') {
				tokens.push({ type: TokenType.Name, value: raw, position: start })
				continue
			}

			const upper = raw.toUpperCase()
			if ((upper === 'TRUE' || upper === 'FALSE') && formula.charAt(pos) !== '(') {
				tokens.push({ type: TokenType.Boolean, value: upper, position: start })
				continue
			}

			if (formula.charAt(pos) === '(') {
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
			const next = formula.charAt(pos)
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
			if (formula.charAt(pos) === '=') {
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
			while (pos < formula.length && depth > 0) {
				const bc = formula.charAt(pos)
				if (bc === '[') depth++
				else if (bc === ']') depth--
				if (depth > 0) pos++
			}
			if (pos < formula.length) pos++
			tokens.push({ type: TokenType.Name, value: formula.slice(start, pos), position: start })
			continue
		}

		pos++
	}

	tokens.push({ type: TokenType.EOF, value: '', position: pos })
	return tokens
}
