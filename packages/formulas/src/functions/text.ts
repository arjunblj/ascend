import type { CellValue, ScalarCellValue } from '@ascend/schema'
import {
	arrayValue,
	booleanValue,
	EMPTY,
	errorValue,
	isEmpty,
	isError,
	numberValue,
	stringValue,
	topLeftScalar,
} from '@ascend/schema'
import { serialToDate } from './date.ts'
import type { EvalArg, FunctionDef, FunctionEvalContext } from './index.ts'
import { numArg, toNum } from './math/helpers.ts'

function fn(
	name: string,
	minArgs: number,
	maxArgs: number,
	evaluate: (args: EvalArg[], ctx?: FunctionEvalContext) => CellValue,
): FunctionDef {
	return { name, minArgs, maxArgs, volatile: false, evaluate }
}

function cvStr(v: CellValue): string {
	v = topLeftScalar(v)
	switch (v.kind) {
		case 'empty':
			return ''
		case 'number':
			return String(v.value)
		case 'string':
			return v.value
		case 'boolean':
			return v.value ? 'TRUE' : 'FALSE'
		case 'error':
			return v.value
		case 'date':
			return String(v.serial)
		case 'richText':
			return v.runs.map((r) => r.text).join('')
	}
}

function strArg(arg: EvalArg | undefined): string | CellValue {
	const v = arg?.value ?? EMPTY
	if (isError(v)) return v
	return cvStr(v)
}

function dollarNumArg(arg: EvalArg | undefined): number | CellValue {
	const v = topLeftScalar(arg?.value ?? EMPTY)
	if (isError(v)) return v
	if (v.kind !== 'string' && v.kind !== 'richText') return numArg(arg)
	const text = v.kind === 'richText' ? v.runs.map((run) => run.text).join('') : v.value
	let trimmed = text.trim()
	if (trimmed === '') return errorValue('#VALUE!')
	let negative = false
	if (trimmed.startsWith('(') && trimmed.endsWith(')')) {
		negative = true
		trimmed = trimmed.slice(1, -1).trim()
	}
	if (trimmed.startsWith('-')) {
		negative = !negative
		trimmed = trimmed.slice(1).trim()
	}
	const n = Number(trimmed.replace(/[$€£¥,]/g, ''))
	if (Number.isNaN(n)) return errorValue('#VALUE!')
	return negative ? -n : n
}

function proper(s: string): string {
	let result = ''
	let cap = true
	for (const ch of s) {
		if (/[a-zA-Z]/.test(ch)) {
			result += cap ? ch.toUpperCase() : ch.toLowerCase()
			cap = false
		} else {
			result += ch
			cap = true
		}
	}
	return result
}

function findTextSlice(
	text: string,
	delimiter: string,
	instanceNum: number,
	matchMode: number,
): { start: number; end: number } | null {
	if (delimiter === '') return null
	const haystack = matchMode === 1 ? text.toLowerCase() : text
	const needle = matchMode === 1 ? delimiter.toLowerCase() : delimiter
	const matches: Array<{ start: number; end: number }> = []
	let searchFrom = 0
	while (searchFrom <= haystack.length) {
		const start = haystack.indexOf(needle, searchFrom)
		if (start === -1) break
		matches.push({ start, end: start + delimiter.length })
		searchFrom = start + Math.max(1, delimiter.length)
	}
	if (matches.length === 0) return null
	if (instanceNum > 0) return matches[instanceNum - 1] ?? null
	const reverseIndex = matches.length + instanceNum
	return reverseIndex >= 0 ? (matches[reverseIndex] ?? null) : null
}

function textSplitPoint(
	args: EvalArg[],
): { text: string; delimiter: string; instanceNum: number; matchMode: number } | CellValue {
	const text = strArg(args[0])
	if (typeof text !== 'string') return text
	const delimiter = strArg(args[1])
	if (typeof delimiter !== 'string') return delimiter
	const instanceRaw = args[2] ? numArg(args[2]) : 1
	if (typeof instanceRaw !== 'number') return instanceRaw
	const instanceNum = Math.trunc(instanceRaw)
	if (instanceNum === 0) return errorValue('#VALUE!')
	const matchModeRaw = args[3] ? numArg(args[3]) : 0
	if (typeof matchModeRaw !== 'number') return matchModeRaw
	const matchMode = Math.trunc(matchModeRaw)
	if (matchMode !== 0 && matchMode !== 1) return errorValue('#VALUE!')
	return { text, delimiter, instanceNum, matchMode }
}

function splitByDelimiter(text: string, delimiter: string, matchMode: number): string[] {
	if (delimiter === '') return [...text]
	if (matchMode === 0) return text.split(delimiter)
	const loweredText = text.toLowerCase()
	const loweredDelimiter = delimiter.toLowerCase()
	const parts: string[] = []
	let start = 0
	let searchFrom = 0
	while (searchFrom <= text.length) {
		const index = loweredText.indexOf(loweredDelimiter, searchFrom)
		if (index === -1) {
			parts.push(text.slice(start))
			break
		}
		parts.push(text.slice(start, index))
		start = index + delimiter.length
		searchFrom = start
	}
	return parts
}

const MONTH_NAMES = [
	'January',
	'February',
	'March',
	'April',
	'May',
	'June',
	'July',
	'August',
	'September',
	'October',
	'November',
	'December',
]

const DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

function isDateFormat(fmt: string): boolean {
	if (/\[(h+|m+|s+)\]/i.test(fmt)) return true
	const stripped = fmt.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '')
	for (const run of stripped.match(/[A-Za-z]+(?:\/[A-Za-z]+)?/g) ?? []) {
		if (/^(AM\/PM|A\/P)$/i.test(run)) return true
		if (/^(y{1,4}|m{1,4}|d{1,4}|h{1,2}|s{1,2})$/i.test(run)) return true
	}
	return false
}

function serialToTime(serial: number): { hours: number; minutes: number; seconds: number } {
	const frac = serial - Math.floor(serial)
	const totalSeconds = Math.round(frac * 86400)
	const hours = Math.floor(totalSeconds / 3600)
	const minutes = Math.floor((totalSeconds % 3600) / 60)
	const seconds = totalSeconds % 60
	return { hours, minutes, seconds }
}

function secondsWithFraction(serial: number): number {
	const frac = serial - Math.floor(serial)
	const totalSeconds = frac * 86400
	return totalSeconds - Math.floor(totalSeconds / 60) * 60
}

function elapsedSecondsWithFraction(serial: number): number {
	return serial * 86400
}

function dayOfWeek(serial: number, dateSystem: '1900' | '1904'): number {
	const parts = serialToDate(Math.floor(serial), dateSystem)
	if (!parts) return 0
	const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day))
	if (parts.year >= 0 && parts.year < 100) date.setUTCFullYear(parts.year)
	return date.getUTCDay()
}

function formatDate(serial: number, fmt: string, dateSystem: '1900' | '1904' = '1900'): string {
	const parts = serialToDate(Math.floor(serial), dateSystem)
	if (!parts) return ''
	const { hours, minutes, seconds } = serialToTime(serial)
	const secondFraction = secondsWithFraction(serial)
	const elapsedSecondFraction = elapsedSecondsWithFraction(serial)
	const totalSeconds = Math.round(serial * 86400)
	const totalMinutes = Math.floor(totalSeconds / 60)
	const totalHours = Math.floor(totalSeconds / 3600)

	const hasAmPm = /AM\/PM|am\/pm|A\/P|a\/p/.test(fmt)
	const hour12 = hours === 0 ? 12 : hours > 12 ? hours - 12 : hours
	const ampm = hours < 12 ? 'AM' : 'PM'

	let result = ''
	let i = 0
	let prevTokenIsTime = false
	let elapsedSecondsContext = false
	while (i < fmt.length) {
		if (fmt[i] === '"') {
			const end = fmt.indexOf('"', i + 1)
			if (end === -1) {
				result += fmt.slice(i + 1)
				break
			}
			result += fmt.slice(i + 1, end)
			i = end + 1
			continue
		}

		const elapsedToken = fmt.slice(i).match(/^\[(h+|m+|s+)\]/i)
		if (elapsedToken) {
			const token = (elapsedToken[1] ?? '').toLowerCase()
			if (token.startsWith('h')) {
				result += String(totalHours)
				prevTokenIsTime = true
			} else if (token.startsWith('m')) {
				result += String(totalMinutes)
				prevTokenIsTime = true
			} else {
				result += String(totalSeconds)
				prevTokenIsTime = true
				elapsedSecondsContext = true
			}
			i += elapsedToken[0].length
			if (token.startsWith('s')) {
				const fractionMatch = fmt.slice(i).match(/^\.([0]+)/)
				if (fractionMatch) {
					const digits = fractionMatch[1]?.length ?? 0
					const rounded = elapsedSecondFraction.toFixed(digits)
					const fraction = rounded.split('.')[1] ?? ''.padEnd(digits, '0')
					result += `.${fraction}`
					i += fractionMatch[0].length
				}
			}
			continue
		}

		let matched = false
		for (const tok of ['AM/PM', 'am/pm', 'A/P', 'a/p']) {
			if (fmt.slice(i, i + tok.length) === tok) {
				result += tok.includes('AM')
					? ampm
					: tok.includes('am')
						? ampm.toLowerCase()
						: tok.includes('A')
							? ampm === 'AM'
								? 'A'
								: 'P'
							: ampm === 'AM'
								? 'a'
								: 'p'
				i += tok.length
				matched = true
				break
			}
		}
		if (matched) continue

		const remaining = fmt.slice(i)
		const tokenMatch = remaining.match(
			/^(yyyy|yyy|yy|y|ee|e|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s)/i,
		)
		if (tokenMatch) {
			const tok = tokenMatch[0] as string
			const tokLower = tok.toLowerCase()
			switch (tokLower) {
				case 'yyyy':
				case 'yyy':
					result += String(parts.year)
					prevTokenIsTime = false
					break
				case 'yy':
				case 'y':
					result += String(parts.year % 100).padStart(2, '0')
					prevTokenIsTime = false
					break
				case 'ee':
				case 'e':
					result += String(parts.year)
					prevTokenIsTime = false
					break
				case 'mmmm':
					result += MONTH_NAMES[parts.month - 1] ?? ''
					prevTokenIsTime = false
					break
				case 'mmm':
					result += (MONTH_NAMES[parts.month - 1] ?? '').slice(0, 3)
					prevTokenIsTime = false
					break
				case 'mm':
					if (prevTokenIsTime) {
						result += String(minutes).padStart(2, '0')
					} else {
						result += String(parts.month).padStart(2, '0')
					}
					break
				case 'm':
					if (prevTokenIsTime) {
						result += String(minutes)
					} else {
						result += String(parts.month)
					}
					break
				case 'dddd':
					result += DAY_NAMES[dayOfWeek(serial, dateSystem)] ?? ''
					prevTokenIsTime = false
					break
				case 'ddd':
					result += (DAY_NAMES[dayOfWeek(serial, dateSystem)] ?? '').slice(0, 3)
					prevTokenIsTime = false
					break
				case 'dd':
					result += String(parts.day).padStart(2, '0')
					prevTokenIsTime = false
					break
				case 'd':
					result += String(parts.day)
					prevTokenIsTime = false
					break
				case 'hh':
					result += String(hasAmPm ? hour12 : hours).padStart(2, '0')
					prevTokenIsTime = true
					break
				case 'h':
					result += String(hasAmPm ? hour12 : hours)
					prevTokenIsTime = true
					break
				case 'ss':
					result += elapsedSecondsContext
						? String(totalSeconds).padStart(2, '0')
						: String(seconds).padStart(2, '0')
					prevTokenIsTime = true
					break
				case 's':
					result += elapsedSecondsContext ? String(totalSeconds) : String(seconds)
					prevTokenIsTime = true
					break
			}
			i += tok.length
			if (tokLower === 's' || tokLower === 'ss') {
				const fractionMatch = fmt.slice(i).match(/^\.([0]+)/)
				if (fractionMatch) {
					const digits = fractionMatch[1]?.length ?? 0
					const rounded = secondFraction.toFixed(digits)
					const fraction = rounded.split('.')[1] ?? ''.padEnd(digits, '0')
					result += `.${fraction}`
					i += fractionMatch[0].length
				}
			}
			continue
		}

		result += fmt[i]
		i++
	}
	return result
}

const NUM_FMT_CHARS = new Set(['0', '#', '?', '.', ',', '%'])

type NumberFormatToken =
	| { kind: 'placeholder'; char: '0' | '#' | '?' }
	| { kind: 'literal'; text: string; source?: 'raw' | 'escaped' | 'underscore' }

interface FormatSection {
	raw: string
	format: string
	condition?: { op: '<' | '<=' | '>' | '>=' | '=' | '<>'; value: number }
}

function splitFormatSections(fmt: string): string[] {
	const sections: string[] = []
	let current = ''
	let inQuote = false
	let bracketDepth = 0
	for (let i = 0; i < fmt.length; i++) {
		const ch = fmt[i] ?? ''
		if (ch === '"' && bracketDepth === 0) {
			inQuote = !inQuote
			current += ch
			continue
		}
		if (!inQuote && ch === '[') {
			bracketDepth++
			current += ch
			continue
		}
		if (!inQuote && ch === ']' && bracketDepth > 0) {
			bracketDepth--
			current += ch
			continue
		}
		if (!inQuote && bracketDepth === 0 && ch === ';') {
			sections.push(current)
			current = ''
			continue
		}
		if (ch === '\\' && i + 1 < fmt.length) {
			current += ch + (fmt[i + 1] ?? '')
			i++
			continue
		}
		current += ch
	}
	sections.push(current)
	return sections
}

function parseConditionToken(token: string): FormatSection['condition'] {
	const match = /^(<=|>=|<>|=|<|>)\s*(-?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)$/.exec(
		token.trim(),
	)
	if (!match) return undefined
	const value = Number(match[2])
	if (Number.isNaN(value)) return undefined
	return { op: match[1] as NonNullable<FormatSection['condition']>['op'], value }
}

function parseFormatSection(raw: string): FormatSection {
	let format = ''
	let condition: FormatSection['condition']
	const colorToken = /^(black|blue|cyan|green|magenta|red|white|yellow|color\d+)$/i
	for (let i = 0; i < raw.length; i++) {
		const ch = raw[i] ?? ''
		if (ch === '[') {
			const end = raw.indexOf(']', i + 1)
			if (end !== -1) {
				const token = raw.slice(i + 1, end)
				const parsedCondition = parseConditionToken(token)
				if (parsedCondition) {
					condition = parsedCondition
				} else if (token.startsWith('$')) {
					const currency = token.split('-')[0]?.slice(1) ?? ''
					if (currency.length > 0) format += currency
				} else if (!colorToken.test(token)) {
					format += `[${token}]`
				}
				i = end
				continue
			}
		}
		format += ch
	}
	return { raw, format, ...(condition ? { condition } : {}) }
}

function conditionMatches(
	value: number,
	condition: NonNullable<FormatSection['condition']>,
): boolean {
	switch (condition.op) {
		case '<':
			return value < condition.value
		case '<=':
			return value <= condition.value
		case '>':
			return value > condition.value
		case '>=':
			return value >= condition.value
		case '=':
			return value === condition.value
		case '<>':
			return value !== condition.value
	}
}

function hasNegativeAffix(format: string): boolean {
	const { before, after } = parseFormatSegments(format)
	return before.includes('-') || before.includes('(') || after.includes('-') || after.includes(')')
}

function selectNumericSection(
	value: number,
	fmt: string,
): { format: string; autoSign: boolean; absValue: number } {
	const sections = splitFormatSections(fmt).map(parseFormatSection)
	const conditioned = sections.some((section) => section.condition)
	if (conditioned) {
		for (const section of sections) {
			if (section.condition && conditionMatches(value, section.condition)) {
				const autoSign = value < 0 && !hasNegativeAffix(section.format)
				return {
					format: section.format,
					autoSign,
					absValue: autoSign ? Math.abs(value) : value,
				}
			}
		}
		const fallback = sections.find((section) => !section.condition)
		return {
			format: fallback?.format ?? 'General',
			autoSign: fallback !== undefined && value < 0,
			absValue: fallback !== undefined ? Math.abs(value) : value,
		}
	}
	if (sections.length === 1)
		return { format: sections[0]?.format ?? fmt, autoSign: value < 0, absValue: Math.abs(value) }
	if (sections.length === 2) {
		if (value < 0)
			return {
				format: sections[1]?.format ?? sections[0]?.format ?? fmt,
				autoSign: false,
				absValue: Math.abs(value),
			}
		return { format: sections[0]?.format ?? fmt, autoSign: false, absValue: value }
	}
	if (sections.length >= 3) {
		if (value > 0) return { format: sections[0]?.format ?? fmt, autoSign: false, absValue: value }
		if (value < 0)
			return {
				format: sections[1]?.format ?? sections[0]?.format ?? fmt,
				autoSign: false,
				absValue: Math.abs(value),
			}
		return {
			format: sections[2]?.format ?? sections[0]?.format ?? fmt,
			autoSign: false,
			absValue: 0,
		}
	}
	return { format: fmt, autoSign: value < 0, absValue: Math.abs(value) }
}

function parseFormatSegments(fmt: string): { before: string; numFmt: string; after: string } {
	let before = ''
	let numFmt = ''
	let after = ''
	let phase: 'before' | 'num' | 'after' = 'before'
	let i = 0

	while (i < fmt.length) {
		if (fmt[i] === '"') {
			const end = fmt.indexOf('"', i + 1)
			const literal = end === -1 ? fmt.slice(i + 1) : fmt.slice(i + 1, end)
			if (phase === 'before') before += literal
			else if (phase === 'num') {
				phase = 'after'
				after += literal
			} else {
				after += literal
			}
			i = end === -1 ? fmt.length : end + 1
			continue
		}
		if (fmt[i] === '\\' && i + 1 < fmt.length) {
			const literal = fmt[i + 1] ?? ''
			if (phase === 'before') before += literal
			else if (phase === 'num') {
				phase = 'after'
				after += literal
			} else {
				after += literal
			}
			i += 2
			continue
		}
		if (fmt[i] === '_' && i + 1 < fmt.length) {
			i += 2
			continue
		}
		if (fmt[i] === '*' && i + 1 < fmt.length) {
			i += 2
			continue
		}

		const ch = fmt[i] as string
		if (NUM_FMT_CHARS.has(ch)) {
			if (phase === 'before') phase = 'num'
			if (phase === 'after') {
				numFmt += after
				after = ''
				phase = 'num'
			}
			numFmt += ch
		} else if (phase === 'before') {
			before += ch
		} else if (phase === 'num') {
			phase = 'after'
			after += ch
		} else {
			after += ch
		}
		i++
	}

	return { before, numFmt, after }
}

function tokenizeNumberFormat(fmt: string): NumberFormatToken[] {
	const tokens: NumberFormatToken[] = []
	for (let i = 0; i < fmt.length; i++) {
		const ch = fmt[i] ?? ''
		if (ch === '"') {
			const end = fmt.indexOf('"', i + 1)
			tokens.push({ kind: 'literal', text: end === -1 ? fmt.slice(i + 1) : fmt.slice(i + 1, end) })
			i = end === -1 ? fmt.length : end
			continue
		}
		if (ch === '\\' && i + 1 < fmt.length) {
			tokens.push({ kind: 'literal', text: fmt[i + 1] ?? '', source: 'escaped' })
			i++
			continue
		}
		if (ch === '_' && i + 1 < fmt.length) {
			tokens.push({ kind: 'literal', text: ' ', source: 'underscore' })
			i++
			continue
		}
		if (ch === '*' && i + 1 < fmt.length) {
			i++
			continue
		}
		if (ch === '0' || ch === '#' || ch === '?') {
			tokens.push({ kind: 'placeholder', char: ch })
		} else {
			tokens.push({ kind: 'literal', text: ch, source: 'raw' })
		}
	}
	return tokens
}

function tokenHasPlaceholder(
	token: NumberFormatToken,
): token is Extract<NumberFormatToken, { kind: 'placeholder' }> {
	return token.kind === 'placeholder'
}

function hasPlaceholder(tokens: readonly NumberFormatToken[]): boolean {
	return tokens.some(tokenHasPlaceholder)
}

function renderPlaceholderTokens(
	tokens: readonly NumberFormatToken[],
	digits: string,
	options: { showZero?: boolean; trimTrailingLiteralsWhenBlank?: boolean } = {},
): string {
	const placeholderIndexes: number[] = []
	for (let i = 0; i < tokens.length; i++) {
		if (tokens[i]?.kind === 'placeholder') placeholderIndexes.push(i)
	}
	if (placeholderIndexes.length === 0) {
		return tokens.map((token) => (token.kind === 'literal' ? token.text : '')).join('')
	}

	let renderDigits = digits.replace(/^0+(?=\d)/, '')
	if (renderDigits === '0' && !options.showZero) renderDigits = ''
	if (renderDigits === '' && options.showZero) renderDigits = '0'

	const assigned = new Map<number, string>()
	let digitIndex = renderDigits.length - 1
	for (let i = placeholderIndexes.length - 1; i >= 0; i--) {
		const tokenIndex = placeholderIndexes[i] as number
		const token = tokens[tokenIndex]
		if (digitIndex >= 0) {
			assigned.set(tokenIndex, renderDigits[digitIndex] ?? '')
			digitIndex--
		} else if (token?.kind === 'placeholder') {
			if (token.char === '0') assigned.set(tokenIndex, '0')
			else if (token.char === '?') assigned.set(tokenIndex, ' ')
			else assigned.set(tokenIndex, '')
		}
	}

	let prefix = ''
	if (digitIndex >= 0) prefix = renderDigits.slice(0, digitIndex + 1)

	let renderedAnyDigit = false
	let out = ''
	for (let i = 0; i < tokens.length; i++) {
		const token = tokens[i]
		if (!token) continue
		if (token.kind === 'placeholder') {
			const text = assigned.get(i) ?? ''
			if (/\d/.test(text)) renderedAnyDigit = true
			out += text
		} else {
			out += token.text
		}
	}

	if (prefix !== '') {
		const firstPlaceholder = placeholderIndexes[0] ?? 0
		let charOffset = 0
		for (let i = 0; i < firstPlaceholder; i++) {
			const token = tokens[i]
			if (token?.kind === 'literal') charOffset += token.text.length
		}
		out = out.slice(0, charOffset) + prefix + out.slice(charOffset)
		renderedAnyDigit = true
	}

	if (options.trimTrailingLiteralsWhenBlank && !renderedAnyDigit) {
		const lastPlaceholder = placeholderIndexes.at(-1) ?? -1
		let trimStart = out.length
		let offset = 0
		for (let i = 0; i < tokens.length; i++) {
			const token = tokens[i]
			if (!token) continue
			const nextOffset =
				offset + (token.kind === 'literal' ? token.text.length : (assigned.get(i) ?? '').length)
			if (i > lastPlaceholder && token.kind === 'literal') trimStart = Math.min(trimStart, offset)
			offset = nextOffset
		}
		return out.slice(0, trimStart)
	}

	return out
}

function literalText(tokens: readonly NumberFormatToken[]): string {
	return tokens.map((token) => (token.kind === 'literal' ? token.text : '')).join('')
}

function blankDisplayWidth(text: string): string {
	return text.replace(/\S/g, ' ')
}

function blankPlaceholderText(tokens: readonly NumberFormatToken[]): string {
	return tokens.map((token) => (token.kind === 'placeholder' ? ' ' : token.text)).join('')
}

function stripTrailingWholeSeparator(text: string): { text: string; padding: string } {
	return text.endsWith(':') ? { text: text.slice(0, -1), padding: ' ' } : { text, padding: '' }
}

function bestFraction(
	value: number,
	maxDenominator: number,
): { numerator: number; denominator: number } {
	if (value === 0) return { numerator: 0, denominator: 1 }
	let bestNumerator = 0
	let bestDenominator = 1
	let bestError = Number.POSITIVE_INFINITY
	for (let denominator = 1; denominator <= maxDenominator; denominator++) {
		const numerator = Math.round(value * denominator)
		const error = Math.abs(value - numerator / denominator)
		if (error < bestError - Number.EPSILON) {
			bestNumerator = numerator
			bestDenominator = denominator
			bestError = error
			if (error === 0) break
		}
	}
	return { numerator: bestNumerator, denominator: bestDenominator }
}

function splitMixedFractionLeft(tokens: readonly NumberFormatToken[]): {
	whole: readonly NumberFormatToken[]
	numerator: readonly NumberFormatToken[]
	preSlash: readonly NumberFormatToken[]
} | null {
	let end = tokens.length - 1
	while (end >= 0 && tokens[end]?.kind !== 'placeholder') end--
	if (end < 0) return null
	let start = end
	while (start > 0 && tokens[start - 1]?.kind === 'placeholder') start--
	return {
		whole: tokens.slice(0, start),
		numerator: tokens.slice(start, end + 1),
		preSlash: tokens.slice(end + 1),
	}
}

function splitDenominatorRight(tokens: readonly NumberFormatToken[]): {
	postSlash: readonly NumberFormatToken[]
	denominator: readonly NumberFormatToken[]
	after: readonly NumberFormatToken[]
} | null {
	let start = 0
	while (start < tokens.length && tokens[start]?.kind !== 'placeholder') start++
	if (start >= tokens.length) return null
	let end = start
	while (end + 1 < tokens.length && tokens[end + 1]?.kind === 'placeholder') end++
	return {
		postSlash: tokens.slice(0, start),
		denominator: tokens.slice(start, end + 1),
		after: tokens.slice(end + 1),
	}
}

function tryFormatFractionNumber(
	value: number,
	fmt: string,
	showSign: boolean,
): string | undefined {
	if (fmt.includes('.') || /[Ee]/.test(fmt)) return undefined
	const tokens = tokenizeNumberFormat(fmt)
	const slashIndex = tokens.findIndex(
		(token, index) =>
			token.kind === 'literal' &&
			token.text === '/' &&
			token.source === 'raw' &&
			hasPlaceholder(tokens.slice(0, index)) &&
			hasPlaceholder(tokens.slice(index + 1)),
	)
	if (slashIndex < 0) return undefined

	const left = tokens.slice(0, slashIndex)
	const right = tokens.slice(slashIndex + 1)
	const denominatorSplit = splitDenominatorRight(right)
	if (!denominatorSplit) return undefined

	const mixedSplit = splitMixedFractionLeft(left)
	const hasExplicitWholeSeparator =
		mixedSplit?.whole.some(
			(token) =>
				token.kind === 'literal' &&
				token.source !== 'underscore' &&
				token.text !== '' &&
				token.text !== '|',
		) ?? false
	const numeratorTokens = hasExplicitWholeSeparator && mixedSplit ? mixedSplit.numerator : left
	const wholeTokens = hasExplicitWholeSeparator && mixedSplit ? mixedSplit.whole : []
	const preSlashTokens = hasExplicitWholeSeparator && mixedSplit ? mixedSplit.preSlash : []
	const denominatorPlaceholders = denominatorSplit.denominator.filter(tokenHasPlaceholder).length
	const maxDenominator = 10 ** Math.max(1, denominatorPlaceholders) - 1
	const abs = Math.abs(value)
	const whole = hasExplicitWholeSeparator ? Math.floor(abs) : 0
	const fractionValue = hasExplicitWholeSeparator ? abs - whole : abs
	let { numerator, denominator } = bestFraction(fractionValue, maxDenominator)
	let renderedWhole = ''
	if (hasExplicitWholeSeparator) {
		if (numerator >= denominator) {
			numerator -= denominator
			renderedWhole = renderPlaceholderTokens(wholeTokens, String(whole + 1), {
				showZero: true,
				trimTrailingLiteralsWhenBlank: false,
			})
		} else {
			renderedWhole = renderPlaceholderTokens(wholeTokens, String(whole), {
				showZero: whole !== 0 || numerator === 0,
				trimTrailingLiteralsWhenBlank: numerator !== 0,
			})
		}
	}

	const sign = showSign ? '-' : ''
	if (numerator === 0) {
		if (hasExplicitWholeSeparator) {
			const numeratorHasRequiredZero = numeratorTokens.some(
				(token) => token.kind === 'placeholder' && token.char === '0',
			)
			if (numeratorHasRequiredZero) {
				return (
					sign +
					renderedWhole +
					renderPlaceholderTokens(numeratorTokens, '0', { showZero: true }) +
					literalText(preSlashTokens) +
					'/' +
					literalText(denominatorSplit.postSlash) +
					renderPlaceholderTokens(denominatorSplit.denominator, '1', { showZero: true }) +
					literalText(denominatorSplit.after)
				)
			}
			const hasQuestionPadding = [
				...wholeTokens,
				...numeratorTokens,
				...denominatorSplit.denominator,
			].some((token) => token.kind === 'placeholder' && token.char === '?')
			const strippedWhole = stripTrailingWholeSeparator(renderedWhole)
			if (hasQuestionPadding) {
				const zeroPadding = [
					blankPlaceholderText(numeratorTokens),
					literalText(preSlashTokens),
					'/',
					literalText(denominatorSplit.postSlash),
					blankPlaceholderText(denominatorSplit.denominator),
				].join('')
				return (
					sign +
					strippedWhole.text +
					strippedWhole.padding +
					blankDisplayWidth(zeroPadding) +
					literalText(denominatorSplit.after)
				)
			}
			return sign + strippedWhole.text + literalText(denominatorSplit.after)
		}
		return (
			sign +
			renderPlaceholderTokens(numeratorTokens, '0', { showZero: true }) +
			'/' +
			literalText(denominatorSplit.postSlash) +
			renderPlaceholderTokens(denominatorSplit.denominator, '1', { showZero: true }) +
			literalText(denominatorSplit.after)
		)
	}

	const renderedNumerator = renderPlaceholderTokens(numeratorTokens, String(numerator), {
		showZero: true,
	})
	const renderedDenominator = renderPlaceholderTokens(
		denominatorSplit.denominator,
		String(denominator),
		{
			showZero: true,
		},
	)
	return (
		sign +
		renderedWhole +
		renderedNumerator +
		literalText(preSlashTokens) +
		'/' +
		literalText(denominatorSplit.postSlash) +
		renderedDenominator +
		literalText(denominatorSplit.after)
	)
}

function tryFormatScientificNumber(
	value: number,
	fmt: string,
	showSign: boolean,
): string | undefined {
	const tokens = tokenizeNumberFormat(fmt)
	const markerIndex = tokens.findIndex((token, index) => {
		const next = tokens[index + 1]
		return (
			token.kind === 'literal' &&
			token.source === 'raw' &&
			/^[Ee]$/.test(token.text) &&
			next?.kind === 'literal' &&
			next.source === 'raw' &&
			(next.text === '+' || next.text === '-')
		)
	})
	if (markerIndex < 0) return undefined
	const signToken = tokens[markerIndex + 1]
	if (signToken?.kind !== 'literal') return undefined
	const mantissaTokens = tokens
		.slice(0, markerIndex)
		.filter((token) => !(token.kind === 'literal' && token.source === 'raw' && token.text === ','))
	const exponentTokens = tokens
		.slice(markerIndex + 2)
		.filter((token) => !(token.kind === 'literal' && token.source === 'raw' && token.text === ','))
	const mantissaPlaceholders = mantissaTokens.filter(tokenHasPlaceholder)
	const exponentPlaceholders = exponentTokens.filter(tokenHasPlaceholder)
	if (mantissaPlaceholders.length === 0 || exponentPlaceholders.length === 0) return undefined

	const decimalIndex = mantissaTokens.findIndex(
		(token) => token.kind === 'literal' && token.source === 'raw' && token.text === '.',
	)
	const integerTokens = decimalIndex >= 0 ? mantissaTokens.slice(0, decimalIndex) : mantissaTokens
	const decimalTokens = decimalIndex >= 0 ? mantissaTokens.slice(decimalIndex + 1) : []
	const integerPlaces = Math.max(1, integerTokens.filter(tokenHasPlaceholder).length)
	const decimalPlaces = decimalTokens.filter(tokenHasPlaceholder).length

	const abs = Math.abs(value)
	let exponent = abs === 0 ? 0 : Math.floor(Math.log10(abs) / integerPlaces) * integerPlaces
	let mantissa = abs === 0 ? 0 : abs / 10 ** exponent
	const overflowThreshold = 10 ** integerPlaces
	if (Number(mantissa.toFixed(decimalPlaces)) >= overflowThreshold) {
		exponent += integerPlaces
		mantissa = abs / 10 ** exponent
	}

	let [integerDigits = '0', decimalDigits = ''] = mantissa.toFixed(decimalPlaces).split('.')
	for (let i = decimalDigits.length - 1; i >= 0; i--) {
		const placeholder = decimalTokens.filter(tokenHasPlaceholder)[i]
		if (decimalDigits[i] !== '0' || placeholder?.char !== '#') break
		decimalDigits = decimalDigits.slice(0, -1)
	}

	const renderedInteger = renderPlaceholderTokens(integerTokens, integerDigits, { showZero: true })
	let renderedDecimal = ''
	let digitIndex = 0
	for (const token of decimalTokens) {
		if (token.kind === 'placeholder') {
			const digit = decimalDigits[digitIndex]
			digitIndex++
			if (digit !== undefined) renderedDecimal += digit
			else if (token.char === '0') renderedDecimal += '0'
			else if (token.char === '?') renderedDecimal += ' '
		} else {
			renderedDecimal += token.text
		}
	}
	if (!/[0-9]/.test(renderedDecimal)) renderedDecimal = ''

	const exponentSign =
		signToken.text === '+' ? (exponent >= 0 ? '+' : '-') : exponent < 0 ? '-' : ''
	const minExponentDigits = exponentPlaceholders.filter((token) => token.char === '0').length
	const exponentDigits = String(Math.abs(exponent)).padStart(minExponentDigits, '0')
	let exponentOut = ''
	let exponentDigitIndex = 0
	let exponentPlaceholderIndex = 0
	let signInserted = false
	for (const token of exponentTokens) {
		if (token.kind === 'placeholder') {
			if (!signInserted) {
				exponentOut += exponentSign
				signInserted = true
			}
			const remainingPlaceholderCount = exponentPlaceholders.length - exponentPlaceholderIndex
			const remainingDigits = exponentDigits.length - exponentDigitIndex
			const take = Math.max(1, remainingDigits - remainingPlaceholderCount + 1)
			const digit = exponentDigits.slice(exponentDigitIndex, exponentDigitIndex + take)
			exponentDigitIndex += take
			exponentPlaceholderIndex++
			if (digit !== '') exponentOut += digit
			else if (token.char === '0') exponentOut += '0'
			else if (token.char === '?') exponentOut += ' '
		} else {
			exponentOut += token.text
		}
	}
	if (!signInserted) exponentOut += exponentSign + exponentDigits

	const valueSign = showSign && value < 0 ? '-' : ''
	const decimalPoint = renderedDecimal === '' ? '' : '.'
	const marker = tokens[markerIndex]?.kind === 'literal' ? tokens[markerIndex].text : 'E'
	return valueSign + renderedInteger + decimalPoint + renderedDecimal + marker + exponentOut
}

function renderTextFormat(value: string, code: string): string | undefined {
	const rawSections = splitFormatSections(code.trim())
	const rawSection = rawSections.length >= 4 ? rawSections[3] : (rawSections[0] ?? code)
	const section = parseFormatSection(rawSection ?? '').format
	const { numFmt } = parseFormatSegments(section)
	if (!section.includes('@') && numFmt !== '') return undefined

	let result = ''
	for (let i = 0; i < section.length; i++) {
		const ch = section[i] ?? ''
		if (ch === '"') {
			const end = section.indexOf('"', i + 1)
			if (end === -1) {
				result += section.slice(i + 1)
				break
			}
			result += section.slice(i + 1, end)
			i = end
			continue
		}
		if (ch === '\\' && i + 1 < section.length) {
			result += section[i + 1] ?? ''
			i++
			continue
		}
		if ((ch === '_' || ch === '*') && i + 1 < section.length) {
			i++
			continue
		}
		result += ch === '@' ? value : ch
	}
	return result
}

function formatGeneralTextNumber(value: number): string {
	if (!Number.isFinite(value)) return String(value)
	if (Object.is(value, -0) || value === 0) return '0'
	const abs = Math.abs(value)
	if (abs >= 1e11 || abs < 1e-9) {
		return value
			.toExponential(14)
			.replace(/(\.\d*?)0+e/i, '$1E')
			.replace(/\.E/i, 'E')
			.replace(/e/i, 'E')
	}
	return String(value)
}

export function formatNumber(
	value: number,
	code: string,
	dateSystem: '1900' | '1904' = '1900',
): string {
	const { format: selectedFormat, autoSign, absValue } = selectNumericSection(value, code.trim())
	const fmt = selectedFormat.trim()

	if (fmt === '@') return String(value)

	if (fmt === '') return ''

	if (fmt === 'General') return formatGeneralTextNumber(value)

	if (isDateFormat(fmt)) return formatDate(value, fmt, dateSystem)

	const fraction = tryFormatFractionNumber(absValue, fmt, autoSign && value < 0)
	if (fraction !== undefined) return fraction

	const scientificValue = autoSign && value < 0 ? -absValue : absValue
	const scientific = tryFormatScientificNumber(scientificValue, fmt, autoSign && value < 0)
	if (scientific !== undefined) return scientific

	const { before, numFmt, after } = parseFormatSegments(fmt)
	if (numFmt === '') return before + after
	const percentCount = (numFmt.match(/%/g) || []).length
	const withoutPercent = numFmt.replace(/%/g, '')
	let trailingCommas = 0
	for (let i = withoutPercent.length - 1; i >= 0 && withoutPercent[i] === ','; i--) {
		trailingCommas++
	}
	const numericFmt = withoutPercent.slice(0, withoutPercent.length - trailingCommas)
	const hasComma = numericFmt.includes(',')
	const dotIdx = numericFmt.indexOf('.')
	const intFmt = dotIdx >= 0 ? numericFmt.slice(0, dotIdx) : numericFmt
	const decFmt = dotIdx >= 0 ? numericFmt.slice(dotIdx + 1) : ''

	const minIntDigits = (intFmt.match(/0/g) || []).length
	const decPlaces = decFmt.replace(/[^0#?]/g, '').length

	const scaledValue = (absValue * 100 ** percentCount) / 1000 ** trailingCommas
	const abs = Math.abs(scaledValue)
	const roundedAbs = Math.round((abs + Number.EPSILON) * 10 ** decPlaces) / 10 ** decPlaces
	const fixed = roundedAbs.toFixed(decPlaces)
	const [rawInt = '', rawDec] = fixed.split('.')
	const sign = autoSign && value < 0 ? '-' : ''

	let intStr = rawInt
	if (intStr.length < minIntDigits) {
		intStr = intStr.padStart(minIntDigits, '0')
	}
	if (minIntDigits === 0 && abs === 0 && intFmt.length > 0) {
		intStr = ''
	}

	let decStr = rawDec ?? ''
	if (decFmt.length > 0) {
		while (decStr.length > 0 && decFmt[decStr.length - 1] === '#' && decStr.endsWith('0')) {
			decStr = decStr.slice(0, -1)
		}
	}

	if (hasComma) {
		intStr = intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
	}
	if (hasComma && intFmt.includes('?')) {
		const groupingCommas = (intStr.match(/,/g) || []).length
		const minDisplayWidth = intFmt.length + groupingCommas
		if (groupingCommas > 0 && intStr.length < minDisplayWidth) {
			intStr = intStr.padStart(minDisplayWidth, ' ')
		}
	}

	return before + sign + intStr + (decStr ? `.${decStr}` : '') + after + '%'.repeat(percentCount)
}

export function formatTextValue(
	value: CellValue,
	code: string,
	dateSystem: '1900' | '1904' = '1900',
): CellValue {
	if (isError(value)) return value
	if (code.trim() === '') return stringValue('')
	const scalar = topLeftScalar(value)
	if (
		code.trim().toLowerCase() === 'general' &&
		scalar.kind !== 'number' &&
		scalar.kind !== 'date'
	) {
		return stringValue(cvStr(scalar))
	}
	if (scalar.kind === 'string' || scalar.kind === 'boolean' || scalar.kind === 'richText') {
		const formatted = renderTextFormat(cvStr(scalar), code)
		if (formatted !== undefined) return stringValue(formatted)
	}
	const n = toNum(value)
	if (typeof n !== 'number') return errorValue('#VALUE!')
	return stringValue(formatNumber(n, code, dateSystem))
}

const MAC_ROMAN_CHAR_OVERRIDES = new Map<number, string>([[240, '\uf8ff']])
const MAC_ROMAN_CODE_OVERRIDES = new Map<string, number>([['\uf8ff', 240]])

function leftText(args: EvalArg[]): CellValue {
	const s = strArg(args[0])
	if (typeof s !== 'string') return s
	const n = args.length >= 2 ? numArg(args[1]) : 1
	if (typeof n !== 'number') return n
	if (n < 0) return errorValue('#VALUE!')
	return stringValue(s.slice(0, Math.trunc(n)))
}

function rightText(args: EvalArg[]): CellValue {
	const s = strArg(args[0])
	if (typeof s !== 'string') return s
	const n = args.length >= 2 ? numArg(args[1]) : 1
	if (typeof n !== 'number') return n
	if (n < 0) return errorValue('#VALUE!')
	const count = Math.trunc(n)
	if (count === 0) return stringValue('')
	return stringValue(count >= s.length ? s : s.slice(-count))
}

function midText(args: EvalArg[]): CellValue {
	const s = strArg(args[0])
	if (typeof s !== 'string') return s
	const start = numArg(args[1])
	if (typeof start !== 'number') return start
	const len = numArg(args[2])
	if (typeof len !== 'number') return len
	if (start < 1 || len < 0) return errorValue('#VALUE!')
	const st = Math.trunc(start) - 1
	return stringValue(s.slice(st, st + Math.trunc(len)))
}

function lenText(args: EvalArg[]): CellValue {
	const s = strArg(args[0])
	if (typeof s !== 'string') return s
	return numberValue(s.length)
}

function findText(args: EvalArg[], caseSensitive: boolean): CellValue {
	const findText = strArg(args[0])
	if (typeof findText !== 'string') return findText
	const within = strArg(args[1])
	if (typeof within !== 'string') return within
	const startNum = args.length >= 3 ? numArg(args[2]) : 1
	if (typeof startNum !== 'number') return startNum
	if (startNum < 1) return errorValue('#VALUE!')
	const start = Math.trunc(startNum) - 1
	const haystack = caseSensitive ? within : within.toLowerCase()
	const needle = caseSensitive ? findText : findText.toLowerCase()
	const idx = haystack.indexOf(needle, start)
	return idx === -1 ? errorValue('#VALUE!') : numberValue(idx + 1)
}

function replaceText(args: EvalArg[]): CellValue {
	const text = strArg(args[0])
	if (typeof text !== 'string') return text
	const startN = numArg(args[1])
	if (typeof startN !== 'number') return startN
	const numChars = numArg(args[2])
	if (typeof numChars !== 'number') return numChars
	const newT = strArg(args[3])
	if (typeof newT !== 'string') return newT
	const start = Math.trunc(startN)
	const count = Math.trunc(numChars)
	if (start < 1 || count < 0) return errorValue('#VALUE!')
	const s = start - 1
	return stringValue(text.slice(0, s) + newT + text.slice(s + count))
}

function ascText(args: EvalArg[]): CellValue {
	const text = strArg(args[0])
	if (typeof text !== 'string') return text
	let result = ''
	for (const char of text) {
		const code = char.codePointAt(0) ?? 0
		if (code === 0x3000) {
			result += ' '
		} else if (code >= 0xff01 && code <= 0xff5e) {
			result += String.fromCharCode(code - 0xfee0)
		} else {
			result += char
		}
	}
	return stringValue(result)
}

export const textFunctions: FunctionDef[] = [
	fn('CONCATENATE', 1, 255, (args) => {
		let result = ''
		for (const arg of args) {
			const v = arg.value ?? EMPTY
			if (isError(v)) return v
			result += cvStr(v)
		}
		return stringValue(result)
	}),

	fn('CONCAT', 1, 255, (args) => {
		let result = ''
		for (const arg of args) {
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						result += cvStr(cell)
					}
				}
			} else {
				const v = arg.value ?? EMPTY
				if (isError(v)) return v
				result += cvStr(v)
			}
		}
		return stringValue(result)
	}),

	fn('TEXTJOIN', 3, 255, (args) => {
		const dv = args[0]?.value ?? EMPTY
		if (isError(dv)) return dv
		const delim = cvStr(dv)

		const iev = args[1]?.value ?? EMPTY
		if (isError(iev)) return iev
		const ignoreEmpty = iev.kind === 'boolean' ? iev.value : iev.kind !== 'empty'

		const parts: string[] = []
		for (let i = 2; i < args.length; i++) {
			const arg = args[i]
			if (!arg) continue
			if (arg.kind === 'range' && arg.values) {
				for (const row of arg.values) {
					for (const cell of row) {
						if (isError(cell)) return cell
						if (ignoreEmpty && (isEmpty(cell) || (cell.kind === 'string' && cell.value === '')))
							continue
						parts.push(cvStr(cell))
					}
				}
			} else {
				const v = arg.value ?? EMPTY
				if (isError(v)) return v
				if (ignoreEmpty && (isEmpty(v) || (v.kind === 'string' && v.value === ''))) continue
				parts.push(cvStr(v))
			}
		}
		return stringValue(parts.join(delim))
	}),

	fn('LEFT', 1, 2, leftText),
	fn('LEFTB', 1, 2, leftText),

	fn('RIGHT', 1, 2, rightText),
	fn('RIGHTB', 1, 2, rightText),

	fn('MID', 3, 3, midText),
	fn('MIDB', 3, 3, midText),

	fn('LEN', 1, 1, lenText),
	fn('LENB', 1, 1, lenText),

	fn('TRIM', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		return stringValue(s.trim().replace(/\s+/g, ' '))
	}),

	fn('UPPER', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		return stringValue(s.toUpperCase())
	}),

	fn('LOWER', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		return stringValue(s.toLowerCase())
	}),

	fn('EXACT', 2, 2, (args) => {
		const left = strArg(args[0])
		if (typeof left !== 'string') return left
		const right = strArg(args[1])
		if (typeof right !== 'string') return right
		return booleanValue(left === right)
	}),

	fn('PROPER', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		return stringValue(proper(s))
	}),

	fn('FIND', 2, 3, (args) => findText(args, true)),
	fn('FINDB', 2, 3, (args) => findText(args, true)),

	fn('SEARCH', 2, 3, (args) => findText(args, false)),
	fn('SEARCHB', 2, 3, (args) => findText(args, false)),

	fn('TEXTBEFORE', 2, 6, (args) => {
		const parsed = textSplitPoint(args)
		if ('kind' in parsed) return parsed
		const match = findTextSlice(parsed.text, parsed.delimiter, parsed.instanceNum, parsed.matchMode)
		if (!match) {
			const matchEnd = args[4] ? numArg(args[4]) : 0
			if (typeof matchEnd !== 'number') return matchEnd
			if (Math.trunc(matchEnd) !== 0) return stringValue(parsed.text)
			return (args[5]?.value ?? errorValue('#N/A')) as CellValue
		}
		return stringValue(parsed.text.slice(0, match.start))
	}),

	fn('TEXTAFTER', 2, 6, (args) => {
		const parsed = textSplitPoint(args)
		if ('kind' in parsed) return parsed
		const match = findTextSlice(parsed.text, parsed.delimiter, parsed.instanceNum, parsed.matchMode)
		if (!match) {
			const matchEnd = args[4] ? numArg(args[4]) : 0
			if (typeof matchEnd !== 'number') return matchEnd
			if (Math.trunc(matchEnd) !== 0) return stringValue('')
			return (args[5]?.value ?? errorValue('#N/A')) as CellValue
		}
		return stringValue(parsed.text.slice(match.end))
	}),

	fn('TEXTSPLIT', 2, 6, (args) => {
		const text = strArg(args[0])
		if (typeof text !== 'string') return text
		const colDelimiter = strArg(args[1])
		if (typeof colDelimiter !== 'string') return colDelimiter
		const rowDelimiter = args[2] ? strArg(args[2]) : undefined
		if (typeof rowDelimiter !== 'string' && rowDelimiter !== undefined) return rowDelimiter
		const ignoreEmptyValue = args[3]?.value ?? booleanValue(false)
		if (isError(ignoreEmptyValue)) return ignoreEmptyValue
		const ignoreEmpty =
			ignoreEmptyValue.kind === 'boolean'
				? ignoreEmptyValue.value
				: ignoreEmptyValue.kind === 'number'
					? ignoreEmptyValue.value !== 0
					: ignoreEmptyValue.kind !== 'empty'
		const matchModeRaw = args[4] ? numArg(args[4]) : 0
		if (typeof matchModeRaw !== 'number') return matchModeRaw
		const matchMode = Math.trunc(matchModeRaw)
		if (matchMode !== 0 && matchMode !== 1) return errorValue('#VALUE!')
		const padWith = topLeftScalar(args[5]?.value ?? errorValue('#N/A'))

		const rowParts = rowDelimiter ? splitByDelimiter(text, rowDelimiter, matchMode) : [text]
		const filteredRows = ignoreEmpty ? rowParts.filter((part) => part !== '') : rowParts
		const splitRows = filteredRows.map((rowText) => {
			const cols = splitByDelimiter(rowText, colDelimiter, matchMode)
			return ignoreEmpty ? cols.filter((part) => part !== '') : cols
		})
		if (splitRows.length === 0) return arrayValue([[padWith]])
		const maxCols = splitRows.reduce((max, row) => Math.max(max, row.length), 0)
		return arrayValue(
			splitRows.map((row): ScalarCellValue[] =>
				Array.from({ length: maxCols }, (_, index): ScalarCellValue => {
					const value = row[index]
					return value === undefined ? padWith : topLeftScalar(stringValue(value))
				}),
			),
		)
	}),

	fn('SUBSTITUTE', 3, 4, (args) => {
		const text = strArg(args[0])
		if (typeof text !== 'string') return text
		const oldT = strArg(args[1])
		if (typeof oldT !== 'string') return oldT
		const newT = strArg(args[2])
		if (typeof newT !== 'string') return newT
		if (oldT === '') return stringValue(text)

		if (args.length >= 4) {
			const inst = numArg(args[3])
			if (typeof inst !== 'number') return inst
			const n = Math.trunc(inst)
			if (n < 1) return errorValue('#VALUE!')
			let count = 0
			let pos = 0
			let result = ''
			while (pos <= text.length) {
				const idx = text.indexOf(oldT, pos)
				if (idx === -1) {
					result += text.slice(pos)
					break
				}
				count++
				if (count === n) {
					result += text.slice(pos, idx) + newT + text.slice(idx + oldT.length)
					return stringValue(result)
				}
				result += text.slice(pos, idx + oldT.length)
				pos = idx + (oldT.length || 1)
			}
			return stringValue(result)
		}

		return stringValue(oldT === '' ? text : text.split(oldT).join(newT))
	}),

	fn('REPLACE', 4, 4, replaceText),
	fn('REPLACEB', 4, 4, replaceText),

	fn('TEXT', 2, 2, (args, ctx) => {
		const v = args[0]?.value ?? EMPTY
		const fmt = strArg(args[1])
		if (typeof fmt !== 'string') return fmt
		return formatTextValue(v, fmt, ctx?.dateSystem ?? '1900')
	}),

	fn('VALUE', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		const trimmed = s.trim()
		if (trimmed === '') return numberValue(0)
		if (trimmed.endsWith('%')) {
			const n = Number(trimmed.slice(0, -1))
			return Number.isNaN(n) ? errorValue('#VALUE!') : numberValue(n / 100)
		}
		const cleaned = trimmed.replace(/[$€£¥,]/g, '')
		const n = Number(cleaned)
		return Number.isNaN(n) ? errorValue('#VALUE!') : numberValue(n)
	}),

	fn('CHAR', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const code = Math.trunc(n)
		if (code < 1 || code > 65535) return errorValue('#VALUE!')
		const override = MAC_ROMAN_CHAR_OVERRIDES.get(code)
		if (override) return stringValue(override)
		return stringValue(String.fromCharCode(code))
	}),

	fn('CODE', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		if (s.length === 0) return errorValue('#VALUE!')
		const override = MAC_ROMAN_CODE_OVERRIDES.get(s[0] ?? '')
		if (override !== undefined) return numberValue(override)
		return numberValue(s.charCodeAt(0))
	}),

	fn('ASC', 1, 1, ascText),
	fn('PHONETIC', 1, 1, () => errorValue('#N/A')),

	fn('REPT', 2, 2, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		const n = numArg(args[1])
		if (typeof n !== 'number') return n
		const times = Math.trunc(n)
		if (times < 0) return errorValue('#VALUE!')
		return stringValue(s.repeat(times))
	}),

	fn('CLEAN', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		return stringValue([...s].filter((c) => c.charCodeAt(0) > 31).join(''))
	}),

	fn('T', 1, 1, (args) => {
		const v = topLeftScalar(args[0]?.value ?? EMPTY)
		if (v.kind === 'error') return v
		if (v.kind === 'string') return stringValue(v.value)
		if (v.kind === 'richText') return stringValue(v.runs.map((r) => r.text).join(''))
		return stringValue('')
	}),

	fn('UNICHAR', 1, 1, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const code = Math.trunc(n)
		if (code < 1 || code > 1114111) return errorValue('#VALUE!')
		return stringValue(String.fromCodePoint(code))
	}),

	fn('UNICODE', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		if (s.length === 0) return errorValue('#VALUE!')
		return numberValue(s.codePointAt(0) ?? s.charCodeAt(0))
	}),

	fn('FIXED', 1, 3, (args) => {
		const n = numArg(args[0])
		if (typeof n !== 'number') return n
		const decimals = args.length >= 2 ? numArg(args[1]) : 2
		if (typeof decimals !== 'number') return decimals
		const ncv = args.length >= 3 ? topLeftScalar(args[2]?.value ?? EMPTY) : EMPTY
		if (isError(ncv)) return ncv
		const noCommas =
			ncv.kind === 'boolean' ? ncv.value : ncv.kind === 'number' ? ncv.value !== 0 : false
		const d = Math.trunc(decimals)
		let text: string
		if (d < 0) {
			const factor = 10 ** -d
			text = String(Math.round(n / factor) * factor)
		} else {
			text = n.toFixed(d)
		}
		if (!noCommas) {
			const neg = text.startsWith('-')
			const raw = neg ? text.slice(1) : text
			const [intPart = '', decPart] = raw.split('.')
			const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
			text = (neg ? '-' : '') + withCommas + (decPart !== undefined ? `.${decPart}` : '')
		}
		return stringValue(text)
	}),

	fn('DOLLAR', 1, 2, (args) => {
		const n = dollarNumArg(args[0])
		if (typeof n !== 'number') return n
		const decimals = args.length >= 2 ? numArg(args[1]) : 2
		if (typeof decimals !== 'number') return decimals
		const d = Math.trunc(decimals)
		let text: string
		if (d < 0) {
			const factor = 10 ** -d
			text = String(Math.round(n / factor) * factor)
		} else {
			text = n.toFixed(d)
		}
		const neg = text.startsWith('-')
		const raw = neg ? text.slice(1) : text
		const [intPart = '', decPart] = raw.split('.')
		const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
		const formatted = `$${withCommas}${decPart !== undefined ? `.${decPart}` : ''}`
		text = neg ? `(${formatted})` : formatted
		return stringValue(text)
	}),

	fn('ARRAYTOTEXT', 1, 2, (args) => {
		const format = args.length >= 2 ? numArg(args[1]) : 0
		if (typeof format !== 'number') return format
		const strict = Math.trunc(format) === 1
		const arg = args[0]
		if (!arg) return stringValue('')
		const rows: readonly (readonly CellValue[])[] | null =
			arg.kind === 'range' && arg.values
				? arg.values
				: arg.value.kind === 'array'
					? arg.value.rows
					: null
		if (rows) {
			if (strict) {
				const rowStrs: string[] = []
				for (const row of rows) {
					const cellStrs: string[] = []
					for (const cell of row) {
						if (isError(cell)) return cell
						cellStrs.push(cell.kind === 'string' ? `"${cell.value}"` : cvStr(cell))
					}
					rowStrs.push(cellStrs.join(','))
				}
				return stringValue(`{${rowStrs.join(';')}}`)
			}
			const parts: string[] = []
			for (const row of rows) {
				for (const cell of row) {
					if (isError(cell)) return cell
					if (!isEmpty(cell)) parts.push(cvStr(cell))
				}
			}
			return stringValue(parts.join(', '))
		}
		const v = topLeftScalar(arg.value ?? EMPTY)
		if (isError(v)) return v
		if (strict && v.kind === 'string') return stringValue(`"${v.value}"`)
		return stringValue(cvStr(v))
	}),

	fn('VALUETOTEXT', 1, 2, (args) => {
		const v = topLeftScalar(args[0]?.value ?? EMPTY)
		if (isError(v)) return v
		const format = args.length >= 2 ? numArg(args[1]) : 0
		if (typeof format !== 'number') return format
		if (Math.trunc(format) === 1 && v.kind === 'string') return stringValue(`"${v.value}"`)
		return stringValue(cvStr(v))
	}),

	fn('HYPERLINK', 1, 2, (args) => {
		const url = strArg(args[0])
		if (typeof url !== 'string') return url
		const display = args.length > 1 ? strArg(args[1]) : url
		if (typeof display !== 'string') return display
		return stringValue(display)
	}),

	fn('ENCODEURL', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		return stringValue(encodeURIComponent(s))
	}),

	fn('REGEXTEST', 2, 2, (args) => {
		const text = strArg(args[0])
		if (typeof text !== 'string') return text
		const pattern = strArg(args[1])
		if (typeof pattern !== 'string') return pattern
		try {
			return booleanValue(new RegExp(pattern).test(text))
		} catch {
			return errorValue('#VALUE!')
		}
	}),

	fn('REGEXEXTRACT', 2, 3, (args) => {
		const text = strArg(args[0])
		if (typeof text !== 'string') return text
		const pattern = strArg(args[1])
		if (typeof pattern !== 'string') return pattern
		const returnMode = args.length > 2 ? numArg(args[2]) : 0
		if (typeof returnMode !== 'number') return returnMode
		try {
			const re = new RegExp(pattern, 'g')
			if (Math.trunc(returnMode) === 0) {
				const match = re.exec(text)
				return match ? stringValue(match[0] as string) : errorValue('#N/A')
			}
			const matches: string[] = []
			for (const m of text.matchAll(re)) matches.push(m[0] as string)
			if (matches.length === 0) return errorValue('#N/A')
			if (matches.length === 1) return stringValue(matches[0] as string)
			return arrayValue(matches.map((v) => [topLeftScalar(stringValue(v))]))
		} catch {
			return errorValue('#VALUE!')
		}
	}),

	fn('REGEXREPLACE', 3, 3, (args) => {
		const text = strArg(args[0])
		if (typeof text !== 'string') return text
		const pattern = strArg(args[1])
		if (typeof pattern !== 'string') return pattern
		const replacement = strArg(args[2])
		if (typeof replacement !== 'string') return replacement
		try {
			return stringValue(text.replace(new RegExp(pattern, 'g'), replacement))
		} catch {
			return errorValue('#VALUE!')
		}
	}),

	fn('NUMBERVALUE', 1, 3, (args) => {
		const text = strArg(args[0])
		if (typeof text !== 'string') return text
		const decimalSep = args[1] ? strArg(args[1]) : '.'
		if (typeof decimalSep !== 'string') return decimalSep
		const groupSep = args[2] ? strArg(args[2]) : ','
		if (typeof groupSep !== 'string') return groupSep
		let cleaned = text.trim()
		cleaned = cleaned.split(groupSep).join('')
		cleaned = cleaned.split(decimalSep).join('.')
		const n = Number(cleaned)
		return Number.isNaN(n) ? errorValue('#VALUE!') : numberValue(n)
	}),
]
