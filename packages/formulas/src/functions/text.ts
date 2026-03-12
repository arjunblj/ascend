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
import type { EvalArg, FunctionDef } from './index.ts'
import { numArg, toNum } from './math/helpers.ts'

function fn(
	name: string,
	minArgs: number,
	maxArgs: number,
	evaluate: (args: EvalArg[]) => CellValue,
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

const DATE_TOKEN_RE = /yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|am\/pm|A\/P|a\/p/g
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
	const stripped = fmt.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '')
	return DATE_TOKEN_RE.test(stripped)
}

function serialToTime(serial: number): { hours: number; minutes: number; seconds: number } {
	const frac = serial - Math.floor(serial)
	const totalSeconds = Math.round(frac * 86400)
	const hours = Math.floor(totalSeconds / 3600)
	const minutes = Math.floor((totalSeconds % 3600) / 60)
	const seconds = totalSeconds % 60
	return { hours, minutes, seconds }
}

function dayOfWeek(serial: number): number {
	return (Math.floor(serial) + 6) % 7
}

function formatDate(serial: number, fmt: string): string {
	const parts = serialToDate(serial)
	if (!parts) return ''
	const { hours, minutes, seconds } = serialToTime(serial)

	const hasAmPm = /AM\/PM|am\/pm|A\/P|a\/p/.test(fmt)
	const hour12 = hours === 0 ? 12 : hours > 12 ? hours - 12 : hours
	const ampm = hours < 12 ? 'AM' : 'PM'

	let result = ''
	let i = 0
	let prevTokenIsTime = false
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
		const tokenMatch = remaining.match(/^(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s)/i)
		if (tokenMatch) {
			const tok = tokenMatch[0] as string
			const tokLower = tok.toLowerCase()
			switch (tokLower) {
				case 'yyyy':
					result += String(parts.year)
					prevTokenIsTime = false
					break
				case 'yy':
					result += String(parts.year % 100).padStart(2, '0')
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
					result += DAY_NAMES[dayOfWeek(serial)] ?? ''
					prevTokenIsTime = false
					break
				case 'ddd':
					result += (DAY_NAMES[dayOfWeek(serial)] ?? '').slice(0, 3)
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
					result += String(seconds).padStart(2, '0')
					prevTokenIsTime = false
					break
				case 's':
					result += String(seconds)
					prevTokenIsTime = false
					break
			}
			i += tok.length
			continue
		}

		result += fmt[i]
		i++
	}
	return result
}

const NUM_FMT_CHARS = new Set(['0', '#', '.', ','])

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

function formatNumber(value: number, code: string): string {
	const fmt = code.trim()

	if (fmt === '@') return String(value)

	if (fmt === '' || fmt === 'General') return String(value)

	if (isDateFormat(fmt)) return formatDate(value, fmt)

	if (/[eE]\+/i.test(fmt)) {
		const decMatch = fmt.match(/\.(0+)[eE]\+/i)
		const dec = decMatch ? (decMatch[1] as string).length : 2
		return value.toExponential(dec).replace('e+', 'E+').replace('e-', 'E-')
	}

	if (fmt.includes('%')) {
		const pctFmt = fmt.replace(/%/g, '').replace(/"[^"]*"/g, '')
		const dec = pctFmt.includes('.') ? (pctFmt.split('.')[1] || '').replace(/[^0#]/g, '').length : 0
		return `${(value * 100).toFixed(dec)}%`
	}

	const { before, numFmt, after } = parseFormatSegments(fmt)

	const hasComma = numFmt.includes(',')
	const dotIdx = numFmt.indexOf('.')
	const intFmt = dotIdx >= 0 ? numFmt.slice(0, dotIdx) : numFmt
	const decFmt = dotIdx >= 0 ? numFmt.slice(dotIdx + 1) : ''

	const minIntDigits = (intFmt.match(/0/g) || []).length
	const decPlaces = decFmt.replace(/[^0#]/g, '').length

	const abs = Math.abs(value)
	const fixed = abs.toFixed(decPlaces)
	const [rawInt = '', rawDec] = fixed.split('.')
	const sign = value < 0 ? '-' : ''

	let intStr = rawInt
	if (intStr.length < minIntDigits) {
		intStr = intStr.padStart(minIntDigits, '0')
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

	return before + sign + intStr + (decStr ? `.${decStr}` : '') + after
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

	fn('LEFT', 1, 2, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		const n = args.length >= 2 ? numArg(args[1]) : 1
		if (typeof n !== 'number') return n
		if (n < 0) return errorValue('#VALUE!')
		return stringValue(s.slice(0, Math.trunc(n)))
	}),

	fn('RIGHT', 1, 2, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		const n = args.length >= 2 ? numArg(args[1]) : 1
		if (typeof n !== 'number') return n
		if (n < 0) return errorValue('#VALUE!')
		const count = Math.trunc(n)
		return stringValue(count >= s.length ? s : s.slice(-count))
	}),

	fn('MID', 3, 3, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		const start = numArg(args[1])
		if (typeof start !== 'number') return start
		const len = numArg(args[2])
		if (typeof len !== 'number') return len
		if (start < 1 || len < 0) return errorValue('#VALUE!')
		const st = Math.trunc(start) - 1
		return stringValue(s.slice(st, st + Math.trunc(len)))
	}),

	fn('LEN', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		return numberValue(s.length)
	}),

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

	fn('FIND', 2, 3, (args) => {
		const findText = strArg(args[0])
		if (typeof findText !== 'string') return findText
		const within = strArg(args[1])
		if (typeof within !== 'string') return within
		const startNum = args.length >= 3 ? numArg(args[2]) : 1
		if (typeof startNum !== 'number') return startNum
		if (startNum < 1) return errorValue('#VALUE!')
		const idx = within.indexOf(findText, Math.trunc(startNum) - 1)
		return idx === -1 ? errorValue('#VALUE!') : numberValue(idx + 1)
	}),

	fn('SEARCH', 2, 3, (args) => {
		const findText = strArg(args[0])
		if (typeof findText !== 'string') return findText
		const within = strArg(args[1])
		if (typeof within !== 'string') return within
		const startNum = args.length >= 3 ? numArg(args[2]) : 1
		if (typeof startNum !== 'number') return startNum
		if (startNum < 1) return errorValue('#VALUE!')
		const idx = within.toLowerCase().indexOf(findText.toLowerCase(), Math.trunc(startNum) - 1)
		return idx === -1 ? errorValue('#VALUE!') : numberValue(idx + 1)
	}),

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

	fn('REPLACE', 4, 4, (args) => {
		const text = strArg(args[0])
		if (typeof text !== 'string') return text
		const startN = numArg(args[1])
		if (typeof startN !== 'number') return startN
		const numChars = numArg(args[2])
		if (typeof numChars !== 'number') return numChars
		const newT = strArg(args[3])
		if (typeof newT !== 'string') return newT
		const s = Math.trunc(startN) - 1
		return stringValue(text.slice(0, s) + newT + text.slice(s + Math.trunc(numChars)))
	}),

	fn('TEXT', 2, 2, (args) => {
		const v = args[0]?.value ?? EMPTY
		if (isError(v)) return v
		const fmt = strArg(args[1])
		if (typeof fmt !== 'string') return fmt
		const n = toNum(v)
		if (typeof n !== 'number') return errorValue('#VALUE!')
		return stringValue(formatNumber(n, fmt))
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
		return stringValue(String.fromCharCode(code))
	}),

	fn('CODE', 1, 1, (args) => {
		const s = strArg(args[0])
		if (typeof s !== 'string') return s
		if (s.length === 0) return errorValue('#VALUE!')
		return numberValue(s.charCodeAt(0))
	}),

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
		const n = numArg(args[0])
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
		text = `${neg ? '-' : ''}$${withCommas}${decPart !== undefined ? `.${decPart}` : ''}`
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
