import type { CellValue } from '@ascend/schema'
import {
	booleanValue,
	EMPTY,
	errorValue,
	isEmpty,
	isError,
	numberValue,
	stringValue,
} from '@ascend/schema'
import type { EvalArg, FunctionDef } from './index.ts'

function fn(
	name: string,
	minArgs: number,
	maxArgs: number,
	evaluate: (args: EvalArg[]) => CellValue,
): FunctionDef {
	return { name, minArgs, maxArgs, volatile: false, evaluate }
}

function toNum(v: CellValue): number | CellValue {
	switch (v.kind) {
		case 'empty':
			return 0
		case 'number':
			return v.value
		case 'string': {
			if (v.value.trim() === '') return 0
			const n = Number(v.value)
			return Number.isNaN(n) ? errorValue('#VALUE!') : n
		}
		case 'boolean':
			return v.value ? 1 : 0
		case 'error':
			return v
		case 'date':
			return v.serial
		case 'richText':
			return errorValue('#VALUE!')
	}
}

function numArg(arg: EvalArg | undefined): number | CellValue {
	return toNum(arg?.value ?? EMPTY)
}

function cvStr(v: CellValue): string {
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

function formatNumber(value: number, code: string): string {
	const fmt = code.trim()

	if (fmt.includes('%')) {
		const pctFmt = fmt.replace(/%/g, '')
		const dec = pctFmt.includes('.') ? (pctFmt.split('.')[1] || '').replace(/[^0#]/g, '').length : 0
		return `${(value * 100).toFixed(dec)}%`
	}

	const hasComma = fmt.includes(',')
	const dec = fmt.includes('.') ? (fmt.split('.')[1] || '').replace(/[^0#]/g, '').length : 0
	const abs = Math.abs(value)
	const fixed = abs.toFixed(dec)
	const [intPart = '', decPart] = fixed.split('.')
	const sign = value < 0 ? '-' : ''

	if (hasComma) {
		const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
		return sign + withCommas + (decPart ? `.${decPart}` : '')
	}

	return sign + intPart + (decPart ? `.${decPart}` : '')
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
]
