import type { CellValue } from '@ascend/schema'
import { errorValue, numberValue, stringValue } from '@ascend/schema'
import { cellOf, type EvalArg, numArg, registerFunction } from './registry.ts'

function strArg(arg: EvalArg | undefined): string | CellValue {
	const v = cellOf(arg)
	if (v.kind === 'error') return v
	if (v.kind === 'string') return v.value
	if (v.kind === 'number') return String(Math.trunc(v.value))
	if (v.kind === 'boolean') return v.value ? '1' : '0'
	if (v.kind === 'empty') return ''
	return errorValue('#VALUE!')
}

const BIN_RE = /^[01]{1,10}$/
function parseBin(s: string): number | null {
	const t = s.trim().toUpperCase()
	if (!BIN_RE.test(t)) return null
	if (t.length > 10) return null
	let n = 0
	for (let i = 0; i < t.length; i++) {
		n = (n << 1) | (t[i] === '1' ? 1 : 0)
	}
	if (t.length === 10 && t[0] === '1') {
		n = n - (1 << 10)
	}
	return n
}

function bin2dec(args: EvalArg[]): CellValue {
	const s = strArg(args[0])
	if (typeof s !== 'string') return s
	const n = parseBin(s)
	return n === null ? errorValue('#NUM!') : numberValue(n)
}

const DEC2BIN_MIN = -512
const DEC2BIN_MAX = 511
function dec2bin(args: EvalArg[]): CellValue {
	const n = numArg(args[0])
	if (typeof n !== 'number') return n
	const x = Math.trunc(n)
	if (x < DEC2BIN_MIN || x > DEC2BIN_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + (1 << 10)) & 0x3ff : x
	let s = abs.toString(2)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

const HEX_RE = /^[0-9A-Fa-f]{1,10}$/
function parseHex(s: string): number | null {
	const t = s.trim().toUpperCase()
	if (!HEX_RE.test(t)) return null
	if (t.length > 10) return null
	let n = Number.parseInt(t, 16)
	if (t.length === 10 && (t[0] ?? '') >= '8') {
		n = n - 0x10000000000
	}
	return n
}

function hex2dec(args: EvalArg[]): CellValue {
	const s = strArg(args[0])
	if (typeof s !== 'string') return s
	const n = parseHex(s)
	return n === null ? errorValue('#NUM!') : numberValue(n)
}

const DEC2HEX_MIN = -549_755_813_888
const DEC2HEX_MAX = 549_755_813_887
function dec2hex(args: EvalArg[]): CellValue {
	const n = numArg(args[0])
	if (typeof n !== 'number') return n
	const x = Math.trunc(n)
	if (x < DEC2HEX_MIN || x > DEC2HEX_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + 0x10000000000) >>> 0 : x
	let s = abs.toString(16).toUpperCase()
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

const OCT_RE = /^[0-7]{1,10}$/
function parseOct(s: string): number | null {
	const t = s.trim()
	if (!OCT_RE.test(t)) return null
	if (t.length > 10) return null
	let n = Number.parseInt(t, 8)
	if (t.length === 10 && (t[0] ?? '') >= '4') {
		n = n - 0x40000000
	}
	return n
}

function oct2dec(args: EvalArg[]): CellValue {
	const s = strArg(args[0])
	if (typeof s !== 'string') return s
	const n = parseOct(s)
	return n === null ? errorValue('#NUM!') : numberValue(n)
}

const DEC2OCT_MIN = -536_870_912
const DEC2OCT_MAX = 536_870_911
function dec2oct(args: EvalArg[]): CellValue {
	const n = numArg(args[0])
	if (typeof n !== 'number') return n
	const x = Math.trunc(n)
	if (x < DEC2OCT_MIN || x > DEC2OCT_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + 0x40000000) >>> 0 : x
	let s = abs.toString(8)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function bin2hex(args: EvalArg[]): CellValue {
	const dec = bin2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const x = Math.trunc(dec.value)
	const abs = x < 0 ? (x + 0x10000000000) >>> 0 : x
	let s = abs.toString(16).toUpperCase()
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function bin2oct(args: EvalArg[]): CellValue {
	const dec = bin2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const x = Math.trunc(dec.value)
	const abs = x < 0 ? (x + 0x40000000) >>> 0 : x
	let s = abs.toString(8)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function hex2bin(args: EvalArg[]): CellValue {
	const dec = hex2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const x = Math.trunc(dec.value)
	if (x < DEC2BIN_MIN || x > DEC2BIN_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + (1 << 10)) & 0x3ff : x
	let s = abs.toString(2)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function hex2oct(args: EvalArg[]): CellValue {
	const dec = hex2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const x = Math.trunc(dec.value)
	if (x < DEC2OCT_MIN || x > DEC2OCT_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + 0x40000000) >>> 0 : x
	let s = abs.toString(8)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function oct2bin(args: EvalArg[]): CellValue {
	const dec = oct2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const x = Math.trunc(dec.value)
	if (x < DEC2BIN_MIN || x > DEC2BIN_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + (1 << 10)) & 0x3ff : x
	let s = abs.toString(2)
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function oct2hex(args: EvalArg[]): CellValue {
	const dec = oct2dec(args)
	if (dec.kind === 'error') return dec
	if (dec.kind !== 'number') return dec
	const x = Math.trunc(dec.value)
	if (x < DEC2HEX_MIN || x > DEC2HEX_MAX) return errorValue('#NUM!')
	const places = args.length >= 2 ? numArg(args[1]) : undefined
	if (typeof places === 'object' && places.kind === 'error') return places
	const abs = x < 0 ? (x + 0x10000000000) >>> 0 : x
	let s = abs.toString(16).toUpperCase()
	if (typeof places === 'number' && places > 0) {
		const p = Math.trunc(places)
		if (s.length > p) return errorValue('#NUM!')
		s = s.padStart(p, '0')
	}
	return stringValue(s)
}

function delta(args: EvalArg[]): CellValue {
	const n1 = numArg(args[0])
	if (typeof n1 !== 'number') return n1
	const n2 = args.length >= 2 ? numArg(args[1]) : 0
	if (typeof n2 !== 'number') return n2
	return numberValue(n1 === n2 ? 1 : 0)
}

function gestep(args: EvalArg[]): CellValue {
	const num = numArg(args[0])
	if (typeof num !== 'number') return num
	const step = args.length >= 2 ? numArg(args[1]) : 0
	if (typeof step !== 'number') return step
	return numberValue(num >= step ? 1 : 0)
}

registerFunction({ name: 'BIN2DEC', minArgs: 1, maxArgs: 1, evaluate: bin2dec })
registerFunction({ name: 'DEC2BIN', minArgs: 1, maxArgs: 2, evaluate: dec2bin })
registerFunction({ name: 'HEX2DEC', minArgs: 1, maxArgs: 1, evaluate: hex2dec })
registerFunction({ name: 'DEC2HEX', minArgs: 1, maxArgs: 2, evaluate: dec2hex })
registerFunction({ name: 'OCT2DEC', minArgs: 1, maxArgs: 1, evaluate: oct2dec })
registerFunction({ name: 'DEC2OCT', minArgs: 1, maxArgs: 2, evaluate: dec2oct })
registerFunction({ name: 'BIN2HEX', minArgs: 1, maxArgs: 2, evaluate: bin2hex })
registerFunction({ name: 'BIN2OCT', minArgs: 1, maxArgs: 2, evaluate: bin2oct })
registerFunction({ name: 'HEX2BIN', minArgs: 1, maxArgs: 2, evaluate: hex2bin })
registerFunction({ name: 'HEX2OCT', minArgs: 1, maxArgs: 2, evaluate: hex2oct })
registerFunction({ name: 'OCT2BIN', minArgs: 1, maxArgs: 2, evaluate: oct2bin })
registerFunction({ name: 'OCT2HEX', minArgs: 1, maxArgs: 2, evaluate: oct2hex })
registerFunction({ name: 'DELTA', minArgs: 1, maxArgs: 2, evaluate: delta })
registerFunction({ name: 'GESTEP', minArgs: 1, maxArgs: 2, evaluate: gestep })

const MAX_48BIT = 2 ** 48 - 1

function bitIntArg(arg: EvalArg | undefined): number | CellValue {
	const n = numArg(arg)
	if (typeof n !== 'number') return n
	const x = Math.trunc(n)
	if (x < 0 || x > MAX_48BIT) return errorValue('#NUM!')
	return x
}

registerFunction({
	name: 'BITAND',
	minArgs: 2,
	maxArgs: 2,
	evaluate(args) {
		const a = bitIntArg(args[0])
		if (typeof a !== 'number') return a
		const b = bitIntArg(args[1])
		if (typeof b !== 'number') return b
		return numberValue(Number(BigInt(a) & BigInt(b)))
	},
})

registerFunction({
	name: 'BITOR',
	minArgs: 2,
	maxArgs: 2,
	evaluate(args) {
		const a = bitIntArg(args[0])
		if (typeof a !== 'number') return a
		const b = bitIntArg(args[1])
		if (typeof b !== 'number') return b
		return numberValue(Number(BigInt(a) | BigInt(b)))
	},
})

registerFunction({
	name: 'BITXOR',
	minArgs: 2,
	maxArgs: 2,
	evaluate(args) {
		const a = bitIntArg(args[0])
		if (typeof a !== 'number') return a
		const b = bitIntArg(args[1])
		if (typeof b !== 'number') return b
		return numberValue(Number(BigInt(a) ^ BigInt(b)))
	},
})

registerFunction({
	name: 'BITLSHIFT',
	minArgs: 2,
	maxArgs: 2,
	evaluate(args) {
		const n = bitIntArg(args[0])
		if (typeof n !== 'number') return n
		const shift = numArg(args[1])
		if (typeof shift !== 'number') return shift
		const s = Math.trunc(shift)
		if (Math.abs(s) > 53) return errorValue('#NUM!')
		const result = s >= 0 ? Number(BigInt(n) << BigInt(s)) : Number(BigInt(n) >> BigInt(-s))
		if (result < 0 || result > MAX_48BIT) return errorValue('#NUM!')
		return numberValue(result)
	},
})

registerFunction({
	name: 'BITRSHIFT',
	minArgs: 2,
	maxArgs: 2,
	evaluate(args) {
		const n = bitIntArg(args[0])
		if (typeof n !== 'number') return n
		const shift = numArg(args[1])
		if (typeof shift !== 'number') return shift
		const s = Math.trunc(shift)
		if (Math.abs(s) > 53) return errorValue('#NUM!')
		const result = s >= 0 ? Number(BigInt(n) >> BigInt(s)) : Number(BigInt(n) << BigInt(-s))
		if (result < 0 || result > MAX_48BIT) return errorValue('#NUM!')
		return numberValue(result)
	},
})
