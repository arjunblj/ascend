import type { CellValue } from '@ascend/schema'
import { errorValue, numberValue } from '@ascend/schema'
import {
	cellOf,
	type EvalArg,
	type FunctionEvalContext,
	getRange,
	numArg,
	registerFunction,
	toNumber,
} from './registry.ts'

const MS_PER_DAY = 86_400_000

function makeUTC(y: number, m: number, d: number): Date {
	const date = new Date(Date.UTC(y, m - 1, d))
	if (y >= 0 && y < 100) date.setUTCFullYear(y)
	return date
}

const EPOCH_1900_MS = makeUTC(1900, 1, 1).getTime()
const EPOCH_1904_MS = makeUTC(1904, 1, 1).getTime()

export function dateToSerial(
	year: number,
	month: number,
	day: number,
	dateSystem: '1900' | '1904' = '1900',
): number {
	const ms = makeUTC(year, month, day).getTime()
	if (dateSystem === '1904') {
		return Math.floor((ms - EPOCH_1904_MS) / MS_PER_DAY)
	}

	const days = Math.floor((ms - EPOCH_1900_MS) / MS_PER_DAY)
	let serial = days + 1
	if (serial >= 60) serial += 1
	return serial
}

interface DateParts {
	year: number
	month: number
	day: number
}

function daysSince1900ToYMD(days: number): DateParts {
	let y = 1900
	let d = Math.floor(days)
	while (d >= 365) {
		const daysInYear = isLeapYear(y) ? 366 : 365
		if (d >= daysInYear) {
			d -= daysInYear
			y++
		} else break
	}
	let m = 1
	for (; m <= 12; m++) {
		const dim = daysInMonth(y, m)
		if (d < dim) break
		d -= dim
	}
	return { year: y, month: m, day: d + 1 }
}

export function serialToDate(
	serial: number,
	dateSystem: '1900' | '1904' = '1900',
): DateParts | null {
	if (dateSystem === '1904') {
		if (serial < 0) return null
		const d = new Date(EPOCH_1904_MS + serial * MS_PER_DAY)
		return {
			year: d.getUTCFullYear(),
			month: d.getUTCMonth() + 1,
			day: d.getUTCDate(),
		}
	}

	if (serial < 1) return null
	if (serial === 60) return { year: 1900, month: 2, day: 29 }
	const days = serial < 60 ? serial - 1 : serial - 2
	return daysSince1900ToYMD(days)
}

function isLeapYear(y: number): boolean {
	return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0
}

const MONTH_DAYS = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31] as const

function daysInMonth(y: number, m: number): number {
	return m === 2 && isLeapYear(y) ? 29 : (MONTH_DAYS[m] ?? 30)
}

function currentDateSystem(ctx: FunctionEvalContext | undefined): '1900' | '1904' {
	return ctx?.dateSystem ?? '1900'
}

function serialDayIndex(serial: number, dateSystem: '1900' | '1904'): number | null {
	if (dateSystem === '1904') {
		const parts = serialToDate(serial, dateSystem)
		if (!parts) return null
		return makeUTC(parts.year, parts.month, parts.day).getUTCDay()
	}

	if (serial < 1) return null
	return (((serial - 1) % 7) + 7) % 7
}

function isWeekend(serial: number, dateSystem: '1900' | '1904'): boolean {
	const di = serialDayIndex(serial, dateSystem)
	return di === 0 || di === 6
}

function getHolidays(arg: EvalArg | undefined): Set<number> {
	const set = new Set<number>()
	if (!arg) return set
	for (const row of getRange(arg)) {
		for (const cell of row) {
			const n = toNumber(cell)
			if (n !== null) set.add(Math.floor(n))
		}
	}
	return set
}

function addMonths(parts: DateParts, months: number): { year: number; month: number } {
	let newMonth = parts.month + Math.trunc(months)
	let newYear = parts.year
	newYear += Math.floor((newMonth - 1) / 12)
	newMonth = ((((newMonth - 1) % 12) + 12) % 12) + 1
	return { year: newYear, month: newMonth }
}

function parseWeekendDays(arg: EvalArg | undefined): Set<number> | CellValue {
	if (!arg) return new Set([0, 6])
	const v = cellOf(arg)
	if (v.kind === 'error') return v
	if (v.kind === 'string' && v.value.length === 7) {
		if (!/^[01]{7}$/.test(v.value)) return errorValue('#VALUE!')
		const set = new Set<number>()
		for (let i = 0; i < 7; i++) {
			if (v.value[i] === '1') set.add((i + 1) % 7)
		}
		if (set.size === 7) return errorValue('#VALUE!')
		return set
	}
	const n = toNumber(v)
	if (n === null) return errorValue('#VALUE!')
	const code = Math.trunc(n)
	const patterns: Record<number, number[]> = {
		1: [0, 6],
		2: [0, 1],
		3: [1, 2],
		4: [2, 3],
		5: [3, 4],
		6: [4, 5],
		7: [5, 6],
		11: [0],
		12: [1],
		13: [2],
		14: [3],
		15: [4],
		16: [5],
		17: [6],
	}
	const days = patterns[code]
	if (!days) return errorValue('#NUM!')
	return new Set(days)
}

// --- Implementations ---

function dateFn(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const y = numArg(args[0])
	if (typeof y !== 'number') return y
	const m = numArg(args[1])
	if (typeof m !== 'number') return m
	const d = numArg(args[2])
	if (typeof d !== 'number') return d

	let year = Math.trunc(y)
	if (year >= 0 && year <= 1899) year += 1900
	const dateSystem = currentDateSystem(ctx)
	const serial = dateToSerial(year, Math.trunc(m), Math.trunc(d), dateSystem)
	return serial < (dateSystem === '1904' ? 0 : 1) ? errorValue('#NUM!') : numberValue(serial)
}

function today(_args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const d = ctx?.today ?? new Date()
	return numberValue(
		dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate(), currentDateSystem(ctx)),
	)
}

function nowFn(_args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const d = ctx?.now ?? new Date()
	const serial = dateToSerial(
		d.getFullYear(),
		d.getMonth() + 1,
		d.getDate(),
		currentDateSystem(ctx),
	)
	const frac = (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400
	return numberValue(serial + frac)
}

function yearFn(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const s = numArg(args[0])
	if (typeof s !== 'number') return s
	const parts = serialToDate(Math.floor(s), currentDateSystem(ctx))
	return parts ? numberValue(parts.year) : errorValue('#NUM!')
}

function monthFn(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const s = numArg(args[0])
	if (typeof s !== 'number') return s
	const parts = serialToDate(Math.floor(s), currentDateSystem(ctx))
	return parts ? numberValue(parts.month) : errorValue('#NUM!')
}

function dayFn(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const s = numArg(args[0])
	if (typeof s !== 'number') return s
	const parts = serialToDate(Math.floor(s), currentDateSystem(ctx))
	return parts ? numberValue(parts.day) : errorValue('#NUM!')
}

function hourFn(args: EvalArg[]): CellValue {
	const s = numArg(args[0])
	if (typeof s !== 'number') return s
	const frac = Math.abs(s) - Math.floor(Math.abs(s))
	return numberValue(Math.floor(Math.round(frac * 86400) / 3600) % 24)
}

function minuteFn(args: EvalArg[]): CellValue {
	const s = numArg(args[0])
	if (typeof s !== 'number') return s
	const frac = Math.abs(s) - Math.floor(Math.abs(s))
	const totalSec = Math.round(frac * 86400)
	return numberValue(Math.floor(totalSec / 60) % 60)
}

function secondFn(args: EvalArg[]): CellValue {
	const s = numArg(args[0])
	if (typeof s !== 'number') return s
	const frac = Math.abs(s) - Math.floor(Math.abs(s))
	return numberValue(Math.round(frac * 86400) % 60)
}

function timeFn(args: EvalArg[]): CellValue {
	const h = numArg(args[0])
	if (typeof h !== 'number') return h
	const m = numArg(args[1])
	if (typeof m !== 'number') return m
	const s = numArg(args[2])
	if (typeof s !== 'number') return s
	const totalSeconds = Math.trunc(h) * 3600 + Math.trunc(m) * 60 + Math.trunc(s)
	const wrapped = ((totalSeconds % 86400) + 86400) % 86400
	return numberValue(wrapped / 86400)
}

function timevalue(args: EvalArg[]): CellValue {
	const v = cellOf(args[0])
	if (v.kind === 'error') return v
	if (v.kind !== 'string') return errorValue('#VALUE!')
	const match = /^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?\s*$/i.exec(v.value)
	if (!match) return errorValue('#VALUE!')
	let hour = Number(match[1])
	const minute = Number(match[2])
	const second = Number(match[3] ?? '0')
	if (Number.isNaN(hour) || Number.isNaN(minute) || Number.isNaN(second)) {
		return errorValue('#VALUE!')
	}
	const meridiem = match[4]?.toUpperCase()
	if (meridiem === 'AM' && hour === 12) hour = 0
	if (meridiem === 'PM' && hour < 12) hour += 12
	if (hour < 0 || minute < 0 || minute >= 60 || second < 0 || second >= 60) {
		return errorValue('#VALUE!')
	}
	return numberValue((hour * 3600 + minute * 60 + second) / 86400)
}

function datevalue(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const v = cellOf(args[0])
	if (v.kind === 'error') return v
	if (v.kind !== 'string') return errorValue('#VALUE!')
	const parsed = parseDateText(v.value)
	if (!parsed) return errorValue('#VALUE!')
	return numberValue(dateToSerial(parsed.year, parsed.month, parsed.day, currentDateSystem(ctx)))
}

function parseDateText(value: string): DateParts | null {
	const text = value.trim()
	const iso = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T\s].*)?$/.exec(text)
	if (iso) return validateDateParts(Number(iso[1]), Number(iso[2]), Number(iso[3]))

	const us = /^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/.exec(text)
	if (us) {
		const yearRaw = Number(us[3])
		const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw
		return validateDateParts(year, Number(us[1]), Number(us[2]))
	}

	return null
}

function validateDateParts(year: number, month: number, day: number): DateParts | null {
	if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null
	if (month < 1 || month > 12) return null
	if (day < 1 || day > daysInMonth(year, month)) return null
	return { year, month, day }
}

function datedif(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const sn = numArg(args[0])
	if (typeof sn !== 'number') return sn
	const en = numArg(args[1])
	if (typeof en !== 'number') return en
	const uv = cellOf(args[2])
	if (uv.kind === 'error') return uv
	if (uv.kind !== 'string') return errorValue('#VALUE!')
	if (sn > en) return errorValue('#NUM!')

	const dateSystem = currentDateSystem(ctx)
	const sp = serialToDate(Math.floor(sn), dateSystem)
	const ep = serialToDate(Math.floor(en), dateSystem)
	if (!sp || !ep) return errorValue('#NUM!')

	switch (uv.value.toUpperCase()) {
		case 'Y': {
			let yrs = ep.year - sp.year
			if (ep.month < sp.month || (ep.month === sp.month && ep.day < sp.day)) yrs--
			return numberValue(yrs)
		}
		case 'M': {
			let mos = (ep.year - sp.year) * 12 + (ep.month - sp.month)
			if (ep.day < sp.day) mos--
			return numberValue(mos)
		}
		case 'D':
			return numberValue(Math.floor(en) - Math.floor(sn))
		case 'MD': {
			let days = ep.day - sp.day
			if (days < 0) {
				const pm = ep.month === 1 ? 12 : ep.month - 1
				const py = ep.month === 1 ? ep.year - 1 : ep.year
				days += daysInMonth(py, pm)
			}
			return numberValue(days)
		}
		case 'YM': {
			let mos = ep.month - sp.month
			if (ep.day < sp.day) mos--
			if (mos < 0) mos += 12
			return numberValue(mos)
		}
		case 'YD': {
			const doy = (p: DateParts) => {
				let d = p.day
				for (let i = 1; i < p.month; i++) d += daysInMonth(p.year, i)
				return d
			}
			let days = doy(ep) - doy(sp)
			if (days < 0) days += isLeapYear(ep.year - 1) ? 366 : 365
			return numberValue(days)
		}
		default:
			return errorValue('#NUM!')
	}
}

function edate(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const sn = numArg(args[0])
	if (typeof sn !== 'number') return sn
	const mn = numArg(args[1])
	if (typeof mn !== 'number') return mn

	const dateSystem = currentDateSystem(ctx)
	const parts = serialToDate(Math.floor(sn), dateSystem)
	if (!parts) return errorValue('#NUM!')

	const { year, month } = addMonths(parts, mn)
	const maxDay = daysInMonth(year, month)
	const serial = dateToSerial(year, month, Math.min(parts.day, maxDay), dateSystem)
	return serial < (dateSystem === '1904' ? 0 : 1) ? errorValue('#NUM!') : numberValue(serial)
}

function eomonth(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const sn = numArg(args[0])
	if (typeof sn !== 'number') return sn
	const mn = numArg(args[1])
	if (typeof mn !== 'number') return mn

	const dateSystem = currentDateSystem(ctx)
	const parts = serialToDate(Math.floor(sn), dateSystem)
	if (!parts) return errorValue('#NUM!')

	const { year, month } = addMonths(parts, mn)
	const serial = dateToSerial(year, month, daysInMonth(year, month), dateSystem)
	return serial < (dateSystem === '1904' ? 0 : 1) ? errorValue('#NUM!') : numberValue(serial)
}

function weekday(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const sn = numArg(args[0])
	if (typeof sn !== 'number') return sn
	const serial = Math.floor(sn)
	const dateSystem = currentDateSystem(ctx)
	if (serial < (dateSystem === '1904' ? 0 : 1)) return errorValue('#NUM!')

	const rt = args.length > 1 ? numArg(args[1]) : 1
	if (typeof rt !== 'number') return rt

	const dayIndex = serialDayIndex(serial, dateSystem)
	if (dayIndex === null) return errorValue('#NUM!')
	const startDays: Record<number, number> = {
		1: 0,
		2: 1,
		3: 1,
		11: 1,
		12: 2,
		13: 3,
		14: 4,
		15: 5,
		16: 6,
		17: 0,
	}
	const start = startDays[rt]
	if (start === undefined) return errorValue('#NUM!')

	const shifted = (dayIndex - start + 7) % 7
	return numberValue(rt === 3 ? shifted : shifted + 1)
}

function weeknum(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const sn = numArg(args[0])
	if (typeof sn !== 'number') return sn
	const serial = Math.floor(sn)
	const dateSystem = currentDateSystem(ctx)
	if (serial < (dateSystem === '1904' ? 0 : 1)) return errorValue('#NUM!')

	const rt = args.length > 1 ? numArg(args[1]) : 1
	if (typeof rt !== 'number') return rt

	const parts = serialToDate(serial, dateSystem)
	if (!parts) return errorValue('#NUM!')

	const jan1 = dateToSerial(parts.year, 1, 1, dateSystem)
	const jan1Di = serialDayIndex(jan1, dateSystem)
	if (jan1Di === null) return errorValue('#NUM!')
	const weekStarts: Record<number, number> = {
		1: 0,
		2: 1,
		11: 1,
		12: 2,
		13: 3,
		14: 4,
		15: 5,
		16: 6,
		17: 0,
	}
	const ws = weekStarts[rt]
	if (ws === undefined) return errorValue('#NUM!')

	const dayOfYear = serial - jan1
	const offset = (jan1Di - ws + 7) % 7
	return numberValue(Math.floor((dayOfYear + offset) / 7) + 1)
}

function networkdays(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const sn = numArg(args[0])
	if (typeof sn !== 'number') return sn
	const en = numArg(args[1])
	if (typeof en !== 'number') return en

	let start = Math.floor(sn)
	let end = Math.floor(en)
	const dateSystem = currentDateSystem(ctx)
	const holidays = args.length > 2 ? getHolidays(args[2]) : new Set<number>()

	const sign = start <= end ? 1 : -1
	if (start > end) [start, end] = [end, start]

	let count = 0
	for (let d = start; d <= end; d++) {
		if (!isWeekend(d, dateSystem) && !holidays.has(d)) count++
	}
	return numberValue(count * sign)
}

function workdayFn(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const sn = numArg(args[0])
	if (typeof sn !== 'number') return sn
	const dn = numArg(args[1])
	if (typeof dn !== 'number') return dn

	let serial = Math.floor(sn)
	let days = Math.trunc(dn)
	const dateSystem = currentDateSystem(ctx)
	const holidays = args.length > 2 ? getHolidays(args[2]) : new Set<number>()

	const step = days > 0 ? 1 : -1
	days = Math.abs(days)
	while (days > 0) {
		serial += step
		if (!isWeekend(serial, dateSystem) && !holidays.has(serial)) days--
	}
	return numberValue(serial)
}

// --- Registration ---

registerFunction({ name: 'DATE', minArgs: 3, maxArgs: 3, evaluate: dateFn })
registerFunction({
	name: 'TODAY',
	minArgs: 0,
	maxArgs: 0,
	evaluate: today,
	volatile: true,
})
registerFunction({
	name: 'NOW',
	minArgs: 0,
	maxArgs: 0,
	evaluate: nowFn,
	volatile: true,
})
registerFunction({ name: 'YEAR', minArgs: 1, maxArgs: 1, evaluate: yearFn })
registerFunction({ name: 'MONTH', minArgs: 1, maxArgs: 1, evaluate: monthFn })
registerFunction({ name: 'DAY', minArgs: 1, maxArgs: 1, evaluate: dayFn })
registerFunction({ name: 'HOUR', minArgs: 1, maxArgs: 1, evaluate: hourFn })
registerFunction({
	name: 'MINUTE',
	minArgs: 1,
	maxArgs: 1,
	evaluate: minuteFn,
})
registerFunction({
	name: 'SECOND',
	minArgs: 1,
	maxArgs: 1,
	evaluate: secondFn,
})
registerFunction({ name: 'TIME', minArgs: 3, maxArgs: 3, evaluate: timeFn })
registerFunction({
	name: 'TIMEVALUE',
	minArgs: 1,
	maxArgs: 1,
	evaluate: timevalue,
})
registerFunction({
	name: 'DATEVALUE',
	minArgs: 1,
	maxArgs: 1,
	evaluate: datevalue,
})
registerFunction({
	name: 'DATEDIF',
	minArgs: 3,
	maxArgs: 3,
	evaluate: datedif,
})
registerFunction({ name: 'EDATE', minArgs: 2, maxArgs: 2, evaluate: edate })
registerFunction({
	name: 'EOMONTH',
	minArgs: 2,
	maxArgs: 2,
	evaluate: eomonth,
})
registerFunction({
	name: 'WEEKDAY',
	minArgs: 1,
	maxArgs: 2,
	evaluate: weekday,
})
registerFunction({
	name: 'WEEKNUM',
	minArgs: 1,
	maxArgs: 2,
	evaluate: weeknum,
})
registerFunction({
	name: 'NETWORKDAYS',
	minArgs: 2,
	maxArgs: 3,
	evaluate: networkdays,
})
registerFunction({
	name: 'WORKDAY',
	minArgs: 2,
	maxArgs: 3,
	evaluate: workdayFn,
})
registerFunction({
	name: 'DAYS360',
	minArgs: 2,
	maxArgs: 3,
	evaluate: days360Fn,
})
registerFunction({
	name: 'YEARFRAC',
	minArgs: 2,
	maxArgs: 3,
	evaluate: yearfracFn,
})

function days360US(
	start: DateParts,
	end: DateParts,
	daysInMonthFn: (y: number, m: number) => number,
): [DateParts, DateParts] {
	let d1 = start.day
	const m1 = start.month
	const y1 = start.year
	let d2 = end.day
	let m2 = end.month
	let y2 = end.year

	const lastDayStart = d1 === daysInMonthFn(y1, m1)
	const lastDayEnd = d2 === daysInMonthFn(y2, m2)

	if (lastDayStart) d1 = 30
	if (lastDayEnd) {
		if (d1 < 30) {
			d2 = 1
			m2++
			if (m2 > 12) {
				m2 = 1
				y2++
			}
		} else {
			d2 = 30
		}
	}

	return [
		{ year: y1, month: m1, day: d1 },
		{ year: y2, month: m2, day: d2 },
	]
}

function days360European(start: DateParts, end: DateParts): [DateParts, DateParts] {
	let d1 = start.day
	let d2 = end.day
	if (d1 === 31) d1 = 30
	if (d2 === 31) d2 = 30
	return [
		{ year: start.year, month: start.month, day: d1 },
		{ year: end.year, month: end.month, day: d2 },
	]
}

function days360Count(adj: [DateParts, DateParts]): number {
	const [a, b] = adj
	return 360 * (b.year - a.year) + 30 * (b.month - a.month) + (b.day - a.day)
}

function days360Fn(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const sn = numArg(args[0])
	if (typeof sn !== 'number') return sn
	const en = numArg(args[1])
	if (typeof en !== 'number') return en

	const dateSystem = currentDateSystem(ctx)
	const sp = serialToDate(Math.floor(sn), dateSystem)
	const ep = serialToDate(Math.floor(en), dateSystem)
	if (!sp || !ep) return errorValue('#NUM!')

	const method = args.length > 2 ? numArg(args[2]) : 0
	if (typeof method !== 'number') return method

	const adj = method === 1 ? days360European(sp, ep) : days360US(sp, ep, daysInMonth)
	return numberValue(days360Count(adj))
}

function yearfracFn(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const sn = numArg(args[0])
	if (typeof sn !== 'number') return sn
	const en = numArg(args[1])
	if (typeof en !== 'number') return en

	const basis = args.length > 2 ? numArg(args[2]) : 0
	if (typeof basis !== 'number') return basis
	if (basis < 0 || basis > 4) return errorValue('#NUM!')

	const dateSystem = currentDateSystem(ctx)
	const sp = serialToDate(Math.floor(sn), dateSystem)
	const ep = serialToDate(Math.floor(en), dateSystem)
	if (!sp || !ep) return errorValue('#NUM!')

	let days: number
	let divisor: number

	if (basis === 0 || basis === 4) {
		const adj = basis === 4 ? days360European(sp, ep) : days360US(sp, ep, daysInMonth)
		days = days360Count(adj)
		divisor = 360
	} else if (basis === 1) {
		const startSerial = Math.floor(sn)
		const endSerial = Math.floor(en)
		let total = 0
		for (let y = sp.year; y <= ep.year; y++) {
			const yearStart = dateToSerial(y, 1, 1, dateSystem)
			const yearEnd = dateToSerial(y, 12, 31, dateSystem)
			const overlapStart = Math.max(startSerial, yearStart)
			const overlapEnd = Math.min(endSerial, yearEnd)
			const overlapDays = Math.max(0, overlapEnd - overlapStart)
			const daysInYear = isLeapYear(y) ? 366 : 365
			total += overlapDays / daysInYear
		}
		return numberValue(total)
	} else if (basis === 2) {
		days = Math.floor(en) - Math.floor(sn)
		divisor = 360
	} else {
		days = Math.floor(en) - Math.floor(sn)
		divisor = 365
	}

	return numberValue(days / divisor)
}

function isoWeekNum(args: EvalArg[], ctx?: FunctionEvalContext): CellValue {
	const s = numArg(args[0])
	if (typeof s !== 'number') return s
	const serial = Math.trunc(s)
	const ds = ctx?.dateSystem ?? '1900'
	const parts = serialToDate(serial, ds)
	if (!parts) return errorValue('#NUM!')
	const d = makeUTC(parts.year, parts.month, parts.day)
	const dayOfWeek = d.getUTCDay() || 7
	d.setUTCDate(d.getUTCDate() + 4 - dayOfWeek)
	const yearStart = makeUTC(d.getUTCFullYear(), 1, 1)
	const weekNo = Math.ceil(((d.getTime() - yearStart.getTime()) / MS_PER_DAY + 1) / 7)
	return numberValue(weekNo)
}

registerFunction({
	name: 'ISOWEEKNUM',
	minArgs: 1,
	maxArgs: 1,
	evaluate: isoWeekNum,
})
registerFunction({
	name: 'DAYS',
	minArgs: 2,
	maxArgs: 2,
	evaluate: (args) => {
		const end = numArg(args[0])
		if (typeof end !== 'number') return end
		const start = numArg(args[1])
		if (typeof start !== 'number') return start
		return numberValue(Math.floor(end) - Math.floor(start))
	},
})
registerFunction({
	name: 'NETWORKDAYS.INTL',
	minArgs: 2,
	maxArgs: 4,
	evaluate: (args, ctx?) => {
		const sn = numArg(args[0])
		if (typeof sn !== 'number') return sn
		const en = numArg(args[1])
		if (typeof en !== 'number') return en
		const dateSystem = currentDateSystem(ctx)
		const weekendResult = parseWeekendDays(args[2])
		if ('kind' in weekendResult) return weekendResult
		const holidays = args.length > 3 ? getHolidays(args[3]) : new Set<number>()
		let start = Math.floor(sn)
		let end = Math.floor(en)
		const sign = start <= end ? 1 : -1
		if (start > end) [start, end] = [end, start]
		let count = 0
		for (let d = start; d <= end; d++) {
			const di = serialDayIndex(d, dateSystem)
			if (di !== null && !weekendResult.has(di) && !holidays.has(d)) count++
		}
		return numberValue(count * sign)
	},
})
registerFunction({
	name: 'WORKDAY.INTL',
	minArgs: 2,
	maxArgs: 4,
	evaluate: (args, ctx?) => {
		const sn = numArg(args[0])
		if (typeof sn !== 'number') return sn
		const dn = numArg(args[1])
		if (typeof dn !== 'number') return dn
		const dateSystem = currentDateSystem(ctx)
		const weekendResult = parseWeekendDays(args[2])
		if ('kind' in weekendResult) return weekendResult
		const holidays = args.length > 3 ? getHolidays(args[3]) : new Set<number>()
		let serial = Math.floor(sn)
		let days = Math.trunc(dn)
		const step = days > 0 ? 1 : -1
		days = Math.abs(days)
		while (days > 0) {
			serial += step
			const di = serialDayIndex(serial, dateSystem)
			if (di !== null && !weekendResult.has(di) && !holidays.has(serial)) days--
		}
		return numberValue(serial)
	},
})
