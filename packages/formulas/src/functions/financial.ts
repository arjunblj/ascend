import type { CellValue } from '@ascend/schema'
import { errorValue, numberValue } from '@ascend/schema'
import { dateToSerial, serialToDate } from './date.ts'
import type { EvalArg, FunctionDef, FunctionEvalContext } from './registry.ts'
import { collectNumbers, getRange, numArg, toNumber } from './registry.ts'

function num(arg: EvalArg | undefined): number | CellValue {
	return numArg(arg)
}

function annuityTerms(rate: number, nper: number): { factor: number; pow: number } {
	const q = 1 + rate
	if (Number.isInteger(nper) && nper >= 0 && nper <= 10000) {
		let factor = 0
		let term = 1
		for (let period = 0; period < nper; period++) {
			factor += term
			term *= q
		}
		return { factor, pow: term }
	}
	const pow = q ** nper
	return { factor: (pow - 1) / rate, pow }
}

function pmt(rate: number, nper: number, pv: number, fv: number, type: number): number {
	if (rate === 0) return -(pv + fv) / nper
	const { factor, pow } = annuityTerms(rate, nper)
	return -(fv + pv * pow) / ((1 + rate * type) * factor)
}

function fvCalc(rate: number, nper: number, pmtVal: number, pv: number, type: number): number {
	if (rate === 0) return -(pv + pmtVal * nper)
	const temp = (1 + rate) ** nper
	return -(pv * temp + ((pmtVal * (1 + rate * type)) / rate) * (temp - 1))
}

function pvCalc(rate: number, nper: number, pmtVal: number, fv: number, type: number): number {
	if (rate === 0) return -(fv + pmtVal * nper)
	const temp = (1 + rate) ** nper
	const fact = ((1 + rate * type) * (temp - 1)) / rate
	return -(fv + pmtVal * fact) / temp
}

function nperCalc(rate: number, pmtVal: number, pv: number, fv: number, type: number): number {
	if (rate === 0) return -(pv + fv) / pmtVal
	const z = (pmtVal * (1 + rate * type)) / rate
	return Math.log((-fv + z) / (pv + z)) / Math.log(1 + rate)
}

function ipmtCalc(
	rate: number,
	per: number,
	nper: number,
	pv: number,
	fv: number,
	type: number,
): number {
	const pmtVal = pmt(rate, nper, pv, fv, type)
	if (type === 1 && per === 1) return 0
	let result = fvCalc(rate, per - 1, pmtVal, pv, type) * rate
	if (type === 1) result /= 1 + rate
	return result
}

function rateNewton(
	nper: number,
	pmtVal: number,
	pv: number,
	fv: number,
	type: number,
	guess: number,
): number {
	const maxIter = 150
	const tol = 1e-10
	let rate = guess
	for (let i = 0; i < maxIter; i++) {
		if (rate <= -1) return Number.NaN
		const t1 = (1 + rate) ** nper
		const t2 = (1 + rate) ** (nper - 1)
		const g = fv + t1 * pv + (pmtVal * (t1 - 1) * (rate * type + 1)) / rate
		const gp =
			nper * t2 * pv -
			(pmtVal * (t1 - 1) * (rate * type + 1)) / (rate * rate) +
			(nper * pmtVal * t2 * (rate * type + 1)) / rate +
			(pmtVal * (t1 - 1) * type) / rate
		if (Math.abs(gp) < 1e-20) return Number.NaN
		const next = rate - g / gp
		if (Math.abs(next - rate) < tol) return next
		rate = next
	}
	return Number.NaN
}

interface DatedCashFlows {
	readonly values: readonly number[]
	readonly dates: readonly number[]
}

function datedCashFlows(
	valuesArg: EvalArg | undefined,
	datesArg: EvalArg | undefined,
): DatedCashFlows | CellValue {
	const valueCells = getRange(valuesArg).flat()
	const dateCells = getRange(datesArg).flat()
	if (valueCells.length !== dateCells.length || valueCells.length === 0) return errorValue('#NUM!')
	const firstDateCell = dateCells[0]
	if (!firstDateCell) return errorValue('#VALUE!')
	const firstDate = serialDateArg(firstDateCell)
	if (typeof firstDate !== 'number') return firstDate

	const values: number[] = []
	const dates: number[] = []
	let hasPositive = false
	let hasNegative = false
	for (let index = 0; index < valueCells.length; index++) {
		const valueCell = valueCells[index]
		const dateCell = dateCells[index]
		if (!valueCell || !dateCell) return errorValue('#VALUE!')
		const value = toNumber(valueCell)
		if (value === null || !Number.isFinite(value)) return errorValue('#VALUE!')
		const date = serialDateArg(dateCell)
		if (typeof date !== 'number') return date
		if (date < firstDate) return errorValue('#NUM!')
		if (value > 0) hasPositive = true
		if (value < 0) hasNegative = true
		values.push(value)
		dates.push(date)
	}
	if (!hasPositive || !hasNegative) return errorValue('#NUM!')
	return { values, dates }
}

function serialDateArg(value: CellValue): number | CellValue {
	const date = toNumber(value)
	if (date === null || !Number.isFinite(date)) return errorValue('#VALUE!')
	return Math.trunc(date)
}

function vdbCalc(
	cost: number,
	salvage: number,
	life: number,
	startPeriod: number,
	endPeriod: number,
	factor: number,
	noSwitch: boolean,
): number {
	let totalDep = 0
	let bookValue = cost
	const periods = Math.ceil(endPeriod)
	for (let p = 0; p < periods; p++) {
		let dep = (bookValue * factor) / life
		if (!noSwitch) {
			const remaining = life - p
			if (remaining > 0) {
				const slnDep = (bookValue - salvage) / remaining
				if (slnDep > dep) dep = slnDep
			}
		}
		dep = Math.max(0, Math.min(dep, bookValue - salvage))
		const overlap = Math.max(0, Math.min(p + 1, endPeriod) - Math.max(p, startPeriod))
		totalDep += dep * overlap
		bookValue -= dep
	}
	return totalDep
}

function currentDateSystem(ctx: FunctionEvalContext | undefined): '1900' | '1904' {
	return ctx?.dateSystem ?? '1900'
}

function daysInMonth(year: number, month: number): number {
	return new Date(Date.UTC(year, month, 0)).getUTCDate()
}

function isLeapYear(year: number): boolean {
	return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0
}

function addMonthsPreserveEom(
	parts: { year: number; month: number; day: number },
	deltaMonths: number,
	preserveEom: boolean,
): { year: number; month: number; day: number } {
	const monthIndex = parts.year * 12 + (parts.month - 1) + deltaMonths
	const year = Math.floor(monthIndex / 12)
	const month = (((monthIndex % 12) + 12) % 12) + 1
	const day = preserveEom ? daysInMonth(year, month) : Math.min(parts.day, daysInMonth(year, month))
	return { year, month, day }
}

function parseDateArg(serial: number, dateSystem: '1900' | '1904') {
	return serialToDate(Math.floor(serial), dateSystem)
}

function isLastDayOfFebruary(parts: { year: number; month: number; day: number }): boolean {
	return parts.month === 2 && parts.day === daysInMonth(parts.year, parts.month)
}

function dayCount30Us(
	start: { year: number; month: number; day: number },
	end: { year: number; month: number; day: number },
): number {
	let d1 = start.day
	let d2 = end.day
	const startWasLastDayOfFebruary = isLastDayOfFebruary(start)
	const endWasLastDayOfFebruary = isLastDayOfFebruary(end)
	if (startWasLastDayOfFebruary) d1 = 30
	if (endWasLastDayOfFebruary && startWasLastDayOfFebruary) d2 = 30
	if (d1 === 31) d1 = 30
	if (d2 === 31 && d1 >= 30) d2 = 30
	return (end.year - start.year) * 360 + (end.month - start.month) * 30 + (d2 - d1)
}

function dayCount30Eu(
	start: { year: number; month: number; day: number },
	end: { year: number; month: number; day: number },
): number {
	const d1 = start.day === 31 ? 30 : start.day
	const d2 = end.day === 31 ? 30 : end.day
	return (end.year - start.year) * 360 + (end.month - start.month) * 30 + (d2 - d1)
}

function dayCountBasis(
	startSerial: number,
	endSerial: number,
	basis: number,
	dateSystem: '1900' | '1904',
): number {
	const start = parseDateArg(startSerial, dateSystem)
	const end = parseDateArg(endSerial, dateSystem)
	if (!start || !end) return Number.NaN
	switch (basis) {
		case 0:
			return dayCount30Us(start, end)
		case 4:
			return dayCount30Eu(start, end)
		default:
			return Math.floor(endSerial) - Math.floor(startSerial)
	}
}

function yearFracActualActual(
	startSerial: number,
	endSerial: number,
	dateSystem: '1900' | '1904',
): number {
	if (endSerial <= startSerial) return 0
	const start = parseDateArg(startSerial, dateSystem)
	const end = parseDateArg(endSerial, dateSystem)
	if (!start || !end) return Number.NaN
	if (start.year === end.year) {
		return (Math.floor(endSerial) - Math.floor(startSerial)) / (isLeapYear(start.year) ? 366 : 365)
	}
	let total = 0
	const startYearEnd = dateToSerial(start.year, 12, 31, dateSystem)
	total += (startYearEnd + 1 - Math.floor(startSerial)) / (isLeapYear(start.year) ? 366 : 365)
	for (let year = start.year + 1; year < end.year; year++) total += 1
	const endYearStart = dateToSerial(end.year, 1, 1, dateSystem)
	total += (Math.floor(endSerial) - endYearStart) / (isLeapYear(end.year) ? 366 : 365)
	return total
}

function yearFracFromBasis(
	startSerial: number,
	endSerial: number,
	basis: number,
	dateSystem: '1900' | '1904',
): number {
	const days = dayCountBasis(startSerial, endSerial, basis, dateSystem)
	switch (basis) {
		case 0:
		case 2:
		case 4:
			return days / 360
		case 3:
			return days / 365
		default:
			return yearFracActualActual(startSerial, endSerial, dateSystem)
	}
}

function validateBasis(basis: number): number | null {
	const value = Math.trunc(basis)
	return value >= 0 && value <= 4 ? value : null
}

function validateFrequency(frequency: number): number | null {
	const value = Math.trunc(frequency)
	return value === 1 || value === 2 || value === 4 ? value : null
}

function resolveCouponWindow(
	settlement: number,
	maturity: number,
	frequency: number,
	dateSystem: '1900' | '1904',
): { prev: number; next: number; monthsPerCoupon: number } | null {
	const maturityParts = parseDateArg(maturity, dateSystem)
	if (!maturityParts) return null
	const preserveEom = maturityParts.day === daysInMonth(maturityParts.year, maturityParts.month)
	const monthsPerCoupon = 12 / frequency
	let nextParts = maturityParts
	let nextSerial = dateToSerial(nextParts.year, nextParts.month, nextParts.day, dateSystem)
	while (true) {
		const prevParts = addMonthsPreserveEom(nextParts, -monthsPerCoupon, preserveEom)
		const prevSerial = dateToSerial(prevParts.year, prevParts.month, prevParts.day, dateSystem)
		if (prevSerial <= settlement) return { prev: prevSerial, next: nextSerial, monthsPerCoupon }
		nextParts = prevParts
		nextSerial = prevSerial
	}
}

function couponPeriodDays(
	prev: number,
	next: number,
	basis: number,
	dateSystem: '1900' | '1904',
): number {
	switch (basis) {
		case 2:
		case 3:
			return Math.floor(next) - Math.floor(prev)
		default:
			return dayCountBasis(prev, next, basis, dateSystem)
	}
}

function couponCountRemaining(
	settlement: number,
	maturity: number,
	frequency: number,
	dateSystem: '1900' | '1904',
): number {
	const window = resolveCouponWindow(settlement, maturity, frequency, dateSystem)
	if (!window) return 0
	const maturityParts = parseDateArg(maturity, dateSystem)
	if (!maturityParts) return 0
	const preserveEom = maturityParts.day === daysInMonth(maturityParts.year, maturityParts.month)
	let count = 0
	let current = window.next
	let currentParts = parseDateArg(current, dateSystem)
	while (currentParts && current <= maturity) {
		count++
		currentParts = addMonthsPreserveEom(currentParts, window.monthsPerCoupon, preserveEom)
		current = dateToSerial(currentParts.year, currentParts.month, currentParts.day, dateSystem)
	}
	return count
}

function regularBondPrice(
	settlement: number,
	maturity: number,
	rate: number,
	yld: number,
	redemption: number,
	frequency: number,
	basis: number,
	dateSystem: '1900' | '1904',
): number {
	const coupon = (100 * rate) / frequency
	const window = resolveCouponWindow(settlement, maturity, frequency, dateSystem)
	if (!window) return Number.NaN
	const a = dayCountBasis(window.prev, settlement, basis, dateSystem)
	const e = couponPeriodDays(window.prev, window.next, basis, dateSystem)
	const dsc = dayCountBasis(settlement, window.next, basis, dateSystem)
	const n = couponCountRemaining(settlement, maturity, frequency, dateSystem)
	const q = 1 + yld / frequency
	if (n <= 1) {
		const dsr = e - a
		return (coupon + redemption) / q ** (dsr / e) - (coupon * a) / e
	}
	let price = redemption / q ** (n - 1 + dsc / e)
	for (let k = 1; k <= n; k++) price += coupon / q ** (k - 1 + dsc / e)
	return price - (coupon * a) / e
}

function oddFirstCouponAmount(
	issue: number,
	firstCoupon: number,
	rate: number,
	basis: number,
	dateSystem: '1900' | '1904',
): number {
	return 100 * rate * yearFracFromBasis(issue, firstCoupon, basis, dateSystem)
}

function oddFirstAccrued(
	issue: number,
	settlement: number,
	rate: number,
	basis: number,
	dateSystem: '1900' | '1904',
): number {
	return 100 * rate * yearFracFromBasis(issue, settlement, basis, dateSystem)
}

function previousCouponDate(
	couponDate: number,
	maturity: number,
	frequency: number,
	dateSystem: '1900' | '1904',
): number | null {
	const couponParts = parseDateArg(couponDate, dateSystem)
	const maturityParts = parseDateArg(maturity, dateSystem)
	if (!couponParts || !maturityParts) return null
	const preserveEom = maturityParts.day === daysInMonth(maturityParts.year, maturityParts.month)
	const previous = addMonthsPreserveEom(couponParts, -(12 / frequency), preserveEom)
	return dateToSerial(previous.year, previous.month, previous.day, dateSystem)
}

function couponScheduleFromFirst(
	firstCoupon: number,
	maturity: number,
	frequency: number,
	dateSystem: '1900' | '1904',
): number[] | null {
	const maturityParts = parseDateArg(maturity, dateSystem)
	if (!maturityParts) return null
	const preserveEom = maturityParts.day === daysInMonth(maturityParts.year, maturityParts.month)
	const monthsPerCoupon = 12 / frequency
	const schedule: number[] = [firstCoupon]
	let current = firstCoupon
	while (current < maturity) {
		const currentParts = parseDateArg(current, dateSystem)
		if (!currentParts) return null
		const nextParts = addMonthsPreserveEom(currentParts, monthsPerCoupon, preserveEom)
		current = dateToSerial(nextParts.year, nextParts.month, nextParts.day, dateSystem)
		if (current <= maturity) schedule.push(current)
	}
	return schedule
}

function oddShortFirstBondPrice(
	settlement: number,
	maturity: number,
	issue: number,
	firstCoupon: number,
	rate: number,
	yld: number,
	redemption: number,
	frequency: number,
	basis: number,
	dateSystem: '1900' | '1904',
): number {
	const previous = previousCouponDate(firstCoupon, maturity, frequency, dateSystem)
	const schedule = couponScheduleFromFirst(firstCoupon, maturity, frequency, dateSystem)
	if (previous === null || !schedule || schedule.length === 0) return Number.NaN
	const coupon = (100 * rate) / frequency
	const a = dayCountBasis(issue, settlement, basis, dateSystem)
	const dsc = dayCountBasis(settlement, firstCoupon, basis, dateSystem)
	const dfc = dayCountBasis(issue, firstCoupon, basis, dateSystem)
	const e = couponPeriodDays(previous, firstCoupon, basis, dateSystem)
	if (e <= 0) return Number.NaN
	const q = 1 + yld / frequency
	const n = schedule.length
	let price = redemption / q ** (n - 1 + dsc / e) + (coupon * (dfc / e)) / q ** (dsc / e)
	for (let k = 2; k <= n; k++) price += coupon / q ** (k - 1 + dsc / e)
	return price - (coupon * a) / e
}

function oddFirstBondPrice(
	settlement: number,
	maturity: number,
	issue: number,
	firstCoupon: number,
	rate: number,
	yld: number,
	redemption: number,
	frequency: number,
	basis: number,
	dateSystem: '1900' | '1904',
): number {
	const previous = previousCouponDate(firstCoupon, maturity, frequency, dateSystem)
	if (previous === null) return Number.NaN
	if (issue >= previous) {
		return oddShortFirstBondPrice(
			settlement,
			maturity,
			issue,
			firstCoupon,
			rate,
			yld,
			redemption,
			frequency,
			basis,
			dateSystem,
		)
	}
	const schedule = couponScheduleFromFirst(firstCoupon, maturity, frequency, dateSystem)
	if (!schedule) return Number.NaN
	const q = 1 + yld / frequency
	const regularCoupon = (100 * rate) / frequency
	const firstAmount = oddFirstCouponAmount(issue, firstCoupon, rate, basis, dateSystem)
	const accrued = oddFirstAccrued(issue, settlement, rate, basis, dateSystem)
	let price = -accrued
	for (let i = 0; i < schedule.length; i++) {
		const payDate = schedule[i]
		if (payDate === undefined) continue
		const t = yearFracFromBasis(settlement, payDate, basis, dateSystem) * frequency
		const amount = i === 0 ? firstAmount : regularCoupon
		price += amount / q ** t
	}
	const lastDate = schedule[schedule.length - 1]
	if (lastDate !== undefined) {
		const t = yearFracFromBasis(settlement, lastDate, basis, dateSystem) * frequency
		price += redemption / q ** t
	}
	return price
}

function oddLastBondPrice(
	settlement: number,
	maturity: number,
	lastInterest: number,
	rate: number,
	yld: number,
	redemption: number,
	frequency: number,
	basis: number,
	dateSystem: '1900' | '1904',
): number {
	const c = (100 * rate) / frequency
	const dc = frequency * yearFracFromBasis(lastInterest, maturity, basis, dateSystem)
	const dsc = frequency * yearFracFromBasis(settlement, maturity, basis, dateSystem)
	const a = frequency * yearFracFromBasis(lastInterest, settlement, basis, dateSystem)
	return (redemption + c * dc) / (1 + (yld / frequency) * dsc) - c * a
}

function amorCoeff(rate: number): number | null {
	const life = 1 / rate
	if (life > 0 && life < 3) return null
	if (life >= 3 && life <= 4) return 1.5
	if (life >= 5 && life <= 6) return 2
	if (life > 6) return 2.5
	return null
}

export const financialFunctions: FunctionDef[] = [
	{
		name: 'PMT',
		minArgs: 3,
		maxArgs: 5,
		evaluate(args) {
			const r = num(args[0])
			if (typeof r !== 'number') return r
			const n = num(args[1])
			if (typeof n !== 'number') return n
			const p = num(args[2])
			if (typeof p !== 'number') return p
			const f = args[3] ? num(args[3]) : 0
			if (typeof f !== 'number') return f
			const t = args[4] ? num(args[4]) : 0
			if (typeof t !== 'number') return t
			return numberValue(pmt(r, n, p, f, t))
		},
	},
	{
		name: 'FV',
		minArgs: 3,
		maxArgs: 5,
		evaluate(args) {
			const r = num(args[0])
			if (typeof r !== 'number') return r
			const n = num(args[1])
			if (typeof n !== 'number') return n
			const p = num(args[2])
			if (typeof p !== 'number') return p
			const pv = args[3] ? num(args[3]) : 0
			if (typeof pv !== 'number') return pv
			const t = args[4] ? num(args[4]) : 0
			if (typeof t !== 'number') return t
			return numberValue(fvCalc(r, n, p, pv, t))
		},
	},
	{
		name: 'PV',
		minArgs: 3,
		maxArgs: 5,
		evaluate(args) {
			const r = num(args[0])
			if (typeof r !== 'number') return r
			const n = num(args[1])
			if (typeof n !== 'number') return n
			const p = num(args[2])
			if (typeof p !== 'number') return p
			const fv = args[3] ? num(args[3]) : 0
			if (typeof fv !== 'number') return fv
			const t = args[4] ? num(args[4]) : 0
			if (typeof t !== 'number') return t
			return numberValue(pvCalc(r, n, p, fv, t))
		},
	},
	{
		name: 'NPER',
		minArgs: 3,
		maxArgs: 5,
		evaluate(args) {
			const r = num(args[0])
			if (typeof r !== 'number') return r
			const p = num(args[1])
			if (typeof p !== 'number') return p
			const pv = num(args[2])
			if (typeof pv !== 'number') return pv
			const fv = args[3] ? num(args[3]) : 0
			if (typeof fv !== 'number') return fv
			const t = args[4] ? num(args[4]) : 0
			if (typeof t !== 'number') return t
			const result = nperCalc(r, p, pv, fv, t)
			if (!Number.isFinite(result)) return errorValue('#NUM!')
			return numberValue(result)
		},
	},
	{
		name: 'RATE',
		minArgs: 3,
		maxArgs: 6,
		evaluate(args) {
			const n = num(args[0])
			if (typeof n !== 'number') return n
			const p = num(args[1])
			if (typeof p !== 'number') return p
			const pv = num(args[2])
			if (typeof pv !== 'number') return pv
			const fv = args[3] ? num(args[3]) : 0
			if (typeof fv !== 'number') return fv
			const t = args[4] ? num(args[4]) : 0
			if (typeof t !== 'number') return t
			const guess = args[5] ? num(args[5]) : 0.1
			if (typeof guess !== 'number') return guess
			const result = rateNewton(n, p, pv, fv, t, guess)
			if (Number.isNaN(result)) return errorValue('#NUM!')
			return numberValue(result)
		},
	},
	{
		name: 'IPMT',
		minArgs: 4,
		maxArgs: 6,
		evaluate(args) {
			const rate = num(args[0])
			if (typeof rate !== 'number') return rate
			const per = num(args[1])
			if (typeof per !== 'number') return per
			const nper = num(args[2])
			if (typeof nper !== 'number') return nper
			const pv = num(args[3])
			if (typeof pv !== 'number') return pv
			const fv = args[4] ? num(args[4]) : 0
			if (typeof fv !== 'number') return fv
			const type = args[5] ? num(args[5]) : 0
			if (typeof type !== 'number') return type
			if (per < 1 || per > nper) return errorValue('#NUM!')
			return numberValue(ipmtCalc(rate, per, nper, pv, fv, type))
		},
	},
	{
		name: 'PPMT',
		minArgs: 4,
		maxArgs: 6,
		evaluate(args) {
			const rate = num(args[0])
			if (typeof rate !== 'number') return rate
			const per = num(args[1])
			if (typeof per !== 'number') return per
			const nper = num(args[2])
			if (typeof nper !== 'number') return nper
			const pv = num(args[3])
			if (typeof pv !== 'number') return pv
			const fv = args[4] ? num(args[4]) : 0
			if (typeof fv !== 'number') return fv
			const type = args[5] ? num(args[5]) : 0
			if (typeof type !== 'number') return type
			if (per < 1 || per > nper) return errorValue('#NUM!')
			return numberValue(pmt(rate, nper, pv, fv, type) - ipmtCalc(rate, per, nper, pv, fv, type))
		},
	},
	{
		name: 'NPV',
		minArgs: 2,
		maxArgs: 255,
		evaluate(args) {
			const rate = num(args[0])
			if (typeof rate !== 'number') return rate
			const nums = collectNumbers(args.slice(1))
			if (!Array.isArray(nums)) return nums
			let npv = 0
			const r1 = rate + 1
			let trate = r1
			for (let i = 0; i < nums.length; i++) {
				npv += (nums[i] ?? 0) / trate
				trate *= r1
			}
			return numberValue(npv)
		},
	},
	{
		name: 'IRR',
		minArgs: 1,
		maxArgs: 2,
		evaluate(args) {
			const first = args[0]
			if (!first) return errorValue('#VALUE!')
			const nums = collectNumbers([first])
			if (!Array.isArray(nums)) return nums
			if (nums.length < 2) return errorValue('#NUM!')
			const guess = args[1] ? num(args[1]) : 0.1
			if (typeof guess !== 'number') return guess
			let rate = guess
			const maxIter = 100
			const tol = 1e-10
			for (let i = 0; i < maxIter; i++) {
				let f = 0
				let df = 0
				for (let j = 0; j < nums.length; j++) {
					const v = nums[j] ?? 0
					f += v / (1 + rate) ** j
					if (j > 0) df -= (j * v) / (1 + rate) ** (j + 1)
				}
				if (Math.abs(df) < 1e-20) return errorValue('#NUM!')
				const next = rate - f / df
				if (Math.abs(next - rate) < tol) return numberValue(next)
				rate = next
			}
			return errorValue('#NUM!')
		},
	},
	{
		name: 'SLN',
		minArgs: 3,
		maxArgs: 3,
		evaluate(args) {
			const cost = num(args[0])
			if (typeof cost !== 'number') return cost
			const salvage = num(args[1])
			if (typeof salvage !== 'number') return salvage
			const life = num(args[2])
			if (typeof life !== 'number') return life
			if (life <= 0) return errorValue('#NUM!')
			return numberValue((cost - salvage) / life)
		},
	},
	{
		name: 'SYD',
		minArgs: 4,
		maxArgs: 4,
		evaluate(args) {
			const cost = num(args[0])
			if (typeof cost !== 'number') return cost
			const salvage = num(args[1])
			if (typeof salvage !== 'number') return salvage
			const life = num(args[2])
			if (typeof life !== 'number') return life
			const per = num(args[3])
			if (typeof per !== 'number') return per
			if (life <= 0 || per < 1 || per > life) return errorValue('#NUM!')
			const sumOfYears = (life * (life + 1)) / 2
			return numberValue(((cost - salvage) * (life - per + 1)) / sumOfYears)
		},
	},
	{
		name: 'DDB',
		minArgs: 4,
		maxArgs: 5,
		evaluate(args) {
			const cost = num(args[0])
			if (typeof cost !== 'number') return cost
			const salvage = num(args[1])
			if (typeof salvage !== 'number') return salvage
			const life = num(args[2])
			if (typeof life !== 'number') return life
			const period = num(args[3])
			if (typeof period !== 'number') return period
			const factor = args[4] ? num(args[4]) : 2
			if (typeof factor !== 'number') return factor
			if (cost < 0 || salvage < 0 || life <= 0 || period < 1 || period > life || factor <= 0) {
				return errorValue('#NUM!')
			}
			let bookValue = cost
			for (let p = 1; p < period; p++) {
				const dep = Math.min((bookValue * factor) / life, bookValue - salvage)
				bookValue -= dep
				if (bookValue <= salvage) return numberValue(0)
			}
			const dep = Math.min((bookValue * factor) / life, bookValue - salvage)
			return numberValue(dep)
		},
	},
	{
		name: 'DOLLARDE',
		minArgs: 2,
		maxArgs: 2,
		evaluate(args) {
			const fractionalDollar = num(args[0])
			if (typeof fractionalDollar !== 'number') return fractionalDollar
			const fraction = num(args[1])
			if (typeof fraction !== 'number') return fraction
			const frac = Math.trunc(fraction)
			if (frac <= 0) return errorValue('#NUM!')
			const intPart = Math.trunc(fractionalDollar)
			const fracPart = fractionalDollar - intPart
			const digits = Math.ceil(Math.log10(frac))
			const divisor = 10 ** digits
			const numerator = Math.round(fracPart * divisor)
			return numberValue(intPart + numerator / frac)
		},
	},
	{
		name: 'DOLLARFR',
		minArgs: 2,
		maxArgs: 2,
		evaluate(args) {
			const decimalDollar = num(args[0])
			if (typeof decimalDollar !== 'number') return decimalDollar
			const fraction = num(args[1])
			if (typeof fraction !== 'number') return fraction
			const frac = Math.trunc(fraction)
			if (frac <= 0) return errorValue('#NUM!')
			const intPart = Math.trunc(decimalDollar)
			const fracPart = decimalDollar - intPart
			const numerator = Math.round(fracPart * frac)
			const digits = Math.ceil(Math.log10(frac))
			const divisor = 10 ** digits
			return numberValue(intPart + numerator / divisor)
		},
	},
	{
		name: 'XNPV',
		minArgs: 3,
		maxArgs: 3,
		evaluate(args) {
			const rate = num(args[0])
			if (typeof rate !== 'number') return rate
			if (!Number.isFinite(rate) || rate <= -1) return errorValue('#NUM!')
			const cashFlows = datedCashFlows(args[1], args[2])
			if ('kind' in cashFlows) return cashFlows
			const d0 = cashFlows.dates[0] ?? 0
			let npv = 0
			for (let i = 0; i < cashFlows.values.length; i++) {
				npv += (cashFlows.values[i] ?? 0) / (1 + rate) ** (((cashFlows.dates[i] ?? d0) - d0) / 365)
			}
			return numberValue(npv)
		},
	},
	{
		name: 'XIRR',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const cashFlows = datedCashFlows(args[0], args[1])
			if ('kind' in cashFlows) return cashFlows
			if (cashFlows.values.length < 2) return errorValue('#NUM!')
			const guess = args[2] ? num(args[2]) : 0.1
			if (typeof guess !== 'number') return guess
			if (!Number.isFinite(guess) || guess <= -1) return errorValue('#NUM!')
			const d0 = cashFlows.dates[0] ?? 0
			const v = cashFlows.values
			const t = cashFlows.dates.map((date) => (date - d0) / 365)
			let rate = guess
			const maxIter = 100
			const tol = 1e-10
			for (let iter = 0; iter < maxIter; iter++) {
				if (rate <= -1) return errorValue('#NUM!')
				let f = 0
				let df = 0
				for (let i = 0; i < v.length; i++) {
					const ti = t[i] ?? 0
					const vi = v[i] ?? 0
					const exp = (1 + rate) ** ti
					f += vi / exp
					df -= (ti * vi) / ((1 + rate) * exp)
				}
				if (Math.abs(df) < 1e-20) return errorValue('#NUM!')
				const next = rate - f / df
				if (Math.abs(next - rate) < tol) return numberValue(next)
				rate = next
			}
			return errorValue('#NUM!')
		},
	},
	{
		name: 'DB',
		minArgs: 4,
		maxArgs: 5,
		evaluate(args) {
			const cost = num(args[0])
			if (typeof cost !== 'number') return cost
			const salvage = num(args[1])
			if (typeof salvage !== 'number') return salvage
			const life = num(args[2])
			if (typeof life !== 'number') return life
			const period = num(args[3])
			if (typeof period !== 'number') return period
			const month = args[4] ? num(args[4]) : 12
			if (typeof month !== 'number') return month
			const iLife = Math.trunc(life)
			const iPeriod = Math.trunc(period)
			const iMonth = Math.trunc(month)
			if (cost < 0 || salvage < 0 || iLife <= 0 || iPeriod < 1 || iMonth < 1 || iMonth > 12) {
				return errorValue('#NUM!')
			}
			const maxPeriod = iMonth < 12 ? iLife + 1 : iLife
			if (iPeriod > maxPeriod) return errorValue('#NUM!')
			if (cost === 0) return numberValue(0)
			const rate = Math.round((1 - (salvage / cost) ** (1 / iLife)) * 1000) / 1000
			let bookValue = cost
			for (let p = 1; p < iPeriod; p++) {
				if (p === 1) bookValue -= (cost * rate * iMonth) / 12
				else bookValue -= bookValue * rate
			}
			let dep: number
			if (iPeriod === 1) dep = (cost * rate * iMonth) / 12
			else if (iPeriod === iLife + 1) dep = (bookValue * rate * (12 - iMonth)) / 12
			else dep = bookValue * rate
			return numberValue(dep)
		},
	},
	{
		name: 'VDB',
		minArgs: 5,
		maxArgs: 7,
		evaluate(args) {
			const cost = num(args[0])
			if (typeof cost !== 'number') return cost
			const salvage = num(args[1])
			if (typeof salvage !== 'number') return salvage
			const life = num(args[2])
			if (typeof life !== 'number') return life
			const startPeriod = num(args[3])
			if (typeof startPeriod !== 'number') return startPeriod
			const endPeriod = num(args[4])
			if (typeof endPeriod !== 'number') return endPeriod
			const factor = args[5] ? num(args[5]) : 2
			if (typeof factor !== 'number') return factor
			const noSwitchArg = args[6] ? num(args[6]) : 0
			if (typeof noSwitchArg !== 'number') return noSwitchArg
			if (
				cost < 0 ||
				salvage < 0 ||
				life <= 0 ||
				startPeriod < 0 ||
				endPeriod < startPeriod ||
				factor <= 0
			) {
				return errorValue('#NUM!')
			}
			return numberValue(
				vdbCalc(cost, salvage, life, startPeriod, endPeriod, factor, noSwitchArg !== 0),
			)
		},
	},
	{
		name: 'MIRR',
		minArgs: 3,
		maxArgs: 3,
		evaluate(args) {
			const first = args[0]
			if (!first) return errorValue('#VALUE!')
			const nums = collectNumbers([first])
			if (!Array.isArray(nums)) return nums
			if (nums.length < 2) return errorValue('#DIV/0!')
			const fRate = num(args[1])
			if (typeof fRate !== 'number') return fRate
			const rRate = num(args[2])
			if (typeof rRate !== 'number') return rRate
			const n = nums.length
			let pvNeg = 0
			let fvPos = 0
			let hasNeg = false
			let hasPos = false
			for (let i = 0; i < n; i++) {
				const v = nums[i] ?? 0
				if (v < 0) {
					pvNeg += v / (1 + fRate) ** i
					hasNeg = true
				} else if (v > 0) {
					fvPos += v * (1 + rRate) ** (n - 1 - i)
					hasPos = true
				}
			}
			if (!hasNeg || !hasPos) return errorValue('#DIV/0!')
			const result = (-fvPos / pvNeg) ** (1 / (n - 1)) - 1
			if (!Number.isFinite(result)) return errorValue('#NUM!')
			return numberValue(result)
		},
	},
	{
		name: 'ISPMT',
		minArgs: 4,
		maxArgs: 4,
		evaluate(args) {
			const rate = num(args[0])
			if (typeof rate !== 'number') return rate
			const per = num(args[1])
			if (typeof per !== 'number') return per
			const nper = num(args[2])
			if (typeof nper !== 'number') return nper
			const pv = num(args[3])
			if (typeof pv !== 'number') return pv
			if (nper === 0) return errorValue('#DIV/0!')
			return numberValue(pv * rate * (per / nper - 1))
		},
	},
	{
		name: 'CUMIPMT',
		minArgs: 6,
		maxArgs: 6,
		evaluate(args) {
			const rate = num(args[0])
			if (typeof rate !== 'number') return rate
			const nper = num(args[1])
			if (typeof nper !== 'number') return nper
			const pv = num(args[2])
			if (typeof pv !== 'number') return pv
			const sp = num(args[3])
			if (typeof sp !== 'number') return sp
			const ep = num(args[4])
			if (typeof ep !== 'number') return ep
			const type = num(args[5])
			if (typeof type !== 'number') return type
			const iSp = Math.trunc(sp)
			const iEp = Math.trunc(ep)
			const iType = Math.trunc(type)
			if (
				rate <= 0 ||
				nper <= 0 ||
				pv <= 0 ||
				iSp < 1 ||
				iEp < iSp ||
				(iType !== 0 && iType !== 1)
			) {
				return errorValue('#NUM!')
			}
			let total = 0
			for (let p = iSp; p <= iEp; p++) {
				total += ipmtCalc(rate, p, nper, pv, 0, iType)
			}
			return numberValue(total)
		},
	},
	{
		name: 'CUMPRINC',
		minArgs: 6,
		maxArgs: 6,
		evaluate(args) {
			const rate = num(args[0])
			if (typeof rate !== 'number') return rate
			const nper = num(args[1])
			if (typeof nper !== 'number') return nper
			const pv = num(args[2])
			if (typeof pv !== 'number') return pv
			const sp = num(args[3])
			if (typeof sp !== 'number') return sp
			const ep = num(args[4])
			if (typeof ep !== 'number') return ep
			const type = num(args[5])
			if (typeof type !== 'number') return type
			const iSp = Math.trunc(sp)
			const iEp = Math.trunc(ep)
			const iType = Math.trunc(type)
			if (
				rate <= 0 ||
				nper <= 0 ||
				pv <= 0 ||
				iSp < 1 ||
				iEp < iSp ||
				(iType !== 0 && iType !== 1)
			) {
				return errorValue('#NUM!')
			}
			const pmtVal = pmt(rate, nper, pv, 0, iType)
			let total = 0
			for (let p = iSp; p <= iEp; p++) {
				total += pmtVal - ipmtCalc(rate, p, nper, pv, 0, iType)
			}
			return numberValue(total)
		},
	},
	{
		name: 'EFFECT',
		minArgs: 2,
		maxArgs: 2,
		evaluate(args) {
			const nominalRate = num(args[0])
			if (typeof nominalRate !== 'number') return nominalRate
			const npery = num(args[1])
			if (typeof npery !== 'number') return npery
			const iNpery = Math.trunc(npery)
			if (nominalRate <= 0 || iNpery < 1) return errorValue('#NUM!')
			return numberValue((1 + nominalRate / iNpery) ** iNpery - 1)
		},
	},
	{
		name: 'NOMINAL',
		minArgs: 2,
		maxArgs: 2,
		evaluate(args) {
			const effectRate = num(args[0])
			if (typeof effectRate !== 'number') return effectRate
			const npery = num(args[1])
			if (typeof npery !== 'number') return npery
			const iNpery = Math.trunc(npery)
			if (effectRate <= 0 || iNpery < 1) return errorValue('#NUM!')
			return numberValue(iNpery * ((1 + effectRate) ** (1 / iNpery) - 1))
		},
	},
	{
		name: 'PDURATION',
		minArgs: 3,
		maxArgs: 3,
		evaluate(args) {
			const rate = num(args[0])
			if (typeof rate !== 'number') return rate
			const pv = num(args[1])
			if (typeof pv !== 'number') return pv
			const fv = num(args[2])
			if (typeof fv !== 'number') return fv
			if (rate <= 0 || pv <= 0 || fv <= 0) return errorValue('#NUM!')
			return numberValue((Math.log(fv) - Math.log(pv)) / Math.log(1 + rate))
		},
	},
	{
		name: 'RRI',
		minArgs: 3,
		maxArgs: 3,
		evaluate(args) {
			const nper = num(args[0])
			if (typeof nper !== 'number') return nper
			const pv = num(args[1])
			if (typeof pv !== 'number') return pv
			const fv = num(args[2])
			if (typeof fv !== 'number') return fv
			if (nper <= 0 || pv === 0) return errorValue('#NUM!')
			const result = (fv / pv) ** (1 / nper) - 1
			if (!Number.isFinite(result)) return errorValue('#NUM!')
			return numberValue(result)
		},
	},
	{
		name: 'FVSCHEDULE',
		minArgs: 2,
		maxArgs: 2,
		evaluate(args) {
			const principal = num(args[0])
			if (typeof principal !== 'number') return principal
			const second = args[1]
			if (!second) return errorValue('#VALUE!')
			const rates = collectNumbers([second])
			if (!Array.isArray(rates)) return rates
			let fv = principal
			for (const rate of rates) {
				fv *= 1 + rate
			}
			return numberValue(fv)
		},
	},
	{
		name: 'ACCRINT',
		minArgs: 6,
		maxArgs: 8,
		evaluate(args, ctx) {
			const issue = num(args[0])
			if (typeof issue !== 'number') return issue
			const firstInterest = num(args[1])
			if (typeof firstInterest !== 'number') return firstInterest
			const settlement = num(args[2])
			if (typeof settlement !== 'number') return settlement
			const rate = num(args[3])
			if (typeof rate !== 'number') return rate
			const par = num(args[4])
			if (typeof par !== 'number') return par
			const frequency = num(args[5])
			if (typeof frequency !== 'number') return frequency
			const basis = args[6] ? num(args[6]) : 0
			if (typeof basis !== 'number') return basis
			const calcMethod = args[7] ? num(args[7]) : 1
			if (typeof calcMethod !== 'number') return calcMethod
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (f === null || b === null || issue >= settlement || rate <= 0 || par <= 0) {
				return errorValue('#NUM!')
			}
			const ds = currentDateSystem(ctx)
			if (Math.trunc(calcMethod) === 0) {
				const window = resolveCouponWindow(settlement, firstInterest, f, ds)
				if (!window) return errorValue('#VALUE!')
				const a = dayCountBasis(window.prev, settlement, b, ds)
				const e = couponPeriodDays(window.prev, window.next, b, ds)
				return numberValue(((par * rate) / f) * (a / e))
			}
			return numberValue(par * rate * yearFracFromBasis(issue, settlement, b, ds))
		},
	},
	{
		name: 'COUPPCD',
		minArgs: 3,
		maxArgs: 4,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const frequency = num(args[2])
			if (typeof frequency !== 'number') return frequency
			const basis = args[3] ? num(args[3]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (f === null || b === null || settlement >= maturity) return errorValue('#NUM!')
			const window = resolveCouponWindow(settlement, maturity, f, currentDateSystem(ctx))
			return window ? numberValue(window.prev) : errorValue('#VALUE!')
		},
	},
	{
		name: 'COUPNCD',
		minArgs: 3,
		maxArgs: 4,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const frequency = num(args[2])
			if (typeof frequency !== 'number') return frequency
			const basis = args[3] ? num(args[3]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (f === null || b === null || settlement >= maturity) return errorValue('#NUM!')
			const window = resolveCouponWindow(settlement, maturity, f, currentDateSystem(ctx))
			return window ? numberValue(window.next) : errorValue('#VALUE!')
		},
	},
	{
		name: 'COUPDAYBS',
		minArgs: 3,
		maxArgs: 4,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const frequency = num(args[2])
			if (typeof frequency !== 'number') return frequency
			const basis = args[3] ? num(args[3]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (f === null || b === null || settlement >= maturity) return errorValue('#NUM!')
			const ds = currentDateSystem(ctx)
			const window = resolveCouponWindow(settlement, maturity, f, ds)
			return window
				? numberValue(dayCountBasis(window.prev, settlement, b, ds))
				: errorValue('#VALUE!')
		},
	},
	{
		name: 'COUPDAYS',
		minArgs: 3,
		maxArgs: 4,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const frequency = num(args[2])
			if (typeof frequency !== 'number') return frequency
			const basis = args[3] ? num(args[3]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (f === null || b === null || settlement >= maturity) return errorValue('#NUM!')
			const ds = currentDateSystem(ctx)
			const window = resolveCouponWindow(settlement, maturity, f, ds)
			return window
				? numberValue(couponPeriodDays(window.prev, window.next, b, ds))
				: errorValue('#VALUE!')
		},
	},
	{
		name: 'COUPDAYSNC',
		minArgs: 3,
		maxArgs: 4,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const frequency = num(args[2])
			if (typeof frequency !== 'number') return frequency
			const basis = args[3] ? num(args[3]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (f === null || b === null || settlement >= maturity) return errorValue('#NUM!')
			const ds = currentDateSystem(ctx)
			const window = resolveCouponWindow(settlement, maturity, f, ds)
			return window
				? numberValue(dayCountBasis(settlement, window.next, b, ds))
				: errorValue('#VALUE!')
		},
	},
	{
		name: 'COUPNUM',
		minArgs: 3,
		maxArgs: 4,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const frequency = num(args[2])
			if (typeof frequency !== 'number') return frequency
			const basis = args[3] ? num(args[3]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (f === null || b === null || settlement >= maturity) return errorValue('#NUM!')
			return numberValue(couponCountRemaining(settlement, maturity, f, currentDateSystem(ctx)))
		},
	},
	{
		name: 'ACCRINTM',
		minArgs: 4,
		maxArgs: 5,
		evaluate(args, ctx) {
			const issue = num(args[0])
			if (typeof issue !== 'number') return issue
			const settlement = num(args[1])
			if (typeof settlement !== 'number') return settlement
			const rate = num(args[2])
			if (typeof rate !== 'number') return rate
			const par = num(args[3])
			if (typeof par !== 'number') return par
			const basis = args[4] ? num(args[4]) : 0
			if (typeof basis !== 'number') return basis
			const b = validateBasis(basis)
			if (b === null || issue >= settlement || rate <= 0 || par <= 0) return errorValue('#NUM!')
			return numberValue(
				par * rate * yearFracFromBasis(issue, settlement, b, currentDateSystem(ctx)),
			)
		},
	},
	{
		name: 'PRICE',
		minArgs: 6,
		maxArgs: 7,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const rate = num(args[2])
			if (typeof rate !== 'number') return rate
			const yld = num(args[3])
			if (typeof yld !== 'number') return yld
			const redemption = num(args[4])
			if (typeof redemption !== 'number') return redemption
			const frequency = num(args[5])
			if (typeof frequency !== 'number') return frequency
			const basis = args[6] ? num(args[6]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (
				f === null ||
				b === null ||
				settlement >= maturity ||
				rate < 0 ||
				yld < 0 ||
				redemption <= 0
			)
				return errorValue('#NUM!')
			return numberValue(
				regularBondPrice(settlement, maturity, rate, yld, redemption, f, b, currentDateSystem(ctx)),
			)
		},
	},
	{
		name: 'YIELD',
		minArgs: 6,
		maxArgs: 7,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const rate = num(args[2])
			if (typeof rate !== 'number') return rate
			const pr = num(args[3])
			if (typeof pr !== 'number') return pr
			const redemption = num(args[4])
			if (typeof redemption !== 'number') return redemption
			const frequency = num(args[5])
			if (typeof frequency !== 'number') return frequency
			const basis = args[6] ? num(args[6]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (
				f === null ||
				b === null ||
				settlement >= maturity ||
				rate < 0 ||
				pr <= 0 ||
				redemption <= 0
			)
				return errorValue('#NUM!')
			const ds = currentDateSystem(ctx)
			const window = resolveCouponWindow(settlement, maturity, f, ds)
			if (!window) return errorValue('#VALUE!')
			const coupon = (100 * rate) / f
			const a = dayCountBasis(window.prev, settlement, b, ds)
			const e = couponPeriodDays(window.prev, window.next, b, ds)
			const dsr = e - a
			const n = couponCountRemaining(settlement, maturity, f, ds)
			if (n <= 1) {
				const result =
					((redemption + coupon - (pr + (coupon * a) / e)) / (pr + (coupon * a) / e)) *
					f *
					(e / dsr)
				return numberValue(result)
			}
			let guess = rate || 0.05
			for (let i = 0; i < 100; i++) {
				const price = regularBondPrice(settlement, maturity, rate, guess, redemption, f, b, ds)
				const delta = 1e-6
				const price2 = regularBondPrice(
					settlement,
					maturity,
					rate,
					guess + delta,
					redemption,
					f,
					b,
					ds,
				)
				const deriv = (price2 - price) / delta
				if (Math.abs(deriv) < 1e-12) break
				const next = guess - (price - pr) / deriv
				if (Math.abs(next - guess) < 1e-10) return numberValue(next)
				guess = next
			}
			return errorValue('#NUM!')
		},
	},
	{
		name: 'DURATION',
		minArgs: 5,
		maxArgs: 6,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const couponRate = num(args[2])
			if (typeof couponRate !== 'number') return couponRate
			const yld = num(args[3])
			if (typeof yld !== 'number') return yld
			const frequency = num(args[4])
			if (typeof frequency !== 'number') return frequency
			const basis = args[5] ? num(args[5]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (f === null || b === null || settlement >= maturity || couponRate < 0 || yld < 0)
				return errorValue('#NUM!')
			const ds = currentDateSystem(ctx)
			const window = resolveCouponWindow(settlement, maturity, f, ds)
			if (!window) return errorValue('#VALUE!')
			const n = couponCountRemaining(settlement, maturity, f, ds)
			const coupon = (100 * couponRate) / f
			const e = couponPeriodDays(window.prev, window.next, b, ds)
			const dsc = dayCountBasis(settlement, window.next, b, ds)
			const q = 1 + yld / f
			let pvTotal = 0
			let weighted = 0
			for (let k = 1; k <= n; k++) {
				const time = (k - 1 + dsc / e) / f
				const cash = k === n ? coupon + 100 : coupon
				const pv = cash / q ** (k - 1 + dsc / e)
				pvTotal += pv
				weighted += time * pv
			}
			return numberValue(weighted / pvTotal)
		},
	},
	{
		name: 'MDURATION',
		minArgs: 5,
		maxArgs: 6,
		evaluate(args, ctx) {
			const duration = (
				financialFunctions.find((fn) => fn.name === 'DURATION') as FunctionDef
			).evaluate(args, ctx)
			if (duration.kind !== 'number') return duration
			const frequency = num(args[4])
			if (typeof frequency !== 'number') return frequency
			const yld = num(args[3])
			if (typeof yld !== 'number') return yld
			return numberValue(duration.value / (1 + yld / frequency))
		},
	},
	{
		name: 'PRICEDISC',
		minArgs: 4,
		maxArgs: 5,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const discount = num(args[2])
			if (typeof discount !== 'number') return discount
			const redemption = num(args[3])
			if (typeof redemption !== 'number') return redemption
			const basis = args[4] ? num(args[4]) : 0
			if (typeof basis !== 'number') return basis
			const b = validateBasis(basis)
			if (b === null || settlement >= maturity || discount <= 0 || redemption <= 0)
				return errorValue('#NUM!')
			const yf = yearFracFromBasis(settlement, maturity, b, currentDateSystem(ctx))
			return numberValue(redemption * (1 - discount * yf))
		},
	},
	{
		name: 'YIELDDISC',
		minArgs: 4,
		maxArgs: 5,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const pr = num(args[2])
			if (typeof pr !== 'number') return pr
			const redemption = num(args[3])
			if (typeof redemption !== 'number') return redemption
			const basis = args[4] ? num(args[4]) : 0
			if (typeof basis !== 'number') return basis
			const b = validateBasis(basis)
			if (b === null || settlement >= maturity || pr <= 0 || redemption <= 0)
				return errorValue('#NUM!')
			const yf = yearFracFromBasis(settlement, maturity, b, currentDateSystem(ctx))
			return numberValue((redemption / pr - 1) / yf)
		},
	},
	{
		name: 'PRICEMAT',
		minArgs: 5,
		maxArgs: 6,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const issue = num(args[2])
			if (typeof issue !== 'number') return issue
			const rate = num(args[3])
			if (typeof rate !== 'number') return rate
			const yld = num(args[4])
			if (typeof yld !== 'number') return yld
			const basis = args[5] ? num(args[5]) : 0
			if (typeof basis !== 'number') return basis
			const b = validateBasis(basis)
			if (b === null || issue >= settlement || settlement >= maturity || rate < 0 || yld < 0)
				return errorValue('#NUM!')
			const ds = currentDateSystem(ctx)
			const dim = yearFracFromBasis(issue, maturity, b, ds)
			const dsm = yearFracFromBasis(settlement, maturity, b, ds)
			const a = yearFracFromBasis(issue, settlement, b, ds)
			return numberValue(100 * ((1 + dim * rate) / (1 + dsm * yld) - a * rate))
		},
	},
	{
		name: 'YIELDMAT',
		minArgs: 5,
		maxArgs: 6,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const issue = num(args[2])
			if (typeof issue !== 'number') return issue
			const rate = num(args[3])
			if (typeof rate !== 'number') return rate
			const pr = num(args[4])
			if (typeof pr !== 'number') return pr
			const basis = args[5] ? num(args[5]) : 0
			if (typeof basis !== 'number') return basis
			const b = validateBasis(basis)
			if (b === null || issue >= settlement || settlement >= maturity || rate < 0 || pr <= 0)
				return errorValue('#NUM!')
			const ds = currentDateSystem(ctx)
			const dim = yearFracFromBasis(issue, maturity, b, ds)
			const dsm = yearFracFromBasis(settlement, maturity, b, ds)
			const a = yearFracFromBasis(issue, settlement, b, ds)
			return numberValue(((1 + dim * rate) / (pr / 100 + a * rate) - 1) / dsm)
		},
	},
	{
		name: 'RECEIVED',
		minArgs: 4,
		maxArgs: 5,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const investment = num(args[2])
			if (typeof investment !== 'number') return investment
			const discount = num(args[3])
			if (typeof discount !== 'number') return discount
			const basis = args[4] ? num(args[4]) : 0
			if (typeof basis !== 'number') return basis
			const b = validateBasis(basis)
			if (b === null || settlement >= maturity || investment <= 0 || discount <= 0)
				return errorValue('#NUM!')
			const yf = yearFracFromBasis(settlement, maturity, b, currentDateSystem(ctx))
			return numberValue(investment / (1 - discount * yf))
		},
	},
	{
		name: 'ODDFPRICE',
		minArgs: 8,
		maxArgs: 9,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const issue = num(args[2])
			if (typeof issue !== 'number') return issue
			const firstCoupon = num(args[3])
			if (typeof firstCoupon !== 'number') return firstCoupon
			const rate = num(args[4])
			if (typeof rate !== 'number') return rate
			const yld = num(args[5])
			if (typeof yld !== 'number') return yld
			const redemption = num(args[6])
			if (typeof redemption !== 'number') return redemption
			const frequency = num(args[7])
			if (typeof frequency !== 'number') return frequency
			const basis = args[8] ? num(args[8]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (
				f === null ||
				b === null ||
				!(maturity > firstCoupon && firstCoupon > settlement && settlement > issue) ||
				rate < 0 ||
				yld < 0 ||
				redemption <= 0
			) {
				return errorValue('#NUM!')
			}
			return numberValue(
				oddFirstBondPrice(
					settlement,
					maturity,
					issue,
					firstCoupon,
					rate,
					yld,
					redemption,
					f,
					b,
					currentDateSystem(ctx),
				),
			)
		},
	},
	{
		name: 'ODDFYIELD',
		minArgs: 8,
		maxArgs: 9,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const issue = num(args[2])
			if (typeof issue !== 'number') return issue
			const firstCoupon = num(args[3])
			if (typeof firstCoupon !== 'number') return firstCoupon
			const rate = num(args[4])
			if (typeof rate !== 'number') return rate
			const pr = num(args[5])
			if (typeof pr !== 'number') return pr
			const redemption = num(args[6])
			if (typeof redemption !== 'number') return redemption
			const frequency = num(args[7])
			if (typeof frequency !== 'number') return frequency
			const basis = args[8] ? num(args[8]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (
				f === null ||
				b === null ||
				!(maturity > firstCoupon && firstCoupon > settlement && settlement > issue) ||
				rate < 0 ||
				pr <= 0 ||
				redemption <= 0
			) {
				return errorValue('#NUM!')
			}
			const ds = currentDateSystem(ctx)
			let guess = rate || 0.05
			for (let i = 0; i < 100; i++) {
				const price = oddFirstBondPrice(
					settlement,
					maturity,
					issue,
					firstCoupon,
					rate,
					guess,
					redemption,
					f,
					b,
					ds,
				)
				const delta = 1e-6
				const price2 = oddFirstBondPrice(
					settlement,
					maturity,
					issue,
					firstCoupon,
					rate,
					guess + delta,
					redemption,
					f,
					b,
					ds,
				)
				const deriv = (price2 - price) / delta
				if (Math.abs(deriv) < 1e-12) break
				const next = guess - (price - pr) / deriv
				if (Math.abs(next - guess) < 1e-10) return numberValue(next)
				guess = next
			}
			return errorValue('#NUM!')
		},
	},
	{
		name: 'ODDLPRICE',
		minArgs: 7,
		maxArgs: 8,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const lastInterest = num(args[2])
			if (typeof lastInterest !== 'number') return lastInterest
			const rate = num(args[3])
			if (typeof rate !== 'number') return rate
			const yld = num(args[4])
			if (typeof yld !== 'number') return yld
			const redemption = num(args[5])
			if (typeof redemption !== 'number') return redemption
			const frequency = num(args[6])
			if (typeof frequency !== 'number') return frequency
			const basis = args[7] ? num(args[7]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (
				f === null ||
				b === null ||
				!(maturity > settlement && settlement > lastInterest) ||
				rate < 0 ||
				yld < 0 ||
				redemption <= 0
			) {
				return errorValue('#NUM!')
			}
			return numberValue(
				oddLastBondPrice(
					settlement,
					maturity,
					lastInterest,
					rate,
					yld,
					redemption,
					f,
					b,
					currentDateSystem(ctx),
				),
			)
		},
	},
	{
		name: 'ODDLYIELD',
		minArgs: 7,
		maxArgs: 8,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const lastInterest = num(args[2])
			if (typeof lastInterest !== 'number') return lastInterest
			const rate = num(args[3])
			if (typeof rate !== 'number') return rate
			const pr = num(args[4])
			if (typeof pr !== 'number') return pr
			const redemption = num(args[5])
			if (typeof redemption !== 'number') return redemption
			const frequency = num(args[6])
			if (typeof frequency !== 'number') return frequency
			const basis = args[7] ? num(args[7]) : 0
			if (typeof basis !== 'number') return basis
			const f = validateFrequency(frequency)
			const b = validateBasis(basis)
			if (
				f === null ||
				b === null ||
				!(maturity > settlement && settlement > lastInterest) ||
				rate < 0 ||
				pr <= 0 ||
				redemption <= 0
			) {
				return errorValue('#NUM!')
			}
			const ds = currentDateSystem(ctx)
			const c = (100 * rate) / f
			const dc = f * yearFracFromBasis(lastInterest, maturity, b, ds)
			const dsc = f * yearFracFromBasis(settlement, maturity, b, ds)
			const a = f * yearFracFromBasis(lastInterest, settlement, b, ds)
			return numberValue(((redemption + c * dc - (pr + c * a)) / (pr + c * a)) * (f / dsc))
		},
	},
	{
		name: 'AMORLINC',
		minArgs: 6,
		maxArgs: 7,
		evaluate(args, ctx) {
			const cost = num(args[0])
			if (typeof cost !== 'number') return cost
			const purchased = num(args[1])
			if (typeof purchased !== 'number') return purchased
			const firstPeriod = num(args[2])
			if (typeof firstPeriod !== 'number') return firstPeriod
			const salvage = num(args[3])
			if (typeof salvage !== 'number') return salvage
			const period = num(args[4])
			if (typeof period !== 'number') return period
			const rate = num(args[5])
			if (typeof rate !== 'number') return rate
			const basis = args[6] ? num(args[6]) : 0
			if (typeof basis !== 'number') return basis
			const b = validateBasis(basis)
			const p = Math.trunc(period)
			if (
				b === null ||
				b === 2 ||
				rate <= 0 ||
				p < 0 ||
				salvage < 0 ||
				salvage > cost ||
				purchased > firstPeriod
			) {
				return errorValue('#NUM!')
			}
			const ds = currentDateSystem(ctx)
			const first = cost * rate * yearFracFromBasis(purchased, firstPeriod, b, ds)
			const full = cost * rate
			if (p === 0) return numberValue(first)
			const remaining = cost - salvage - first
			if (remaining <= 0) return numberValue(0)
			if (remaining >= full * p) return numberValue(full)
			return numberValue(Math.max(0, remaining - full * (p - 1)))
		},
	},
	{
		name: 'AMORDEGRC',
		minArgs: 6,
		maxArgs: 7,
		evaluate(args, ctx) {
			const cost = num(args[0])
			if (typeof cost !== 'number') return cost
			const purchased = num(args[1])
			if (typeof purchased !== 'number') return purchased
			const firstPeriod = num(args[2])
			if (typeof firstPeriod !== 'number') return firstPeriod
			const salvage = num(args[3])
			if (typeof salvage !== 'number') return salvage
			const period = num(args[4])
			if (typeof period !== 'number') return period
			const rate = num(args[5])
			if (typeof rate !== 'number') return rate
			const basis = args[6] ? num(args[6]) : 0
			if (typeof basis !== 'number') return basis
			const b = validateBasis(basis)
			const p = Math.trunc(period)
			const coeff = amorCoeff(rate)
			if (
				b === null ||
				b === 2 ||
				coeff === null ||
				rate <= 0 ||
				p < 0 ||
				salvage < 0 ||
				salvage > cost ||
				purchased > firstPeriod
			) {
				return errorValue('#NUM!')
			}
			const ds = currentDateSystem(ctx)
			const first = Math.round(
				cost * rate * coeff * yearFracFromBasis(purchased, firstPeriod, b, ds),
			)
			if (p === 0) return numberValue(first)
			let book = cost - first
			for (let i = 1; i <= p; i++) {
				const dep = Math.min(Math.round(book * rate * coeff), Math.max(0, book - salvage))
				if (i === p) return numberValue(dep)
				book -= dep
				if (book <= salvage) return numberValue(0)
			}
			return numberValue(0)
		},
	},
	{
		name: 'TBILLPRICE',
		minArgs: 3,
		maxArgs: 3,
		evaluate(args) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const discount = num(args[2])
			if (typeof discount !== 'number') return discount
			const dsm = Math.floor(maturity) - Math.floor(settlement)
			if (settlement >= maturity || discount <= 0 || dsm > 365) return errorValue('#NUM!')
			return numberValue(100 * (1 - (discount * dsm) / 360))
		},
	},
	{
		name: 'TBILLEQ',
		minArgs: 3,
		maxArgs: 3,
		evaluate(args) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const discount = num(args[2])
			if (typeof discount !== 'number') return discount
			const dsm = Math.floor(maturity) - Math.floor(settlement)
			if (settlement >= maturity || discount <= 0 || dsm > 365) return errorValue('#NUM!')
			return numberValue((365 * discount) / (360 - discount * dsm))
		},
	},
	{
		name: 'TBILLYIELD',
		minArgs: 3,
		maxArgs: 3,
		evaluate(args) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const pr = num(args[2])
			if (typeof pr !== 'number') return pr
			const dsm = Math.floor(maturity) - Math.floor(settlement)
			if (settlement >= maturity || pr <= 0 || dsm > 365) return errorValue('#NUM!')
			return numberValue(((100 - pr) / pr) * (360 / dsm))
		},
	},
	{
		name: 'DISC',
		minArgs: 4,
		maxArgs: 5,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const pr = num(args[2])
			if (typeof pr !== 'number') return pr
			const redemption = num(args[3])
			if (typeof redemption !== 'number') return redemption
			const basis = args[4] ? num(args[4]) : 0
			if (typeof basis !== 'number') return basis
			const b = validateBasis(basis)
			if (b === null || settlement >= maturity || pr <= 0 || redemption <= 0) {
				return errorValue('#NUM!')
			}
			const yf = yearFracFromBasis(settlement, maturity, b, currentDateSystem(ctx))
			if (yf <= 0 || !Number.isFinite(yf)) return errorValue('#NUM!')
			return numberValue((redemption - pr) / redemption / yf)
		},
	},
	{
		name: 'INTRATE',
		minArgs: 4,
		maxArgs: 5,
		evaluate(args, ctx) {
			const settlement = num(args[0])
			if (typeof settlement !== 'number') return settlement
			const maturity = num(args[1])
			if (typeof maturity !== 'number') return maturity
			const investment = num(args[2])
			if (typeof investment !== 'number') return investment
			const redemption = num(args[3])
			if (typeof redemption !== 'number') return redemption
			const basis = args[4] ? num(args[4]) : 0
			if (typeof basis !== 'number') return basis
			const b = validateBasis(basis)
			if (b === null || settlement >= maturity || investment <= 0 || redemption <= 0) {
				return errorValue('#NUM!')
			}
			const yf = yearFracFromBasis(settlement, maturity, b, currentDateSystem(ctx))
			if (yf <= 0 || !Number.isFinite(yf)) return errorValue('#NUM!')
			return numberValue((redemption - investment) / investment / yf)
		},
	},
]
