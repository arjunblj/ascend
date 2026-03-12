import type { CellValue } from '@ascend/schema'
import { errorValue, numberValue } from '@ascend/schema'
import type { EvalArg, FunctionDef } from './registry.ts'
import { collectNumbers, getRange, numArg, toNumber } from './registry.ts'

function num(arg: EvalArg | undefined): number | CellValue {
	return numArg(arg)
}

function pmt(rate: number, nper: number, pv: number, fv: number, type: number): number {
	if (rate === 0) return -(pv + fv) / nper
	const temp = (1 + rate) ** nper
	const fact = ((1 + rate * type) * (temp - 1)) / rate
	return -(fv + pv * temp) / fact
}

function fvCalc(rate: number, nper: number, pmtVal: number, pv: number, type: number): number {
	if (rate === 0) return -(pv + pmtVal * nper)
	const temp = (1 + rate) ** nper
	return -(pv * temp + (pmtVal * (1 + rate * type) * (temp - 1)) / rate)
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

function yearFracBasis(dsm: number, basis: number): number {
	const b = Math.trunc(basis)
	if (b === 2 || b === 0 || b === 4) return dsm / 360
	return dsm / 365
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
			const values = getRange(args[1]).flat()
			const dates = getRange(args[2]).flat()
			if (values.length !== dates.length || values.length === 0) return errorValue('#NUM!')
			const first = dates[0]
			if (!first) return errorValue('#VALUE!')
			const d0 = toNumber(first)
			if (d0 === null) return errorValue('#VALUE!')
			let npv = 0
			for (let i = 0; i < values.length; i++) {
				const vi = values[i]
				const di = dates[i]
				if (!vi || !di) return errorValue('#VALUE!')
				const v = toNumber(vi)
				const d = toNumber(di)
				if (v === null || d === null) return errorValue('#VALUE!')
				npv += v / (1 + rate) ** ((d - d0) / 365)
			}
			return numberValue(npv)
		},
	},
	{
		name: 'XIRR',
		minArgs: 2,
		maxArgs: 3,
		evaluate(args) {
			const values = getRange(args[0]).flat()
			const dates = getRange(args[1]).flat()
			if (values.length !== dates.length || values.length < 2) return errorValue('#NUM!')
			const guess = args[2] ? num(args[2]) : 0.1
			if (typeof guess !== 'number') return guess
			const firstDate = dates[0]
			if (!firstDate) return errorValue('#VALUE!')
			const d0 = toNumber(firstDate)
			if (d0 === null) return errorValue('#VALUE!')
			const v: number[] = []
			const t: number[] = []
			for (let i = 0; i < values.length; i++) {
				const vi = values[i]
				const di = dates[i]
				if (!vi || !di) return errorValue('#VALUE!')
				const vn = toNumber(vi)
				const dn = toNumber(di)
				if (vn === null || dn === null) return errorValue('#VALUE!')
				v.push(vn)
				t.push((dn - d0) / 365)
			}
			let rate = guess
			const maxIter = 100
			const tol = 1e-10
			for (let iter = 0; iter < maxIter; iter++) {
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
		name: 'DISC',
		minArgs: 4,
		maxArgs: 5,
		evaluate(args) {
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
			const iBasis = Math.trunc(basis)
			if (iBasis < 0 || iBasis > 4) return errorValue('#NUM!')
			const dsm = Math.floor(maturity) - Math.floor(settlement)
			if (dsm <= 0 || pr <= 0 || redemption <= 0) return errorValue('#NUM!')
			const yf = yearFracBasis(dsm, iBasis)
			return numberValue((redemption - pr) / redemption / yf)
		},
	},
	{
		name: 'INTRATE',
		minArgs: 4,
		maxArgs: 5,
		evaluate(args) {
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
			const iBasis = Math.trunc(basis)
			if (iBasis < 0 || iBasis > 4) return errorValue('#NUM!')
			const dsm = Math.floor(maturity) - Math.floor(settlement)
			if (dsm <= 0 || investment <= 0 || redemption <= 0) return errorValue('#NUM!')
			const yf = yearFracBasis(dsm, iBasis)
			return numberValue((redemption - investment) / investment / yf)
		},
	},
]
