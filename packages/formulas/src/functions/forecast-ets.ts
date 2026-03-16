import type { CellValue } from '@ascend/schema'
import { errorValue, numberValue } from '@ascend/schema'
import type { EvalArg, FunctionDef } from './registry.ts'
import { getRange, numArg } from './registry.ts'

function collectPairedETS(
	arg1: EvalArg | undefined,
	arg2: EvalArg | undefined,
): [number[], number[]] | CellValue {
	const r1 = getRange(arg1)
	const r2 = getRange(arg2)
	const rows1 = r1.length
	const cols1 = r1[0]?.length ?? 0
	if (rows1 !== r2.length || cols1 !== (r2[0]?.length ?? 0)) return errorValue('#N/A')
	const a: number[] = []
	const b: number[] = []
	for (let r = 0; r < rows1; r++) {
		for (let c = 0; c < cols1; c++) {
			const v1 = r1[r]?.[c]
			const v2 = r2[r]?.[c]
			if (!v1 || !v2) continue
			if (v1.kind === 'error') return v1
			if (v2.kind === 'error') return v2
			const n1 = v1.kind === 'number' ? v1.value : v1.kind === 'date' ? v1.serial : null
			const n2 = v2.kind === 'number' ? v2.value : v2.kind === 'date' ? v2.serial : null
			if (n1 !== null && n2 !== null) {
				a.push(n1)
				b.push(n2)
			}
		}
	}
	if (a.length < 3) return errorValue('#VALUE!')
	return [a, b]
}

function at(arr: number[], i: number): number {
	return arr[i] as number
}

function prepareSeries(
	rawVals: number[],
	rawTimes: number[],
	dataCompletion: number,
	aggregation: number,
): { values: number[]; timeline: number[]; stepSize: number } | CellValue {
	const idx = rawTimes.map((_, i) => i)
	idx.sort((a, b) => at(rawTimes, a) - at(rawTimes, b))
	const sv = idx.map((i) => at(rawVals, i))
	const st = idx.map((i) => at(rawTimes, i))

	const seen = new Map<number, { sum: number; count: number }>()
	for (let i = 0; i < st.length; i++) {
		const t = at(st, i)
		const v = at(sv, i)
		const e = seen.get(t)
		if (e) {
			e.sum += v
			e.count++
		} else {
			seen.set(t, { sum: v, count: 1 })
		}
	}
	const timeline: number[] = []
	const values: number[] = []
	for (const [t, { sum, count }] of seen) {
		const bucket = idx
			.filter((i) => at(rawTimes, i) === t)
			.map((i) => at(rawVals, i))
			.sort((a, b) => a - b)
		timeline.push(t)
		switch (aggregation) {
			case 2:
			case 3:
				values.push(count)
				break
			case 4:
				values.push(at(bucket, bucket.length - 1))
				break
			case 5:
				values.push(
					bucket.length % 2 === 1
						? at(bucket, Math.floor(bucket.length / 2))
						: (at(bucket, bucket.length / 2 - 1) + at(bucket, bucket.length / 2)) / 2,
				)
				break
			case 6:
				values.push(at(bucket, 0))
				break
			case 7:
				values.push(sum)
				break
			default:
				values.push(sum / count)
				break
		}
	}

	if (timeline.length < 2) return { values, timeline, stepSize: 1 }
	const diffs: number[] = []
	for (let i = 1; i < timeline.length; i++) {
		const d = at(timeline, i) - at(timeline, i - 1)
		if (d > 0) diffs.push(d)
	}
	if (diffs.length === 0) return { values, timeline, stepSize: 1 }
	diffs.sort((a, b) => a - b)
	const stepSize = at(diffs, 0)
	if (stepSize <= 0) return errorValue('#NUM!')
	for (const diff of diffs) {
		const multiple = diff / stepSize
		if (Math.abs(multiple - Math.round(multiple)) > 1e-9) return errorValue('#NUM!')
	}
	const completedTimeline: number[] = [at(timeline, 0)]
	const completedValues: number[] = [at(values, 0)]
	for (let i = 1; i < timeline.length; i++) {
		const prevT = at(timeline, i - 1)
		const prevV = at(values, i - 1)
		const nextT = at(timeline, i)
		const nextV = at(values, i)
		const gap = Math.round((nextT - prevT) / stepSize)
		for (let g = 1; g < gap; g++) {
			completedTimeline.push(prevT + g * stepSize)
			if (dataCompletion === 0) {
				completedValues.push(0)
			} else {
				completedValues.push(prevV + ((nextV - prevV) * g) / gap)
			}
		}
		completedTimeline.push(nextT)
		completedValues.push(nextV)
	}
	return { values: completedValues, timeline: completedTimeline, stepSize }
}

function detectPeriod(values: number[]): number {
	const n = values.length
	if (n < 6) return 1
	const mean = values.reduce((a, b) => a + b, 0) / n
	let v0 = 0
	for (const v of values) v0 += (v - mean) ** 2
	if (v0 === 0) return 1
	const maxLag = Math.min(Math.floor(n / 2), n - 1)
	const acf: number[] = []
	for (let lag = 1; lag <= maxLag; lag++) {
		let s = 0
		for (let i = 0; i < n - lag; i++) s += (at(values, i) - mean) * (at(values, i + lag) - mean)
		acf.push(s / v0)
	}
	const thr = 1.96 / Math.sqrt(n)
	for (let i = 1; i < acf.length - 1; i++) {
		if (at(acf, i) > at(acf, i - 1) && at(acf, i) > at(acf, i + 1) && at(acf, i) > thr) {
			const p = i + 1
			if (p >= 2 && n >= 2 * p) return p
		}
	}
	return 1
}

interface ETSModel {
	alpha: number
	beta: number
	gamma: number
	level: number
	trend: number
	seasonal: number[]
	period: number
	stepSize: number
	mse: number
	residuals: number[]
}

function hwRun(
	values: number[],
	alpha: number,
	beta: number,
	gamma: number,
	m: number,
): { mse: number; level: number; trend: number; seasonal: number[]; residuals: number[] } {
	const n = values.length
	const sea = m > 1
	let L: number
	let T: number
	const S = new Array<number>(sea ? m : 0)

	if (sea && n >= 2 * m) {
		let s1 = 0
		for (let i = 0; i < m; i++) s1 += at(values, i)
		L = s1 / m
		let s2 = 0
		for (let i = 0; i < m; i++) s2 += (at(values, i + m) - at(values, i)) / m
		T = s2 / m
		for (let i = 0; i < m; i++) S[i] = L !== 0 ? at(values, i) / L : 1
	} else if (sea) {
		L = at(values, 0)
		T = n >= 2 ? at(values, 1) - at(values, 0) : 0
		for (let i = 0; i < m; i++) {
			S[i] = L !== 0 ? at(values, i % n) / L || 0.001 : 1
		}
	} else {
		L = at(values, 0)
		T = n >= 2 ? at(values, 1) - at(values, 0) : 0
	}

	let sse = 0
	let cnt = 0
	const residuals: number[] = []
	for (let t = 0; t < n; t++) {
		const y = at(values, t)
		const fc = sea ? (L + T) * (at(S, t % m) || 1) : L + T
		const e = y - fc
		residuals.push(e)
		if (t > 0) {
			sse += e * e
			cnt++
		}
		const pL = L
		if (sea) {
			const si = at(S, t % m) || 1
			L = alpha * (y / si) + (1 - alpha) * (pL + T)
			T = beta * (L - pL) + (1 - beta) * T
			S[t % m] = gamma * (y / (L || 1)) + (1 - gamma) * si
		} else {
			L = alpha * y + (1 - alpha) * (pL + T)
			T = beta * (L - pL) + (1 - beta) * T
		}
	}
	return { mse: cnt > 0 ? sse / cnt : 0, level: L, trend: T, seasonal: S, residuals }
}

function fitETS(values: number[], period: number): ETSModel {
	const aGrid = [0.05, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]
	const bGrid = [0.01, 0.05, 0.1, 0.2, 0.3, 0.5]
	const gGrid = period > 1 ? [0.01, 0.05, 0.1, 0.2, 0.3, 0.5] : [0]

	let bestMSE = Number.POSITIVE_INFINITY
	let bestA = 0.3
	let bestB = 0.05
	let bestG = 0
	let best = hwRun(values, 0.3, 0.05, 0, period)
	bestMSE = best.mse

	for (const a of aGrid) {
		for (const b of bGrid) {
			for (const g of gGrid) {
				const r = hwRun(values, a, b, g, period)
				if (r.mse < bestMSE) {
					bestMSE = r.mse
					bestA = a
					bestB = b
					bestG = g
					best = r
				}
			}
		}
	}

	return {
		alpha: bestA,
		beta: bestB,
		gamma: bestG,
		level: best.level,
		trend: best.trend,
		seasonal: best.seasonal,
		period,
		stepSize: 0,
		mse: bestMSE,
		residuals: best.residuals,
	}
}

function buildETSModel(
	valuesArg: EvalArg | undefined,
	timelineArg: EvalArg | undefined,
	seasonalityParam: number,
	dataCompletionParam = 1,
	aggregationParam = 1,
): [ETSModel, number[], number[]] | CellValue {
	const paired = collectPairedETS(valuesArg, timelineArg)
	if (!Array.isArray(paired)) return paired
	const prepared = prepareSeries(
		paired[0],
		paired[1],
		Math.trunc(dataCompletionParam) === 0 ? 0 : 1,
		Math.trunc(aggregationParam) || 1,
	)
	if ('kind' in prepared) return prepared
	const { values, timeline, stepSize } = prepared
	if (values.length < 3) return errorValue('#VALUE!')
	if (seasonalityParam !== Math.trunc(seasonalityParam)) return errorValue('#NUM!')
	let period = Math.trunc(seasonalityParam)
	if (period === 1) period = detectPeriod(values)
	else if (period === 0) period = 1
	if (period < 0 || period > 8760) return errorValue('#NUM!')
	if (period > 1 && values.length < 2 * period) return errorValue('#NUM!')
	const model = fitETS(values, period)
	model.stepSize = stepSize
	return [model, values, timeline]
}

function normSInv(p: number): number {
	const y = p - 0.5
	if (Math.abs(y) < 0.42) {
		const r = y * y
		return (
			(y * (((-25.44106049637 * r + 41.39119773534) * r + -18.61500062529) * r + 2.50662823884)) /
			((((3.13082909833 * r + -21.06224101826) * r + 23.08336743743) * r + -8.4735109309) * r + 1)
		)
	}
	let r = y > 0 ? 1 - p : p
	r = Math.log(-Math.log(r))
	let z =
		0.3374754822726147 +
		r *
			(0.9761690190917186 +
				r *
					(0.1607979714918209 +
						r *
							(0.0276438810333863 +
								r *
									(0.0038405729373609 +
										r *
											(0.0003951896511919 +
												r *
													(0.0000321767881768 +
														r * (0.0000002888167364 + r * 0.0000003960315187)))))))
	if (y < 0) z = -z
	return z
}

function etsForecast(model: ETSModel, timeline: number[], target: number): number {
	const lastT = at(timeline, timeline.length - 1)
	const h = model.stepSize !== 0 ? (target - lastT) / model.stepSize : 1
	if (model.period > 1) {
		const n = timeline.length
		const hR = Math.round(h)
		const sIdx = (((n - 1 + hR) % model.period) + model.period) % model.period
		return (model.level + h * model.trend) * (at(model.seasonal, sIdx) || 1)
	}
	return model.level + h * model.trend
}

function forecastEtsFn(args: EvalArg[]): CellValue {
	const target = numArg(args[0])
	if (typeof target !== 'number') return target
	const seasonality = args.length > 3 ? numArg(args[3]) : 1
	if (typeof seasonality !== 'number') return seasonality
	const dataCompletion = args.length > 4 ? numArg(args[4]) : 1
	if (typeof dataCompletion !== 'number') return dataCompletion
	const aggregation = args.length > 5 ? numArg(args[5]) : 1
	if (typeof aggregation !== 'number') return aggregation
	const result = buildETSModel(args[1], args[2], seasonality, dataCompletion, aggregation)
	if (!Array.isArray(result)) return result
	const [model, , timeline] = result
	if (target <= at(timeline, timeline.length - 1)) return errorValue('#NUM!')
	if (model.stepSize !== 0) {
		const step = (target - at(timeline, timeline.length - 1)) / model.stepSize
		if (Math.abs(step - Math.round(step)) > 1e-9) return errorValue('#NUM!')
	}
	return numberValue(etsForecast(model, timeline, target))
}

function forecastEtsSeasonalityFn(args: EvalArg[]): CellValue {
	const paired = collectPairedETS(args[0], args[1])
	if (!Array.isArray(paired)) return paired
	const dataCompletion = args.length > 2 ? numArg(args[2]) : 1
	if (typeof dataCompletion !== 'number') return dataCompletion
	const aggregation = args.length > 3 ? numArg(args[3]) : 1
	if (typeof aggregation !== 'number') return aggregation
	const prepared = prepareSeries(
		paired[0],
		paired[1],
		Math.trunc(dataCompletion) === 0 ? 0 : 1,
		Math.trunc(aggregation) || 1,
	)
	if ('kind' in prepared) return prepared
	const { values } = prepared
	return numberValue(detectPeriod(values))
}

function forecastEtsConfintFn(args: EvalArg[]): CellValue {
	const target = numArg(args[0])
	if (typeof target !== 'number') return target
	const confidence = args.length > 3 ? numArg(args[3]) : 0.95
	if (typeof confidence !== 'number') return confidence
	if (confidence <= 0 || confidence >= 1) return errorValue('#NUM!')
	const seasonality = args.length > 4 ? numArg(args[4]) : 1
	if (typeof seasonality !== 'number') return seasonality
	const dataCompletion = args.length > 5 ? numArg(args[5]) : 1
	if (typeof dataCompletion !== 'number') return dataCompletion
	const aggregation = args.length > 6 ? numArg(args[6]) : 1
	if (typeof aggregation !== 'number') return aggregation
	const result = buildETSModel(args[1], args[2], seasonality, dataCompletion, aggregation)
	if (!Array.isArray(result)) return result
	const [model, , timeline] = result
	if (target <= at(timeline, timeline.length - 1)) return errorValue('#NUM!')
	if (model.stepSize !== 0) {
		const step = (target - at(timeline, timeline.length - 1)) / model.stepSize
		if (Math.abs(step - Math.round(step)) > 1e-9) return errorValue('#NUM!')
	}
	const lastT = at(timeline, timeline.length - 1)
	const h = model.stepSize !== 0 ? Math.max(1, Math.abs((target - lastT) / model.stepSize)) : 1
	const rmse = Math.sqrt(model.mse)
	const z = normSInv(1 - (1 - confidence) / 2)
	return numberValue(z * rmse * Math.sqrt(h))
}

function forecastEtsStatFn(args: EvalArg[]): CellValue {
	const stArg = numArg(args[2])
	if (typeof stArg !== 'number') return stArg
	const st = Math.trunc(stArg)
	if (st < 1 || st > 8) return errorValue('#NUM!')
	const seasonality = args.length > 3 ? numArg(args[3]) : 1
	if (typeof seasonality !== 'number') return seasonality
	const dataCompletion = args.length > 4 ? numArg(args[4]) : 1
	if (typeof dataCompletion !== 'number') return dataCompletion
	const aggregation = args.length > 5 ? numArg(args[5]) : 1
	if (typeof aggregation !== 'number') return aggregation
	const result = buildETSModel(args[0], args[1], seasonality, dataCompletion, aggregation)
	if (!Array.isArray(result)) return result
	const [model, values] = result
	const res = model.residuals.slice(1)
	const n = res.length
	if (n === 0) return errorValue('#NUM!')
	switch (st) {
		case 1:
			return numberValue(model.alpha)
		case 2:
			return numberValue(model.beta)
		case 3:
			return numberValue(model.gamma)
		case 4: {
			let naiveSum = 0
			for (let i = 1; i < values.length; i++)
				naiveSum += Math.abs(at(values, i) - at(values, i - 1))
			const naiveMae = values.length > 1 ? naiveSum / (values.length - 1) : 1
			let maeSum = 0
			for (const r of res) maeSum += Math.abs(r)
			return numberValue(naiveMae !== 0 ? maeSum / n / naiveMae : 0)
		}
		case 5: {
			let sum = 0
			for (let i = 1; i < values.length && i < model.residuals.length; i++) {
				const actual = at(values, i)
				const forecast = actual - at(model.residuals, i)
				const denom = (Math.abs(actual) + Math.abs(forecast)) / 2
				if (denom !== 0) sum += Math.abs(at(model.residuals, i)) / denom
			}
			return numberValue((sum / n) * 100)
		}
		case 6: {
			let sum = 0
			for (const r of res) sum += Math.abs(r)
			return numberValue(sum / n)
		}
		case 7:
			return numberValue(Math.sqrt(model.mse))
		case 8:
			return numberValue(model.stepSize)
		default:
			return errorValue('#NUM!')
	}
}

export const etsFunctions: FunctionDef[] = [
	{ name: 'FORECAST.ETS', minArgs: 3, maxArgs: 6, evaluate: forecastEtsFn },
	{
		name: 'FORECAST.ETS.SEASONALITY',
		minArgs: 2,
		maxArgs: 4,
		evaluate: forecastEtsSeasonalityFn,
	},
	{ name: 'FORECAST.ETS.CONFINT', minArgs: 3, maxArgs: 7, evaluate: forecastEtsConfintFn },
	{ name: 'FORECAST.ETS.STAT', minArgs: 3, maxArgs: 6, evaluate: forecastEtsStatFn },
]
