import { EMPTY, errorValue } from '@ascend/schema'
import {
	type EvalArg,
	type FunctionEvalContext,
	functionRegistry,
} from '../../packages/formulas/src/index.ts'

const ctx: FunctionEvalContext = {
	now: new Date('2026-01-01T00:00:00Z'),
	today: new Date('2026-01-01T00:00:00Z'),
	randomSeed: 1,
	locale: 'en-US',
	dateSystem: '1900',
	exactLookupCache: undefined,
	lookupVectorCache: undefined,
	aggregateRangeCache: undefined,
}

const ERROR_CODES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'] as const

const ERROR_INSPECTORS = new Set([
	'ERROR.TYPE',
	'ISBLANK',
	'ISERR',
	'ISERROR',
	'ISFORMULA',
	'ISLOGICAL',
	'ISNA',
	'ISNONTEXT',
	'ISNUMBER',
	'ISREF',
	'ISTEXT',
	'TYPE',
])

function isIntentionalNonPropagation(name: string, position: number, code: string): boolean {
	if (ERROR_INSPECTORS.has(name)) return true
	if (name === 'IF') return position === 1
	if (name === 'IFS') return position % 2 === 1
	if (name === 'IFERROR') return position === 0 || position === 1
	if (name === 'IFNA') return (position === 0 && code === '#N/A') || position === 1
	if (name === 'COUNT' || name === 'COUNTA') return true
	if (name === 'INDIRECT' || name === 'OFFSET') return true
	if (name === 'LET') return position < 2
	if (name === 'MAP' || name === 'BYROW' || name === 'BYCOL') return position === 1
	if (name === 'REDUCE' || name === 'SCAN') return position === 1 || position === 2
	return false
}

export function runErrorPropagationAudit(): { unexpected: number; registered: number } {
	const nonPropagating: Array<{
		name: string
		position: number
		code: string
		result: string
		minArgs: number
		maxArgs: number
	}> = []
	let exactSlots = 0
	let maskedSlots = 0
	for (const fn of [...functionRegistry.values()].sort((a, b) => a.name.localeCompare(b.name))) {
		if (fn.minArgs <= 0) continue
		for (let pos = 0; pos < fn.minArgs; pos++) {
			const refArgs: EvalArg[] = Array.from({ length: fn.minArgs }, () => ({ value: EMPTY }))
			refArgs[pos] = { value: errorValue('#REF!') }
			const refResult = fn.evaluate(refArgs, ctx)
			if (refResult.kind !== 'error') {
				nonPropagating.push({
					name: fn.name,
					position: pos,
					code: '#REF!',
					result: refResult.kind,
					minArgs: fn.minArgs,
					maxArgs: fn.maxArgs,
				})
				continue
			}
			if (refResult.value !== '#REF!') {
				maskedSlots++
				continue
			}
			exactSlots++
			for (const code of ERROR_CODES) {
				const args: EvalArg[] = Array.from({ length: fn.minArgs }, () => ({ value: EMPTY }))
				args[pos] = { value: errorValue(code) }
				const result = fn.evaluate(args, ctx)
				if (result.kind !== 'error' || result.value !== code) {
					nonPropagating.push({
						name: fn.name,
						position: pos,
						code,
						result: result.kind === 'error' ? result.value : result.kind,
						minArgs: fn.minArgs,
						maxArgs: fn.maxArgs,
					})
				}
			}
		}
	}

	console.log('Formula Error Propagation Audit')
	console.log('='.repeat(72))
	console.log(`registered functions: ${String(functionRegistry.size)}`)
	console.log(`exactly audited arg slots: ${String(exactSlots)}`)
	console.log(`masked arg slots skipped: ${String(maskedSlots)}`)
	console.log(`non-propagating arg/error cases: ${String(nonPropagating.length)}`)
	const grouped = new Map<string, number>()
	for (const item of nonPropagating) grouped.set(item.name, (grouped.get(item.name) ?? 0) + 1)
	const top = [...grouped.entries()].sort((a, b) => b[1] - a[1]).slice(0, 60)
	for (const [name, count] of top) console.log(`${name}: ${count}`)

	let unexpected = 0
	for (const item of nonPropagating) {
		if (!isIntentionalNonPropagation(item.name, item.position, item.code)) {
			if (unexpected < 20) {
				console.log(
					`unexpected ${item.name} arg ${String(item.position + 1)} ${item.code} -> ${item.result}`,
				)
			}
			unexpected++
		}
	}
	console.log('-'.repeat(72))
	console.log(`error codes checked: ${String(ERROR_CODES.length)}`)
	console.log(`unexpected non-propagating arg/error cases: ${String(unexpected)}`)
	return { unexpected, registered: functionRegistry.size }
}

function main(): void {
	const { unexpected } = runErrorPropagationAudit()
	if (unexpected > 0) process.exit(1)
}

if (import.meta.main) {
	main()
}
