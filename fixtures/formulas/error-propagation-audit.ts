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

const INTENTIONAL_NON_PROPAGATORS = new Map<string, number>([
	['BYCOL', 1],
	['BYROW', 1],
	['ERROR.TYPE', 1],
	['IF', 1],
	['IFERROR', 2],
	['IFNA', 1],
	['ISBLANK', 1],
	['ISERR', 1],
	['ISERROR', 1],
	['ISLOGICAL', 1],
	['ISNA', 1],
	['ISNONTEXT', 1],
	['ISNUMBER', 1],
	['ISREF', 1],
	['ISTEXT', 1],
	['LET', 2],
	['MAP', 1],
	['REDUCE', 2],
	['SCAN', 2],
	['TYPE', 1],
])

function main(): void {
	const nonPropagating: Array<{
		name: string
		position: number
		minArgs: number
		maxArgs: number
	}> = []
	for (const fn of [...functionRegistry.values()].sort((a, b) => a.name.localeCompare(b.name))) {
		if (fn.minArgs <= 0) continue
		for (let pos = 0; pos < fn.minArgs; pos++) {
			const args: EvalArg[] = Array.from({ length: fn.minArgs }, () => ({ value: EMPTY }))
			args[pos] = { value: errorValue('#REF!') }
			const result = fn.evaluate(args, ctx)
			if (result.kind !== 'error') {
				nonPropagating.push({
					name: fn.name,
					position: pos,
					minArgs: fn.minArgs,
					maxArgs: fn.maxArgs,
				})
			}
		}
	}

	console.log('Formula Error Propagation Audit')
	console.log('='.repeat(72))
	console.log(`registered functions: ${functionRegistry.size}`)
	console.log(`non-propagating arg slots: ${nonPropagating.length}`)
	const grouped = new Map<string, number>()
	for (const item of nonPropagating) grouped.set(item.name, (grouped.get(item.name) ?? 0) + 1)
	const top = [...grouped.entries()].sort((a, b) => b[1] - a[1]).slice(0, 60)
	for (const [name, count] of top) console.log(`${name}: ${count}`)

	let unexpected = 0
	for (const [name, count] of grouped) {
		const allowed = INTENTIONAL_NON_PROPAGATORS.get(name) ?? 0
		if (count > allowed) unexpected += count - allowed
	}
	console.log('-'.repeat(72))
	console.log(`intentional exceptions tracked: ${INTENTIONAL_NON_PROPAGATORS.size}`)
	console.log(`unexpected non-propagating slots: ${unexpected}`)
	if (unexpected > 0) process.exit(1)
}

main()
