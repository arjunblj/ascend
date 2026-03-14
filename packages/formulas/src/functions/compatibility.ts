import { booleanValue, errorValue } from '@ascend/schema'
import type { EvalArg, FunctionDef, FunctionEvalContext } from './registry.ts'
import { functionRegistry } from './registry.ts'

function aliasFunction(
	name: string,
	target: string,
	minArgs: number,
	maxArgs: number,
	mapArgs?: (args: EvalArg[]) => EvalArg[],
): FunctionDef {
	return {
		name,
		minArgs,
		maxArgs,
		evaluate(args: EvalArg[], ctx?: FunctionEvalContext) {
			const targetDef = functionRegistry.get(target)
			if (!targetDef) return errorValue('#NAME?')
			return targetDef.evaluate(mapArgs ? mapArgs(args) : args, ctx)
		},
	}
}

function appendBooleanArg(value: boolean) {
	return (args: EvalArg[]): EvalArg[] => [...args, { value: booleanValue(value) }]
}

function tdistCompat(args: EvalArg[], ctx?: FunctionEvalContext) {
	const tailsArg = args[2]?.value
	if (!tailsArg || tailsArg.kind === 'error') return tailsArg ?? errorValue('#VALUE!')
	const tails = tailsArg.kind === 'number' ? tailsArg.value : Number.NaN
	if (tails === 1) {
		const targetDef = functionRegistry.get('T.DIST.RT')
		return targetDef?.evaluate(args.slice(0, 2), ctx) ?? errorValue('#NAME?')
	}
	if (tails === 2) {
		const targetDef = functionRegistry.get('T.DIST.2T')
		return targetDef?.evaluate(args.slice(0, 2), ctx) ?? errorValue('#NAME?')
	}
	return errorValue('#NUM!')
}

export const compatibilityFunctions: FunctionDef[] = [
	aliasFunction('BETADIST', 'BETA.DIST', 4, 6),
	aliasFunction('BETAINV', 'BETA.INV', 3, 5),
	aliasFunction('BINOMDIST', 'BINOM.DIST', 4, 4),
	aliasFunction('CHIDIST', 'CHISQ.DIST.RT', 2, 2),
	aliasFunction('CHIINV', 'CHISQ.INV.RT', 2, 2),
	aliasFunction('CONFIDENCE', 'CONFIDENCE.NORM', 3, 3),
	aliasFunction('COVAR', 'COVARIANCE.P', 2, 2),
	aliasFunction('CRITBINOM', 'BINOM.INV', 3, 3),
	aliasFunction('EXPONDIST', 'EXPON.DIST', 3, 3),
	aliasFunction('FDIST', 'F.DIST.RT', 3, 3),
	aliasFunction('FINV', 'F.INV.RT', 3, 3),
	aliasFunction('GAMMADIST', 'GAMMA.DIST', 4, 4),
	aliasFunction('GAMMAINV', 'GAMMA.INV', 3, 3),
	aliasFunction('HYPGEOMDIST', 'HYPGEOM.DIST', 4, 4, appendBooleanArg(false)),
	aliasFunction('LOGINV', 'LOGNORM.INV', 3, 3),
	aliasFunction('LOGNORMDIST', 'LOGNORM.DIST', 3, 3, appendBooleanArg(true)),
	aliasFunction('NEGBINOMDIST', 'NEGBINOM.DIST', 3, 3, appendBooleanArg(false)),
	aliasFunction('NORMDIST', 'NORM.DIST', 4, 4),
	aliasFunction('NORMINV', 'NORM.INV', 3, 3),
	aliasFunction('NORMSDIST', 'NORM.S.DIST', 1, 1, appendBooleanArg(true)),
	aliasFunction('NORMSINV', 'NORM.S.INV', 1, 1),
	aliasFunction('PERCENTRANK', 'PERCENTRANK.INC', 2, 3),
	aliasFunction('POISSON', 'POISSON.DIST', 3, 3),
	{
		name: 'TDIST',
		minArgs: 3,
		maxArgs: 3,
		evaluate: tdistCompat,
	},
	aliasFunction('TINV', 'T.INV.2T', 2, 2),
	aliasFunction('WEIBULL', 'WEIBULL.DIST', 4, 4),
]
