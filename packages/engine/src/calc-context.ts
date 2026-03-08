export interface CalcContext {
	readonly now: Date
	readonly today: Date
	readonly randomSeed: number
	readonly locale: string
	readonly dateSystem: '1900' | '1904'
	readonly iterativeCalc: {
		readonly enabled: boolean
		readonly maxIterations: number
		readonly maxChange: number
	}
}

export function defaultCalcContext(): CalcContext {
	const now = new Date()
	const today = new Date(now.getFullYear(), now.getMonth(), now.getDate())
	return {
		now,
		today,
		randomSeed: 42,
		locale: 'en-US',
		dateSystem: '1900',
		iterativeCalc: {
			enabled: false,
			maxIterations: 100,
			maxChange: 0.001,
		},
	}
}
