import type { CellValue } from '@ascend/schema'

export interface StatusSummary {
	readonly count: number
	readonly numericCount: number
	readonly sum: number
	readonly min: number | null
	readonly max: number | null
	readonly average: number | null
}

export function summarizeValues(values: readonly CellValue[]): StatusSummary {
	let count = 0
	let numericCount = 0
	let sum = 0
	let min: number | null = null
	let max: number | null = null
	for (const value of values) {
		if (value.kind === 'empty') continue
		count++
		if (value.kind === 'number' || value.kind === 'date') {
			const numericValue = value.kind === 'number' ? value.value : value.serial
			numericCount++
			sum += numericValue
			min = min === null ? numericValue : Math.min(min, numericValue)
			max = max === null ? numericValue : Math.max(max, numericValue)
		}
	}
	return {
		count,
		numericCount,
		sum,
		min,
		max,
		average: numericCount > 0 ? sum / numericCount : null,
	}
}
