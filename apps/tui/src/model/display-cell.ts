import type { CellValue } from '@ascend/schema'
import { formatDisplayCellValue } from '@ascend/sdk'

export function displayCellValue(value: CellValue): string {
	return formatDisplayCellValue(value)
}

export function parseInputValue(input: string): string | number | boolean {
	const trimmed = input.trim()
	if (trimmed.toUpperCase() === 'TRUE') return true
	if (trimmed.toUpperCase() === 'FALSE') return false
	const numberValue = Number(trimmed)
	if (trimmed !== '' && !Number.isNaN(numberValue)) return numberValue
	return input
}
