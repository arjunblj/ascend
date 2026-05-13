import { cachedParseFormula, normalizeFormulaInput } from '@ascend/formulas'
import type { InputValue, Operation } from '@ascend/schema'

export interface CellInputParseOptions {
	readonly parseDates?: boolean
}

export type ParsedCellInput =
	| {
			readonly kind: 'blank'
			readonly value: null
			readonly reason: 'blank'
	  }
	| {
			readonly kind: 'value'
			readonly value: Exclude<InputValue, null>
			readonly valueKind: 'string' | 'number' | 'boolean' | 'date'
			readonly reason: 'escaped-text' | 'text' | 'number' | 'percent' | 'boolean' | 'date-iso'
	  }
	| {
			readonly kind: 'formula'
			readonly formula: string
			readonly parseOk: boolean
			readonly parseError?: string
			readonly reason: 'formula'
	  }

export function parseCellInput(text: string, options: CellInputParseOptions = {}): ParsedCellInput {
	if (text.length === 0) return { kind: 'blank', value: null, reason: 'blank' }
	if (text.startsWith("'")) {
		return { kind: 'value', value: text.slice(1), valueKind: 'string', reason: 'escaped-text' }
	}
	if (text.startsWith('=')) {
		const formula = normalizeFormulaInput(text)
		const parsed = cachedParseFormula(formula)
		return parsed.ok
			? { kind: 'formula', formula, parseOk: true, reason: 'formula' }
			: {
					kind: 'formula',
					formula,
					parseOk: false,
					parseError: parsed.error.message,
					reason: 'formula',
				}
	}

	const trimmed = text.trim()
	if (trimmed.length === 0)
		return { kind: 'value', value: text, valueKind: 'string', reason: 'text' }
	const lower = trimmed.toLowerCase()
	if (lower === 'true')
		return { kind: 'value', value: true, valueKind: 'boolean', reason: 'boolean' }
	if (lower === 'false')
		return { kind: 'value', value: false, valueKind: 'boolean', reason: 'boolean' }

	const percent = parsePercentInput(trimmed)
	if (percent !== undefined) {
		return { kind: 'value', value: percent, valueKind: 'number', reason: 'percent' }
	}
	const number = parseNumberInput(trimmed)
	if (number !== undefined) {
		return { kind: 'value', value: number, valueKind: 'number', reason: 'number' }
	}
	if (options.parseDates) {
		const date = parseIsoDateInput(trimmed)
		if (date) return { kind: 'value', value: date, valueKind: 'date', reason: 'date-iso' }
	}
	return { kind: 'value', value: text, valueKind: 'string', reason: 'text' }
}

export function parseCellInputOperation(sheet: string, ref: string, text: string): Operation {
	const parsed = parseCellInput(text)
	if (parsed.kind === 'formula') return { op: 'setFormula', sheet, ref, formula: parsed.formula }
	return {
		op: 'setCells',
		sheet,
		updates: [{ ref, value: parsed.kind === 'blank' ? null : parsed.value }],
	}
}

function parseNumberInput(text: string): number | undefined {
	if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) return undefined
	const value = Number(text)
	return Number.isFinite(value) ? value : undefined
}

function parsePercentInput(text: string): number | undefined {
	if (!text.endsWith('%')) return undefined
	const number = parseNumberInput(text.slice(0, -1).trim())
	return number === undefined ? undefined : number / 100
}

function parseIsoDateInput(text: string): Date | undefined {
	const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text)
	if (!match) return undefined
	const year = Number(match[1])
	const month = Number(match[2])
	const day = Number(match[3])
	const date = new Date(year, month - 1, day)
	if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
		return undefined
	}
	return date
}
