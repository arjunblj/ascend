import { createWorkbook, parseA1, type StyleId } from '../../packages/core/src/index.ts'
import type { CalcContext } from '../../packages/engine/src/index.ts'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import type { CellValue, ExcelError } from '../../packages/schema/src/index.ts'
import {
	booleanValue,
	EMPTY,
	errorValue,
	numberValue,
	stringValue,
	topLeftScalar,
} from '../../packages/schema/src/index.ts'

const SID = 0 as StyleId

export interface FormulaLogicConformanceCase {
	readonly description: string
	readonly setup: Record<string, string | number | boolean>
	readonly setupFormulas?: Record<string, string>
	readonly formula: string
	readonly context?: FormulaLogicContext
	readonly expected: FormulaLogicExpected
}

export interface FormulaLogicConformanceFixture {
	readonly function: string
	readonly cases: readonly FormulaLogicConformanceCase[]
}

export interface FormulaLogicContext {
	readonly dateSystem?: '1900' | '1904'
	readonly now?: string
	readonly today?: string
	readonly randomSeed?: number
	readonly locale?: string
}

export interface FormulaLogicExpected {
	readonly kind: string
	readonly value?: number | string | boolean
	readonly serial?: number
	readonly approx?: number
	readonly tolerance?: number
}

export interface FormulaLogicRecord {
	readonly label: string
	readonly description?: string
	readonly setup: Record<string, string | number | boolean>
	readonly setupFormulas: Record<string, string>
	readonly formula: string
	readonly context: FormulaLogicContext
	readonly expected: FormulaLogicExpected
}

export interface FormulaLogicResult {
	readonly label: string
	readonly pass: boolean
	readonly actual?: CellValue
	readonly expected: FormulaLogicExpected
	readonly error?: string
}

export function emitFormulaLogicTest(
	fixture: FormulaLogicConformanceFixture,
	options: { readonly source?: string } = {},
): string {
	const lines: string[] = [
		`# formula-logictest: ${fixture.function}`,
		...(options.source ? [`# source: ${options.source}`] : []),
		'',
	]
	for (const entry of fixture.cases) {
		const label = slug(`${fixture.function}-${entry.description}`)
		lines.push(`# ${entry.description}`)
		for (const [ref, value] of Object.entries(entry.setup)) {
			lines.push(`setup Sheet1!${ref} ${formatScalar(value)}`)
		}
		for (const [ref, formula] of Object.entries(entry.setupFormulas ?? {})) {
			lines.push(`setup-formula Sheet1!${ref} ${formula}`)
		}
		for (const [key, value] of Object.entries(entry.context ?? {})) {
			if (value !== undefined) lines.push(`context ${key} ${value}`)
		}
		lines.push(`query value label=${label}`)
		lines.push(entry.formula)
		lines.push('----')
		lines.push(formatExpected(entry.expected))
		lines.push('')
	}
	return lines.join('\n')
}

export function parseFormulaLogicTest(text: string): FormulaLogicRecord[] {
	const records = splitRecords(text)
	return records.map(parseFormulaLogicRecord)
}

export function runFormulaLogicTest(text: string): FormulaLogicResult[] {
	return parseFormulaLogicTest(text).map((record) => runFormulaLogicRecord(record))
}

export function runFormulaLogicRecord(
	record: FormulaLogicRecord,
	baseCtx: CalcContext = defaultCalcContext(),
): FormulaLogicResult {
	const wb = createWorkbook()
	const sheet = wb.addSheet('Sheet1')
	for (const [ref, value] of Object.entries(record.setup)) {
		const { row, col } = parseA1(ref)
		sheet.cells.set(row, col, {
			value: inputToCellValue(value),
			formula: null,
			styleId: SID,
		})
	}
	for (const [ref, formula] of Object.entries(record.setupFormulas)) {
		const { row, col } = parseA1(ref)
		sheet.cells.set(row, col, { value: EMPTY, formula, styleId: SID })
	}
	const formulaRow = 10
	const formulaCol = 0
	sheet.cells.set(formulaRow, formulaCol, {
		value: EMPTY,
		formula: record.formula.startsWith('=') ? record.formula.slice(1) : record.formula,
		styleId: SID,
	})
	const ctx: CalcContext = {
		...baseCtx,
		...(record.context.dateSystem ? { dateSystem: record.context.dateSystem } : {}),
		...(record.context.randomSeed !== undefined ? { randomSeed: record.context.randomSeed } : {}),
		...(record.context.locale ? { locale: record.context.locale } : {}),
		...(record.context.now ? { now: new Date(record.context.now) } : {}),
		...(record.context.today ? { today: new Date(record.context.today) } : {}),
	}
	wb.calcSettings = {
		...wb.calcSettings,
		dateSystem: ctx.dateSystem,
		iterativeCalc: ctx.iterativeCalc,
	}
	const result = recalculate(wb, ctx)
	if (result.errors.length > 0) {
		return {
			label: record.label,
			pass: false,
			expected: record.expected,
			error: result.errors[0]?.error.message ?? 'formula evaluation failed',
		}
	}
	const cell = sheet.cells.get(formulaRow, formulaCol)
	const actual = topLeftScalar(cell?.value ?? EMPTY)
	return {
		label: record.label,
		pass: expectedMatches(actual, record.expected),
		actual,
		expected: record.expected,
	}
}

function splitRecords(text: string): string[] {
	const records: string[] = []
	let current: string[] = []
	for (const rawLine of text.split(/\r?\n/)) {
		const line = rawLine.trimEnd()
		if (line.trim() === '') {
			if (current.some((entry) => entry.trim() && !entry.trimStart().startsWith('#'))) {
				records.push(current.join('\n'))
			}
			current = []
			continue
		}
		current.push(line)
	}
	if (current.some((entry) => entry.trim() && !entry.trimStart().startsWith('#'))) {
		records.push(current.join('\n'))
	}
	return records
}

function parseFormulaLogicRecord(text: string): FormulaLogicRecord {
	const setup: Record<string, string | number | boolean> = {}
	const setupFormulas: Record<string, string> = {}
	const context: Record<string, string> = {}
	let label = ''
	let description: string | undefined
	let formula: string | undefined
	let expected: FormulaLogicExpected | undefined
	const lines = text
		.split(/\r?\n/)
		.map((line) => line.trim())
		.filter((line) => line.length > 0)
	for (let index = 0; index < lines.length; index++) {
		const line = lines[index]
		if (!line) continue
		if (line.startsWith('#')) {
			description = description ?? line.replace(/^#\s*/, '')
			continue
		}
		if (line.startsWith('setup-formula ')) {
			const [, ref, formulaText] = /^setup-formula\s+Sheet1!(\S+)\s+(.+)$/.exec(line) ?? []
			if (!ref || !formulaText) throw new Error(`Invalid setup-formula line: ${line}`)
			setupFormulas[ref] = formulaText
			continue
		}
		if (line.startsWith('setup ')) {
			const [, ref, scalar] = /^setup\s+Sheet1!(\S+)\s+(.+)$/.exec(line) ?? []
			if (!ref || !scalar) throw new Error(`Invalid setup line: ${line}`)
			setup[ref] = parseScalar(scalar)
			continue
		}
		if (line.startsWith('context ')) {
			const [, key, value] = /^context\s+(\S+)\s+(.+)$/.exec(line) ?? []
			if (!key || value === undefined) throw new Error(`Invalid context line: ${line}`)
			context[key] = value
			continue
		}
		if (line.startsWith('query value ')) {
			label = /(?:^|\s)label=([^\s]+)/.exec(line)?.[1] ?? ''
			formula = lines[index + 1]
			if (lines[index + 2] !== '----') throw new Error(`Missing result separator for ${label}`)
			expected = parseExpected(lines[index + 3] ?? '')
			index += 3
			continue
		}
		throw new Error(`Unknown formula-logictest line: ${line}`)
	}
	if (!label) throw new Error('Formula-logictest record is missing label')
	if (!formula) throw new Error(`Formula-logictest record ${label} is missing formula`)
	if (!expected) throw new Error(`Formula-logictest record ${label} is missing expected result`)
	return {
		label,
		...(description ? { description } : {}),
		setup,
		setupFormulas,
		formula,
		context: parseContext(context),
		expected,
	}
}

function formatScalar(value: string | number | boolean): string {
	if (typeof value === 'number') return `number ${value}`
	if (typeof value === 'boolean') return `boolean ${value ? 'true' : 'false'}`
	return `string ${JSON.stringify(value)}`
}

function parseScalar(text: string): string | number | boolean {
	if (text.startsWith('number ')) return Number(text.slice('number '.length))
	if (text.startsWith('boolean ')) return text.slice('boolean '.length).trim() === 'true'
	if (text.startsWith('string ')) return JSON.parse(text.slice('string '.length)) as string
	throw new Error(`Unsupported scalar: ${text}`)
}

function formatExpected(expected: FormulaLogicExpected): string {
	if (expected.approx !== undefined || expected.tolerance !== undefined) {
		return `${expected.kind} approx ${expected.approx} tolerance ${expected.tolerance}`
	}
	if (expected.kind === 'empty') return 'empty'
	if (expected.kind === 'date') return `date ${expected.serial ?? expected.value}`
	return `${expected.kind} ${JSON.stringify(expected.value)}`
}

function parseExpected(line: string): FormulaLogicExpected {
	if (line === 'empty') return { kind: 'empty' }
	const [kind, ...rest] = line.split(/\s+/)
	if (!kind) throw new Error('Expected result is empty')
	if (rest[0] === 'approx') {
		return { kind, approx: Number(rest[1]), tolerance: Number(rest[3]) }
	}
	if (kind === 'date') return { kind, serial: Number(rest[0]) }
	const raw = rest.join(' ')
	return { kind, value: raw ? (JSON.parse(raw) as string | number | boolean) : undefined }
}

function parseContext(values: Record<string, string>): FormulaLogicContext {
	return {
		...(values.dateSystem === '1900' || values.dateSystem === '1904'
			? { dateSystem: values.dateSystem }
			: {}),
		...(values.now ? { now: values.now } : {}),
		...(values.today ? { today: values.today } : {}),
		...(values.randomSeed !== undefined ? { randomSeed: Number(values.randomSeed) } : {}),
		...(values.locale ? { locale: values.locale } : {}),
	}
}

function inputToCellValue(value: string | number | boolean): CellValue {
	if (typeof value === 'number') return numberValue(value)
	if (typeof value === 'boolean') return booleanValue(value)
	return stringValue(value)
}

function expectedMatches(actual: CellValue, expected: FormulaLogicExpected): boolean {
	if (expected.approx !== undefined && expected.tolerance !== undefined) {
		return (
			actual.kind === 'number' && Math.abs(actual.value - expected.approx) <= expected.tolerance
		)
	}
	const expectedCell = expectedToCellValue(expected)
	if (actual.kind !== expectedCell.kind) return false
	switch (actual.kind) {
		case 'empty':
			return true
		case 'number':
			return expectedCell.kind === 'number' && actual.value === expectedCell.value
		case 'string':
			return expectedCell.kind === 'string' && actual.value === expectedCell.value
		case 'boolean':
			return expectedCell.kind === 'boolean' && actual.value === expectedCell.value
		case 'error':
			return expectedCell.kind === 'error' && actual.value === expectedCell.value
		case 'date':
			return expectedCell.kind === 'date' && actual.serial === expectedCell.serial
		case 'richText':
		case 'array':
			return false
	}
}

function expectedToCellValue(expected: FormulaLogicExpected): CellValue {
	switch (expected.kind) {
		case 'number':
			return numberValue(expected.value as number)
		case 'string':
			return stringValue(expected.value as string)
		case 'boolean':
			return booleanValue(expected.value as boolean)
		case 'error':
			return errorValue(expected.value as ExcelError)
		case 'empty':
			return EMPTY
		case 'date':
			return { kind: 'date', serial: expected.serial ?? (expected.value as number) }
		default:
			throw new Error(`Unknown expected kind: ${expected.kind}`)
	}
}

function slug(text: string): string {
	return text
		.toLowerCase()
		.replace(/[^a-z0-9]+/g, '-')
		.replace(/^-+|-+$/g, '')
		.slice(0, 80)
}
