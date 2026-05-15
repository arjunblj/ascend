#!/usr/bin/env bun
import { existsSync, readFileSync } from 'node:fs'
import { relative } from 'node:path'
import { performance } from 'node:perf_hooks'
import { indexToColumn } from '../../packages/core/src/index.ts'
import { readXlsx } from '../../packages/io-xlsx/src/index.ts'
import {
	type FormulaAssistResult,
	type FormulaPrepareRenameBlockReason,
	formulaAssist,
} from '../../packages/sdk/src/index.ts'

export interface FormulaAssistProofOptions {
	readonly repeat?: number
	readonly warmup?: number
	readonly publicFormulaLimit?: number
	readonly includeTimings?: boolean
	readonly workbookPaths?: readonly string[]
}

export interface FormulaAssistProofFormula {
	readonly source: string
	readonly sheet: string
	readonly ref: string
	readonly formula: string
}

export interface FormulaAssistProofEdgeCase {
	readonly name: string
	readonly formula: string
	readonly cursor: number
	readonly expected: 'ok' | FormulaPrepareRenameBlockReason
	readonly observed: 'ok' | FormulaPrepareRenameBlockReason | 'none'
	readonly role?: string
	readonly referenceKind?: string
	readonly occurrenceCount: number
	readonly passed: boolean
}

export interface FormulaAssistProofResult {
	readonly generatedAt: string
	readonly publicFormulaCount: number
	readonly sampledFormulaCount: number
	readonly staticEdgeCaseCount: number
	readonly parseOkCount: number
	readonly diagnosticFormulaCount: number
	readonly referenceCount: number
	readonly bindingCount: number
	readonly longestFormulaChars: number
	readonly renameOkCount: number
	readonly renameRefusalCounts: Record<FormulaPrepareRenameBlockReason, number>
	readonly edgeCases: readonly FormulaAssistProofEdgeCase[]
	readonly timings: {
		readonly repeat: number
		readonly warmup: number
		readonly sampleCount: number
		readonly medianMs: number
		readonly p95Ms: number
		readonly maxMs: number
	} | null
	readonly publicSources: readonly string[]
	readonly passed: boolean
	readonly boundary: string
}

const DEFAULT_PUBLIC_WORKBOOKS = [
	'fixtures/xlsx/poi/formula_stress_test.xlsx',
	'fixtures/xlsx/poi/FormulaEvalTestData_Copy.xlsx',
	'fixtures/xlsx/poi/evaluate_formula_with_structured_table_references.xlsx',
	'fixtures/xlsx/closedxml/Misc_FormulasWithEvaluation.xlsx',
] as const

const EDGE_CASES = [
	{
		name: 'let-local-binding',
		formula: '=LET(x,1,x+1)',
		cursor: 5,
		expected: 'ok',
	},
	{
		name: 'let-shadowed-inner-binding',
		formula: '=LET(x,1,LET(x,2,x)+x)',
		cursor: 17,
		expected: 'ok',
	},
	{
		name: 'let-shadowed-outer-use',
		formula: '=LET(x,1,LET(x,2,x)+x)',
		cursor: 20,
		expected: 'ok',
	},
	{
		name: 'defined-name-refusal',
		formula: '=Budget+1',
		cursor: 3,
		expected: 'workbook-context-required',
	},
	{
		name: 'table-name-refusal',
		formula: '=SUM(Sales[Amount])',
		cursor: 7,
		expected: 'workbook-context-required',
	},
	{
		name: 'table-column-refusal',
		formula: '=SUM(Sales[Amount])',
		cursor: 13,
		expected: 'workbook-context-required',
	},
	{
		name: 'external-reference-refusal',
		formula: '=[Book.xlsx]Sheet1!A1',
		cursor: 5,
		expected: 'reference-target-not-renameable',
	},
	{
		name: 'three-dimensional-reference-refusal',
		formula: '=SUM(Sheet1:Sheet3!A1)',
		cursor: 9,
		expected: 'reference-target-not-renameable',
	},
	{
		name: 'spill-reference-refusal',
		formula: '=A1#',
		cursor: 2,
		expected: 'reference-target-not-renameable',
	},
	{
		name: 'function-token-refusal',
		formula: '=SUM(A1)',
		cursor: 2,
		expected: 'no-symbol-at-cursor',
	},
] as const

export function runFormulaAssistProof(
	options: FormulaAssistProofOptions = {},
): FormulaAssistProofResult {
	const repeat = options.repeat ?? 8
	const warmup = options.warmup ?? 2
	const includeTimings = options.includeTimings ?? true
	const publicFormulas = collectPublicFormulas(options)
	const sampledFormulas =
		options.publicFormulaLimit === undefined
			? publicFormulas
			: publicFormulas.slice(0, options.publicFormulaLimit)
	const sampledResults = sampledFormulas.map((entry) =>
		assistForFormula(entry.formula, firstUsefulCursor(entry.formula)),
	)
	const edgeCases = EDGE_CASES.map(runEdgeCase)
	const edgeResults = EDGE_CASES.map((entry) => assistForFormula(entry.formula, entry.cursor))
	const allResults = [...sampledResults, ...edgeResults]
	const refusalCounts = blankRefusalCounts()
	for (const result of allResults) {
		const target = result.renameTarget
		if (target && !target.ok && target.reason) refusalCounts[target.reason] += 1
	}
	const timingSamples = includeTimings
		? measureAssistLatency(
				[
					...sampledFormulas.map((entry) => ({
						formula: entry.formula,
						cursor: firstUsefulCursor(entry.formula),
					})),
					...EDGE_CASES.map((entry) => ({ formula: entry.formula, cursor: entry.cursor })),
				],
				repeat,
				warmup,
			)
		: null
	const timings =
		timingSamples === null
			? null
			: {
					repeat,
					warmup,
					sampleCount: timingSamples.length,
					medianMs: percentile(timingSamples, 0.5),
					p95Ms: percentile(timingSamples, 0.95),
					maxMs: Math.max(...timingSamples),
				}
	const passed =
		sampledFormulas.length > 0 &&
		edgeCases.every((entry) => entry.passed) &&
		edgeCases.some((entry) => entry.expected === 'ok' && entry.observed === 'ok') &&
		Object.values(refusalCounts).every((count) => count > 0)
	return {
		generatedAt: new Date().toISOString(),
		publicFormulaCount: publicFormulas.length,
		sampledFormulaCount: sampledFormulas.length,
		staticEdgeCaseCount: EDGE_CASES.length,
		parseOkCount: allResults.filter((result) => result.diagnostics.parseOk).length,
		diagnosticFormulaCount: allResults.filter((result) => result.diagnostics.diagnostics.length > 0)
			.length,
		referenceCount: allResults.reduce((sum, result) => sum + result.references.length, 0),
		bindingCount: allResults.reduce((sum, result) => sum + result.bindings.length, 0),
		longestFormulaChars: Math.max(
			0,
			...sampledFormulas.map((entry) => entry.formula.length),
			...EDGE_CASES.map((entry) => entry.formula.length),
		),
		renameOkCount: allResults.filter((result) => result.renameTarget?.ok).length,
		renameRefusalCounts: refusalCounts,
		edgeCases,
		timings,
		publicSources: [...new Set(publicFormulas.map((entry) => entry.source))],
		passed,
		boundary:
			'This proves formula-local assist latency and rejection-first prepareRename classification. It does not apply workbook edits, resolve workbook names, or claim safe cross-workbook/table rename.',
	}
}

export function formulaAssistProofMarkdown(result: FormulaAssistProofResult): string {
	const timingLines = result.timings
		? [
				`Median assist latency: ${formatMs(result.timings.medianMs)} ms`,
				`P95 assist latency: ${formatMs(result.timings.p95Ms)} ms`,
				`Max assist latency: ${formatMs(result.timings.maxMs)} ms`,
			]
		: ['Timings disabled for this run.']
	return [
		'# Formula Assist Corpus Proof',
		'',
		`Generated: ${result.generatedAt}`,
		`Boundary: ${result.boundary}`,
		'',
		'## Corpus',
		'',
		`Public formulas discovered: ${result.publicFormulaCount}`,
		`Sampled formulas: ${result.sampledFormulaCount}`,
		`Static rejection-first edge cases: ${result.staticEdgeCaseCount}`,
		`Public sources: ${result.publicSources.join(', ')}`,
		'',
		'## Results',
		'',
		`Parse OK formulas: ${result.parseOkCount}`,
		`Diagnostic formulas: ${result.diagnosticFormulaCount}`,
		`Reference spans: ${result.referenceCount}`,
		`Binding roles: ${result.bindingCount}`,
		`Longest formula chars: ${result.longestFormulaChars}`,
		`Prepare-rename OK targets: ${result.renameOkCount}`,
		`Prepare-rename refusals: ${JSON.stringify(result.renameRefusalCounts)}`,
		...timingLines,
		`All proof checks passed: ${result.passed}`,
		'',
		'## Edge Cases',
		'',
		'| Case | Expected | Observed | Role | Reference kind | Occurrences | Passed |',
		'| --- | --- | --- | --- | --- | ---: | --- |',
		...result.edgeCases.map(
			(entry) =>
				`| ${entry.name} | ${entry.expected} | ${entry.observed} | ${entry.role ?? ''} | ${entry.referenceKind ?? ''} | ${entry.occurrenceCount} | ${entry.passed} |`,
		),
	].join('\n')
}

function collectPublicFormulas(
	options: FormulaAssistProofOptions,
): readonly FormulaAssistProofFormula[] {
	const workbookPaths = options.workbookPaths ?? DEFAULT_PUBLIC_WORKBOOKS
	const formulas: FormulaAssistProofFormula[] = []
	for (const workbookPath of workbookPaths) {
		if (!existsSync(workbookPath)) continue
		const result = readXlsx(readFileSync(workbookPath), { mode: 'full' })
		if (!result.ok) continue
		for (const sheet of result.value.workbook.sheets) {
			for (const [row, col, cell] of sheet.cells.iterate()) {
				if (!cell.formula) continue
				formulas.push({
					source: relative(process.cwd(), workbookPath),
					sheet: sheet.name,
					ref: `${indexToColumn(col)}${row + 1}`,
					formula: cell.formula,
				})
			}
		}
	}
	return formulas
}

function runEdgeCase(edgeCase: (typeof EDGE_CASES)[number]): FormulaAssistProofEdgeCase {
	const result = assistForFormula(edgeCase.formula, edgeCase.cursor)
	const target = result.renameTarget
	const observed = target?.ok ? 'ok' : (target?.reason ?? 'none')
	return {
		name: edgeCase.name,
		formula: edgeCase.formula,
		cursor: edgeCase.cursor,
		expected: edgeCase.expected,
		observed,
		role: target?.role?.role,
		referenceKind: target?.reference?.kind,
		occurrenceCount: target?.occurrences.length ?? 0,
		passed: observed === edgeCase.expected,
	}
}

function assistForFormula(formula: string, cursor: number): FormulaAssistResult {
	return formulaAssist(formula, {
		cursor,
		prefix: 'SU',
		completionLimit: 5,
		functionName: 'SUM',
		reference: 'C1',
		replaceReferenceAtCursor: true,
		cycleReference: true,
	})
}

function firstUsefulCursor(formula: string): number {
	const initial = formulaAssist(formula, { cursor: 0 })
	const reference = initial.references[0]
	if (reference) return Math.floor((reference.start + reference.end) / 2)
	const binding = initial.bindings[0]
	if (binding) return Math.floor((binding.start + binding.end) / 2)
	const token = initial.tokens.find((entry) => entry.className !== 'whitespace')
	if (token) return Math.floor((token.start + token.end) / 2)
	return Math.floor(formula.length / 2)
}

function measureAssistLatency(
	inputs: readonly { readonly formula: string; readonly cursor: number }[],
	repeat: number,
	warmup: number,
): readonly number[] {
	const samples: number[] = []
	for (let round = 0; round < warmup + repeat; round++) {
		for (const input of inputs) {
			const start = performance.now()
			assistForFormula(input.formula, input.cursor)
			const elapsed = performance.now() - start
			if (round >= warmup) samples.push(elapsed)
		}
	}
	return samples
}

function percentile(values: readonly number[], fraction: number): number {
	if (values.length === 0) return 0
	const sorted = [...values].sort((a, b) => a - b)
	const index = Math.min(sorted.length - 1, Math.ceil(sorted.length * fraction) - 1)
	return sorted[index] ?? 0
}

function blankRefusalCounts(): Record<FormulaPrepareRenameBlockReason, number> {
	return {
		'no-symbol-at-cursor': 0,
		'workbook-context-required': 0,
		'reference-target-not-renameable': 0,
	}
}

function formatMs(value: number): string {
	return value.toFixed(4)
}

function parseArgs(): FormulaAssistProofOptions & { readonly json: boolean } {
	const repeat = readNumberFlag('--repeat')
	const warmup = readNumberFlag('--warmup')
	const publicFormulaLimit = readNumberFlag('--public-formula-limit')
	return {
		repeat,
		warmup,
		publicFormulaLimit,
		includeTimings: !process.argv.includes('--no-timings'),
		json: process.argv.includes('--json'),
	}
}

function readNumberFlag(name: string): number | undefined {
	const index = process.argv.indexOf(name)
	if (index === -1) return undefined
	const value = Number(process.argv[index + 1])
	return Number.isFinite(value) ? value : undefined
}

if (import.meta.main) {
	const args = parseArgs()
	const result = runFormulaAssistProof(args)
	console.log(args.json ? JSON.stringify(result, null, 2) : formulaAssistProofMarkdown(result))
	process.exit(result.passed ? 0 : 1)
}
