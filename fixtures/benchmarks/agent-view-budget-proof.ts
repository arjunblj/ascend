import { readFileSync } from 'node:fs'
import {
	type AgentViewResult,
	AscendWorkbook,
	type Operation,
} from '../../packages/sdk/src/index.ts'

export interface AgentViewBudgetProofCase {
	readonly name: string
	readonly fixture: string
	readonly sheet: string
	readonly range: string
	readonly budget: number
	readonly prepareWorkbook: () => Promise<AscendWorkbook>
}

export interface AgentViewBudgetProofCaseResult {
	readonly name: string
	readonly fixture: string
	readonly sheet: string
	readonly range: string
	readonly requestedApproxTokens: number
	readonly fullApproxTokens: number
	readonly budgetedApproxTokens: number
	readonly unbudgetedApproxTokens: number
	readonly compressionRatio: number
	readonly withinBudget: boolean
	readonly deterministic: boolean
	readonly truncated: boolean
	readonly omittedSampleRows: number
	readonly omittedColumnSampleValues: number
	readonly omittedFormulaPatterns: number
	readonly shapePreserved: boolean
	readonly columnCountPreserved: boolean
	readonly formulaFactsPreserved: boolean
	readonly inspectMetadataSignals: readonly string[]
}

export interface AgentViewBudgetProofResult {
	readonly generatedAt: string
	readonly cases: readonly AgentViewBudgetProofCaseResult[]
	readonly allDeterministic: boolean
	readonly allShapePreserved: boolean
	readonly allOmissionsCounted: boolean
}

export async function runAgentViewBudgetProof(): Promise<AgentViewBudgetProofResult> {
	const cases = defaultAgentViewBudgetProofCases()
	const results: AgentViewBudgetProofCaseResult[] = []
	for (const proofCase of cases) results.push(await runAgentViewBudgetProofCase(proofCase))
	return {
		generatedAt: new Date().toISOString(),
		cases: results,
		allDeterministic: results.every((entry) => entry.deterministic),
		allShapePreserved: results.every((entry) => entry.shapePreserved && entry.columnCountPreserved),
		allOmissionsCounted: results.every(
			(entry) =>
				!entry.truncated ||
				entry.omittedSampleRows + entry.omittedColumnSampleValues + entry.omittedFormulaPatterns >
					0,
		),
	}
}

export function defaultAgentViewBudgetProofCases(): AgentViewBudgetProofCase[] {
	return [
		{
			name: 'dense-table',
			fixture: 'generated dense table',
			sheet: 'Sheet1',
			range: 'A1:H40',
			budget: 512,
			prepareWorkbook: denseTableWorkbook,
		},
		{
			name: 'wide-sparse',
			fixture: 'generated wide sparse sheet',
			sheet: 'Sheet1',
			range: 'A1:Z40',
			budget: 384,
			prepareWorkbook: wideSparseWorkbook,
		},
		{
			name: 'formula-heavy',
			fixture: 'generated formula sheet',
			sheet: 'Sheet1',
			range: 'A1:F60',
			budget: 512,
			prepareWorkbook: formulaHeavyWorkbook,
		},
		{
			name: 'metadata-heavy',
			fixture: 'generated metadata sheet',
			sheet: 'Sheet1',
			range: 'A1:F24',
			budget: 448,
			prepareWorkbook: metadataHeavyWorkbook,
		},
		{
			name: 'public-formula-stress',
			fixture: 'fixtures/xlsx/poi/formula_stress_test.xlsx',
			sheet: 'Finance',
			range: 'A1:L62',
			budget: 640,
			prepareWorkbook: openFixture('fixtures/xlsx/poi/formula_stress_test.xlsx'),
		},
	]
}

export function agentViewBudgetProofMarkdown(result: AgentViewBudgetProofResult): string {
	return [
		'# Agent View Budget Proof',
		'',
		`Generated: ${result.generatedAt}`,
		'Boundary: token counts use Ascend approximate JSON-byte estimates, not model-specific tokenization. Omitted evidence is intentionally absent and must be recovered through narrower reads or unbudgeted views.',
		'',
		'| Case | Fixture | Range | Requested | Full tokens | Budgeted tokens | Unbudgeted recorded | Ratio | Within budget | Deterministic | Truncated | Omitted rows | Omitted values | Omitted formulas | Shape preserved | Metadata signals |',
		'| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |',
		...result.cases.map(markdownRow),
		'',
		`All deterministic: ${result.allDeterministic}`,
		`All shape preserved: ${result.allShapePreserved}`,
		`All truncations counted: ${result.allOmissionsCounted}`,
	].join('\n')
}

async function runAgentViewBudgetProofCase(
	proofCase: AgentViewBudgetProofCase,
): Promise<AgentViewBudgetProofCaseResult> {
	const wb = await proofCase.prepareWorkbook()
	const full = requireAgentView(wb.agentView(proofCase.sheet, proofCase.range), proofCase)
	const budgeted = requireAgentView(
		wb.agentView(proofCase.sheet, proofCase.range, { maxApproxTokens: proofCase.budget }),
		proofCase,
	)
	const repeat = requireAgentView(
		wb.agentView(proofCase.sheet, proofCase.range, { maxApproxTokens: proofCase.budget }),
		proofCase,
	)
	if (!budgeted.budget) throw new Error(`${proofCase.name} did not return budget metadata`)
	const inspect = wb.inspect()
	const omitted =
		budgeted.budget.omittedSampleRows +
		budgeted.budget.omittedColumnSampleValues +
		budgeted.budget.omittedFormulaPatterns
	return {
		name: proofCase.name,
		fixture: proofCase.fixture,
		sheet: proofCase.sheet,
		range: proofCase.range,
		requestedApproxTokens: budgeted.budget.requestedApproxTokens,
		fullApproxTokens: estimateApproxTokens(full),
		budgetedApproxTokens: budgeted.budget.estimatedApproxTokens,
		unbudgetedApproxTokens: budgeted.budget.unbudgetedApproxTokens,
		compressionRatio: roundRatio(
			budgeted.budget.estimatedApproxTokens / estimateApproxTokens(full),
		),
		withinBudget: budgeted.budget.estimatedApproxTokens <= budgeted.budget.requestedApproxTokens,
		deterministic: stableJson(budgeted) === stableJson(repeat),
		truncated: budgeted.budget.truncated,
		omittedSampleRows: budgeted.budget.omittedSampleRows,
		omittedColumnSampleValues: budgeted.budget.omittedColumnSampleValues,
		omittedFormulaPatterns: budgeted.budget.omittedFormulaPatterns,
		shapePreserved:
			budgeted.rowCount === full.rowCount &&
			budgeted.colCount === full.colCount &&
			budgeted.nonEmptyCount === full.nonEmptyCount,
		columnCountPreserved: budgeted.columns.length === full.columns.length,
		formulaFactsPreserved:
			budgeted.formulaCount === full.formulaCount &&
			budgeted.distinctFunctions.join(',') === full.distinctFunctions.join(','),
		inspectMetadataSignals: metadataSignals(inspect, omitted),
	}
}

function requireAgentView(
	view: AgentViewResult | undefined,
	proofCase: AgentViewBudgetProofCase,
): AgentViewResult {
	if (!view) throw new Error(`${proofCase.name} did not produce an agent view`)
	return view
}

async function denseTableWorkbook(): Promise<AscendWorkbook> {
	const wb = AscendWorkbook.create()
	const updates: { ref: string; value: string | number }[] = []
	for (let row = 1; row <= 40; row++) {
		for (let col = 0; col < 8; col++) {
			updates.push({
				ref: `${column(col)}${row}`,
				value: row === 1 ? `Metric ${col + 1}` : row * (col + 1),
			})
		}
	}
	applyExact(wb, [{ op: 'setCells', sheet: 'Sheet1', updates }])
	return wb
}

async function wideSparseWorkbook(): Promise<AscendWorkbook> {
	const wb = AscendWorkbook.create()
	applyExact(wb, [
		{
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [
				{ ref: 'A1', value: 'Scenario' },
				{ ref: 'Z1', value: 'Owner' },
				{ ref: 'M20', value: 42 },
				{ ref: 'Z40', value: 'done' },
			],
		},
	])
	return wb
}

async function formulaHeavyWorkbook(): Promise<AscendWorkbook> {
	const wb = AscendWorkbook.create()
	const values: { ref: string; value: string | number }[] = []
	const formulas: Operation[] = []
	for (let row = 1; row <= 60; row++) {
		values.push({ ref: `A${row}`, value: row })
		values.push({ ref: `B${row}`, value: row % 5 })
		formulas.push({
			op: 'setFormula',
			sheet: 'Sheet1',
			ref: `C${row}`,
			formula: `SUM(A${row}:B${row})`,
		})
		formulas.push({
			op: 'setFormula',
			sheet: 'Sheet1',
			ref: `D${row}`,
			formula: `IF(B${row}>2,C${row},0)`,
		})
	}
	applyExact(wb, [{ op: 'setCells', sheet: 'Sheet1', updates: values }, ...formulas])
	return wb
}

async function metadataHeavyWorkbook(): Promise<AscendWorkbook> {
	const wb = AscendWorkbook.create()
	const updates: { ref: string; value: string | number }[] = []
	for (let row = 1; row <= 24; row++) {
		for (let col = 0; col < 6; col++)
			updates.push({ ref: `${column(col)}${row}`, value: `${row}:${col}` })
	}
	applyExact(wb, [
		{ op: 'setCells', sheet: 'Sheet1', updates },
		{ op: 'setComment', sheet: 'Sheet1', ref: 'A2', text: 'reviewed', author: 'agent' },
		{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'B2', url: 'https://example.com' },
		{
			op: 'setDataValidation',
			sheet: 'Sheet1',
			range: 'C2:C20',
			rule: { type: 'list', formula1: '"yes,no"' },
		},
		{
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'D2:D20',
			rule: { type: 'expression', formula: 'D2<>""', priority: 1 },
		},
	])
	return wb
}

function openFixture(path: string): () => Promise<AscendWorkbook> {
	return async () => AscendWorkbook.open(readFileSync(path))
}

function metadataSignals(
	inspect: ReturnType<AscendWorkbook['inspect']>,
	omittedCount: number,
): string[] {
	const signals: string[] = []
	if (inspect.commentCount) signals.push(`comments=${inspect.commentCount}`)
	if (inspect.hyperlinkCount) signals.push(`hyperlinks=${inspect.hyperlinkCount}`)
	if (inspect.dataValidationCount) signals.push(`validations=${inspect.dataValidationCount}`)
	if (inspect.conditionalFormatCount)
		signals.push(`conditionalFormats=${inspect.conditionalFormatCount}`)
	if (omittedCount > 0) signals.push(`omittedEvidence=${omittedCount}`)
	return signals
}

function applyExact(workbook: AscendWorkbook, ops: readonly Operation[]): void {
	const result = workbook.apply(ops)
	if (result.errors.length > 0)
		throw new Error(result.errors.map((error) => error.message).join('\n'))
}

function estimateApproxTokens(value: unknown): number {
	return Math.ceil(new TextEncoder().encode(JSON.stringify(value)).byteLength / 4)
}

function stableJson(value: unknown): string {
	return JSON.stringify(value)
}

function roundRatio(value: number): number {
	return Number(value.toFixed(3))
}

function column(index: number): string {
	let current = index + 1
	let label = ''
	while (current > 0) {
		const mod = (current - 1) % 26
		label = String.fromCharCode(65 + mod) + label
		current = Math.floor((current - mod - 1) / 26)
	}
	return label
}

function markdownRow(row: AgentViewBudgetProofCaseResult): string {
	return [
		row.name,
		`\`${row.fixture}\``,
		`${row.sheet}!${row.range}`,
		String(row.requestedApproxTokens),
		String(row.fullApproxTokens),
		String(row.budgetedApproxTokens),
		String(row.unbudgetedApproxTokens),
		row.compressionRatio.toFixed(3),
		String(row.withinBudget),
		String(row.deterministic),
		String(row.truncated),
		String(row.omittedSampleRows),
		String(row.omittedColumnSampleValues),
		String(row.omittedFormulaPatterns),
		String(row.shapePreserved && row.columnCountPreserved && row.formulaFactsPreserved),
		row.inspectMetadataSignals.join(', ') || 'none',
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

if (import.meta.main) {
	const result = await runAgentViewBudgetProof()
	if (process.argv.includes('--json')) console.log(JSON.stringify(result, null, 2))
	else {
		console.log(agentViewBudgetProofMarkdown(result))
		console.error(`Generated agent-view budget proof over ${result.cases.length} cases.`)
	}
}
