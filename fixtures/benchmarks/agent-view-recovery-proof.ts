import {
	type AgentViewBudgetProofCase,
	defaultAgentViewBudgetProofCases,
} from './agent-view-budget-proof.ts'

export interface AgentViewRecoveryProofCaseResult {
	readonly name: string
	readonly fixture: string
	readonly sheet: string
	readonly range: string
	readonly requestedApproxTokens: number
	readonly budgetedApproxTokens: number
	readonly unbudgetedApproxTokens: number
	readonly truncated: boolean
	readonly withinBudget: boolean
	readonly omittedSampleRows: number
	readonly omittedColumnSampleValues: number
	readonly omittedFormulaPatterns: number
	readonly budgetMetadataHasLocations: boolean
	readonly sampleRowLocatorExact: boolean
	readonly columnSampleLocatorExact: boolean
	readonly narrowSampleRowsRecovered: boolean
	readonly narrowSampleRange?: string
	readonly nextFormulaPatternExampleRef?: string
	readonly nextFormulaPatternRecovered: boolean
	readonly unbudgetedSameRangeRecovered: boolean
	readonly recoveredSampleRows: number
	readonly recoveredColumnSampleValues: number
	readonly recoveredFormulaPatterns: number
	readonly recoveryAction: string
}

export interface AgentViewRecoveryProofResult {
	readonly generatedAt: string
	readonly cases: readonly AgentViewRecoveryProofCaseResult[]
	readonly allUnbudgetedRecoveriesExact: boolean
	readonly allBudgetMetadataHasLocators: boolean
	readonly allSampleRowLocatorsExact: boolean
	readonly allColumnSampleLocatorsExact: boolean
	readonly allNarrowSampleRowRecoveriesExact: boolean
	readonly allFormulaPatternExampleRecoveriesExact: boolean
}

interface AgentViewShape {
	readonly sampleRows: number
	readonly columnSampleValues: number
	readonly formulaPatterns: number
}

export async function runAgentViewRecoveryProof(): Promise<AgentViewRecoveryProofResult> {
	const results: AgentViewRecoveryProofCaseResult[] = []
	for (const proofCase of defaultAgentViewBudgetProofCases()) {
		results.push(await runAgentViewRecoveryProofCase(proofCase))
	}
	return {
		generatedAt: new Date().toISOString(),
		cases: results,
		allUnbudgetedRecoveriesExact: results.every((entry) => entry.unbudgetedSameRangeRecovered),
		allBudgetMetadataHasLocators: results.every((entry) => entry.budgetMetadataHasLocations),
		allSampleRowLocatorsExact: results.every((entry) => entry.sampleRowLocatorExact),
		allColumnSampleLocatorsExact: results.every((entry) => entry.columnSampleLocatorExact),
		allNarrowSampleRowRecoveriesExact: results.every((entry) => entry.narrowSampleRowsRecovered),
		allFormulaPatternExampleRecoveriesExact: results.every(
			(entry) => entry.nextFormulaPatternRecovered,
		),
	}
}

export function agentViewRecoveryProofMarkdown(result: AgentViewRecoveryProofResult): string {
	return [
		'# Agent View Omitted Evidence Recovery Proof',
		'',
		`Generated: ${result.generatedAt}`,
		'Boundary: budget metadata carries compact omitted sample-row and column-sample locators plus formula-pattern example refs. Same-range unbudgeted recovery is exact; sample-row locators and formula-pattern examples can drive narrower follow-up reads, while full omitted-occurrence recovery still needs richer provenance.',
		'',
		'| Case | Fixture | Range | Requested | Budgeted tokens | Unbudgeted tokens | Within budget | Truncated | Omitted rows | Omitted values | Omitted formulas | Has locations | Row locators exact | Column locators exact | Narrow sample range | Narrow rows recovered | Formula example | Formula recovered | Same-range recovered | Recovered rows | Recovered values | Recovered formulas | Recovery action |',
		'| --- | --- | --- | ---: | ---: | ---: | --- | --- | ---: | ---: | ---: | --- | --- | --- | --- | --- | --- | --- | --- | ---: | ---: | ---: | --- |',
		...result.cases.map(markdownRow),
		'',
		`All same-range unbudgeted recoveries exact: ${result.allUnbudgetedRecoveriesExact}`,
		`All budget metadata has locators: ${result.allBudgetMetadataHasLocators}`,
		`All sample-row locators exact: ${result.allSampleRowLocatorsExact}`,
		`All column-sample locators exact: ${result.allColumnSampleLocatorsExact}`,
		`All narrow sample-row recoveries exact: ${result.allNarrowSampleRowRecoveriesExact}`,
		`All formula-pattern example recoveries exact: ${result.allFormulaPatternExampleRecoveriesExact}`,
	].join('\n')
}

async function runAgentViewRecoveryProofCase(
	proofCase: AgentViewBudgetProofCase,
): Promise<AgentViewRecoveryProofCaseResult> {
	const wb = await proofCase.prepareWorkbook()
	const full = requireAgentView(wb.agentView(proofCase.sheet, proofCase.range), proofCase.name)
	const budgeted = requireAgentView(
		wb.agentView(proofCase.sheet, proofCase.range, { maxApproxTokens: proofCase.budget }),
		proofCase.name,
	)
	const recovered = requireAgentView(wb.agentView(proofCase.sheet, proofCase.range), proofCase.name)
	if (!budgeted.budget) throw new Error(`${proofCase.name} did not return budget metadata`)
	const fullShape = shapeOf(full)
	const budgetedShape = shapeOf(budgeted)
	const recoveredShape = shapeOf(recovered)
	const omittedSampleRows = fullShape.sampleRows - budgetedShape.sampleRows
	const omittedColumnSampleValues = fullShape.columnSampleValues - budgetedShape.columnSampleValues
	const omittedFormulaPatterns = fullShape.formulaPatterns - budgetedShape.formulaPatterns
	const omittedEvidence = budgeted.budget.omittedEvidence
	const narrowSampleRange = omittedEvidence?.sampleRows
		? rangeForRows(full, omittedEvidence.sampleRows.firstRow, omittedEvidence.sampleRows.lastRow)
		: undefined
	const narrowSampleView = narrowSampleRange
		? requireAgentView(wb.agentView(proofCase.sheet, narrowSampleRange), proofCase.name)
		: undefined
	const nextFormulaPatternExampleRef = omittedEvidence?.formulaPatterns?.nextExampleRef
	const nextFormulaPatternView = nextFormulaPatternExampleRef
		? requireAgentView(wb.agentView(proofCase.sheet, nextFormulaPatternExampleRef), proofCase.name)
		: undefined
	return {
		name: proofCase.name,
		fixture: proofCase.fixture,
		sheet: proofCase.sheet,
		range: proofCase.range,
		requestedApproxTokens: budgeted.budget.requestedApproxTokens,
		budgetedApproxTokens: budgeted.budget.estimatedApproxTokens,
		unbudgetedApproxTokens: budgeted.budget.unbudgetedApproxTokens,
		truncated: budgeted.budget.truncated,
		withinBudget: budgeted.budget.estimatedApproxTokens <= budgeted.budget.requestedApproxTokens,
		omittedSampleRows,
		omittedColumnSampleValues,
		omittedFormulaPatterns,
		budgetMetadataHasLocations: hasOmissionLocations(budgeted.budget),
		sampleRowLocatorExact: (omittedEvidence?.sampleRows?.count ?? 0) === omittedSampleRows,
		columnSampleLocatorExact:
			(omittedEvidence?.columnSampleValues?.omittedValues ?? 0) === omittedColumnSampleValues,
		narrowSampleRowsRecovered:
			omittedSampleRows === 0 || narrowSampleView?.samples.length === omittedSampleRows,
		...(narrowSampleRange ? { narrowSampleRange } : {}),
		...(nextFormulaPatternExampleRef ? { nextFormulaPatternExampleRef } : {}),
		nextFormulaPatternRecovered:
			omittedFormulaPatterns === 0 || (nextFormulaPatternView?.formulaPatterns.length ?? 0) > 0,
		unbudgetedSameRangeRecovered: stableJson(full) === stableJson(recovered),
		recoveredSampleRows: recoveredShape.sampleRows - budgetedShape.sampleRows,
		recoveredColumnSampleValues:
			recoveredShape.columnSampleValues - budgetedShape.columnSampleValues,
		recoveredFormulaPatterns: recoveredShape.formulaPatterns - budgetedShape.formulaPatterns,
		recoveryAction:
			omittedSampleRows + omittedColumnSampleValues + omittedFormulaPatterns > 0
				? `Use omittedEvidence sample-row and column-sample locators for narrower follow-up reads, or run unbudgeted agentView for ${proofCase.sheet}!${proofCase.range}.`
				: 'No omitted evidence to recover.',
	}
}

function requireAgentView(
	view: ReturnType<Awaited<ReturnType<AgentViewBudgetProofCase['prepareWorkbook']>>['agentView']>,
	caseName: string,
) {
	if (!view) throw new Error(`${caseName} did not produce an agent view`)
	return view
}

function shapeOf(view: ReturnType<typeof requireAgentView>): AgentViewShape {
	return {
		sampleRows: view.samples.length,
		columnSampleValues: view.columns.reduce(
			(total, column) => total + column.sampleValues.length,
			0,
		),
		formulaPatterns: view.formulaPatterns.length,
	}
}

function hasOmissionLocations(budget: Record<string, unknown>): boolean {
	return typeof budget.omittedEvidence === 'object' && budget.omittedEvidence !== null
}

function rangeForRows(
	view: ReturnType<typeof requireAgentView>,
	firstRow: number,
	lastRow: number,
): string {
	const startCol = columnLabel(view.range.start.col)
	const endCol = columnLabel(view.range.end.col)
	return `${startCol}${firstRow + 1}:${endCol}${lastRow + 1}`
}

function columnLabel(index: number): string {
	let current = index + 1
	let label = ''
	while (current > 0) {
		const mod = (current - 1) % 26
		label = String.fromCharCode(65 + mod) + label
		current = Math.floor((current - mod - 1) / 26)
	}
	return label
}

function stableJson(value: unknown): string {
	return JSON.stringify(value)
}

function markdownRow(row: AgentViewRecoveryProofCaseResult): string {
	return [
		row.name,
		`\`${row.fixture}\``,
		`${row.sheet}!${row.range}`,
		String(row.requestedApproxTokens),
		String(row.budgetedApproxTokens),
		String(row.unbudgetedApproxTokens),
		String(row.withinBudget),
		String(row.truncated),
		String(row.omittedSampleRows),
		String(row.omittedColumnSampleValues),
		String(row.omittedFormulaPatterns),
		String(row.budgetMetadataHasLocations),
		String(row.sampleRowLocatorExact),
		String(row.columnSampleLocatorExact),
		row.narrowSampleRange ?? 'n/a',
		String(row.narrowSampleRowsRecovered),
		row.nextFormulaPatternExampleRef ?? 'n/a',
		String(row.nextFormulaPatternRecovered),
		String(row.unbudgetedSameRangeRecovered),
		String(row.recoveredSampleRows),
		String(row.recoveredColumnSampleValues),
		String(row.recoveredFormulaPatterns),
		row.recoveryAction,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

if (import.meta.main) {
	const result = await runAgentViewRecoveryProof()
	if (process.argv.includes('--json')) console.log(JSON.stringify(result, null, 2))
	else {
		console.log(agentViewRecoveryProofMarkdown(result))
		console.error(`Generated agent-view recovery proof over ${result.cases.length} cases.`)
	}
}
