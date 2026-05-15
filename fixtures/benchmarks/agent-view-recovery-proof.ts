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
	readonly allBudgetMetadataCountOnly: boolean
	readonly narrowerRecoveryRequiresLocatorMetadata: boolean
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
		allBudgetMetadataCountOnly: results.every((entry) => !entry.budgetMetadataHasLocations),
		narrowerRecoveryRequiresLocatorMetadata: results.some((entry) => entry.truncated),
	}
}

export function agentViewRecoveryProofMarkdown(result: AgentViewRecoveryProofResult): string {
	return [
		'# Agent View Omitted Evidence Recovery Proof',
		'',
		`Generated: ${result.generatedAt}`,
		'Boundary: budget metadata currently carries omission counts, not omitted row/column/formula locations. Same-range unbudgeted recovery is exact; automated narrower recovery needs locator metadata before it can be product-proofed.',
		'',
		'| Case | Fixture | Range | Requested | Budgeted tokens | Unbudgeted tokens | Within budget | Truncated | Omitted rows | Omitted values | Omitted formulas | Has locations | Same-range recovered | Recovered rows | Recovered values | Recovered formulas | Recovery action |',
		'| --- | --- | --- | ---: | ---: | ---: | --- | --- | ---: | ---: | ---: | --- | --- | ---: | ---: | ---: | --- |',
		...result.cases.map(markdownRow),
		'',
		`All same-range unbudgeted recoveries exact: ${result.allUnbudgetedRecoveriesExact}`,
		`All budget metadata count-only: ${result.allBudgetMetadataCountOnly}`,
		`Narrower recovery requires locator metadata: ${result.narrowerRecoveryRequiresLocatorMetadata}`,
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
		unbudgetedSameRangeRecovered: stableJson(full) === stableJson(recovered),
		recoveredSampleRows: recoveredShape.sampleRows - budgetedShape.sampleRows,
		recoveredColumnSampleValues:
			recoveredShape.columnSampleValues - budgetedShape.columnSampleValues,
		recoveredFormulaPatterns: recoveredShape.formulaPatterns - budgetedShape.formulaPatterns,
		recoveryAction:
			omittedSampleRows + omittedColumnSampleValues + omittedFormulaPatterns > 0
				? `Run unbudgeted agentView for ${proofCase.sheet}!${proofCase.range}, or choose a human/agent-selected narrower range; budget metadata does not yet locate omitted evidence.`
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
	return Object.keys(budget).some(
		(key) =>
			key === 'omittedRows' ||
			key === 'omittedColumns' ||
			key === 'omittedRanges' ||
			key === 'omittedFormulaPatternRanges',
	)
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
