#!/usr/bin/env bun
/**
 * Golden-path safe edit for coding agents.
 * Usage: bun run examples/agent-safe-edit.ts <input.xlsx> [output.xlsx]
 */
import { existsSync, readFileSync } from 'node:fs'
import { mkdir } from 'node:fs/promises'
import { dirname } from 'node:path'
import {
	Ascend,
	createAgentWorkflowProofSummary,
	createPreparedAgentPlan,
	inspectWorkbookOpenPlan,
} from '@ascend/sdk'

const input = process.argv[2]
if (!input) {
	console.error('Usage: bun run examples/agent-safe-edit.ts <input.xlsx> [output.xlsx]')
	process.exit(1)
}
const output =
	process.argv[3] ??
	(input.toLowerCase().endsWith('.xlsx')
		? input.replace(/\.xlsx$/i, '.agent.xlsx')
		: `${input}.agent.xlsx`)

if (!existsSync(input)) {
	await mkdir(dirname(input), { recursive: true }).catch(() => {})
	const seed = Ascend.create()
	seed.apply([
		{
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [
				{ ref: 'A1', value: 'Revenue' },
				{ ref: 'A2', value: 120 },
				{ ref: 'A3', value: 150 },
				{ ref: 'A4', value: 180 },
			],
		},
	])
	await seed.save(input)
}

const openPlan = inspectWorkbookOpenPlan(new Uint8Array(readFileSync(input)), {
	intent: 'edit-plan',
})
if (openPlan.reviewBeforeHydration) {
	console.log(
		JSON.stringify(
			{
				ok: false,
				step: 'open-plan',
				recommendedLoadOptions: openPlan.recommendedLoadOptions,
				riskFeatures: openPlan.riskFeatures,
				nextActions: [
					'Review trust, package graph, and active-content findings before reading workbook cells.',
					'Promote to editable only after the source is trusted.',
				],
			},
			null,
			2,
		),
	)
	process.exit(1)
}

const inspected = await Ascend.open(input, openPlan.recommendedLoadOptions)
const trustReport = inspected.trustReport({ maxFindings: 20 })
const sheet = inspected.inspect().sheets[0]?.name ?? 'Sheet1'
const readWindow = inspected.sheet(sheet)?.readWindow('A1:B4', { rowLimit: 4 })
const ops = [
	{
		op: 'setFormula' as const,
		sheet,
		ref: 'B2',
		formula: '=SUM(A2:A4)',
	},
]

const prepared = await createPreparedAgentPlan(input, ops)
if (!prepared.plan.preview.wouldSucceed) {
	console.log(
		JSON.stringify(
			{
				ok: false,
				step: 'plan',
				inputSha256: prepared.plan.inputSha256,
				planDigest: prepared.plan.planDigest,
				errors: prepared.plan.preview.errors,
				nextActions: prepared.plan.modelOutput.nextActions,
			},
			null,
			2,
		),
	)
	process.exit(1)
}

if (prepared.plan.needsApproval) {
	console.log(
		JSON.stringify(
			{
				ok: false,
				step: 'approval',
				approvals: prepared.plan.approvals,
				nextActions: prepared.plan.modelOutput.nextActions,
			},
			null,
			2,
		),
	)
	process.exit(1)
}

const committed = await prepared.commit({ output })
const reopened = await Ascend.open(output)
const reopenedCell = reopened.sheet(sheet)?.cell('B2')
const reopenedCheck = reopened.check()
const reopenedLint = reopened.lint()
const proofBundle = createAgentWorkflowProofSummary(prepared.plan, committed, {
	defaultSheetName: sheet,
	preflightGates: [
		{
			gate: 'open-plan',
			ok: !openPlan.reviewBeforeHydration,
			evidence: {
				mode: openPlan.recommendedLoadOptions.mode,
				riskFeatureCount: openPlan.riskFeatures.length,
			},
		},
		{
			gate: 'trust',
			ok: trustReport.trust !== 'unsafe',
			evidence: {
				trust: trustReport.trust,
				posture: trustReport.posture,
				findingCount: trustReport.summary.findingCount,
			},
		},
	],
})

console.log(
	JSON.stringify(
		{
			ok: true,
			workflow: 'open-plan-trust-inspect-read-plan-prepared-commit-verify-repair',
			input: {
				file: input,
				openPlan: {
					recommendedLoadOptions: openPlan.recommendedLoadOptions,
					reviewBeforeHydration: openPlan.reviewBeforeHydration,
					riskFeatureCount: openPlan.riskFeatures.length,
					reasons: openPlan.reasons,
				},
				trust: {
					trust: trustReport.trust,
					posture: trustReport.posture,
					findingCount: trustReport.summary.findingCount,
					codes: trustReport.findings.map((finding) => finding.code),
					nextActions: trustReport.nextActions,
				},
				sheet,
				read: {
					ref: readWindow?.ref,
					cellCount: readWindow?.cells.length ?? 0,
				},
				operationCount: prepared.operationCount,
				inputSha256: prepared.inputSha256,
				planDigest: prepared.planDigest,
			},
			plan: {
				wouldSucceed: prepared.plan.preview.wouldSucceed,
				changedCells: prepared.plan.preview.cellChanges.map((cell) => cell.ref),
				approvalCount: prepared.plan.approvals.length,
				writePolicyDiagnostics: prepared.plan.writePolicy.diagnostics.length,
				nextActions: prepared.plan.modelOutput.nextActions,
			},
			commit: {
				output: committed.output,
				outputSha256: committed.outputSha256,
				postWriteValid: committed.postWrite.valid,
				auditsPassed: committed.postWrite.auditsPassed,
				checkValid: committed.postWrite.check.valid,
				lintClean: committed.postWrite.lint.clean,
			},
			verify: {
				reopened: true,
				cell: {
					ref: `${sheet}!B2`,
					formula: reopenedCell?.formula ?? null,
					value: reopenedCell?.value ?? null,
				},
				checkValid: reopenedCheck.valid,
				checkIssueCount: reopenedCheck.issues.length,
				lintClean: reopenedLint.clean,
				lintWarningCount: reopenedLint.warnings.length,
				commands: {
					check: `ascend check ${output} --json`,
					diff: `ascend diff ${input} ${output} --json`,
					repair: `ascend repair-plan ${output} --json`,
				},
			},
			proofBundle,
		},
		null,
		2,
	),
)
