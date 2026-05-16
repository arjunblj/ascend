#!/usr/bin/env bun
/**
 * Safe edit workflow that runs from an installed @ascend/sdk package.
 * Usage after install:
 *   node_modules/.bin/ascend-sdk-safe-edit <input.xlsx> [output.xlsx]
 */
import { existsSync, readFileSync } from 'node:fs'
import { mkdir } from 'node:fs/promises'
import { dirname } from 'node:path'
import { Ascend, createPreparedAgentPlan, inspectWorkbookOpenPlan } from '@ascend/sdk'

const input = process.argv[2]
if (!input) {
	console.error('Usage: node_modules/.bin/ascend-sdk-safe-edit <input.xlsx> [output.xlsx]')
	process.exit(1)
}

const output =
	process.argv[3] ??
	(input.toLowerCase().endsWith('.xlsx')
		? input.replace(/\.xlsx$/i, '.ascend-safe-edit.xlsx')
		: `${input}.ascend-safe-edit.xlsx`)

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
					'Review workbook risks before hydrating cells or planning edits.',
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

if (!prepared.plan.preview.wouldSucceed || prepared.plan.needsApproval) {
	console.log(
		JSON.stringify(
			{
				ok: false,
				step: prepared.plan.needsApproval ? 'approval' : 'plan',
				inputSha256: prepared.inputSha256,
				planDigest: prepared.planDigest,
				errors: prepared.plan.preview.errors,
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
const changedCells = prepared.plan.preview.cellChanges.map((cell) => ({
	ref: `${sheet}!${cell.ref}`,
	before: cell.before,
	after: cell.after,
	formulaBefore: cell.formulaBefore,
	formulaAfter: cell.formulaAfter,
}))
const safetyGates = [
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
	{
		gate: 'plan',
		ok: prepared.plan.preview.wouldSucceed && prepared.plan.approvals.length === 0,
		evidence: {
			planDigest: prepared.planDigest,
			changedCellCount: changedCells.length,
			approvalCount: prepared.plan.approvals.length,
		},
	},
	{
		gate: 'commit',
		ok: committed.postWrite.valid && committed.postWrite.auditsPassed,
		evidence: {
			outputSha256: committed.outputSha256,
			postWriteValid: committed.postWrite.valid,
			auditsPassed: committed.postWrite.auditsPassed,
		},
	},
	{
		gate: 'reopen-verify',
		ok: reopenedCheck.valid && reopenedLint.clean,
		evidence: {
			reopened: true,
			checkIssueCount: reopenedCheck.issues.length,
			lintWarningCount: reopenedLint.warnings.length,
		},
	},
] as const
const safeToUse = safetyGates.every((gate) => gate.ok)

console.log(
	JSON.stringify(
		{
			ok: true,
			workflow: 'installed-sdk-open-plan-trust-inspect-read-plan-commit-reopen-verify',
			install: {
				package: '@ascend/sdk',
				example: 'node_modules/.bin/ascend-sdk-safe-edit',
			},
			input: {
				file: input,
				openPlan: {
					recommendedLoadOptions: openPlan.recommendedLoadOptions,
					reviewBeforeHydration: openPlan.reviewBeforeHydration,
					riskFeatureCount: openPlan.riskFeatures.length,
				},
				trust: {
					trust: trustReport.trust,
					posture: trustReport.posture,
					findingCount: trustReport.summary.findingCount,
				},
				read: {
					ref: readWindow?.ref,
					cellCount: readWindow?.cells.length ?? 0,
				},
				inputSha256: prepared.inputSha256,
				planDigest: prepared.planDigest,
			},
			plan: {
				wouldSucceed: prepared.plan.preview.wouldSucceed,
				changedCells: prepared.plan.preview.cellChanges.map((cell) => cell.ref),
				approvalCount: prepared.plan.approvals.length,
			},
			commit: {
				output: committed.output,
				outputSha256: committed.outputSha256,
				postWriteValid: committed.postWrite.valid,
				auditsPassed: committed.postWrite.auditsPassed,
			},
			verify: {
				reopened: true,
				checkValid: reopenedCheck.valid,
				checkIssueCount: reopenedCheck.issues.length,
				lintClean: reopenedLint.clean,
				lintWarningCount: reopenedLint.warnings.length,
				cell: {
					ref: `${sheet}!B2`,
					formula: reopenedCell?.formula ?? null,
					value: reopenedCell?.value ?? null,
				},
			},
			proofBundle: {
				safeToUse,
				whatChanged: changedCells,
				whySafe: safetyGates,
				evidence: {
					inputSha256: prepared.inputSha256,
					planDigest: prepared.planDigest,
					outputSha256: committed.outputSha256,
					reopened: true,
					checkValid: reopenedCheck.valid,
					lintClean: reopenedLint.clean,
					postWriteValid: committed.postWrite.valid,
					auditsPassed: committed.postWrite.auditsPassed,
				},
			},
		},
		null,
		2,
	),
)
