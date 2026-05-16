import { existsSync, readFileSync } from 'node:fs'
import { mkdir } from 'node:fs/promises'
import { dirname } from 'node:path'
import { ascendError } from '@ascend/schema'
import {
	Ascend,
	createAgentWorkflowProofSummary,
	createPreparedAgentPlan,
	inspectWorkbookOpenPlan,
} from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'

export const usage = `Usage: ascend example-safe-edit <input.xlsx> [output.xlsx] [flags]

  Run a packaged inspect -> plan -> commit -> reopen -> verify workflow example.
  If <input.xlsx> does not exist, the command creates a small seed workbook first.

Flags:
  --json    Output as JSON
`

export async function exampleSafeEditCommand(
	args: string[],
	flags: Map<string, string>,
): Promise<number> {
	const input = args[0]
	if (!input) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required example-safe-edit input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'example-safe-edit',
					required: ['input.xlsx'],
					workflow: ['open-plan', 'inspect', 'plan', 'commit', 'reopen', 'verify'],
				},
				suggestedFix: 'Run ascend example-safe-edit ./demo.xlsx ./demo.out.xlsx --json.',
			}),
			flags,
		)
		return 1
	}

	const output =
		args[1] ??
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
		cliError(
			ascendError(
				'VALIDATION_ERROR',
				'Safe edit example requires workbook review before hydration',
				{
					retryable: true,
					retryStrategy: 'modified',
					details: {
						step: 'open-plan',
						recommendedLoadOptions: openPlan.recommendedLoadOptions,
						riskFeatures: openPlan.riskFeatures,
						nextActions: [
							'Review workbook risks before hydrating cells or planning edits.',
							'Promote to editable only after the source is trusted.',
						],
					},
					suggestedFix: 'Review the open-plan risks, then use explicit plan and commit commands.',
				},
			),
			flags,
		)
		return 1
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
		cliError(
			ascendError('VALIDATION_ERROR', 'Safe edit example plan did not pass without approvals', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					step: prepared.plan.needsApproval ? 'approval' : 'plan',
					inputSha256: prepared.inputSha256,
					planDigest: prepared.planDigest,
					errors: prepared.plan.preview.errors,
					approvals: prepared.plan.approvals,
					nextActions: prepared.plan.modelOutput.nextActions,
				},
				suggestedFix:
					'Use ascend plan and ascend commit with explicit approvals for this workbook.',
			}),
			flags,
		)
		return 1
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
	const payload = {
		workflow: 'installed-cli-open-plan-trust-inspect-read-plan-commit-reopen-verify',
		install: {
			package: '@ascend/cli',
			example: 'ascend example-safe-edit',
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
		proofBundle,
	}

	if (flags.has('json')) console.log(jsonOut(payload))
	else printSummary(payload)
	return proofBundle.safeToUse ? 0 : 1
}

function printSummary(payload: {
	readonly commit: { readonly output: string; readonly outputSha256: string }
	readonly proofBundle: {
		readonly safeToUse: boolean
		readonly whatChanged: readonly { ref: string }[]
	}
}) {
	console.log(heading('Ascend Safe Edit Example'))
	console.log(bullet('Output', payload.commit.output))
	console.log(bullet('Output SHA-256', payload.commit.outputSha256))
	console.log(bullet('Safe to use', String(payload.proofBundle.safeToUse)))
	console.log(
		bullet('Changed cells', payload.proofBundle.whatChanged.map((cell) => cell.ref).join(', ')),
	)
}
