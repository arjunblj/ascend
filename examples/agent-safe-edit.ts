#!/usr/bin/env bun
/**
 * Golden-path safe edit for coding agents.
 * Usage: bun run examples/agent-safe-edit.ts <input.xlsx> [output.xlsx]
 */
import { existsSync } from 'node:fs'
import { mkdir } from 'node:fs/promises'
import { dirname } from 'node:path'
import { Ascend, createPreparedAgentPlan } from '@ascend/sdk'

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

const inspected = await Ascend.open(input)
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

console.log(
	JSON.stringify(
		{
			ok: true,
			workflow: 'inspect-read-plan-prepared-commit-verify-repair',
			input: {
				file: input,
				sheet,
				read: {
					ref: readWindow?.ref,
					cellCount: readWindow?.cells.length ?? 0,
					changeToken: readWindow?.snapshot.changeToken,
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
				reopen: `ascend check ${output} --json`,
				diff: `ascend diff ${input} ${output} --json`,
				repair: `ascend repair-plan ${output} --json`,
			},
		},
		null,
		2,
	),
)
