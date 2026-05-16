#!/usr/bin/env bun
/**
 * Golden-path safe edit through the Ascend HTTP API fetch surface.
 * Usage: bun run examples/agent-safe-edit-http.ts <input.xlsx> [output.xlsx]
 */
import { existsSync } from 'node:fs'
import { mkdir } from 'node:fs/promises'
import { dirname } from 'node:path'
import { createApiFetch } from '@ascend/api'
import { Ascend } from '@ascend/sdk'

const input = process.argv[2]
if (!input) {
	console.error('Usage: bun run examples/agent-safe-edit-http.ts <input.xlsx> [output.xlsx]')
	process.exit(1)
}

const output =
	process.argv[3] ??
	(input.toLowerCase().endsWith('.xlsx')
		? input.replace(/\.xlsx$/i, '.api-agent.xlsx')
		: `${input}.api-agent.xlsx`)

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

type ApiEnvelope<T = Record<string, unknown>> =
	| { ok: true; data: T }
	| { ok: false; error: Record<string, unknown> }

const apiFetch = createApiFetch()

async function postJson<T>(path: string, body: unknown): Promise<T> {
	const response = await apiFetch(
		new Request(`http://ascend.local${path}`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify(body),
		}),
	)
	const envelope = (await response.json()) as ApiEnvelope<T>
	if (!envelope.ok) {
		console.log(JSON.stringify({ ok: false, step: path, error: envelope.error }, null, 2))
		process.exit(1)
	}
	return envelope.data
}

const openPlan = await postJson<{
	recommendedLoadOptions?: Record<string, unknown>
	reviewBeforeHydration?: boolean
	riskFeatures?: unknown[]
	reasons?: string[]
}>('/open-plan', { file: input, intent: 'edit-plan' })

if (openPlan.reviewBeforeHydration) {
	console.log(
		JSON.stringify(
			{
				ok: false,
				step: 'open-plan',
				recommendedLoadOptions: openPlan.recommendedLoadOptions,
				riskFeatures: openPlan.riskFeatures,
				nextActions: ['Review workbook risks before hydrating cells or planning edits.'],
			},
			null,
			2,
		),
	)
	process.exit(1)
}

const trustReport = await postJson<{
	trust?: string
	posture?: string
	summary?: { findingCount?: number }
	findings?: Array<{ code?: string }>
	nextActions?: string[]
}>('/trust-report', { file: input, maxFindings: 20 })

if (trustReport.trust === 'unsafe') {
	console.log(
		JSON.stringify(
			{
				ok: false,
				step: 'trust-preflight',
				trust: trustReport.trust,
				posture: trustReport.posture,
				nextActions: trustReport.nextActions,
			},
			null,
			2,
		),
	)
	process.exit(1)
}

const inspected = await postJson<{
	sheets?: Array<{ name?: string }>
	load?: Record<string, unknown>
}>('/inspect', { file: input })
const sheet = inspected.sheets?.[0]?.name ?? 'Sheet1'
const read = await postJson<{ cells?: unknown[] }>('/read', {
	file: input,
	range: `${sheet}!A1:B4`,
})
const ops = [
	{
		op: 'setFormula' as const,
		sheet,
		ref: 'B2',
		formula: '=SUM(A2:A4)',
	},
]
const plan = await postJson<{
	inputSha256?: string
	planDigest?: string
	operationCount?: number
	preview?: { wouldSucceed?: boolean; cellChanges?: Array<{ ref?: string }> }
	preparedPlan?: { id?: string }
}>('/plan', { file: input, ops, includePackageActions: true })

if (!plan.preview?.wouldSucceed || !plan.preparedPlan?.id) {
	console.log(
		JSON.stringify(
			{
				ok: false,
				step: 'plan',
				inputSha256: plan.inputSha256,
				planDigest: plan.planDigest,
				preview: plan.preview,
			},
			null,
			2,
		),
	)
	process.exit(1)
}

const committed = await postJson<{
	output?: string
	outputSha256?: string
	postWrite?: {
		valid?: boolean
		auditsPassed?: boolean
		check?: { valid?: boolean }
		lint?: { clean?: boolean }
	}
	proofBundle?: {
		safeToUse?: boolean
		whatChanged?: Array<{ ref?: string }>
		whySafe?: Array<{ gate?: string; ok?: boolean }>
	}
}>('/commit', {
	planHandle: plan.preparedPlan.id,
	output,
	includePackageActions: true,
	includeProofBundle: true,
})
const check = await postJson<{ valid?: boolean; issues?: unknown[] }>('/check', { file: output })
const lint = await postJson<{ clean?: boolean; warnings?: unknown[] }>('/lint', { file: output })
const verifiedRead = await postJson<{
	cells?: Array<{ ref?: string; formula?: string; value?: unknown }>
}>('/read', { file: output, range: `${sheet}!B2:B2` })
const verifiedCell = verifiedRead.cells?.[0]

console.log(
	JSON.stringify(
		{
			ok: true,
			workflow: 'api-open-plan-trust-inspect-read-plan-prepared-commit-reopen-verify',
			input: {
				file: input,
				openPlan: {
					recommendedLoadOptions: openPlan.recommendedLoadOptions,
					reviewBeforeHydration: openPlan.reviewBeforeHydration,
					riskFeatureCount: openPlan.riskFeatures?.length ?? 0,
					reasons: openPlan.reasons,
				},
				trust: {
					trust: trustReport.trust,
					posture: trustReport.posture,
					findingCount: trustReport.summary?.findingCount ?? 0,
					codes: trustReport.findings?.map((finding) => finding.code) ?? [],
					nextActions: trustReport.nextActions ?? [],
				},
				inspect: {
					sheet,
					load: inspected.load,
				},
				read: {
					cellCount: read.cells?.length ?? 0,
				},
			},
			plan: {
				inputSha256: plan.inputSha256,
				planDigest: plan.planDigest,
				operationCount: plan.operationCount,
				changedCells: plan.preview.cellChanges?.map((cell) => cell.ref),
				preparedPlanId: plan.preparedPlan.id,
			},
			commit: {
				output: committed.output,
				outputSha256: committed.outputSha256,
				postWriteValid: committed.postWrite?.valid,
				auditsPassed: committed.postWrite?.auditsPassed,
				checkValid: committed.postWrite?.check?.valid,
				lintClean: committed.postWrite?.lint?.clean,
			},
			proofBundle: committed.proofBundle,
			verify: {
				checkValid: check.valid,
				checkIssueCount: check.issues?.length ?? 0,
				lintClean: lint.clean,
				lintWarningCount: lint.warnings?.length ?? 0,
				cell: verifiedCell,
			},
		},
		null,
		2,
	),
)
