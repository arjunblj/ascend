#!/usr/bin/env bun
/**
 * Golden-path safe edit through the Ascend MCP tool surface.
 * Usage: bun run examples/agent-safe-edit-mcp.ts <input.xlsx> [output.xlsx]
 */
import { existsSync } from 'node:fs'
import { mkdir } from 'node:fs/promises'
import { dirname } from 'node:path'
import { createServer } from '@ascend/mcp'
import { Ascend } from '@ascend/sdk'

const input = process.argv[2]
if (!input) {
	console.error('Usage: bun run examples/agent-safe-edit-mcp.ts <input.xlsx> [output.xlsx]')
	process.exit(1)
}

const output =
	process.argv[3] ??
	(input.toLowerCase().endsWith('.xlsx')
		? input.replace(/\.xlsx$/i, '.mcp-agent.xlsx')
		: `${input}.mcp-agent.xlsx`)

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

type McpEnvelope<T = Record<string, unknown>> =
	| { ok: true; data: T }
	| { ok: false; error: Record<string, unknown> }

const server = createServer()
const tools = (
	server as unknown as {
		_registeredTools: Record<string, { handler?: (args: unknown) => Promise<unknown> }>
	}
)._registeredTools

async function callTool<T>(name: string, args: unknown): Promise<T> {
	const handler = tools[name]?.handler
	if (typeof handler !== 'function') throw new Error(`Missing MCP tool: ${name}`)
	const result = (await handler(args)) as { structuredContent?: McpEnvelope<T> }
	const envelope = result.structuredContent
	if (!envelope) throw new Error(`MCP tool ${name} did not return structuredContent`)
	if (!envelope.ok) {
		console.log(JSON.stringify({ ok: false, tool: name, error: envelope.error }, null, 2))
		process.exit(1)
	}
	return envelope.data
}

const workflow = await callTool<{
	tools?: Record<string, string>
	workflow?: Array<{ step?: string }>
}>('ascend.agent_workflow', {})
if (workflow.tools?.plan !== 'ascend.plan' || workflow.tools?.commit !== 'ascend.commit') {
	console.log(JSON.stringify({ ok: false, step: 'agent_workflow', workflow }, null, 2))
	process.exit(1)
}

const openPlan = await callTool<{
	recommendedLoadOptions?: Record<string, unknown>
	reviewBeforeHydration?: boolean
	riskFeatures?: unknown[]
	reasons?: string[]
}>('ascend.open_plan', { file: input, intent: 'edit-plan' })

if (openPlan.reviewBeforeHydration) {
	console.log(
		JSON.stringify(
			{
				ok: false,
				step: 'open_plan',
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

const trustReport = await callTool<{
	trust?: string
	posture?: string
	summary?: { findingCount?: number }
	findings?: Array<{ code?: string }>
	nextActions?: string[]
}>('ascend.trust_report', { file: input, maxFindings: 20 })

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

const inspected = await callTool<{
	sheets?: Array<{ name?: string }>
	load?: Record<string, unknown>
}>('ascend.inspect', { file: input })
const sheet = inspected.sheets?.[0]?.name ?? 'Sheet1'
const read = await callTool<{ cells?: unknown[] }>('ascend.read', {
	file: input,
	range: `${sheet}!A1:B4`,
	format: 'cells',
})
const ops = [
	{
		op: 'setFormula' as const,
		sheet,
		ref: 'B2',
		formula: '=SUM(A2:A4)',
	},
]
const plan = await callTool<{
	inputSha256?: string
	planDigest?: string
	operationCount?: number
	preview?: { wouldSucceed?: boolean; cellChanges?: Array<{ ref?: string }> }
	preparedPlan?: { id?: string }
}>('ascend.plan', { file: input, ops, includePackageActions: true })

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

const committed = await callTool<{
	output?: string
	outputSha256?: string
	postWrite?: {
		valid?: boolean
		auditsPassed?: boolean
		check?: { valid?: boolean }
		lint?: { clean?: boolean }
		dataConnections?: unknown
		formulaState?: unknown
		security?: unknown
		visuals?: unknown
	}
	proofBundle?: {
		safeToUse?: boolean
		whatChanged?: Array<{ ref?: string }>
		whySafe?: Array<{ gate?: string; ok?: boolean }>
	}
}>('ascend.commit', {
	planHandle: plan.preparedPlan.id,
	output,
	includePackageActions: true,
	includeProofBundle: true,
})
const check = await callTool<{ valid?: boolean; issues?: unknown[] }>('ascend.check', {
	file: output,
})
const lint = await callTool<{ clean?: boolean; warnings?: unknown[] }>('ascend.lint', {
	file: output,
})
const verifiedRead = await callTool<{
	cells?: Array<{ ref?: string; formula?: string; value?: unknown }>
}>('ascend.read', { file: output, range: `${sheet}!B2:B2`, format: 'cells' })
const verifiedCell = verifiedRead.cells?.[0]

console.log(
	JSON.stringify(
		{
			ok: true,
			workflow: 'mcp-open-plan-trust-inspect-read-plan-prepared-commit-reopen-verify',
			discovery: {
				workflowSteps: workflow.workflow?.length ?? 0,
				planTool: workflow.tools.plan,
				commitTool: workflow.tools.commit,
			},
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
			postWriteProof: {
				dataConnections: committed.postWrite?.dataConnections,
				formulaState: committed.postWrite?.formulaState,
				security: committed.postWrite?.security,
				visuals: committed.postWrite?.visuals,
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
