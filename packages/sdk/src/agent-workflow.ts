import { createHash } from 'node:crypto'
import { copyFile, readFile, rename, writeFile } from 'node:fs/promises'
import { dirname, extname, join } from 'node:path'
import type { ExternalReferenceInfo, Workbook } from '@ascend/core'
import type {
	XlsxPackageGraph,
	XlsxPackageGraphFidelityIssue,
	XlsxPackageLossPolicy,
	XlsxPackageOwnerScope,
} from '@ascend/io-xlsx'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
} from '@ascend/io-xlsx'
import { AscendException, ascendError, type FeatureReport, type Operation } from '@ascend/schema'
import { listCapabilities, summarizeCapabilities } from './capabilities.ts'
import { AscendWorkbook } from './workbook.ts'

export interface AgentPlanResult {
	readonly file: string
	readonly inputSha256: string
	readonly operationCount: number
	readonly planDigest: string
	readonly needsApproval: boolean
	readonly approvals: readonly ApprovalRequirement[]
	readonly trace: AgentWorkflowTrace
	readonly modelOutput: AgentModelOutput
	readonly preview: ReturnType<AscendWorkbook['preview']>
	readonly check: ReturnType<AscendWorkbook['check']>
	readonly lint: ReturnType<AscendWorkbook['lint']>
	readonly preservation: ReturnType<AscendWorkbook['writePlanSummary']>
	readonly writePolicy: WritePolicyReport
	readonly unsupportedFeatures: readonly unknown[]
	readonly lossAudit: LossAudit
	readonly packageGraphAudit: PackageGraphAudit
	readonly capabilities: ReturnType<typeof summarizeCapabilities>
}

export interface AgentCommitOptions {
	readonly output?: string
	readonly inPlace?: boolean
	readonly backup?: string
	readonly expectSha256?: string
	readonly allowLoss?: readonly string[] | 'all'
	readonly approvals?: readonly string[] | 'all'
	readonly onProgress?: AgentWorkflowProgressHandler
}

export interface AgentCommitResult {
	readonly file: string
	readonly output: string
	readonly backup?: string
	readonly inputSha256: string
	readonly outputSha256: string
	readonly planDigest: string
	readonly operationCount: number
	readonly approvals: readonly ApprovalRequirement[]
	readonly trace: AgentWorkflowTrace
	readonly modelOutput: AgentModelOutput
	readonly apply: ReturnType<AscendWorkbook['apply']>
	readonly recalc: ReturnType<AscendWorkbook['recalc']> | null
	readonly check: ReturnType<AscendWorkbook['check']>
	readonly lint: ReturnType<AscendWorkbook['lint']>
	readonly preservation: ReturnType<AscendWorkbook['writePlanSummary']>
	readonly writePolicy: WritePolicyReport
	readonly postWrite: AgentPostWriteVerification
	readonly lossAudit: LossAudit
	readonly packageGraphAudit: PackageGraphAudit
}

export interface AgentPostWriteVerification {
	readonly valid: boolean
	readonly output: string
	readonly outputSha256: string
	readonly reopened: true
	readonly check: ReturnType<AscendWorkbook['check']>
	readonly lint: ReturnType<AscendWorkbook['lint']>
	readonly preservation: ReturnType<AscendWorkbook['writePlanSummary']>
	readonly packageGraphAudit: PackageGraphAudit
}

export interface LossAudit {
	readonly ok: boolean
	readonly blockedFeatures: readonly FeatureReport[]
	readonly blockedPackageParts: readonly LossAuditPackagePart[]
	readonly allowedLoss: readonly string[] | 'all'
	readonly policy: 'block-preserved-and-unsupported'
}

export interface LossAuditPackagePart {
	readonly partPath: string
	readonly featureFamily: string
	readonly preservationPolicy: XlsxPackageLossPolicy
	readonly ownerScope: XlsxPackageOwnerScope
	readonly bytePreservationExpected: boolean
	readonly contentType: string
	readonly contentTypeSource: string
	readonly sourceRelationshipPart?: string
	readonly sourceRelationshipId?: string
	readonly sourceRelationshipType?: string
	readonly sourceRelationshipRawTarget?: string
	readonly sourceRelationshipResolvedTarget?: string
	readonly sourceRelationshipTargetMode?: string
	readonly reason: string
}

export interface PackageGraphAudit {
	readonly ok: boolean
	readonly issues: readonly XlsxPackageGraphFidelityIssue[]
	readonly policy: 'read-integrity' | 'safe-edit-roundtrip'
}

export interface WritePolicyReport {
	readonly ok: boolean
	readonly diagnostics: readonly WritePolicyDiagnostic[]
	readonly summary: {
		readonly generatedParts: number
		readonly copiedThroughParts: number
		readonly skippedCapsules: number
		readonly invalidatedSignatures: number
		readonly approvalRequiredFeatures: number
		readonly packageGraphIssues: number
		readonly externalReferences: number
		readonly externalReferenceBindingIssues: number
		readonly x14ConditionalFormatExtensionPayloads: number
		readonly calcChainPolicy: 'not-present' | 'preserved' | 'discarded-for-formula-topology'
	}
}

export interface WritePolicyDiagnostic {
	readonly code:
		| 'generated-replacement-parts'
		| 'copied-through-package-parts'
		| 'skipped-preservation-capsules'
		| 'signature-invalidation'
		| 'calc-chain-preserved'
		| 'calc-chain-discarded'
		| 'active-content-preserved'
		| 'visual-sidecar-preservation-risk'
		| 'conditional-format-extension-preservation'
		| 'external-link-dependency'
		| 'external-link-binding-risk'
		| 'package-graph-audit-issue'
		| 'approval-required-feature'
	readonly severity: 'info' | 'warning' | 'blocker'
	readonly message: string
	readonly suggestedAction: string
	readonly partPaths?: readonly string[]
	readonly packageParts?: readonly WritePolicyPackagePart[]
	readonly featureFamily?: string
	readonly ownerScope?: XlsxPackageOwnerScope
	readonly preservationPolicy?: XlsxPackageLossPolicy
	readonly details?: unknown
}

export interface WritePolicyPackagePart {
	readonly partPath: string
	readonly featureFamily: string
	readonly preservationPolicy: XlsxPackageLossPolicy
	readonly ownerScope: XlsxPackageOwnerScope
	readonly bytePreservationExpected: boolean
	readonly contentType: string
	readonly contentTypeSource: string
	readonly sourceRelationshipPart?: string
	readonly sourceRelationshipId?: string
	readonly sourceRelationshipType?: string
	readonly sourceRelationshipRawType?: string
	readonly sourceRelationshipRawTarget?: string
	readonly sourceRelationshipResolvedTarget?: string
	readonly sourceRelationshipTargetMode?: string
}

export interface AgentWorkflowTrace {
	readonly formatVersion: 1
	readonly kind: 'plan' | 'commit' | 'repair-plan'
	readonly file: string
	readonly inputSha256: string
	readonly outputSha256?: string
	readonly planDigest?: string
	readonly operationCount?: number
	readonly traceDigest: string
	readonly phases: readonly AgentTracePhase[]
	readonly artifacts: readonly AgentTraceArtifact[]
}

export interface AgentTracePhase {
	readonly phase: string
	readonly status: 'ok' | 'warning' | 'blocked' | 'failed'
	readonly summary: string
	readonly count?: number
	readonly refs?: readonly string[]
	readonly details?: unknown
}

export interface AgentTraceArtifact {
	readonly name: string
	readonly digest: string
	readonly summary: string
}

export type AgentWorkflowProgressHandler = (
	event: AgentWorkflowProgressEvent,
) => void | Promise<void>

export interface AgentWorkflowProgressEvent {
	readonly formatVersion: 1
	readonly sequence: number
	readonly kind: AgentWorkflowTrace['kind']
	readonly phase: string
	readonly status: AgentTracePhase['status'] | 'started' | 'skipped'
	readonly summary: string
	readonly count?: number
	readonly refs?: readonly string[]
	readonly details?: unknown
}

export interface AgentModelOutput {
	readonly summary: string
	readonly blocked: boolean
	readonly warnings: readonly string[]
	readonly nextActions: readonly string[]
	readonly digests: {
		readonly inputSha256: string
		readonly outputSha256?: string
		readonly planDigest?: string
		readonly traceDigest: string
	}
	readonly counts: {
		readonly operations?: number
		readonly changedCells?: number
		readonly recalcErrors?: number
		readonly checkIssues?: number
		readonly lintWarnings?: number
		readonly blockedFeatures?: number
		readonly packageGraphIssues?: number
		readonly postWritePackageGraphIssues?: number
		readonly preservationParts?: number
		readonly writePolicyDiagnostics?: number
	}
}

export interface ApprovalRequirement {
	readonly id: string
	readonly kind: 'lossy-write' | 'destructive-operation'
	readonly severity: 'medium' | 'high' | 'critical'
	readonly title: string
	readonly reason: string
	readonly operationIndex?: number
	readonly operation?: string
	readonly feature?: string
	readonly tier?: FeatureReport['tier']
	readonly satisfies: readonly string[]
}

export interface RepairPlanResult {
	readonly file: string
	readonly inputSha256: string
	readonly trace: AgentWorkflowTrace
	readonly modelOutput: AgentModelOutput
	readonly check: ReturnType<AscendWorkbook['check']>
	readonly lint: ReturnType<AscendWorkbook['lint']>
	readonly unsupportedFeatures: readonly unknown[]
	readonly actions: readonly RepairAction[]
}

export interface RepairAction {
	readonly code: string
	readonly title: string
	readonly command?: string
	readonly details?: unknown
}

export async function createAgentPlan(
	file: string,
	ops: readonly Operation[],
	options: { readonly onProgress?: AgentWorkflowProgressHandler } = {},
): Promise<AgentPlanResult> {
	const progress = createProgressEmitter('plan', options.onProgress)
	await progress('hash-input', 'started', 'Hashing input workbook.')
	const inputSha256 = await fileSha256(file)
	await progress('hash-input', 'ok', 'Input workbook hash captured.')
	await progress('load-workbook', 'started', 'Opening workbook.')
	const wb = await AscendWorkbook.open(file)
	await progress('load-workbook', 'ok', 'Workbook opened.')
	await progress('preview', 'started', 'Previewing operations.', { count: ops.length })
	const preview = wb.preview(ops)
	await progressFromPhase(previewPhase(preview), progress)
	const packageGraph = wb.packageGraph()
	await progress('loss-audit', 'started', 'Auditing preserved and unsupported features.')
	const lossAudit = auditLossPolicy(wb.report.features, [], packageGraph)
	await progressFromPhase(lossAuditPhase(lossAudit), progress)
	await progress('package-graph-audit', 'started', 'Auditing package graph fidelity.')
	const packageGraphAudit = auditPackageGraphIntegrity(packageGraph)
	await progressFromPhase(packageGraphAuditPhase(packageGraphAudit), progress)
	await progress('approval-audit', 'started', 'Auditing explicit approval requirements.')
	const approvals = buildApprovalRequirements(wb.report.features, ops)
	await progressFromPhase(approvalPhase(approvals), progress)
	await progress('check', 'started', 'Running structural checks.')
	const check = wb.check()
	await progressFromPhase(checkPhase(check), progress)
	await progress('lint', 'started', 'Running formula lint.')
	const lint = wb.lint()
	await progressFromPhase(lintPhase(lint), progress)
	await progress('preservation-audit', 'started', 'Summarizing package preservation.')
	const preservation = preview.writePlan ?? wb.writePlanSummary()
	await progressFromPhase(preservationPhase(preservation), progress)
	await progress('write-policy', 'started', 'Explaining write preservation and loss policy.')
	const writePolicy = buildWritePolicyReport(
		wb.report.features,
		packageGraph,
		preservation,
		packageGraphAudit,
		wb.getWorkbookModel(),
		wb.inspect().externalReferenceDetails,
	)
	await progressFromPhase(writePolicyPhase(writePolicy), progress)
	const planDigest = digestPlan(inputSha256, ops)
	await progress('finalize', 'ok', 'Plan digest and trace finalized.')
	const trace = finalizeTrace({
		kind: 'plan',
		file,
		inputSha256,
		planDigest,
		operationCount: ops.length,
		phases: [
			okPhase('hash-input', 'Input workbook hash captured.'),
			previewPhase(preview),
			checkPhase(check),
			lintPhase(lint),
			lossAuditPhase(lossAudit),
			packageGraphAuditPhase(packageGraphAudit),
			approvalPhase(approvals),
			preservationPhase(preservation),
			writePolicyPhase(writePolicy),
		],
		artifacts: [
			artifact('ops', ops, `${ops.length} operation(s)`),
			artifact('preview', preview, `${preview.changedCells.length} changed cell(s)`),
			artifact('lossAudit', lossAudit, `${lossAudit.blockedFeatures.length} blocked feature(s)`),
			artifact(
				'packageGraphAudit',
				packageGraphAudit,
				`${packageGraphAudit.issues.length} package graph issue(s)`,
			),
			artifact('approvals', approvals, `${approvals.length} approval requirement(s)`),
			artifact('preservation', preservation, `${preservation.totalParts} package part(s)`),
			artifact(
				'writePolicy',
				writePolicy,
				`${writePolicy.diagnostics.length} write policy diagnostic(s)`,
			),
		],
	})
	return {
		file,
		inputSha256,
		operationCount: ops.length,
		planDigest,
		needsApproval: approvals.length > 0,
		approvals,
		trace,
		modelOutput: modelOutputFromTrace(trace),
		preview,
		check,
		lint,
		preservation,
		writePolicy,
		unsupportedFeatures: wb.report.features,
		lossAudit,
		packageGraphAudit,
		capabilities: summarizeCapabilities(listCapabilities({ gapsOnly: true })),
	}
}

export async function commitAgentPlan(
	file: string,
	ops: readonly Operation[],
	options: AgentCommitOptions = {},
): Promise<AgentCommitResult> {
	const progress = createProgressEmitter('commit', options.onProgress)
	await progress('hash-input', 'started', 'Hashing input workbook.')
	const inputSha256 = await fileSha256(file)
	await progress('hash-input', 'ok', 'Input workbook hash captured.')
	if (options.expectSha256 && options.expectSha256 !== inputSha256) {
		await progress('hash-guard', 'failed', 'Input hash did not match expected SHA-256.', {
			details: { expected: options.expectSha256, actual: inputSha256 },
		})
		throw new AscendException(
			ascendError('VALIDATION_ERROR', 'Input workbook hash does not match --expect-sha256', {
				details: { expected: options.expectSha256, actual: inputSha256 },
				suggestedFix: 'Re-run ascend plan and commit with the new inputSha256.',
			}),
		)
	}
	await progress(
		'hash-guard',
		'ok',
		options.expectSha256
			? 'Input hash matched expected SHA-256.'
			: 'No input hash guard requested.',
	)

	const output = options.inPlace ? file : options.output
	if (!output) {
		await progress('output-policy', 'failed', 'Commit output target is missing.')
		throw new AscendException(
			ascendError('INVALID_ARGUMENT', 'Commit requires --output unless --in-place is set', {
				suggestedFix: 'Pass --output out.xlsx or --in-place with an optional --backup path.',
			}),
		)
	}
	if (!options.inPlace && options.backup) {
		await progress('output-policy', 'failed', '--backup was provided without --in-place.')
		throw new AscendException(
			ascendError('INVALID_ARGUMENT', '--backup is only valid with --in-place commits', {
				suggestedFix:
					'Use --output for non-destructive writes, or add --in-place when creating a backup.',
			}),
		)
	}
	await progress('output-policy', 'ok', `Commit output target resolved to ${output}.`)

	await progress('load-workbook', 'started', 'Opening workbook.')
	const wb = await AscendWorkbook.open(file)
	await progress('load-workbook', 'ok', 'Workbook opened.')
	const packageGraph = wb.packageGraph()
	await progress('approval-audit', 'started', 'Auditing explicit approval requirements.')
	const approvals = buildApprovalRequirements(wb.report.features, ops)
	const effectiveAllowLoss = mergeAllowLoss(
		options.allowLoss,
		approvalSatisfiedLossFeatures(approvals, options.approvals),
	)
	await progress('loss-audit', 'started', 'Auditing preserved and unsupported features.')
	const lossAudit = auditLossPolicy(wb.report.features, effectiveAllowLoss, packageGraph)
	await progressFromPhase(lossAuditPhase(lossAudit), progress)
	await progress('package-graph-audit', 'started', 'Auditing package graph fidelity.')
	const packageGraphAudit = auditPackageGraphIntegrity(packageGraph)
	const blockedApprovals = unsatisfiedApprovalRequirements(approvals, lossAudit, options.approvals)
	await progressFromPhase(packageGraphAuditPhase(packageGraphAudit), progress)
	await progressFromPhase(approvalPhase(approvals, options.approvals, blockedApprovals), progress)
	if (blockedApprovals.length > 0) {
		throw new AscendException(
			ascendError('VALIDATION_ERROR', 'Commit requires explicit approval', {
				details: { approvals: blockedApprovals, lossAudit },
				suggestedFix:
					'Inspect plan.approvals. If the action is intentional, pass --approval <id> for each requirement, --approval all, or use --allow-loss for lossy workbook features.',
			}),
		)
	}
	await progress('apply', 'started', 'Applying operations.', { count: ops.length })
	const apply = wb.apply(ops, { transaction: true })
	await progressFromPhase(applyPhase(apply), progress)
	if (apply.errors.length > 0)
		throw new AscendException(apply.errors[0] ?? ascendError('VALIDATION_ERROR', 'Apply failed'))

	let recalc: ReturnType<AscendWorkbook['recalc']> | null = null
	if (apply.recalcRequired) {
		await progress('recalc', 'started', 'Recalculating formulas.')
		recalc = wb.recalc()
		await progressFromPhase(recalcPhase(recalc), progress)
		if (recalc.errors.length > 0) {
			const first = recalc.errors[0]
			throw new AscendException(
				first
					? ascendError('FORMULA_EVAL_ERROR', `${first.ref}: ${first.error.message}`, {
							refs: [first.ref],
							details: { evalError: first.error },
							suggestedFix:
								'Run ascend plan to inspect the failing formula state before committing.',
						})
					: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed'),
			)
		}
	} else {
		await progressFromPhase(recalcPhase(recalc), progress)
	}

	if (options.inPlace && options.backup) await copyFile(file, options.backup)
	await progress('preservation-audit', 'started', 'Summarizing package preservation.')
	const preservation = wb.writePlanSummary()
	await progressFromPhase(preservationPhase(preservation), progress)
	await progress('write-policy', 'started', 'Explaining write preservation and loss policy.')
	const writePolicy = buildWritePolicyReport(
		wb.report.features,
		packageGraph,
		preservation,
		packageGraphAudit,
		wb.getWorkbookModel(),
		wb.inspect().externalReferenceDetails,
	)
	await progressFromPhase(writePolicyPhase(writePolicy), progress)
	const sourceBytes = await readFile(file)
	await progress('write', 'started', `Writing workbook to ${output}.`)
	await writeWorkbookAtomically(wb, output)
	const outputSha256 = await fileSha256(output)
	await progress('write', 'ok', `Workbook written to ${output}.`)
	await progress('post-write', 'started', 'Reopening written workbook for verification.')
	const postWrite = await verifyWrittenWorkbook(
		output,
		outputSha256,
		packageGraph,
		new Uint8Array(sourceBytes.buffer, sourceBytes.byteOffset, sourceBytes.byteLength),
	)
	await progressFromPhase(postWritePhase(postWrite), progress)
	await progress('check', 'started', 'Running structural checks.')
	const check = wb.check()
	await progressFromPhase(checkPhase(check), progress)
	await progress('lint', 'started', 'Running formula lint.')
	const lint = wb.lint()
	await progressFromPhase(lintPhase(lint), progress)
	const planDigest = digestPlan(inputSha256, ops)
	await progress('finalize', 'ok', 'Commit digest and trace finalized.')
	const trace = finalizeTrace({
		kind: 'commit',
		file,
		inputSha256,
		outputSha256,
		planDigest,
		operationCount: ops.length,
		phases: [
			okPhase('hash-input', 'Input workbook hash captured.'),
			options.expectSha256
				? okPhase('hash-guard', 'Input hash matched expected SHA-256.')
				: okPhase('hash-guard', 'No input hash guard requested.'),
			lossAuditPhase(lossAudit),
			packageGraphAuditPhase(packageGraphAudit),
			approvalPhase(approvals, options.approvals, blockedApprovals),
			applyPhase(apply),
			recalcPhase(recalc),
			preservationPhase(preservation),
			writePolicyPhase(writePolicy),
			okPhase('write', `Workbook written to ${output}.`),
			postWritePhase(postWrite),
			checkPhase(check),
			lintPhase(lint),
		],
		artifacts: [
			artifact('ops', ops, `${ops.length} operation(s)`),
			artifact('apply', apply, `${apply.affectedCells.length} affected cell(s)`),
			artifact(
				'recalc',
				recalc,
				recalc ? `${recalc.changed.length} recalculated cell(s)` : 'not required',
			),
			artifact('lossAudit', lossAudit, `${lossAudit.blockedFeatures.length} blocked feature(s)`),
			artifact(
				'packageGraphAudit',
				packageGraphAudit,
				`${packageGraphAudit.issues.length} package graph issue(s)`,
			),
			artifact('approvals', approvals, `${approvals.length} approval requirement(s)`),
			artifact('preservation', preservation, `${preservation.totalParts} package part(s)`),
			artifact(
				'writePolicy',
				writePolicy,
				`${writePolicy.diagnostics.length} write policy diagnostic(s)`,
			),
			artifact(
				'postWrite',
				postWrite,
				postWrite.valid
					? 'written workbook reopened and verified'
					: 'written workbook failed verification',
			),
		],
	})

	const result: AgentCommitResult = {
		file,
		output,
		...(options.inPlace && options.backup ? { backup: options.backup } : {}),
		inputSha256,
		outputSha256,
		planDigest,
		operationCount: ops.length,
		approvals,
		trace,
		modelOutput: modelOutputFromTrace(trace),
		apply,
		recalc,
		check,
		lint,
		preservation,
		writePolicy,
		postWrite,
		lossAudit,
		packageGraphAudit,
	}
	return result
}

async function verifyWrittenWorkbook(
	output: string,
	outputSha256: string,
	sourceGraph: XlsxPackageGraph,
	sourceBytes: Uint8Array,
): Promise<AgentPostWriteVerification> {
	const reopened = await AscendWorkbook.open(output, { richMetadata: true })
	const check = reopened.check()
	const lint = reopened.lint()
	const preservation = reopened.writePlanSummary()
	const outputBytes = await readFile(output)
	const packageGraphAudit = auditPackageGraphRoundtrip(
		sourceGraph,
		sourceBytes,
		reopened.packageGraph(),
		new Uint8Array(outputBytes.buffer, outputBytes.byteOffset, outputBytes.byteLength),
	)
	return {
		valid: check.valid,
		output,
		outputSha256,
		reopened: true,
		check,
		lint,
		preservation,
		packageGraphAudit,
	}
}

export function auditLossPolicy(
	features: readonly FeatureReport[],
	allowLoss: readonly string[] | 'all' = [],
	packageGraph?: XlsxPackageGraph,
): LossAudit {
	const allowed = allowLoss === 'all' ? 'all' : allowLoss.map((entry) => entry.toLowerCase())
	const blockedFeatures = features.filter((feature) => {
		if (feature.tier !== 'preserved' && feature.tier !== 'unsupported') return false
		if (isSafePackagePreservationFeature(feature)) return false
		if (allowed === 'all') return false
		return !allowed.includes(feature.feature.toLowerCase()) && !allowed.includes(feature.tier)
	})
	const blockedPackageParts = packageGraph
		? buildBlockedPackageParts(packageGraph, blockedFeatures)
		: []
	return {
		ok: blockedFeatures.length === 0,
		blockedFeatures,
		blockedPackageParts,
		allowedLoss: allowLoss,
		policy: 'block-preserved-and-unsupported',
	}
}

export function auditPackageGraphIntegrity(packageGraph: XlsxPackageGraph): PackageGraphAudit {
	const issues = auditXlsxPackageGraphReadIntegrity(packageGraph)
	return {
		ok: issues.length === 0,
		issues,
		policy: 'read-integrity',
	}
}

export function auditPackageGraphRoundtrip(
	sourceGraph: XlsxPackageGraph,
	sourceBytes: Uint8Array,
	outputGraph: XlsxPackageGraph,
	outputBytes: Uint8Array,
): PackageGraphAudit {
	const issues = [
		...auditXlsxPackageGraphReadIntegrity(outputGraph),
		...auditXlsxPackageGraphSafeEditIntegrity(sourceGraph, outputGraph),
		...auditXlsxPackageGraphBytePreservation(sourceGraph, sourceBytes, outputBytes),
	]
	return {
		ok: issues.length === 0,
		issues,
		policy: 'safe-edit-roundtrip',
	}
}

export async function createRepairPlan(file: string): Promise<RepairPlanResult> {
	const inputSha256 = await fileSha256(file)
	const wb = await AscendWorkbook.open(file)
	const check = wb.check()
	const lint = wb.lint()
	const unsupportedFeatures = wb.report.features
	const actions: RepairAction[] = []
	const checkErrors = check.issues.filter((issue) => issue.severity === 'error')
	if (checkErrors.length > 0) {
		actions.push({
			code: 'run-check',
			title: 'Resolve structural check failures before writing.',
			command: `ascend check ${file} --json`,
			details: { errors: checkErrors },
		})
	}
	if (lint.warnings.length > 0) {
		actions.push({
			code: 'run-lint',
			title: 'Resolve formula lint failures before committing formula edits.',
			command: `ascend lint ${file} --json`,
			details: { warnings: lint.warnings },
		})
	}
	if (unsupportedFeatures.length > 0) {
		actions.push({
			code: 'inspect-unsupported',
			title: 'Review preserved or unsupported workbook features before editing.',
			command: `ascend inspect ${file} --json --verbose`,
			details: { features: unsupportedFeatures },
		})
	}
	if (actions.length === 0) {
		actions.push({
			code: 'plan-before-commit',
			title: 'Workbook is ready for the agent edit workflow.',
			command: `ascend plan ${file} --ops ops.json --json`,
		})
	}
	const trace = finalizeTrace({
		kind: 'repair-plan',
		file,
		inputSha256,
		phases: [
			okPhase('hash-input', 'Input workbook hash captured.'),
			checkPhase(check),
			lintPhase(lint),
			featureAuditPhase(unsupportedFeatures),
			actions.length === 0
				? okPhase('repair-actions', 'No repair actions were needed.')
				: warningPhase(
						'repair-actions',
						`${actions.length} repair action(s) suggested.`,
						actions.length,
					),
		],
		artifacts: [
			artifact('check', check, `${check.issues.length} check issue(s)`),
			artifact('lint', lint, `${lint.warnings.length} lint warning(s)`),
			artifact('features', unsupportedFeatures, `${unsupportedFeatures.length} feature(s)`),
			artifact('actions', actions, `${actions.length} action(s)`),
		],
	})
	return {
		file,
		inputSha256,
		trace,
		modelOutput: modelOutputFromTrace(trace),
		check,
		lint,
		unsupportedFeatures,
		actions,
	}
}

export function digestPlan(inputSha256: string, ops: readonly Operation[]): string {
	return sha256Text(stableStringify({ inputSha256, ops }))
}

export function sha256Bytes(bytes: Uint8Array): string {
	return createHash('sha256').update(bytes).digest('hex')
}

async function fileSha256(file: string): Promise<string> {
	const bytes = await readFile(file)
	return sha256Bytes(new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength))
}

function sha256Text(text: string): string {
	return createHash('sha256').update(text).digest('hex')
}

function finalizeTrace(
	trace: Omit<AgentWorkflowTrace, 'formatVersion' | 'traceDigest'>,
): AgentWorkflowTrace {
	const base = { formatVersion: 1 as const, ...trace }
	return { ...base, traceDigest: sha256Text(stableStringify(base)) }
}

function createProgressEmitter(
	kind: AgentWorkflowTrace['kind'],
	handler: AgentWorkflowProgressHandler | undefined,
): (
	phase: string,
	status: AgentWorkflowProgressEvent['status'],
	summary: string,
	extras?: Pick<AgentWorkflowProgressEvent, 'count' | 'refs' | 'details'>,
) => Promise<void> {
	let sequence = 0
	return async (phase, status, summary, extras = {}) => {
		if (!handler) return
		sequence += 1
		await handler({
			formatVersion: 1,
			sequence,
			kind,
			phase,
			status,
			summary,
			...definedProgressExtras(extras),
		})
	}
}

async function progressFromPhase(
	phase: AgentTracePhase,
	emit: ReturnType<typeof createProgressEmitter>,
): Promise<void> {
	await emit(phase.phase, phase.status, phase.summary, {
		...(phase.count !== undefined ? { count: phase.count } : {}),
		...(phase.refs ? { refs: phase.refs } : {}),
		...(phase.details !== undefined ? { details: phase.details } : {}),
	})
}

function definedProgressExtras(
	extras: Pick<AgentWorkflowProgressEvent, 'count' | 'refs' | 'details'>,
): Pick<AgentWorkflowProgressEvent, 'count' | 'refs' | 'details'> {
	return {
		...(extras.count !== undefined ? { count: extras.count } : {}),
		...(extras.refs ? { refs: extras.refs } : {}),
		...(extras.details !== undefined ? { details: extras.details } : {}),
	}
}

function artifact(name: string, value: unknown, summary: string): AgentTraceArtifact {
	return {
		name,
		digest: sha256Text(stableStringify(value)),
		summary,
	}
}

function okPhase(phase: string, summary: string, count?: number): AgentTracePhase {
	return phaseResult(phase, 'ok', summary, count)
}

function warningPhase(phase: string, summary: string, count?: number): AgentTracePhase {
	return phaseResult(phase, 'warning', summary, count)
}

function phaseResult(
	phase: string,
	status: AgentTracePhase['status'],
	summary: string,
	count?: number,
	details?: unknown,
): AgentTracePhase {
	return {
		phase,
		status,
		summary,
		...(count !== undefined ? { count } : {}),
		...(details !== undefined ? { details } : {}),
	}
}

function previewPhase(preview: ReturnType<AscendWorkbook['preview']>): AgentTracePhase {
	if (preview.errors.length > 0)
		return phaseResult(
			'preview',
			'failed',
			`Preview failed with ${preview.errors.length} error(s).`,
			preview.errors.length,
			preview.errors,
		)
	if (preview.warnings.length > 0)
		return phaseResult(
			'preview',
			'warning',
			`Preview succeeded with ${preview.warnings.length} warning(s).`,
			preview.warnings.length,
			preview.warnings,
		)
	return okPhase(
		'preview',
		`${preview.changedCells.length} changed cell(s) previewed.`,
		preview.changedCells.length,
	)
}

function applyPhase(apply: ReturnType<AscendWorkbook['apply']>): AgentTracePhase {
	if (apply.errors.length > 0)
		return phaseResult(
			'apply',
			'failed',
			`Apply failed with ${apply.errors.length} error(s).`,
			apply.errors.length,
			apply.errors,
		)
	if ((apply.warnings?.length ?? 0) > 0)
		return phaseResult(
			'apply',
			'warning',
			`Apply completed with ${apply.warnings?.length ?? 0} warning(s).`,
			apply.warnings?.length ?? 0,
			apply.warnings,
		)
	return okPhase(
		'apply',
		`${apply.affectedCells.length} affected cell(s).`,
		apply.affectedCells.length,
	)
}

function recalcPhase(recalc: ReturnType<AscendWorkbook['recalc']> | null): AgentTracePhase {
	if (!recalc) return okPhase('recalc', 'Recalculation not required.')
	if (recalc.errors.length > 0)
		return phaseResult(
			'recalc',
			'failed',
			`Recalculation failed with ${recalc.errors.length} error(s).`,
			recalc.errors.length,
			recalc.errors,
		)
	return okPhase('recalc', `${recalc.changed.length} recalculated cell(s).`, recalc.changed.length)
}

function checkPhase(check: ReturnType<AscendWorkbook['check']>): AgentTracePhase {
	const errors = check.issues.filter((issue) => issue.severity === 'error')
	if (errors.length > 0)
		return phaseResult(
			'check',
			'failed',
			`Structural check found ${errors.length} error(s).`,
			errors.length,
			errors,
		)
	if (check.issues.length > 0)
		return phaseResult(
			'check',
			'warning',
			`Structural check found ${check.issues.length} issue(s).`,
			check.issues.length,
			check.issues,
		)
	return okPhase('check', 'Structural check passed.', 0)
}

function lintPhase(lint: ReturnType<AscendWorkbook['lint']>): AgentTracePhase {
	if (lint.warnings.length > 0)
		return phaseResult(
			'lint',
			'warning',
			`Formula lint found ${lint.warnings.length} warning(s).`,
			lint.warnings.length,
			lint.warnings,
		)
	return okPhase('lint', 'Formula lint passed.', 0)
}

function lossAuditPhase(lossAudit: LossAudit): AgentTracePhase {
	if (!lossAudit.ok)
		return phaseResult(
			'loss-audit',
			'blocked',
			`${lossAudit.blockedFeatures.length} preserved or unsupported feature(s) require approval.`,
			lossAudit.blockedFeatures.length,
			{
				blockedFeatures: lossAudit.blockedFeatures,
				blockedPackageParts: lossAudit.blockedPackageParts,
			},
		)
	return okPhase('loss-audit', 'No unapproved preserved or unsupported feature loss detected.', 0)
}

function packageGraphAuditPhase(audit: PackageGraphAudit): AgentTracePhase {
	if (!audit.ok) {
		return phaseResult(
			'package-graph-audit',
			'warning',
			`${audit.issues.length} package graph fidelity issue(s) require inspection.`,
			audit.issues.length,
			{ issues: audit.issues },
		)
	}
	return okPhase('package-graph-audit', 'Package graph fidelity audit passed.', 0)
}

function approvalPhase(
	approvals: readonly ApprovalRequirement[],
	granted: readonly string[] | 'all' = [],
	blockedApprovals?: readonly ApprovalRequirement[],
): AgentTracePhase {
	if (approvals.length === 0)
		return okPhase('approval-audit', 'No approval-gated action detected.', 0)
	const unsatisfied =
		blockedApprovals ?? approvals.filter((approval) => !isApprovalSatisfied(approval, granted))
	if (unsatisfied.length > 0) {
		return phaseResult(
			'approval-audit',
			'blocked',
			`${unsatisfied.length} action(s) require explicit approval.`,
			unsatisfied.length,
			unsatisfied,
		)
	}
	return okPhase(
		'approval-audit',
		`${approvals.length} approval requirement(s) satisfied.`,
		approvals.length,
	)
}

function featureAuditPhase(features: readonly unknown[]): AgentTracePhase {
	if (features.length > 0)
		return warningPhase(
			'feature-audit',
			`${features.length} preserved or unsupported feature(s) require review.`,
			features.length,
		)
	return okPhase('feature-audit', 'No preserved or unsupported features detected.', 0)
}

function preservationPhase(
	preservation: ReturnType<AscendWorkbook['writePlanSummary']>,
): AgentTracePhase {
	return okPhase(
		'preservation-audit',
		`${preservation.totalParts} workbook package part(s) planned for write/preservation.`,
		preservation.totalParts,
	)
}

function writePolicyPhase(writePolicy: WritePolicyReport): AgentTracePhase {
	const blockers = writePolicy.diagnostics.filter((diagnostic) => diagnostic.severity === 'blocker')
	if (blockers.length > 0) {
		return phaseResult(
			'write-policy',
			'blocked',
			`${blockers.length} write policy blocker(s) require action before commit.`,
			blockers.length,
			writePolicy,
		)
	}
	const warnings = writePolicy.diagnostics.filter((diagnostic) => diagnostic.severity === 'warning')
	if (warnings.length > 0) {
		return phaseResult(
			'write-policy',
			'warning',
			`${warnings.length} write policy warning(s) require inspection.`,
			warnings.length,
			writePolicy,
		)
	}
	return phaseResult(
		'write-policy',
		'ok',
		'Write policy has no warning or blocker diagnostics.',
		writePolicy.diagnostics.length,
		writePolicy,
	)
}

function postWritePhase(postWrite: AgentPostWriteVerification): AgentTracePhase {
	const checkErrors = postWrite.check.issues.filter((issue) => issue.severity === 'error')
	if (!postWrite.valid || checkErrors.length > 0) {
		return phaseResult(
			'post-write',
			'failed',
			`Written workbook reopened but structural verification found ${checkErrors.length} error(s).`,
			checkErrors.length,
			{ output: postWrite.output, outputSha256: postWrite.outputSha256, errors: checkErrors },
		)
	}
	if (postWrite.check.issues.length > 0 || postWrite.lint.warnings.length > 0) {
		return phaseResult(
			'post-write',
			'warning',
			`Written workbook reopened with ${postWrite.check.issues.length} check issue(s), ${postWrite.lint.warnings.length} lint warning(s), and ${postWrite.packageGraphAudit.issues.length} package graph issue(s).`,
			postWrite.check.issues.length +
				postWrite.lint.warnings.length +
				postWrite.packageGraphAudit.issues.length,
			{
				output: postWrite.output,
				outputSha256: postWrite.outputSha256,
				check: postWrite.check,
				lint: postWrite.lint,
				packageGraphAudit: postWrite.packageGraphAudit,
			},
		)
	}
	if (!postWrite.packageGraphAudit.ok) {
		return phaseResult(
			'post-write',
			'warning',
			`Written workbook reopened but package graph roundtrip audit found ${postWrite.packageGraphAudit.issues.length} issue(s).`,
			postWrite.packageGraphAudit.issues.length,
			{
				output: postWrite.output,
				outputSha256: postWrite.outputSha256,
				packageGraphAudit: postWrite.packageGraphAudit,
			},
		)
	}
	return phaseResult(
		'post-write',
		'ok',
		'Written workbook reopened and passed post-write package graph verification.',
		0,
		{
			output: postWrite.output,
			outputSha256: postWrite.outputSha256,
			packageGraphAudit: postWrite.packageGraphAudit,
		},
	)
}

function modelOutputFromTrace(trace: AgentWorkflowTrace): AgentModelOutput {
	const blocked = trace.phases.some(
		(phase) => phase.status === 'blocked' || phase.status === 'failed',
	)
	const warnings = trace.phases
		.filter((phase) => phase.status === 'warning')
		.map((phase) => `${phase.phase}: ${phase.summary}`)
	const nextActions = buildNextActions(trace)
	const counts: Partial<Record<keyof AgentModelOutput['counts'], number>> = {}
	setCount(counts, 'operations', trace.operationCount)
	setCount(counts, 'changedCells', traceArtifactCount(trace, 'preview'))
	setCount(counts, 'recalcErrors', phaseCount(trace, 'recalc', 'failed'))
	setCount(counts, 'checkIssues', phaseCount(trace, 'check'))
	setCount(counts, 'lintWarnings', phaseCount(trace, 'lint', 'warning'))
	setCount(counts, 'blockedFeatures', phaseCount(trace, 'loss-audit', 'blocked'))
	setCount(counts, 'packageGraphIssues', phaseCount(trace, 'package-graph-audit', 'warning'))
	setCount(counts, 'postWritePackageGraphIssues', postWritePackageGraphIssueCount(trace))
	setCount(counts, 'preservationParts', phaseCount(trace, 'preservation-audit'))
	setCount(counts, 'writePolicyDiagnostics', phaseCount(trace, 'write-policy'))
	return {
		summary: blocked
			? `${trace.kind} needs attention before it can be treated as safe.`
			: `${trace.kind} completed its safety workflow.`,
		blocked,
		warnings,
		nextActions,
		digests: {
			inputSha256: trace.inputSha256,
			...(trace.outputSha256 ? { outputSha256: trace.outputSha256 } : {}),
			...(trace.planDigest ? { planDigest: trace.planDigest } : {}),
			traceDigest: trace.traceDigest,
		},
		counts,
	}
}

function setCount<K extends keyof AgentModelOutput['counts']>(
	counts: Partial<Record<keyof AgentModelOutput['counts'], number>>,
	key: K,
	value: number | undefined,
): void {
	if (value !== undefined) counts[key] = value
}

function buildNextActions(trace: AgentWorkflowTrace): string[] {
	if (
		trace.phases.some((phase) => phase.phase === 'approval-audit' && phase.status === 'blocked')
	) {
		return [
			'Inspect approvals and rerun commit with --approval <id> only if the action is intentional.',
		]
	}
	if (trace.phases.some((phase) => phase.phase === 'loss-audit' && phase.status === 'blocked')) {
		return [
			'Inspect lossAudit.blockedFeatures and lossAudit.blockedPackageParts.',
			'Commit only with explicit allowLoss entries if the write is intentionally lossy.',
		]
	}
	if (trace.phases.some((phase) => phase.status === 'failed')) {
		return ['Inspect failed trace phases and run repair-plan before committing.']
	}
	if ((postWritePackageGraphIssueCount(trace) ?? 0) > 0) {
		return [
			'Inspect postWrite.packageGraphAudit.issues before treating the output workbook as package-fidelity safe.',
			'Compare the output workbook hash and retain the traceDigest with the edit record.',
		]
	}
	if (trace.phases.some((phase) => phase.phase === 'write-policy' && phase.status === 'warning')) {
		return [
			'Inspect writePolicy.diagnostics before committing around invalidated or copied-through package features.',
			...(trace.kind === 'plan' ? ['Commit with this planDigest and inputSha256 when ready.'] : []),
		]
	}
	if (
		trace.phases.some(
			(phase) => phase.phase === 'package-graph-audit' && phase.status === 'warning',
		)
	) {
		return [
			'Inspect packageGraphAudit.issues before editing around malformed or unclassified package structure.',
			...(trace.kind === 'plan' ? ['Commit with this planDigest and inputSha256 when ready.'] : []),
		]
	}
	if (trace.kind === 'plan') return ['Commit with this planDigest and inputSha256 when ready.']
	if (trace.kind === 'commit')
		return ['Verify the output workbook hash and retain the traceDigest.']
	return ['Follow the suggested repair actions in order.']
}

function postWritePackageGraphIssueCount(trace: AgentWorkflowTrace): number | undefined {
	const postWrite = trace.phases.find((phase) => phase.phase === 'post-write')
	const audit = packageGraphAuditFromDetails(postWrite?.details)
	return audit ? audit.issues.length : undefined
}

function packageGraphAuditFromDetails(details: unknown): PackageGraphAudit | undefined {
	if (!details || typeof details !== 'object') return undefined
	const audit = (details as { packageGraphAudit?: unknown }).packageGraphAudit
	if (!audit || typeof audit !== 'object') return undefined
	const issues = (audit as { issues?: unknown }).issues
	const policy = (audit as { policy?: unknown }).policy
	const ok = (audit as { ok?: unknown }).ok
	if (!Array.isArray(issues) || typeof policy !== 'string' || typeof ok !== 'boolean') {
		return undefined
	}
	return audit as PackageGraphAudit
}

function buildWritePolicyReport(
	features: readonly FeatureReport[],
	packageGraph: XlsxPackageGraph,
	preservation: ReturnType<AscendWorkbook['writePlanSummary']>,
	packageGraphAudit: PackageGraphAudit,
	workbook: Workbook,
	externalReferences: readonly ExternalReferenceInfo[] = [],
): WritePolicyReport {
	const diagnostics: WritePolicyDiagnostic[] = []
	const copiedThroughParts = preservation.parts.filter((part) => part.origin !== 'generated')
	const generatedParts = preservation.parts.filter((part) => part.origin === 'generated')
	const skipped = preservation.skippedCapsules
	const packagePartByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
	if (generatedParts.length > 0) {
		diagnostics.push({
			code: 'generated-replacement-parts',
			severity: 'info',
			message: `${generatedParts.length} package part(s) will be generated or regenerated by Ascend.`,
			suggestedAction:
				'Inspect preservation.parts where origin is generated when auditing replacement package XML.',
			partPaths: generatedParts.map((part) => part.path),
		})
	}
	if (copiedThroughParts.length > 0) {
		const packageParts = packagePartDetails(
			packagePartByPath,
			copiedThroughParts.map((part) => part.path),
		)
		diagnostics.push({
			code: 'copied-through-package-parts',
			severity: 'info',
			message: `${copiedThroughParts.length} package part(s) will be copied through from preserved source or capsule bytes.`,
			suggestedAction:
				'Expect byte-preservation audit coverage for copied-through package parts after write.',
			partPaths: copiedThroughParts.map((part) => part.path),
			...(packageParts.length > 0 ? { packageParts } : {}),
		})
	}
	const signatureParts = packageGraph.parts.filter(
		(part) => part.featureFamily === 'preservedSignature',
	)
	const skippedSignatureParts = signatureParts.filter((part) => skipped.includes(part.path))
	if (skippedSignatureParts.length > 0) {
		diagnostics.push({
			code: 'signature-invalidation',
			severity: 'warning',
			message: `${skippedSignatureParts.length} digital signature package part(s) will be omitted because generated edits invalidate signed package content.`,
			suggestedAction:
				'Commit only with explicit approval and re-sign the workbook outside Ascend if a trusted signature is required.',
			partPaths: skippedSignatureParts.map((part) => part.path),
			packageParts: packagePartDetailsFromParts(skippedSignatureParts),
			featureFamily: 'preservedSignature',
			preservationPolicy: 'invalidate-on-edit',
		})
	}
	const calcChainParts = packageGraph.parts.filter(
		(part) => part.featureFamily === 'preservedCalcChain',
	)
	const skippedCalcChainParts = calcChainParts.filter((part) => skipped.includes(part.path))
	const calcChainPolicy =
		calcChainParts.length === 0
			? 'not-present'
			: skippedCalcChainParts.length > 0
				? 'discarded-for-formula-topology'
				: 'preserved'
	if (skippedCalcChainParts.length > 0) {
		diagnostics.push({
			code: 'calc-chain-discarded',
			severity: 'warning',
			message:
				'The imported calculation chain will be discarded because the planned edit changes formula topology or recalculation freshness.',
			suggestedAction:
				'Allow Excel or another spreadsheet application to rebuild calcChain on open; do not treat the old order as authoritative.',
			partPaths: skippedCalcChainParts.map((part) => part.path),
			packageParts: packagePartDetailsFromParts(skippedCalcChainParts),
			featureFamily: 'preservedCalcChain',
			preservationPolicy: 'discard-on-recalc',
		})
	} else if (calcChainParts.length > 0) {
		diagnostics.push({
			code: 'calc-chain-preserved',
			severity: 'info',
			message:
				'The imported calculation chain is planned for preservation because this edit does not alter formula topology.',
			suggestedAction:
				'Post-write package graph audit should confirm the calcChain part and workbook relationship survive.',
			partPaths: calcChainParts.map((part) => part.path),
			packageParts: packagePartDetailsFromParts(calcChainParts),
			featureFamily: 'preservedCalcChain',
			preservationPolicy: 'discard-on-recalc',
		})
	}
	const skippedNonSignatureCalc = skipped.filter(
		(path) =>
			!signatureParts.some((part) => part.path === path) &&
			!calcChainParts.some((part) => part.path === path) &&
			!preservation.parts.some((part) => part.path === path),
	)
	if (skippedNonSignatureCalc.length > 0) {
		const packageParts = packagePartDetails(packagePartByPath, skippedNonSignatureCalc)
		diagnostics.push({
			code: 'skipped-preservation-capsules',
			severity: 'warning',
			message: `${skippedNonSignatureCalc.length} preservation capsule(s) will not be copied into the written package.`,
			suggestedAction:
				'Inspect skippedCapsules and confirm each skipped part is intentionally regenerated or no longer package-reachable.',
			partPaths: skippedNonSignatureCalc,
			...(packageParts.length > 0 ? { packageParts } : {}),
		})
	}
	for (const part of packageGraph.parts.filter((part) =>
		isActiveContentFeature(part.featureFamily),
	)) {
		diagnostics.push({
			code: 'active-content-preserved',
			severity: 'warning',
			message: `${part.featureFamily} part ${part.path} is active or security-sensitive content planned for preservation, not execution.`,
			suggestedAction:
				'Require explicit approval before writing and never imply macro, ActiveX, callback, or protected-payload execution support.',
			partPaths: [part.path],
			packageParts: packagePartDetailsFromParts([part]),
			featureFamily: part.featureFamily,
			ownerScope: part.ownerScope,
			preservationPolicy: part.preservationPolicy,
		})
	}
	const visualFamilies = new Set<string>()
	const visualPartPaths: string[] = []
	for (const part of packageGraph.parts) {
		if (!isVisualOrAnalyticalSidecar(part.featureFamily)) continue
		visualFamilies.add(part.featureFamily)
		visualPartPaths.push(part.path)
	}
	if (visualPartPaths.length > 0) {
		diagnostics.push({
			code: 'visual-sidecar-preservation-risk',
			severity: 'warning',
			message: `${visualPartPaths.length} chart, drawing, pivot, slicer, or timeline sidecar part(s) require post-write preservation audit.`,
			suggestedAction:
				'Inspect postWrite.packageGraphAudit before treating visuals or analytical caches as fidelity-safe.',
			partPaths: visualPartPaths,
			packageParts: packagePartDetails(packagePartByPath, visualPartPaths),
			featureFamily: [...visualFamilies].sort().join(','),
			preservationPolicy: 'preserve-exact',
		})
	}
	const x14ConditionalFormatExtensionPayloads =
		collectX14ConditionalFormatExtensionPayloads(workbook)
	if (x14ConditionalFormatExtensionPayloads.length > 0) {
		const partPaths = uniqueStrings(
			x14ConditionalFormatExtensionPayloads.flatMap((entry) =>
				entry.sheetPartPath ? [entry.sheetPartPath] : [],
			),
		)
		diagnostics.push({
			code: 'conditional-format-extension-preservation',
			severity: 'warning',
			message: `${x14ConditionalFormatExtensionPayloads.length} x14 conditional-format rule extension payload(s) will be preserved inside generated worksheet XML.`,
			suggestedAction:
				'Inspect sheet x14ConditionalFormats before editing conditional-format order or rules, then verify postWrite.packageGraphAudit after write.',
			...(partPaths.length > 0 ? { partPaths } : {}),
			featureFamily: 'x14ConditionalFormatting',
			preservationPolicy: 'generated',
			details: { x14ConditionalFormats: x14ConditionalFormatExtensionPayloads },
		})
	}
	if (externalReferences.length > 0) {
		const externalReferencePartPaths = uniqueStrings(
			externalReferences.map((entry) => entry.partPath),
		)
		diagnostics.push({
			code: 'external-link-dependency',
			severity: 'info',
			message: `${externalReferences.length} external workbook dependency link(s) are preserved by relationship-id metadata.`,
			suggestedAction:
				'Inspect externalReferenceDetails and externalReferenceUsages before rewriting workbook paths or trusting imported formula dependencies.',
			partPaths: externalReferencePartPaths,
			packageParts: packagePartDetails(packagePartByPath, externalReferencePartPaths),
			featureFamily: 'preservedExternalLink',
			preservationPolicy: 'preserve-exact',
		})
	}
	const externalReferenceBindingIssues = externalReferences.filter(
		(entry) =>
			entry.linkBindingStatus !== undefined && entry.linkBindingStatus !== 'externalBookRelId',
	)
	if (externalReferenceBindingIssues.length > 0) {
		const issuePartPaths = uniqueStrings(
			externalReferenceBindingIssues.map((entry) => entry.partPath),
		)
		const fallbackCount = externalReferenceBindingIssues.filter(
			(entry) => entry.linkBindingStatus === 'fallbackPathRelationship',
		).length
		const missingCount = externalReferenceBindingIssues.filter(
			(entry) => entry.linkBindingStatus === 'missingPathRelationship',
		).length
		diagnostics.push({
			code: 'external-link-binding-risk',
			severity: 'warning',
			message: `${externalReferenceBindingIssues.length} external workbook dependency link(s) have ambiguous package binding (${fallbackCount} fallback, ${missingCount} missing).`,
			suggestedAction:
				'Use rewriteExternalLink with a stable partPath/linkRelId selector to repair or intentionally preserve the externalBook relationship binding.',
			partPaths: issuePartPaths,
			packageParts: packagePartDetails(packagePartByPath, issuePartPaths),
			featureFamily: 'preservedExternalLink',
			preservationPolicy: 'preserve-exact',
		})
	}
	if (!packageGraphAudit.ok) {
		const issuePartPaths = uniqueStrings(
			packageGraphAudit.issues.flatMap((issue) => (issue.partPath ? [issue.partPath] : [])),
		)
		const packageParts = packagePartDetails(packagePartByPath, issuePartPaths)
		diagnostics.push({
			code: 'package-graph-audit-issue',
			severity: 'warning',
			message: `${packageGraphAudit.issues.length} package graph issue(s) were found before write planning.`,
			suggestedAction:
				'Inspect packageGraphAudit.issues and resolve or explicitly accept package graph risk before committing.',
			partPaths: issuePartPaths,
			...(packageParts.length > 0 ? { packageParts } : {}),
		})
	}
	for (const feature of features) {
		if (feature.tier !== 'preserved' && feature.tier !== 'unsupported') continue
		if (isSafePackagePreservationFeature(feature)) continue
		const packageParts = packagePartDetails(packagePartByPath, feature.locations)
		diagnostics.push({
			code: 'approval-required-feature',
			severity: 'warning',
			message: `${feature.feature} (${feature.tier}) requires explicit approval before write.`,
			suggestedAction:
				'Inspect plan.approvals and pass only the corresponding approval id or allow-loss entry when intentional.',
			partPaths: feature.locations,
			...(packageParts.length > 0 ? { packageParts } : {}),
			featureFamily: feature.feature,
		})
	}
	const warningsOrBlockers = diagnostics.some((diagnostic) => diagnostic.severity !== 'info')
	return {
		ok: !warningsOrBlockers,
		diagnostics,
		summary: {
			generatedParts: generatedParts.length,
			copiedThroughParts: copiedThroughParts.length,
			skippedCapsules: skipped.length,
			invalidatedSignatures: skippedSignatureParts.length,
			approvalRequiredFeatures: diagnostics.filter(
				(diagnostic) => diagnostic.code === 'approval-required-feature',
			).length,
			packageGraphIssues: packageGraphAudit.issues.length,
			externalReferences: externalReferences.length,
			externalReferenceBindingIssues: externalReferenceBindingIssues.length,
			x14ConditionalFormatExtensionPayloads: x14ConditionalFormatExtensionPayloads.length,
			calcChainPolicy,
		},
	}
}

interface X14ConditionalFormatExtensionPayload {
	readonly sheetName: string
	readonly sheetPartPath?: string
	readonly sqref: string
	readonly index: number
	readonly priority?: number
	readonly type?: string
	readonly preservedAttributeNames: readonly string[]
	readonly preservedChildElements: readonly string[]
}

function collectX14ConditionalFormatExtensionPayloads(
	workbook: Workbook,
): X14ConditionalFormatExtensionPayload[] {
	const payloads: X14ConditionalFormatExtensionPayload[] = []
	for (const sheet of workbook.sheets) {
		for (const format of sheet.x14ConditionalFormats) {
			const preservedAttributeNames = Object.keys(format.preservedRuleAttributes ?? {}).sort()
			const preservedChildElements = uniqueStrings(
				(format.preservedRuleChildXml ?? []).map(preservedChildElementName),
			).sort()
			if (preservedAttributeNames.length === 0 && preservedChildElements.length === 0) {
				continue
			}
			payloads.push({
				sheetName: sheet.name,
				...(sheet.preservedXml?.partPath ? { sheetPartPath: sheet.preservedXml.partPath } : {}),
				sqref: format.sqref,
				index: format.index,
				...(format.priority !== undefined ? { priority: format.priority } : {}),
				...(format.type ? { type: format.type } : {}),
				preservedAttributeNames,
				preservedChildElements,
			})
		}
	}
	return payloads
}

function preservedChildElementName(xml: string): string {
	return /^<\s*([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)/.exec(xml)?.[1] ?? 'unknown'
}

function packagePartDetails(
	packagePartByPath: ReadonlyMap<string, XlsxPackageGraph['parts'][number]>,
	paths: readonly string[],
): WritePolicyPackagePart[] {
	return packagePartDetailsFromParts(
		uniqueStrings(paths)
			.map((path) => packagePartByPath.get(path))
			.filter((part): part is XlsxPackageGraph['parts'][number] => part !== undefined),
	)
}

function packagePartDetailsFromParts(
	parts: readonly XlsxPackageGraph['parts'][number][],
): WritePolicyPackagePart[] {
	return parts.map((part) => ({
		partPath: part.path,
		featureFamily: part.featureFamily,
		preservationPolicy: part.preservationPolicy,
		ownerScope: part.ownerScope,
		bytePreservationExpected: part.bytePreservationExpected,
		contentType: part.contentType,
		contentTypeSource: part.contentTypeSource,
		...(part.sourceRelationshipPart ? { sourceRelationshipPart: part.sourceRelationshipPart } : {}),
		...(part.sourceRelationshipId ? { sourceRelationshipId: part.sourceRelationshipId } : {}),
		...(part.sourceRelationshipType ? { sourceRelationshipType: part.sourceRelationshipType } : {}),
		...(part.sourceRelationshipRawType
			? { sourceRelationshipRawType: part.sourceRelationshipRawType }
			: {}),
		...(part.sourceRelationshipRawTarget
			? { sourceRelationshipRawTarget: part.sourceRelationshipRawTarget }
			: {}),
		...(part.sourceRelationshipResolvedTarget
			? { sourceRelationshipResolvedTarget: part.sourceRelationshipResolvedTarget }
			: {}),
		...(part.sourceRelationshipTargetMode
			? { sourceRelationshipTargetMode: part.sourceRelationshipTargetMode }
			: {}),
	}))
}

function uniqueStrings(values: readonly string[]): string[] {
	return [...new Set(values)]
}

function isActiveContentFeature(featureFamily: string): boolean {
	return (
		featureFamily === 'preservedMacro' ||
		featureFamily === 'preservedActiveX' ||
		featureFamily === 'preservedControl' ||
		featureFamily === 'preservedCustomUi' ||
		featureFamily === 'preservedVendorSecurity'
	)
}

function isVisualOrAnalyticalSidecar(featureFamily: string): boolean {
	return (
		featureFamily === 'preservedDrawing' ||
		featureFamily === 'preservedChart' ||
		featureFamily === 'preservedChartSheet' ||
		featureFamily === 'preservedChartStyle' ||
		featureFamily === 'preservedChartColor' ||
		featureFamily === 'preservedMedia' ||
		featureFamily === 'preservedVml' ||
		featureFamily === 'preservedPivot' ||
		featureFamily === 'preservedSlicer' ||
		featureFamily === 'preservedTimeline'
	)
}

function buildBlockedPackageParts(
	packageGraph: XlsxPackageGraph,
	blockedFeatures: readonly FeatureReport[],
): LossAuditPackagePart[] {
	const blockedFamilies = new Set(blockedFeatures.map((feature) => feature.feature.toLowerCase()))
	if (blockedFamilies.size === 0) return []
	return packageGraph.parts
		.filter((part) => blockedFamilies.has(part.featureFamily.toLowerCase()))
		.map((part) => ({
			partPath: part.path,
			featureFamily: part.featureFamily,
			preservationPolicy: part.preservationPolicy,
			ownerScope: part.ownerScope,
			bytePreservationExpected: part.bytePreservationExpected,
			contentType: part.contentType,
			contentTypeSource: part.contentTypeSource,
			...(part.sourceRelationshipPart
				? { sourceRelationshipPart: part.sourceRelationshipPart }
				: {}),
			...(part.sourceRelationshipId ? { sourceRelationshipId: part.sourceRelationshipId } : {}),
			...(part.sourceRelationshipType
				? { sourceRelationshipType: part.sourceRelationshipType }
				: {}),
			...(part.sourceRelationshipRawTarget
				? { sourceRelationshipRawTarget: part.sourceRelationshipRawTarget }
				: {}),
			...(part.sourceRelationshipResolvedTarget
				? { sourceRelationshipResolvedTarget: part.sourceRelationshipResolvedTarget }
				: {}),
			...(part.sourceRelationshipTargetMode
				? { sourceRelationshipTargetMode: part.sourceRelationshipTargetMode }
				: {}),
			reason: packagePartLossReason(part.preservationPolicy),
		}))
}

function packagePartLossReason(policy: XlsxPackageLossPolicy): string {
	switch (policy) {
		case 'invalidate-on-edit':
			return 'Generated workbook edits invalidate this package feature unless the workbook is re-signed outside Ascend.'
		case 'unknown-review-required':
			return 'Ascend cannot classify this preserved package part, so an agent must explicitly approve any write around it.'
		case 'preserve-exact':
			return 'This package part is expected to be byte-preserved, but writing around preserved features still requires explicit agent approval.'
		case 'discard-on-recalc':
			return 'This package part may be discarded when formula recalculation invalidates the cached dependency state.'
		case 'inspect-only':
			return 'This package part is inspect-only and cannot be safely regenerated by Ascend.'
		case 'generated':
			return 'This package part is regenerated by Ascend during writes.'
	}
}

function buildApprovalRequirements(
	features: readonly FeatureReport[],
	ops: readonly Operation[],
): ApprovalRequirement[] {
	const approvals: ApprovalRequirement[] = []
	for (const feature of features) {
		if (feature.tier !== 'preserved' && feature.tier !== 'unsupported') continue
		if (isSafePackagePreservationFeature(feature)) continue
		const featureKey = feature.feature.toLowerCase()
		const tierKey = feature.tier.toLowerCase()
		approvals.push({
			id: `loss:${featureKey}:${shortDigest(feature.locations)}`,
			kind: 'lossy-write',
			severity: feature.tier === 'unsupported' ? 'critical' : 'high',
			title: `Approve write with ${feature.feature}`,
			reason:
				'Workbook contains preserved or unsupported package features that require explicit approval before writing.',
			feature: feature.feature,
			tier: feature.tier,
			satisfies: [`loss:${featureKey}`, featureKey, tierKey],
		})
	}
	for (const [operationIndex, op] of ops.entries()) {
		const destructive = destructiveOperationApproval(op, operationIndex)
		if (destructive) approvals.push(destructive)
	}
	return approvals
}

function isSafePackagePreservationFeature(feature: FeatureReport): boolean {
	if (feature.feature === 'preservedDocumentProperties' && feature.tier === 'preserved') return true
	if (feature.feature === 'preservedCalcChain' && feature.tier === 'preserved') return true
	if (feature.feature !== 'preservedOther' || feature.tier !== 'preserved') return false
	if (feature.locations.length === 0) return false
	return feature.locations.every(isSafePackagePreservationLocation)
}

function isSafePackagePreservationLocation(location: string): boolean {
	return (
		location === 'docProps/core.xml' ||
		location === 'docProps/app.xml' ||
		location === 'docProps/custom.xml'
	)
}

function destructiveOperationApproval(
	op: Operation,
	operationIndex: number,
): ApprovalRequirement | null {
	switch (op.op) {
		case 'deleteSheet':
			return destructiveApproval(op, operationIndex, 'critical', `Delete sheet "${op.sheet}"`)
		case 'deleteRows':
			return destructiveApproval(
				op,
				operationIndex,
				'high',
				`Delete ${op.count} row(s) from "${op.sheet}"`,
			)
		case 'deleteCols':
			return destructiveApproval(
				op,
				operationIndex,
				'high',
				`Delete ${op.count} column(s) from "${op.sheet}"`,
			)
		case 'clearRange':
			return op.what === 'all'
				? destructiveApproval(
						op,
						operationIndex,
						'high',
						`Clear all content in ${op.sheet}!${op.range}`,
					)
				: null
		case 'deleteTable':
			return destructiveApproval(op, operationIndex, 'medium', `Delete table "${op.table}"`)
		case 'deleteDefinedName':
			return destructiveApproval(op, operationIndex, 'medium', `Delete defined name "${op.name}"`)
		default:
			return null
	}
}

function destructiveApproval(
	op: Operation,
	operationIndex: number,
	severity: ApprovalRequirement['severity'],
	title: string,
): ApprovalRequirement {
	const opKey = op.op.toLowerCase()
	return {
		id: `op:${operationIndex}:${opKey}`,
		kind: 'destructive-operation',
		severity,
		title,
		reason: 'Operation can remove workbook content or metadata and requires explicit approval.',
		operationIndex,
		operation: op.op,
		satisfies: [`op:${operationIndex}`, opKey],
	}
}

function approvalSatisfiedLossFeatures(
	approvals: readonly ApprovalRequirement[],
	granted: readonly string[] | 'all' | undefined,
): readonly string[] {
	if (!granted) return []
	return approvals
		.filter((approval) => approval.kind === 'lossy-write' && isApprovalSatisfied(approval, granted))
		.map((approval) => approval.feature)
		.filter((feature): feature is string => feature !== undefined)
}

function unsatisfiedApprovalRequirements(
	approvals: readonly ApprovalRequirement[],
	lossAudit: LossAudit,
	granted: readonly string[] | 'all' | undefined,
): ApprovalRequirement[] {
	const blockedLossKeys = new Set(
		lossAudit.blockedFeatures.map((feature) => `${feature.feature}:${feature.tier}`.toLowerCase()),
	)
	return approvals.filter((approval) => {
		if (approval.kind === 'lossy-write') {
			const key = `${approval.feature ?? ''}:${approval.tier ?? ''}`.toLowerCase()
			return blockedLossKeys.has(key)
		}
		return !isApprovalSatisfied(approval, granted)
	})
}

function isApprovalSatisfied(
	approval: ApprovalRequirement,
	granted: readonly string[] | 'all' | undefined,
): boolean {
	if (granted === 'all') return true
	const normalized = new Set((granted ?? []).map((entry) => entry.toLowerCase()))
	return (
		normalized.has(approval.id.toLowerCase()) ||
		approval.satisfies.some((entry) => normalized.has(entry.toLowerCase()))
	)
}

function mergeAllowLoss(
	base: readonly string[] | 'all' | undefined,
	extra: readonly string[],
): readonly string[] | 'all' {
	if (base === 'all') return 'all'
	return [...new Set([...(base ?? []), ...extra])]
}

function shortDigest(value: unknown): string {
	return sha256Text(stableStringify(value)).slice(0, 10)
}

function phaseCount(
	trace: AgentWorkflowTrace,
	phaseName: string,
	status?: AgentTracePhase['status'],
): number | undefined {
	const phase = trace.phases.find(
		(candidate) => candidate.phase === phaseName && (status ? candidate.status === status : true),
	)
	return phase?.count
}

function traceArtifactCount(trace: AgentWorkflowTrace, name: string): number | undefined {
	const artifact = trace.artifacts.find((entry) => entry.name === name)
	if (!artifact) return undefined
	const match = /^(\d+)/.exec(artifact.summary)
	return match ? Number(match[1]) : undefined
}

async function writeWorkbookAtomically(wb: AscendWorkbook, output: string): Promise<void> {
	const ext = extname(output)
	const temp = join(dirname(output), `.${Date.now()}.${process.pid}.ascend-tmp${ext}`)
	if (ext === '.csv' || ext === '.tsv') {
		await wb.save(temp)
	} else {
		await writeFile(temp, wb.toBytes())
	}
	await rename(temp, output)
}

function stableStringify(value: unknown): string {
	if (Array.isArray(value)) return `[${value.map((item) => stableStringify(item)).join(',')}]`
	if (value && typeof value === 'object') {
		const entries = Object.entries(value as Record<string, unknown>).sort(([a], [b]) =>
			a.localeCompare(b),
		)
		return `{${entries.map(([key, val]) => `${JSON.stringify(key)}:${stableStringify(val)}`).join(',')}}`
	}
	return JSON.stringify(value)
}
