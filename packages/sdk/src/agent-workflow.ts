import { createHash } from 'node:crypto'
import { copyFile, readFile, rename, writeFile } from 'node:fs/promises'
import { dirname, extname, join } from 'node:path'
import {
	type ChartPartInfo,
	type ExternalReferenceInfo,
	indexToColumn,
	parseRange,
	type RangeRef,
	type SheetDrawingObjectRef,
	type SheetImageRef,
	type Workbook,
} from '@ascend/core'
import { normalizeFormulaInput, parseFormula } from '@ascend/formulas'
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
import { collectFormulaReferences } from './formula-info.ts'
import type { CheckIssue, FormulaReferenceInfo } from './types.ts'
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
		readonly chartSourceIntegrityIssues: number
		readonly legacyCommentLocations: number
		readonly threadedCommentLocations: number
		readonly commentIntegrityIssues: number
		readonly x14ConditionalFormatExtensionPayloads: number
		readonly x14DataValidationExtensionPayloads: number
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
		| 'visual-edit-preservation-risk'
		| 'drawingml-vml-drift-risk'
		| 'chart-source-ref-drift-risk'
		| 'analytics-preservation-risk'
		| 'analytics-pivot-refresh-risk'
		| 'table-preservation-risk'
		| 'legacy-comment-preservation-risk'
		| 'threaded-comment-preservation-risk'
		| 'conditional-format-extension-preservation'
		| 'data-validation-extension-preservation'
		| 'external-link-dependency'
		| 'external-link-binding-risk'
		| 'external-link-package-risk'
		| 'package-graph-audit-issue'
		| 'pre-write-check-error'
		| 'approval-required-feature'
	readonly severity: 'info' | 'warning' | 'blocker'
	readonly message: string
	readonly suggestedAction: string
	readonly locations?: readonly string[]
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
	const writePolicyWorkbook = snapshotWritePolicyWorkbook(wb.getWorkbookModel())
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
		writePolicyWorkbook,
		ops,
		wb.inspect().externalReferenceDetails,
		check.issues,
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
	const writePolicyWorkbook = snapshotWritePolicyWorkbook(wb.getWorkbookModel())
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
	const writePolicyCheck = wb.check()
	const writePolicy = buildWritePolicyReport(
		wb.report.features,
		packageGraph,
		preservation,
		packageGraphAudit,
		writePolicyWorkbook,
		ops,
		wb.inspect().externalReferenceDetails,
		writePolicyCheck.issues,
	)
	await progressFromPhase(writePolicyPhase(writePolicy), progress)
	if (writePolicy.diagnostics.some((diagnostic) => diagnostic.severity === 'blocker')) {
		throw new AscendException(
			ascendError('VALIDATION_ERROR', 'Commit blocked by write policy', {
				details: { writePolicy },
				suggestedFix:
					'Inspect writePolicy.diagnostics and repair blocker diagnostics before committing.',
			}),
		)
	}
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

function snapshotWritePolicyWorkbook(workbook: Workbook): Workbook {
	const snapshot = workbook.clone()
	for (const sheet of snapshot.sheets) sheet.ensureWritable()
	return snapshot
}

function buildWritePolicyReport(
	features: readonly FeatureReport[],
	packageGraph: XlsxPackageGraph,
	preservation: ReturnType<AscendWorkbook['writePlanSummary']>,
	packageGraphAudit: PackageGraphAudit,
	workbook: Workbook,
	operations: readonly Operation[] = [],
	externalReferences: readonly ExternalReferenceInfo[] = [],
	checkIssues: readonly CheckIssue[] = [],
): WritePolicyReport {
	const diagnostics: WritePolicyDiagnostic[] = []
	const copiedThroughParts = preservation.parts.filter((part) => part.origin !== 'generated')
	const generatedParts = preservation.parts.filter((part) => part.origin === 'generated')
	const skipped = preservation.skippedCapsules
	const packagePartByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
	const checkErrors = checkIssues.filter((issue) => issue.severity === 'error')
	if (checkErrors.length > 0) {
		diagnostics.push({
			code: 'pre-write-check-error',
			severity: 'blocker',
			message: `${checkErrors.length} structural check error(s) would be written to the output workbook.`,
			suggestedAction:
				'Run ascend check or repair-plan, fix the reported workbook integrity errors, and re-run plan before committing.',
			details: { checkIssues: checkErrors },
		})
	}
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
	const analyticsPartPaths: string[] = []
	for (const part of packageGraph.parts) {
		if (isVisualSidecar(part.featureFamily)) {
			visualFamilies.add(part.featureFamily)
			visualPartPaths.push(part.path)
		}
		if (isAnalyticalSidecar(part.featureFamily)) analyticsPartPaths.push(part.path)
	}
	const visualWriteRisk = buildVisualWriteRisk(workbook, operations, packageGraph.parts)
	const chartIntegrityIssues = collectChartIntegrityIssues(checkIssues)
	const visualHasOperationRisk =
		visualWriteRisk.relatedOperations.length > 0 ||
		visualWriteRisk.chartSourceRefDrift.length > 0 ||
		visualWriteRisk.drawingmlVmlDrift.length > 0
	const visualPackageGraphIssues = packageGraphAudit.issues.filter((issue) =>
		isVisualPackageGraphIssue(issue),
	)
	const visualIssuePartPaths = uniqueStrings([
		...visualPackageGraphIssues.flatMap((issue) =>
			packageGraphIssuePackagePaths(issue).filter(isVisualPackagePartPath),
		),
		...chartIntegrityIssues.flatMap((issue) =>
			checkIssuePackagePartPaths(issue).filter(isVisualPackagePartPath),
		),
	])
	const visualDiagnosticPartPaths = uniqueStrings([...visualPartPaths, ...visualIssuePartPaths])
	for (const issue of visualPackageGraphIssues) {
		if (issue.featureFamily && isVisualSidecar(issue.featureFamily)) {
			visualFamilies.add(issue.featureFamily)
		}
	}
	if (
		visualDiagnosticPartPaths.length > 0 ||
		visualPackageGraphIssues.length > 0 ||
		chartIntegrityIssues.length > 0
	) {
		const copiedThroughVisualPartPaths = copiedThroughParts
			.map((part) => part.path)
			.filter((path) => visualPartPaths.includes(path))
		const generatedOrReplacementVisualParts = generatedParts
			.filter(
				(part) => visualDiagnosticPartPaths.includes(part.path) || isDrawingMlVisualPath(part.path),
			)
			.map((part) => ({ partPath: part.path, origin: part.origin }))
		const visualWarningIssueCount = visualPackageGraphIssues.filter(
			(issue) => issue.severity === 'error' || issue.severity === 'warning',
		).length
		const visualHasPackageRisk = visualWarningIssueCount > 0
		const visualHasIntegrityRisk = chartIntegrityIssues.length > 0
		const visualPartCount = uniqueStrings(
			visualDiagnosticPartPaths.length > 0
				? visualDiagnosticPartPaths
				: visualPackageGraphIssues.flatMap((issue) => packageGraphIssuePackagePaths(issue)),
		)
		diagnostics.push({
			code: 'visual-sidecar-preservation-risk',
			severity:
				visualHasOperationRisk || visualHasPackageRisk || visualHasIntegrityRisk
					? 'warning'
					: 'info',
			message:
				visualHasOperationRisk || visualHasPackageRisk || visualHasIntegrityRisk
					? `${visualPartCount.length} chart, drawing, image, or VML sidecar package path(s) have operation-scoped write risk, package graph issues, or verify issues.`
					: `${visualPartCount.length} chart, drawing, image, or VML sidecar part(s) will be copied through; no visual or chart-source edit is planned.`,
			suggestedAction:
				visualHasOperationRisk || visualHasPackageRisk || visualHasIntegrityRisk
					? 'Inspect visualInventory, verify chart source refs, and check postWrite.packageGraphAudit before treating visuals as fidelity-safe.'
					: 'No action is required for unrelated cell edits; retain visualInventory/packageGraphAudit with the edit record when visual fidelity is audited.',
			...(visualDiagnosticPartPaths.length > 0 ? { partPaths: visualDiagnosticPartPaths } : {}),
			...(visualDiagnosticPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, visualDiagnosticPartPaths) }
				: {}),
			featureFamily: [...visualFamilies].sort().join(','),
			preservationPolicy: 'preserve-exact',
			details: {
				operationScoped: visualHasOperationRisk,
				packageGraphAudit: {
					policy: packageGraphAudit.policy,
					ok: packageGraphAudit.ok,
					issueCount: packageGraphAudit.issues.length,
					visualIssueCount: visualPackageGraphIssues.length,
					issues: visualPackageGraphIssues,
				},
				verifyIssues: chartIntegrityIssues,
				copiedThroughVisualParts: packagePartDetails(
					packagePartByPath,
					copiedThroughVisualPartPaths,
				),
				generatedOrReplacementVisualParts,
				chartSourceRefs: collectChartSourceRefSummary(workbook),
				drawingModel: collectDrawingModelSummary(workbook, packageGraph.parts),
				recommendedInspection:
					'Inspect visualInventory before visual edits; after write, inspect postWrite.check, postWrite.packageGraphAudit, chart source refs, DrawingML drawing relationships/media bytes, and any separate VML legacy drawing/comment layout.',
				relatedOperations: visualWriteRisk.relatedOperations,
				chartSourceRefDrift: visualWriteRisk.chartSourceRefDrift,
				drawingmlVmlDrift: visualWriteRisk.drawingmlVmlDrift,
			},
		})
	}
	const analyticsPackageGraphIssues = packageGraphAudit.issues.filter((issue) =>
		isAnalyticalPackageGraphIssue(issue),
	)
	if (
		analyticsPartPaths.length > 0 ||
		analyticsPackageGraphIssues.length > 0 ||
		visualWriteRisk.analyticsPivotRefreshRisk.length > 0
	) {
		const copiedThroughAnalyticsPartPaths = copiedThroughParts
			.map((part) => part.path)
			.filter((path) => analyticsPartPaths.includes(path))
		const generatedOrReplacementAnalyticsParts = generatedParts
			.filter(
				(part) => analyticsPartPaths.includes(part.path) || isAnalyticalPackagePartPath(part.path),
			)
			.map((part) => ({ partPath: part.path, origin: part.origin }))
		const operationScoped = visualWriteRisk.analyticsPivotRefreshRisk.length > 0
		const warningIssueCount = analyticsPackageGraphIssues.filter(
			(issue) => issue.severity === 'error' || issue.severity === 'warning',
		).length
		const analyticsPartCount = uniqueStrings([
			...analyticsPartPaths,
			...analyticsPackageGraphIssues.flatMap((issue) => packageGraphIssuePackagePaths(issue)),
		]).length
		diagnostics.push({
			code: 'analytics-preservation-risk',
			severity: operationScoped || warningIssueCount > 0 ? 'warning' : 'info',
			message:
				operationScoped || warningIssueCount > 0
					? `${analyticsPartCount} pivot, slicer, or timeline package part(s) have operation-scoped write risk or package graph issues.`
					: `${analyticsPartCount} pivot, slicer, or timeline package part(s) will be preserved; no analytical cache/filter edit is planned.`,
			suggestedAction:
				'Inspect pivotRefreshPlans, pivotCaches, pivotTables, slicerCaches, timelineCaches, packageGraphAudit.issues, and postWrite.packageGraphAudit before trusting analytical output after write.',
			...(analyticsPartPaths.length > 0 ? { partPaths: analyticsPartPaths } : {}),
			...(analyticsPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, analyticsPartPaths) }
				: {}),
			featureFamily: 'preservedPivot,preservedSlicer,preservedTimeline',
			preservationPolicy: 'preserve-exact',
			details: {
				operationScoped,
				copiedThroughAnalyticsParts: packagePartDetails(
					packagePartByPath,
					copiedThroughAnalyticsPartPaths,
				),
				generatedOrReplacementAnalyticsParts,
				packageGraphAudit: {
					policy: packageGraphAudit.policy,
					ok: packageGraphAudit.ok,
					issueCount: packageGraphAudit.issues.length,
					analyticsIssueCount: analyticsPackageGraphIssues.length,
					issues: analyticsPackageGraphIssues,
				},
				pivotCacheRisks: collectPivotCacheRisks(workbook),
				slicerTimelineCacheDependencies: collectSlicerTimelineCacheDependencies(workbook),
				unsupportedHeadlessRefresh:
					'Ascend can preserve and edit pivot, slicer, and timeline metadata, but cannot refresh pivot caches or recalculate PivotTable output cells headlessly.',
				recommendedInspection: [
					'inspect --detail pivots',
					'inspect --detail slicers',
					'inspect --detail timelines',
					'pivotRefreshPlans',
					'packageGraphAudit.issues',
					'postWrite.packageGraphAudit',
				],
				analyticsPivotRefreshRisk: visualWriteRisk.analyticsPivotRefreshRisk,
			},
		})
	}
	if (visualWriteRisk.relatedOperations.length > 0) {
		const relatedPartPaths = uniqueStrings(
			visualWriteRisk.relatedOperations.flatMap((operation) => operation.partPaths),
		)
		diagnostics.push({
			code: 'visual-edit-preservation-risk',
			severity: visualWriteRisk.relatedOperations.some((operation) => operation.matchCount !== 1)
				? 'warning'
				: 'info',
			message: `${visualWriteRisk.relatedOperations.length} planned drawing, image, or chart operation(s) should be checked against visualInventory selectors before commit.`,
			suggestedAction:
				'Use targetPath/relId/imageIndex for image replacement, drawingPartPath/id/drawingObjectIndex for drawing text, and partPath/chartIndex/seriesIndex for chart source edits; verify postWrite.packageGraphAudit after write.',
			...(relatedPartPaths.length > 0 ? { partPaths: relatedPartPaths } : {}),
			...(relatedPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, relatedPartPaths) }
				: {}),
			preservationPolicy: 'preserve-exact',
			details: {
				relatedOperations: visualWriteRisk.relatedOperations,
			},
		})
	}
	if (visualWriteRisk.drawingmlVmlDrift.length > 0) {
		const driftPartPaths = uniqueStrings(
			visualWriteRisk.drawingmlVmlDrift.flatMap((entry) => entry.partPaths),
		)
		diagnostics.push({
			code: 'drawingml-vml-drift-risk',
			severity: 'warning',
			message: `${visualWriteRisk.drawingmlVmlDrift.length} visual operation(s) target sheet(s) that contain both DrawingML and VML drawing state.`,
			suggestedAction:
				'Select the exact drawing object or image from visualInventory and verify DrawingML plus VML/comment layout metadata after write; image and DrawingML text operations do not imply legacy VML shape updates.',
			...(driftPartPaths.length > 0 ? { partPaths: driftPartPaths } : {}),
			...(driftPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, driftPartPaths) }
				: {}),
			featureFamily: 'preservedDrawing,preservedVml',
			preservationPolicy: 'preserve-exact',
			details: {
				drift: visualWriteRisk.drawingmlVmlDrift,
			},
		})
	}
	if (visualWriteRisk.chartSourceRefDrift.length > 0 || chartIntegrityIssues.length > 0) {
		const driftPartPaths = uniqueStrings([
			...visualWriteRisk.chartSourceRefDrift.flatMap((entry) => entry.partPaths),
			...chartIntegrityIssues.flatMap((issue) =>
				checkIssuePackagePartPaths(issue).filter(isVisualPackagePartPath),
			),
		])
		diagnostics.push({
			code: 'chart-source-ref-drift-risk',
			severity: 'warning',
			message: chartSourceRefRiskMessage(
				visualWriteRisk.chartSourceRefDrift.length,
				chartIntegrityIssues.length,
			),
			suggestedAction:
				'Use setChartSeriesSource with explicit partPath, chartIndex, and seriesIndex when a structural edit should move chart name/category/value refs; repair verify chart source issues before write and verify chart series refs after write.',
			...(driftPartPaths.length > 0 ? { partPaths: driftPartPaths } : {}),
			...(driftPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, driftPartPaths) }
				: {}),
			featureFamily: 'preservedChart',
			preservationPolicy: 'preserve-exact',
			details: {
				chartSourceRefDrift: visualWriteRisk.chartSourceRefDrift,
				verifyIssues: chartIntegrityIssues,
				chartSourceRefs: collectChartSourceRefSummary(workbook),
			},
		})
	}
	if (visualWriteRisk.analyticsPivotRefreshRisk.length > 0) {
		const refreshPartPaths = uniqueStrings(
			visualWriteRisk.analyticsPivotRefreshRisk.flatMap((entry) => entry.partPaths),
		)
		diagnostics.push({
			code: 'analytics-pivot-refresh-risk',
			severity: 'warning',
			message: `${visualWriteRisk.analyticsPivotRefreshRisk.length} planned pivot, slicer, or timeline operation(s) update analytical filter/cache state; linked pivot output can remain stale until refreshed.`,
			suggestedAction:
				'Ascend cannot refresh pivots headlessly; after write, open the workbook in Excel or another pivot-aware engine, refresh affected pivots/caches, and recalculate formulas before trusting analytical output cells.',
			...(refreshPartPaths.length > 0 ? { partPaths: refreshPartPaths } : {}),
			...(refreshPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, refreshPartPaths) }
				: {}),
			featureFamily: 'preservedPivot,preservedSlicer,preservedTimeline',
			preservationPolicy: 'preserve-exact',
			details: {
				operationScoped: true,
				analyticsPivotRefreshRisk: visualWriteRisk.analyticsPivotRefreshRisk,
				unsupportedHeadlessRefresh:
					'Pivot cache refresh and PivotTable output recalculation require Excel or another pivot-aware engine.',
				recommendedInspection: [
					'inspect --detail pivots',
					'inspect --detail slicers',
					'inspect --detail timelines',
					'pivotRefreshPlans',
					'postWrite.packageGraphAudit',
				],
			},
		})
	}
	const tableIntegrityIssues = collectTableIntegrityIssues(checkIssues)
	const tableLocations = collectTableLocations(workbook)
	const tablePackageIssuePartPaths = packageGraphAudit.issues.flatMap((issue) =>
		isTablePackageGraphIssue(issue) ? packageGraphIssuePackagePaths(issue) : [],
	)
	const tablePartPaths = collectTablePartPaths(
		packageGraph.parts,
		tableLocations,
		tableIntegrityIssues,
		tablePackageIssuePartPaths,
	)
	const tablePackageGraphIssues = packageGraphIssuesForParts(packageGraphAudit, tablePartPaths, [
		'preservedTable',
		'preservedQueryTable',
	])
	if (
		tableLocations.length > 0 ||
		tableIntegrityIssues.length > 0 ||
		tablePackageGraphIssues.length > 0
	) {
		const relatedOperations = collectTableRelatedOperations(workbook, operations)
		diagnostics.push({
			code: 'table-preservation-risk',
			severity:
				tableIntegrityIssues.length > 0 ||
				tablePackageGraphIssues.some((issue) => issue.severity === 'error')
					? 'warning'
					: 'info',
			message:
				tablePackageGraphIssues.length > 0 || tableIntegrityIssues.length > 0
					? `${tablePackageGraphIssues.length + tableIntegrityIssues.length} table or queryTable integrity issue(s) require inspection before write.`
					: `${tableLocations.length} table location(s) depend on worksheet table relationships and table/queryTable sidecar preservation.`,
			suggestedAction:
				'Use inspectSheet(sheet).tables for table part paths, queryTable sidecars, filters, sort state, style metadata, and topology preconditions; verify postWrite.check plus postWrite.packageGraphAudit after table or structural edits.',
			locations: tableLocations.map((entry) => entry.location),
			...(tablePartPaths.length > 0 ? { partPaths: tablePartPaths } : {}),
			...(tablePartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, tablePartPaths) }
				: {}),
			featureFamily: 'preservedTable,preservedQueryTable',
			preservationPolicy: 'preserve-exact',
			details: {
				tables: tableLocations,
				relatedOperations,
				packageGraphIssues: tablePackageGraphIssues,
				verifyIssues: tableIntegrityIssues,
				preconditions: [
					'workbook-unique table names',
					'workbook-unique table ids',
					'non-overlapping same-sheet table ranges',
					'table/queryTable relationship binding preserved',
				],
			},
		})
	}
	const commentIntegrityIssues = collectCommentIntegrityIssues(checkIssues)
	const legacyCommentLocations = collectLegacyCommentLocations(workbook)
	const legacyCommentPackageIssuePartPaths = packageGraphAudit.issues.flatMap((issue) =>
		isLegacyCommentPackageGraphIssue(issue) ? packageGraphIssuePackagePaths(issue) : [],
	)
	const legacyCommentPartPaths = collectLegacyCommentPartPaths(
		packageGraph.parts,
		legacyCommentLocations,
		commentIntegrityIssues.legacy,
		legacyCommentPackageIssuePartPaths,
	)
	const legacyCommentPackageGraphIssues = packageGraphIssuesForParts(
		packageGraphAudit,
		legacyCommentPartPaths,
		['preservedComments', 'preservedVml'],
	)
	if (
		legacyCommentLocations.length > 0 ||
		commentIntegrityIssues.legacy.length > 0 ||
		legacyCommentPackageGraphIssues.length > 0
	) {
		const relatedOperations = collectLegacyCommentRelatedOperations(
			legacyCommentLocations,
			operations,
		)
		diagnostics.push({
			code: 'legacy-comment-preservation-risk',
			severity: 'warning',
			message:
				legacyCommentPackageGraphIssues.length > 0 || commentIntegrityIssues.legacy.length > 0
					? `${legacyCommentPackageGraphIssues.length + commentIntegrityIssues.legacy.length} legacy comment package or verify issue(s) require inspection before write.`
					: legacyCommentLocations.length > 0
						? `${legacyCommentLocations.length} legacy comment location(s) depend on comments XML${legacyCommentPartPaths.some((path) => path.includes('/drawings/')) ? ' and VML drawing' : ''} package preservation.`
						: `${commentIntegrityIssues.legacy.length} legacy comment verify issue(s) require inspection before write.`,
			suggestedAction:
				'For text-only edits, prefer setComment on the existing sheet/ref. Inspect inspectSheet(sheet).comments legacyDrawing/VML metadata before layout changes, then verify postWrite.check and postWrite.packageGraphAudit.',
			locations: legacyCommentLocations.map((entry) => entry.location),
			...(legacyCommentPartPaths.length > 0 ? { partPaths: legacyCommentPartPaths } : {}),
			...(legacyCommentPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, legacyCommentPartPaths) }
				: {}),
			featureFamily: 'preservedComments,preservedVml',
			preservationPolicy: 'preserve-exact',
			details: {
				comments: legacyCommentLocations,
				relatedOperations,
				packageGraphIssues: legacyCommentPackageGraphIssues,
				verifyIssues: commentIntegrityIssues.legacy,
				safeTextEdit: 'setComment',
				structuralEditVerification:
					'For row or column inserts/deletes near comments, inspectSheet(sheet).comments legacyDrawing/VML metadata before and after the write.',
			},
		})
	}
	const threadedCommentLocations = collectThreadedCommentLocations(workbook)
	const threadedCommentPackageIssuePartPaths = packageGraphAudit.issues.flatMap((issue) =>
		isThreadedCommentPackageGraphIssue(issue) ? packageGraphIssuePackagePaths(issue) : [],
	)
	const threadedCommentPartPaths = collectThreadedCommentPartPaths(
		packageGraph.parts,
		threadedCommentLocations,
		commentIntegrityIssues.threaded,
		threadedCommentPackageIssuePartPaths,
	)
	const threadedCommentPackageGraphIssues = packageGraphIssuesForParts(
		packageGraphAudit,
		threadedCommentPartPaths,
		['preservedThreadedComments'],
	)
	if (
		threadedCommentLocations.length > 0 ||
		commentIntegrityIssues.threaded.length > 0 ||
		threadedCommentPackageGraphIssues.length > 0
	) {
		const relatedOperations = collectThreadedCommentRelatedOperations(
			threadedCommentLocations,
			operations,
		)
		diagnostics.push({
			code: 'threaded-comment-preservation-risk',
			severity: 'warning',
			message:
				threadedCommentPackageGraphIssues.length > 0 || commentIntegrityIssues.threaded.length > 0
					? `${threadedCommentPackageGraphIssues.length + commentIntegrityIssues.threaded.length} threaded comment package or verify issue(s) require inspection before write.`
					: threadedCommentLocations.length > 0
						? `${threadedCommentLocations.length} threaded comment location(s) depend on threaded comment ids, parent ids, person ids, and persons sidecar preservation.`
						: `${commentIntegrityIssues.threaded.length} threaded comment verify issue(s) require inspection before write.`,
			suggestedAction:
				'For text-only edits, use setThreadedComment with partPath and threadedCommentId from inspectSheet(sheet).threadedComments. Preserve parentId/personId/persons metadata and verify postWrite.check plus postWrite.packageGraphAudit.',
			locations: threadedCommentLocations.map((entry) => entry.location),
			...(threadedCommentPartPaths.length > 0 ? { partPaths: threadedCommentPartPaths } : {}),
			...(threadedCommentPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, threadedCommentPartPaths) }
				: {}),
			featureFamily: 'preservedThreadedComments',
			preservationPolicy: 'preserve-exact',
			details: {
				threadedComments: threadedCommentLocations,
				relatedOperations,
				packageGraphIssues: threadedCommentPackageGraphIssues,
				verifyIssues: commentIntegrityIssues.threaded,
				safeTextEdit: 'setThreadedComment',
				structuralEditVerification:
					'For row or column inserts/deletes near threaded comments, inspectSheet(sheet).threadedComments ids/person metadata before and after the write.',
			},
		})
	}
	const x14ConditionalFormatExtensionPayloads =
		collectX14ConditionalFormatExtensionPayloads(workbook)
	if (x14ConditionalFormatExtensionPayloads.length > 0) {
		const relatedOperations = collectX14ConditionalFormatSemanticEdits(
			workbook,
			operations,
			x14ConditionalFormatExtensionPayloads,
		)
		const partPaths = uniqueStrings(
			x14ConditionalFormatExtensionPayloads.flatMap((entry) =>
				entry.sheetPartPath ? [entry.sheetPartPath] : [],
			),
		)
		diagnostics.push({
			code: 'conditional-format-extension-preservation',
			severity: relatedOperations.length > 0 ? 'warning' : 'info',
			message:
				relatedOperations.length > 0
					? x14ConditionalFormatRiskMessage(
							x14ConditionalFormatExtensionPayloads.length,
							relatedOperations,
						)
					: `${x14ConditionalFormatExtensionPayloads.length} x14 conditional-format extension payload(s) will be preserved as opaque worksheet XML; no conditional-format semantic edit is planned.`,
			suggestedAction: x14ConditionalFormatSuggestedAction(relatedOperations.length > 0),
			...(partPaths.length > 0 ? { partPaths } : {}),
			featureFamily: 'x14ConditionalFormatting',
			preservationPolicy: 'generated',
			details: {
				provenance: 'worksheet-extLst',
				preservationMode: 'opaque-payload-preserved-in-generated-worksheet-xml',
				semanticEditRisk: relatedOperations.length > 0,
				relatedOperations,
				x14ConditionalFormats: x14ConditionalFormatExtensionPayloads,
			},
		})
	}
	const x14DataValidationExtensionPayloads = collectX14DataValidationExtensionPayloads(workbook)
	if (x14DataValidationExtensionPayloads.length > 0) {
		const relatedOperations = collectX14DataValidationSemanticEdits(
			workbook,
			operations,
			x14DataValidationExtensionPayloads,
		)
		const partPaths = uniqueStrings(
			x14DataValidationExtensionPayloads.flatMap((entry) =>
				entry.sheetPartPath ? [entry.sheetPartPath] : [],
			),
		)
		diagnostics.push({
			code: 'data-validation-extension-preservation',
			severity: relatedOperations.length > 0 ? 'warning' : 'info',
			message:
				relatedOperations.length > 0
					? x14DataValidationRiskMessage(
							x14DataValidationExtensionPayloads.length,
							relatedOperations,
						)
					: `${x14DataValidationExtensionPayloads.length} x14 data-validation extension payload(s) will be preserved as opaque worksheet XML; no data-validation semantic edit is planned.`,
			suggestedAction: x14DataValidationSuggestedAction(relatedOperations.length > 0),
			...(partPaths.length > 0 ? { partPaths } : {}),
			featureFamily: 'x14DataValidation',
			preservationPolicy: 'generated',
			details: {
				provenance: 'worksheet-extLst',
				preservationMode: 'opaque-payload-preserved-in-generated-worksheet-xml',
				semanticEditRisk: relatedOperations.length > 0,
				relatedOperations,
				x14DataValidations: x14DataValidationExtensionPayloads,
			},
		})
	}
	const externalLinkRisk = buildExternalLinkRisk(workbook, operations, externalReferences)
	const externalLinkIntegrityIssues = collectExternalLinkIntegrityIssues(checkIssues)
	const externalLinkPackageIssuePartPaths = packageGraphAudit.issues.flatMap((issue) =>
		isExternalLinkPackageGraphIssue(issue) ? packageGraphIssuePackagePaths(issue) : [],
	)
	const externalLinkPartPaths = collectExternalLinkPartPaths(
		packageGraph.parts,
		externalReferences,
		externalLinkIntegrityIssues,
		externalLinkPackageIssuePartPaths,
	)
	const externalLinkPackageGraphIssues = packageGraphIssuesForParts(
		packageGraphAudit,
		externalLinkPartPaths,
		['preservedExternalLink'],
	)
	if (externalReferences.length > 0) {
		const externalReferencePartPaths = uniqueStrings(
			externalReferences.map((entry) => entry.partPath),
		)
		const relatedOperations = externalLinkRisk.relatedOperations.filter(
			(operation) => operation.externalReference?.partPath !== undefined,
		)
		diagnostics.push({
			code: 'external-link-dependency',
			severity: relatedOperations.length > 0 ? 'warning' : 'info',
			message:
				relatedOperations.length > 0
					? `${externalReferences.length} external workbook dependency link(s) are preserved by relationship-id metadata and ${relatedOperations.length} planned operation(s) reference them.`
					: `${externalReferences.length} external workbook dependency link(s) are preserved by relationship-id metadata.`,
			suggestedAction:
				relatedOperations.length > 0
					? 'Inspect details.externalLinks and details.relatedOperations before committing formula or defined-name edits that depend on external workbooks.'
					: 'Inspect externalReferenceDetails and externalReferenceUsages before rewriting workbook paths or trusting imported formula dependencies.',
			partPaths: externalReferencePartPaths,
			packageParts: packagePartDetails(packagePartByPath, externalReferencePartPaths),
			featureFamily: 'preservedExternalLink',
			preservationPolicy: 'preserve-exact',
			details: {
				externalLinks: externalLinkRisk.linkGroups.filter(
					(group) => group.externalReference !== undefined,
				),
				relatedOperations,
				operationScoped: relatedOperations.length > 0,
			},
		})
	}
	if (externalReferences.length === 0 && externalLinkRisk.relatedOperations.length > 0) {
		diagnostics.push({
			code: 'external-link-dependency',
			severity: 'warning',
			message: `${externalLinkRisk.relatedOperations.length} planned operation(s) contain external workbook formula or defined-name references without imported external-link package metadata.`,
			suggestedAction:
				'Keep the formula or defined name text symbolic, inspect externalReferenceUsages after apply, and add external-link package metadata before relying on workbook path rebinding.',
			featureFamily: 'preservedExternalLink',
			preservationPolicy: 'preserve-exact',
			details: {
				externalLinks: externalLinkRisk.linkGroups,
				relatedOperations: externalLinkRisk.relatedOperations,
				operationScoped: true,
			},
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
		const bindingIssueGroups = externalLinkRisk.linkGroups.filter((group) =>
			group.externalReference
				? externalReferenceBindingIssues.some(
						(entry) => entry.partPath === group.externalReference?.partPath,
					)
				: false,
		)
		const relatedOperations = bindingIssueGroups.flatMap((group) => group.relatedOperations)
		diagnostics.push({
			code: 'external-link-binding-risk',
			severity: relatedOperations.length > 0 ? 'warning' : 'info',
			message: `${externalReferenceBindingIssues.length} external workbook dependency link(s) have ambiguous package binding (${fallbackCount} fallback, ${missingCount} missing).`,
			suggestedAction:
				relatedOperations.length > 0
					? 'Use rewriteExternalLink with a stable partPath/linkRelId selector before or with these operation-scoped formula and defined-name edits.'
					: 'Use rewriteExternalLink with a stable partPath/linkRelId selector to repair or intentionally preserve the externalBook relationship binding before editing formulas or names that depend on it.',
			partPaths: issuePartPaths,
			packageParts: packagePartDetails(packagePartByPath, issuePartPaths),
			featureFamily: 'preservedExternalLink',
			preservationPolicy: 'preserve-exact',
			details: {
				bindingIssueCounts: { fallback: fallbackCount, missing: missingCount },
				externalLinks: bindingIssueGroups,
				relatedOperations,
				operationScoped: relatedOperations.length > 0,
				rewriteExternalLinkRecommendations: bindingIssueGroups.map(
					externalLinkRewriteRecommendation,
				),
			},
		})
	}
	if (externalLinkIntegrityIssues.length > 0 || externalLinkPackageGraphIssues.length > 0) {
		diagnostics.push({
			code: 'external-link-package-risk',
			severity: 'warning',
			message: `${externalLinkIntegrityIssues.length + externalLinkPackageGraphIssues.length} external-link package issue(s) require inspection before write.`,
			suggestedAction:
				'Inspect externalReferenceDetails, external-link package relationships, externalBook r:id bindings, and packageGraphAudit before committing; use rewriteExternalLink only with a stable partPath/linkRelId selector.',
			...(externalLinkPartPaths.length > 0 ? { partPaths: externalLinkPartPaths } : {}),
			...(externalLinkPartPaths.length > 0
				? { packageParts: packagePartDetails(packagePartByPath, externalLinkPartPaths) }
				: {}),
			featureFamily: 'preservedExternalLink',
			preservationPolicy: 'preserve-exact',
			details: {
				externalLinks: externalLinkRisk.linkGroups,
				packageGraphIssues: externalLinkPackageGraphIssues,
				verifyIssues: externalLinkIntegrityIssues,
				rewriteExternalLinkRecommendations: externalLinkRisk.linkGroups
					.filter((group) => group.externalReference !== undefined)
					.map(externalLinkRewriteRecommendation),
			},
		})
	}
	const genericPackageGraphIssues = packageGraphAudit.issues.filter(
		(issue) => !isRoutedPackageGraphIssue(issue),
	)
	if (genericPackageGraphIssues.length > 0) {
		const issuePartPaths = uniqueStrings(
			genericPackageGraphIssues.flatMap((issue) => packageGraphIssuePackagePaths(issue)),
		)
		const packageParts = packagePartDetails(packagePartByPath, issuePartPaths)
		diagnostics.push({
			code: 'package-graph-audit-issue',
			severity: 'warning',
			message: `${genericPackageGraphIssues.length} package graph issue(s) were found before write planning.`,
			suggestedAction:
				'Inspect packageGraphAudit.issues and resolve or explicitly accept package graph risk before committing.',
			partPaths: issuePartPaths,
			...(packageParts.length > 0 ? { packageParts } : {}),
			details: { packageGraphIssues: genericPackageGraphIssues },
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
			chartSourceIntegrityIssues: chartIntegrityIssues.length,
			legacyCommentLocations: legacyCommentLocations.length,
			threadedCommentLocations: threadedCommentLocations.length,
			commentIntegrityIssues:
				commentIntegrityIssues.legacy.length + commentIntegrityIssues.threaded.length,
			x14ConditionalFormatExtensionPayloads: x14ConditionalFormatExtensionPayloads.length,
			x14DataValidationExtensionPayloads: x14DataValidationExtensionPayloads.length,
			calcChainPolicy,
		},
	}
}

interface VisualWriteRisk {
	readonly relatedOperations: readonly VisualRelatedOperation[]
	readonly drawingmlVmlDrift: readonly VisualDrawingDriftRisk[]
	readonly chartSourceRefDrift: readonly ChartSourceRefDriftRisk[]
	readonly analyticsPivotRefreshRisk: readonly AnalyticsPivotRefreshRisk[]
}

interface VisualRelatedOperation {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly sheetName?: string | undefined
	readonly targetKind: 'image' | 'drawingText' | 'chartSource'
	readonly selector: Readonly<Record<string, unknown>>
	readonly matchCount: number
	readonly matches: readonly Record<string, unknown>[]
	readonly partPaths: readonly string[]
	readonly recommendation: string
}

interface VisualDrawingDriftRisk {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly sheetName: string
	readonly partPaths: readonly string[]
	readonly drawingMlObjectCount: number
	readonly vmlObjectCount: number
	readonly imageCount: number
	readonly recommendation: string
}

interface ChartSourceRefDriftRisk {
	readonly chartPartPath: string
	readonly sheetName?: string | undefined
	readonly chartIndex: number
	readonly seriesIndex: number
	readonly sourceRefs: readonly ChartSourceRefAudit[]
	readonly relatedOperations: readonly ChartSourceRelatedOperation[]
	readonly partPaths: readonly string[]
	readonly recommendation: Readonly<Record<string, unknown>>
}

interface ChartSourceRefAudit {
	readonly sourceKind: 'nameRef' | 'categoryRef' | 'valueRef'
	readonly ref: string
	readonly referencedSheets: readonly string[]
}

interface ChartSourceRelatedOperation {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly sheetName: string
	readonly rangeImpact: 'rows' | 'columns' | 'sheet'
	readonly at?: number | undefined
	readonly count?: number | undefined
	readonly newName?: string | undefined
}

interface AnalyticsPivotRefreshRisk {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly targetKind: 'pivotCache' | 'pivotFieldItem' | 'slicerCacheItem' | 'timelineRange'
	readonly selector: Readonly<Record<string, unknown>>
	readonly matchCount: number
	readonly matches: readonly Record<string, unknown>[]
	readonly linkedPivotTableNames: readonly string[]
	readonly partPaths: readonly string[]
	readonly recommendation: string
}

function buildVisualWriteRisk(
	workbook: Workbook,
	operations: readonly Operation[],
	parts: readonly XlsxPackageGraph['parts'][number][],
): VisualWriteRisk {
	return {
		relatedOperations: collectVisualRelatedOperations(workbook, operations),
		drawingmlVmlDrift: collectDrawingMlVmlDriftRisks(workbook, operations, parts),
		chartSourceRefDrift: collectChartSourceRefDriftRisks(workbook, operations, parts),
		analyticsPivotRefreshRisk: collectAnalyticsPivotRefreshRisks(workbook, operations, parts),
	}
}

function collectVisualRelatedOperations(
	workbook: Workbook,
	operations: readonly Operation[],
): VisualRelatedOperation[] {
	return operations.flatMap((operation, operationIndex) => {
		switch (operation.op) {
			case 'replaceImage':
			case 'deleteImage':
				return [imageReplacementRecommendation(workbook, operation, operationIndex)]
			case 'insertImage':
				return [imageInsertRecommendation(workbook, operation, operationIndex)]
			case 'setDrawingText':
				return [drawingTextRecommendation(workbook, operation, operationIndex)]
			case 'setChartSeriesSource':
				return [chartSourceReplacementRecommendation(workbook, operation, operationIndex)]
			default:
				return []
		}
	})
}

function imageReplacementRecommendation(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'replaceImage' | 'deleteImage' }>,
	operationIndex: number,
): VisualRelatedOperation {
	const sheet = workbook.getSheet(operation.sheet)
	const matches = sheet
		? sheet.imageRefs
				.map((image, index) => ({ image, index }))
				.filter(({ image, index }) => imageSelectorMatches(image, index, operation))
				.map(({ image, index }) => ({
					index,
					drawingPartPath: image.drawingPartPath,
					targetPath: image.targetPath,
					relId: image.relId,
					...(image.name ? { name: image.name } : {}),
				}))
		: []
	return {
		operationIndex,
		op: operation.op,
		sheetName: operation.sheet,
		targetKind: 'image',
		selector: compactRecord({
			targetPath: operation.targetPath,
			relId: operation.relId,
			name: operation.name,
			imageIndex: operation.imageIndex,
		}),
		matchCount: matches.length,
		matches,
		partPaths: uniqueStrings(matches.flatMap((match) => [match.drawingPartPath, match.targetPath])),
		recommendation:
			matches.length === 1
				? `${operation.op} has a unique image selector; keep targetPath or imageIndex in the edit record and verify the drawing relationship plus media bytes after write.`
				: `Use inspect --detail images or visualInventory to choose one image by targetPath, relId, name, or imageIndex before ${operation.op}.`,
	}
}

function imageInsertRecommendation(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'insertImage' }>,
	operationIndex: number,
): VisualRelatedOperation {
	const sheet = workbook.getSheet(operation.sheet)
	const collisions = sheet
		? sheet.imageRefs
				.map((image, index) => ({ image, index }))
				.filter(
					({ image }) =>
						(operation.targetPath !== undefined && image.targetPath === operation.targetPath) ||
						(operation.relId !== undefined && image.relId === operation.relId),
				)
				.map(({ image, index }) => ({
					index,
					drawingPartPath: image.drawingPartPath,
					targetPath: image.targetPath,
					relId: image.relId,
				}))
		: []
	const drawingPartPath =
		operation.drawingPartPath ?? sheet?.imageRefs[0]?.drawingPartPath ?? 'xl/drawings/drawing1.xml'
	return {
		operationIndex,
		op: operation.op,
		sheetName: operation.sheet,
		targetKind: 'image',
		selector: compactRecord({
			drawingPartPath: operation.drawingPartPath,
			targetPath: operation.targetPath,
			relId: operation.relId,
			name: operation.name,
		}),
		matchCount: collisions.length === 0 ? 1 : collisions.length,
		matches: collisions,
		partPaths: uniqueStrings([
			drawingPartPath,
			...(operation.targetPath ? [operation.targetPath] : []),
		]),
		recommendation:
			collisions.length === 0
				? 'insertImage can allocate missing media and relationship ids; provide an anchor when placement matters and verify the generated drawing relationship after write.'
				: 'insertImage selector collides with existing image identity; omit targetPath/relId or provide unique values before commit.',
	}
}

function drawingTextRecommendation(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setDrawingText' }>,
	operationIndex: number,
): VisualRelatedOperation {
	const sheet = workbook.getSheet(operation.sheet)
	const matches = sheet
		? sheet.drawingObjectRefs
				.map((object, index) => ({ object, index }))
				.filter(({ object, index }) => drawingObjectSelectorMatches(object, index, operation))
				.map(({ object, index }) => ({
					index,
					drawingPartPath: object.drawingPartPath,
					kind: object.kind,
					...(object.source ? { source: object.source } : {}),
					...(object.id !== undefined ? { id: object.id } : {}),
					...(object.name ? { name: object.name } : {}),
					editableText: object.text !== undefined,
				}))
		: []
	return {
		operationIndex,
		op: operation.op,
		sheetName: operation.sheet,
		targetKind: 'drawingText',
		selector: compactRecord({
			drawingPartPath: operation.drawingPartPath,
			id: operation.id,
			name: operation.name,
			drawingObjectIndex: operation.drawingObjectIndex,
		}),
		matchCount: matches.length,
		matches,
		partPaths: uniqueStrings(matches.map((match) => match.drawingPartPath)),
		recommendation:
			matches.length === 1 && matches[0]?.editableText === true
				? 'setDrawingText has a unique text-bearing object; preserve drawingPartPath plus id/name or drawingObjectIndex and verify the drawing part after write.'
				: 'Use inspect --detail visuals or visualInventory to select exactly one text-bearing DrawingML/VML object before setDrawingText.',
	}
}

function chartSourceReplacementRecommendation(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setChartSeriesSource' }>,
	operationIndex: number,
): VisualRelatedOperation {
	const matches = matchingCharts(workbook, operation).map((chart) => ({
		partPath: chart.partPath,
		...(chart.sheetName ? { sheetName: chart.sheetName } : {}),
		seriesCount: chart.series.length,
		seriesExists: chart.series[operation.seriesIndex] !== undefined,
		currentSeries: chart.series[operation.seriesIndex],
		replacement: compactRecord({
			nameRef: operation.nameRef,
			categoryRef: operation.categoryRef,
			valueRef: operation.valueRef,
		}),
	}))
	return {
		operationIndex,
		op: operation.op,
		...(operation.sheet ? { sheetName: operation.sheet } : {}),
		targetKind: 'chartSource',
		selector: compactRecord({
			partPath: operation.partPath,
			sheet: operation.sheet,
			chartIndex: operation.chartIndex,
			seriesIndex: operation.seriesIndex,
		}),
		matchCount: matches.length,
		matches,
		partPaths: uniqueStrings(matches.map((match) => match.partPath)),
		recommendation:
			matches.length === 1 && matches[0]?.seriesExists === true
				? 'setChartSeriesSource has a unique chart and existing series; keep partPath/chartIndex/seriesIndex and verify name/category/value refs after write.'
				: 'Use inspect --detail visuals or visualInventory to select one chart by partPath or chartIndex and an existing seriesIndex before editing source refs.',
	}
}

function collectDrawingMlVmlDriftRisks(
	workbook: Workbook,
	operations: readonly Operation[],
	parts: readonly XlsxPackageGraph['parts'][number][],
): VisualDrawingDriftRisk[] {
	return operations.flatMap((operation, operationIndex) => {
		if (!isSheetVisualOperation(operation)) return []
		const sheet = workbook.getSheet(operation.sheet)
		if (!sheet?.drawingRefs.hasDrawing || !sheet.drawingRefs.hasLegacyDrawing) return []
		const drawingMlObjectCount = sheet.drawingObjectRefs.filter(
			(object) => object.source !== 'vml',
		).length
		const vmlObjectCount = sheet.drawingObjectRefs.filter(
			(object) => object.source === 'vml',
		).length
		return [
			{
				operationIndex,
				op: operation.op,
				sheetName: operation.sheet,
				drawingMlObjectCount,
				vmlObjectCount,
				imageCount: sheet.imageRefs.length,
				partPaths: sheetVisualPartPaths(sheet.imageRefs, sheet.drawingObjectRefs, parts),
				recommendation:
					'DrawingML and VML are separate package graphs; verify the selected object source and legacy VML/comment layout after write.',
			},
		]
	})
}

function collectChartSourceRefDriftRisks(
	workbook: Workbook,
	operations: readonly Operation[],
	parts: readonly XlsxPackageGraph['parts'][number][],
): ChartSourceRefDriftRisk[] {
	const risks: ChartSourceRefDriftRisk[] = []
	for (const [chartIndex, chart] of workbook.chartParts.entries()) {
		for (const [seriesIndex, series] of chart.series.entries()) {
			const sourceRefs = chartSourceRefs(workbook, series, chart.sheetName)
			const referencedSheets = new Set(sourceRefs.flatMap((source) => source.referencedSheets))
			const relatedOperations = chartSourceRelatedOperations(operations, referencedSheets)
			if (sourceRefs.length === 0 || relatedOperations.length === 0) continue
			risks.push({
				chartPartPath: chart.partPath,
				...(chart.sheetName ? { sheetName: chart.sheetName } : {}),
				chartIndex,
				seriesIndex,
				sourceRefs,
				relatedOperations,
				partPaths: chartRelatedPartPaths(chart.partPath, parts),
				recommendation: {
					op: 'setChartSeriesSource',
					partPath: chart.partPath,
					chartIndex,
					seriesIndex,
					current: compactRecord({
						nameRef: series.nameRef,
						categoryRef: series.categoryRef,
						valueRef: series.valueRef,
					}),
				},
			})
		}
	}
	return risks
}

function collectAnalyticsPivotRefreshRisks(
	workbook: Workbook,
	operations: readonly Operation[],
	parts: readonly XlsxPackageGraph['parts'][number][],
): AnalyticsPivotRefreshRisk[] {
	return operations.flatMap((operation, operationIndex) => {
		switch (operation.op) {
			case 'setPivotCache':
				return [pivotCacheRefreshRisk(workbook, operation, operationIndex)]
			case 'setPivotFieldItem':
				return [pivotFieldItemRefreshRisk(workbook, operation, operationIndex)]
			case 'setSlicerCacheItem':
				return [slicerCacheItemRefreshRisk(workbook, operation, operationIndex, parts)]
			case 'setTimelineRange':
				return [timelineRangeRefreshRisk(workbook, operation, operationIndex, parts)]
			default:
				return []
		}
	})
}

function pivotCacheRefreshRisk(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setPivotCache' }>,
	operationIndex: number,
): AnalyticsPivotRefreshRisk {
	const matches = workbook.pivotCaches
		.filter((cache) => pivotCacheSelectorMatches(workbook, cache, operation))
		.map((cache) => {
			const linkedPivots = workbook.pivotTables.filter(
				(pivot) => cache.cacheId !== undefined && pivot.cacheId === cache.cacheId,
			)
			return compactRecord({
				pivotCache: cache.partPath,
				cacheId: cache.cacheId,
				sourceSheet: cache.sourceSheet,
				sourceRef: cache.sourceRef,
				nextSourceSheet: operation.sourceSheet,
				nextSourceRef: operation.sourceRef,
				refreshOnLoad: operation.refreshOnLoad ?? cache.refreshOnLoad,
				invalid: operation.invalid ?? cache.invalid,
				saveData: operation.saveData ?? cache.saveData,
				outputState: pivotCacheOutputState(workbook, cache, linkedPivots),
				linkedPivotTableNames: linkedPivotNames(linkedPivots, []),
				pivotTablePartPaths: linkedPivots.map((pivot) => pivot.partPath),
				pivotCacheRecordsPartPath: cache.recordsPartPath,
			})
		})
	const linkedPivotTableNames = uniqueStrings(
		matches.flatMap((match) =>
			Array.isArray(match.linkedPivotTableNames)
				? match.linkedPivotTableNames.filter((name): name is string => typeof name === 'string')
				: [],
		),
	)
	const partPaths = uniqueStrings(
		matches.flatMap((match) =>
			[
				match.pivotCache,
				match.pivotCacheRecordsPartPath,
				...(Array.isArray(match.pivotTablePartPaths) ? match.pivotTablePartPaths : []),
			].filter((path): path is string => typeof path === 'string' && path.length > 0),
		),
	)
	return {
		operationIndex,
		op: operation.op,
		targetKind: 'pivotCache',
		selector: compactRecord({
			cacheId: operation.cacheId,
			partPath: operation.partPath,
			pivotTable: operation.pivotTable,
			sourceSheet: operation.sourceSheet,
			sourceRef: operation.sourceRef,
			refreshOnLoad: operation.refreshOnLoad,
			invalid: operation.invalid,
			saveData: operation.saveData,
		}),
		matchCount: matches.length,
		matches,
		linkedPivotTableNames,
		partPaths,
		recommendation:
			matches.length === 1
				? 'setPivotCache matched one cache; mark invalid/refreshOnLoad when changing source metadata, then refresh linked pivots in Excel or another pivot-aware engine after write.'
				: 'Use inspect --detail pivots to select one pivot cache by cacheId, partPath, or pivotTable before changing source or freshness metadata.',
	}
}

function pivotFieldItemRefreshRisk(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setPivotFieldItem' }>,
	operationIndex: number,
): AnalyticsPivotRefreshRisk {
	const matches = workbook.pivotTables
		.filter((pivot) => pivotTableSelectorMatches(pivot, operation))
		.map((pivot) => {
			const field = pivot.fields[operation.fieldIndex]
			const item = field?.items?.[operation.itemIndex]
			const cache = pivotCacheForTable(workbook, pivot)
			return compactRecord({
				pivotTable: pivot.name,
				sheetName: pivot.sheetName,
				partPath: pivot.partPath,
				cacheId: pivot.cacheId,
				cachePartPath: cache?.partPath,
				cacheRecordsPartPath: cache?.recordsPartPath,
				fieldIndex: operation.fieldIndex,
				fieldName: field?.name,
				fieldExists: field !== undefined,
				itemIndex: operation.itemIndex,
				itemCaption: item?.caption,
				itemExists: item !== undefined,
			})
		})
	const linkedPivotTableNames = uniqueStrings(
		matches.flatMap((match) =>
			typeof match.pivotTable === 'string' && match.pivotTable.length > 0 ? [match.pivotTable] : [],
		),
	)
	const partPaths = uniqueStrings(
		matches.flatMap((match) =>
			[match.partPath, match.cachePartPath, match.cacheRecordsPartPath].filter(
				(path): path is string => typeof path === 'string' && path.length > 0,
			),
		),
	)
	return {
		operationIndex,
		op: operation.op,
		targetKind: 'pivotFieldItem',
		selector: compactRecord({
			pivotTable: operation.pivotTable,
			partPath: operation.partPath,
			sheet: operation.sheet,
			fieldIndex: operation.fieldIndex,
			itemIndex: operation.itemIndex,
		}),
		matchCount: matches.length,
		matches,
		linkedPivotTableNames,
		partPaths,
		recommendation:
			matches.length === 1 && matches[0]?.fieldExists === true && matches[0]?.itemExists === true
				? 'setPivotFieldItem matched one pivot field item; refresh the linked pivot cache/table in Excel or another pivot-aware engine after write.'
				: 'Use inspect --detail pivots to select one pivot table by name or partPath and an existing fieldIndex/itemIndex before editing filter state.',
	}
}

function slicerCacheItemRefreshRisk(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setSlicerCacheItem' }>,
	operationIndex: number,
	parts: readonly XlsxPackageGraph['parts'][number][],
): AnalyticsPivotRefreshRisk {
	const matches = workbook.slicerCaches
		.filter((cache) => slicerCacheSelectorMatches(cache, operation))
		.map((cache) => {
			const linkedPivots = linkedPivotTables(workbook, cache.pivotTableNames, cache.pivotCacheId)
			const linkedCaches = linkedPivotCaches(workbook, linkedPivots, cache.pivotCacheId)
			const slicerPartPaths = workbook.slicers
				.filter((slicer) => slicer.cacheName === cache.name)
				.map((slicer) => slicer.partPath)
			return compactRecord({
				slicerCache: cache.name,
				partPath: cache.partPath,
				sourceName: cache.sourceName,
				pivotCacheId: cache.pivotCacheId,
				item: operation.item,
				itemExists: cache.items?.some((item) => item.index === operation.item) === true,
				linkedPivotTableNames: linkedPivotNames(linkedPivots, cache.pivotTableNames),
				slicerPartPaths,
				pivotTablePartPaths: linkedPivots.map((pivot) => pivot.partPath),
				pivotCachePartPaths: linkedCaches.map((pivotCache) => pivotCache.partPath),
				pivotCacheRecordsPartPaths: linkedCaches.flatMap((pivotCache) =>
					pivotCache.recordsPartPath ? [pivotCache.recordsPartPath] : [],
				),
			})
		})
	const linkedPivotTableNames = uniqueStrings(
		matches.flatMap((match) =>
			Array.isArray(match.linkedPivotTableNames)
				? match.linkedPivotTableNames.filter((name): name is string => typeof name === 'string')
				: [],
		),
	)
	const partPaths = uniqueStrings([
		...matches.flatMap((match) =>
			[
				match.partPath,
				...(Array.isArray(match.slicerPartPaths) ? match.slicerPartPaths : []),
				...(Array.isArray(match.pivotTablePartPaths) ? match.pivotTablePartPaths : []),
				...(Array.isArray(match.pivotCachePartPaths) ? match.pivotCachePartPaths : []),
				...(Array.isArray(match.pivotCacheRecordsPartPaths)
					? match.pivotCacheRecordsPartPaths
					: []),
			].filter((path): path is string => typeof path === 'string' && path.length > 0),
		),
		...operationScopedAnalyticalPartPaths(operation.partPath, parts),
	])
	return {
		operationIndex,
		op: operation.op,
		targetKind: 'slicerCacheItem',
		selector: compactRecord({
			slicerCache: operation.slicerCache,
			partPath: operation.partPath,
			item: operation.item,
		}),
		matchCount: matches.length,
		matches,
		linkedPivotTableNames,
		partPaths,
		recommendation:
			matches.length === 1
				? 'setSlicerCacheItem matched one slicer cache; refresh linked pivot tables/caches in Excel or another pivot-aware engine after write.'
				: 'Use inspect --detail slicers to select one slicer cache by name or partPath before editing item state.',
	}
}

function timelineRangeRefreshRisk(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setTimelineRange' }>,
	operationIndex: number,
	parts: readonly XlsxPackageGraph['parts'][number][],
): AnalyticsPivotRefreshRisk {
	const matches = workbook.timelineCaches
		.filter((cache) => timelineCacheSelectorMatches(cache, operation))
		.map((cache) => {
			const linkedPivots = linkedPivotTables(workbook, cache.pivotTableNames, cache.pivotCacheId)
			const linkedCaches = linkedPivotCaches(workbook, linkedPivots, cache.pivotCacheId)
			const timelinePartPaths = workbook.timelines
				.filter((timeline) => timeline.cacheName === cache.name)
				.map((timeline) => timeline.partPath)
			return compactRecord({
				timelineCache: cache.name,
				partPath: cache.partPath,
				sourceName: cache.sourceName,
				pivotCacheId: cache.pivotCacheId,
				startDate: operation.startDate,
				endDate: operation.endDate,
				linkedPivotTableNames: linkedPivotNames(linkedPivots, cache.pivotTableNames),
				timelinePartPaths,
				pivotTablePartPaths: linkedPivots.map((pivot) => pivot.partPath),
				pivotCachePartPaths: linkedCaches.map((pivotCache) => pivotCache.partPath),
				pivotCacheRecordsPartPaths: linkedCaches.flatMap((pivotCache) =>
					pivotCache.recordsPartPath ? [pivotCache.recordsPartPath] : [],
				),
			})
		})
	const linkedPivotTableNames = uniqueStrings(
		matches.flatMap((match) =>
			Array.isArray(match.linkedPivotTableNames)
				? match.linkedPivotTableNames.filter((name): name is string => typeof name === 'string')
				: [],
		),
	)
	const partPaths = uniqueStrings([
		...matches.flatMap((match) =>
			[
				match.partPath,
				...(Array.isArray(match.timelinePartPaths) ? match.timelinePartPaths : []),
				...(Array.isArray(match.pivotTablePartPaths) ? match.pivotTablePartPaths : []),
				...(Array.isArray(match.pivotCachePartPaths) ? match.pivotCachePartPaths : []),
				...(Array.isArray(match.pivotCacheRecordsPartPaths)
					? match.pivotCacheRecordsPartPaths
					: []),
			].filter((path): path is string => typeof path === 'string' && path.length > 0),
		),
		...operationScopedAnalyticalPartPaths(operation.partPath, parts),
	])
	return {
		operationIndex,
		op: operation.op,
		targetKind: 'timelineRange',
		selector: compactRecord({
			timelineCache: operation.timelineCache,
			partPath: operation.partPath,
			startDate: operation.startDate,
			endDate: operation.endDate,
		}),
		matchCount: matches.length,
		matches,
		linkedPivotTableNames,
		partPaths,
		recommendation:
			matches.length === 1
				? 'setTimelineRange matched one timeline cache; refresh linked pivot tables/caches in Excel or another pivot-aware engine after write.'
				: 'Use inspect --detail timelines to select one timeline cache by name or partPath before editing date range state.',
	}
}

function imageSelectorMatches(
	image: SheetImageRef,
	index: number,
	operation: Extract<Operation, { op: 'replaceImage' | 'deleteImage' }>,
): boolean {
	if (
		operation.imageIndex === undefined &&
		operation.targetPath === undefined &&
		operation.relId === undefined &&
		operation.name === undefined
	) {
		return false
	}
	if (operation.imageIndex !== undefined && index !== operation.imageIndex) return false
	if (operation.targetPath !== undefined && image.targetPath !== operation.targetPath) return false
	if (operation.relId !== undefined && image.relId !== operation.relId) return false
	if (operation.name !== undefined && image.name !== operation.name) return false
	return true
}

function drawingObjectSelectorMatches(
	object: SheetDrawingObjectRef,
	index: number,
	operation: Extract<Operation, { op: 'setDrawingText' }>,
): boolean {
	if (
		operation.drawingObjectIndex === undefined &&
		operation.drawingPartPath === undefined &&
		operation.id === undefined &&
		operation.name === undefined
	) {
		return false
	}
	if (operation.drawingObjectIndex !== undefined && index !== operation.drawingObjectIndex) {
		return false
	}
	if (
		operation.drawingPartPath !== undefined &&
		object.drawingPartPath !== operation.drawingPartPath
	) {
		return false
	}
	if (operation.id !== undefined && object.id !== operation.id) return false
	if (operation.name !== undefined && object.name !== operation.name) return false
	return true
}

function matchingCharts(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'setChartSeriesSource' }>,
): ChartPartInfo[] {
	let candidates = workbook.chartParts
	if (operation.partPath !== undefined) {
		candidates = candidates.filter((chart) => chart.partPath === operation.partPath)
	}
	if (operation.sheet !== undefined) {
		candidates = candidates.filter((chart) => chart.sheetName === operation.sheet)
	}
	if (operation.chartIndex !== undefined) {
		const chart = candidates[operation.chartIndex]
		return chart ? [chart] : []
	}
	return candidates
}

function pivotTableSelectorMatches(
	pivot: Workbook['pivotTables'][number],
	operation: Extract<Operation, { op: 'setPivotFieldItem' }>,
): boolean {
	if (operation.partPath !== undefined && pivot.partPath !== operation.partPath) return false
	if (operation.pivotTable !== undefined && pivot.name !== operation.pivotTable) return false
	if (operation.sheet !== undefined && pivot.sheetName !== operation.sheet) return false
	return true
}

function pivotCacheSelectorMatches(
	workbook: Workbook,
	cache: Workbook['pivotCaches'][number],
	operation: Extract<Operation, { op: 'setPivotCache' }>,
): boolean {
	if (operation.cacheId !== undefined && cache.cacheId !== operation.cacheId) return false
	if (operation.partPath !== undefined && cache.partPath !== operation.partPath) return false
	if (operation.pivotTable !== undefined) {
		return workbook.pivotTables.some(
			(pivot) =>
				pivot.name === operation.pivotTable &&
				pivot.cacheId !== undefined &&
				pivot.cacheId === cache.cacheId,
		)
	}
	return operation.cacheId !== undefined || operation.partPath !== undefined
}

function pivotCacheForTable(
	workbook: Workbook,
	pivot: Workbook['pivotTables'][number],
): Workbook['pivotCaches'][number] | undefined {
	if (pivot.cacheId === undefined) return undefined
	return workbook.pivotCaches.find((cache) => cache.cacheId === pivot.cacheId)
}

function slicerCacheSelectorMatches(
	cache: Workbook['slicerCaches'][number],
	operation: Extract<Operation, { op: 'setSlicerCacheItem' }>,
): boolean {
	if (operation.partPath !== undefined && cache.partPath !== operation.partPath) return false
	if (
		operation.slicerCache !== undefined &&
		cache.name !== operation.slicerCache &&
		cache.partPath !== operation.slicerCache
	) {
		return false
	}
	return true
}

function timelineCacheSelectorMatches(
	cache: Workbook['timelineCaches'][number],
	operation: Extract<Operation, { op: 'setTimelineRange' }>,
): boolean {
	if (operation.partPath !== undefined && cache.partPath !== operation.partPath) return false
	if (
		operation.timelineCache !== undefined &&
		cache.name !== operation.timelineCache &&
		cache.partPath !== operation.timelineCache
	) {
		return false
	}
	return true
}

function linkedPivotTables(
	workbook: Workbook,
	pivotTableNames: readonly string[],
	pivotCacheId: number | undefined,
): Workbook['pivotTables'] {
	const names = new Set(pivotTableNames)
	return workbook.pivotTables.filter((pivot) => {
		if (pivot.name !== undefined && names.has(pivot.name)) return true
		if (pivotCacheId !== undefined && pivot.cacheId === pivotCacheId) return true
		return false
	})
}

function linkedPivotCaches(
	workbook: Workbook,
	pivots: readonly Workbook['pivotTables'][number][],
	pivotCacheId: number | undefined,
): Workbook['pivotCaches'] {
	const cacheIds = new Set(
		[
			pivotCacheId,
			...pivots.flatMap((pivot) => (pivot.cacheId === undefined ? [] : [pivot.cacheId])),
		].filter((cacheId): cacheId is number => cacheId !== undefined),
	)
	return workbook.pivotCaches.filter((cache) => {
		if (cache.cacheId === undefined) return false
		return cacheIds.has(cache.cacheId)
	})
}

function linkedPivotNames(
	pivots: readonly Workbook['pivotTables'][number][],
	pivotTableNames: readonly string[],
): string[] {
	return uniqueStrings([
		...pivotTableNames,
		...pivots.flatMap((pivot) =>
			pivot.name !== undefined && pivot.name.length > 0 ? [pivot.name] : [],
		),
	])
}

function collectPivotCacheRisks(
	workbook: Workbook,
): ReadonlyArray<Readonly<Record<string, unknown>>> {
	return workbook.pivotCaches.map((cache) => {
		const linkedPivots = workbook.pivotTables.filter(
			(pivot) => cache.cacheId !== undefined && pivot.cacheId === cache.cacheId,
		)
		const outputState = pivotCacheOutputState(workbook, cache, linkedPivots)
		return compactRecord({
			partPath: cache.partPath,
			cacheId: cache.cacheId,
			sourceType: cache.sourceType,
			sourceSheet: cache.sourceSheet,
			sourceRef: cache.sourceRef,
			sourceName: cache.sourceName,
			refreshOnLoad: cache.refreshOnLoad,
			enableRefresh: cache.enableRefresh,
			invalid: cache.invalid,
			saveData: cache.saveData,
			recordCount: cache.recordCount,
			recordsPartPath: cache.recordsPartPath,
			outputState,
			requiresExternalRefresh: outputState !== 'cached',
			linkedPivotTables: linkedPivots.map((pivot) =>
				compactRecord({
					name: pivot.name,
					partPath: pivot.partPath,
					sheetName: pivot.sheetName,
					locationRef: pivot.locationRef,
				}),
			),
			warnings: pivotCacheRiskWarnings(cache, outputState),
			recommendedOps:
				outputState === 'cached'
					? []
					: [
							compactRecord({
								op: 'setPivotCache',
								partPath: cache.partPath,
								cacheId: cache.cacheId,
								refreshOnLoad: true,
								invalid: true,
								saveData: false,
							}),
						],
		})
	})
}

function collectSlicerTimelineCacheDependencies(
	workbook: Workbook,
): ReadonlyArray<Readonly<Record<string, unknown>>> {
	return [
		...workbook.slicerCaches.map((cache) => {
			const linkedPivots = linkedPivotTables(workbook, cache.pivotTableNames, cache.pivotCacheId)
			const linkedCaches = linkedPivotCaches(workbook, linkedPivots, cache.pivotCacheId)
			return compactRecord({
				kind: 'slicerCache',
				name: cache.name,
				partPath: cache.partPath,
				sourceName: cache.sourceName,
				pivotCacheId: cache.pivotCacheId,
				pivotTableNames: linkedPivotNames(linkedPivots, cache.pivotTableNames),
				pivotTablePartPaths: linkedPivots.map((pivot) => pivot.partPath),
				pivotCachePartPaths: linkedCaches.map((pivotCache) => pivotCache.partPath),
				slicerPartPaths: workbook.slicers
					.filter((slicer) => slicer.cacheName === cache.name)
					.map((slicer) => slicer.partPath),
			})
		}),
		...workbook.timelineCaches.map((cache) => {
			const linkedPivots = linkedPivotTables(workbook, cache.pivotTableNames, cache.pivotCacheId)
			const linkedCaches = linkedPivotCaches(workbook, linkedPivots, cache.pivotCacheId)
			return compactRecord({
				kind: 'timelineCache',
				name: cache.name,
				partPath: cache.partPath,
				sourceName: cache.sourceName,
				pivotCacheId: cache.pivotCacheId,
				pivotTableNames: linkedPivotNames(linkedPivots, cache.pivotTableNames),
				pivotTablePartPaths: linkedPivots.map((pivot) => pivot.partPath),
				pivotCachePartPaths: linkedCaches.map((pivotCache) => pivotCache.partPath),
				timelinePartPaths: workbook.timelines
					.filter((timeline) => timeline.cacheName === cache.name)
					.map((timeline) => timeline.partPath),
				state: cache.state,
			})
		}),
	]
}

function pivotCacheOutputState(
	workbook: Workbook,
	cache: Workbook['pivotCaches'][number],
	linkedPivots: readonly Workbook['pivotTables'][number][],
): 'cached' | 'stale' | 'refresh-on-open' | 'not-saved' | 'unknown' {
	if (cache.invalid) return 'stale'
	if (cache.refreshOnLoad) return 'refresh-on-open'
	if (cache.saveData === false) return 'not-saved'
	if (linkedPivots.some((pivot) => pivotSavedOutputState(workbook, pivot) === 'not-saved')) {
		return 'not-saved'
	}
	if (cache.recordsPartPath === undefined && cache.recordCount === undefined) return 'unknown'
	if (linkedPivots.some((pivot) => pivotSavedOutputState(workbook, pivot) === 'unknown')) {
		return 'unknown'
	}
	return 'cached'
}

function pivotSavedOutputState(
	workbook: Workbook,
	pivot: Workbook['pivotTables'][number],
): 'cached' | 'not-saved' | 'unknown' {
	if (!pivot.locationRef) return 'unknown'
	let bounds: RangeRef
	try {
		bounds = parseRange(pivot.locationRef)
	} catch {
		return 'unknown'
	}
	const sheet = workbook.sheets.find((entry) => entry.name === pivot.sheetName)
	if (!sheet) return 'unknown'
	for (let row = bounds.start.row; row <= bounds.end.row; row++) {
		for (let col = bounds.start.col; col <= bounds.end.col; col++) {
			const kind = (sheet.cells.get(row, col)?.value as { kind?: string } | undefined)?.kind
			if (kind !== undefined && kind !== 'empty') return 'cached'
		}
	}
	return 'not-saved'
}

function pivotCacheRiskWarnings(
	cache: Workbook['pivotCaches'][number],
	outputState: 'cached' | 'stale' | 'refresh-on-open' | 'not-saved' | 'unknown',
): string[] {
	const warnings: string[] = []
	if (outputState === 'stale') warnings.push('Pivot cache is marked invalid; output is stale.')
	if (outputState === 'refresh-on-open') {
		warnings.push('Pivot cache requests refresh on open; saved output may change.')
	}
	if (outputState === 'not-saved') {
		warnings.push('Pivot cache records or PivotTable output cells are not fully saved.')
	}
	if (outputState === 'unknown') warnings.push('Pivot cache/output freshness is unknown.')
	if (cache.sourceSheet !== undefined || cache.sourceRef !== undefined) {
		warnings.push(
			'Cache source metadata is inspectable, but output is not recalculated headlessly.',
		)
	}
	return warnings
}

function operationScopedAnalyticalPartPaths(
	partPath: string | undefined,
	parts: readonly XlsxPackageGraph['parts'][number][],
): string[] {
	if (partPath === undefined) return []
	return parts.some((part) => part.path === partPath && isAnalyticalSidecar(part.featureFamily))
		? [partPath]
		: []
}

function isSheetVisualOperation(
	operation: Operation,
): operation is Extract<
	Operation,
	{ op: 'replaceImage' | 'insertImage' | 'deleteImage' | 'setDrawingText' }
> {
	return (
		operation.op === 'replaceImage' ||
		operation.op === 'insertImage' ||
		operation.op === 'deleteImage' ||
		operation.op === 'setDrawingText'
	)
}

function sheetVisualPartPaths(
	imageRefs: readonly SheetImageRef[],
	drawingObjectRefs: readonly SheetDrawingObjectRef[],
	parts: readonly XlsxPackageGraph['parts'][number][],
): string[] {
	const drawingPartPaths = uniqueStrings([
		...imageRefs.map((image) => image.drawingPartPath),
		...drawingObjectRefs.map((object) => object.drawingPartPath),
	])
	const mediaPartPaths = imageRefs.map((image) => image.targetPath)
	const vmlPartPaths = parts
		.filter((part) => part.featureFamily === 'preservedVml')
		.map((part) => part.path)
	return uniqueStrings([...drawingPartPaths, ...mediaPartPaths, ...vmlPartPaths])
}

function collectChartSourceRefSummary(workbook: Workbook): readonly Record<string, unknown>[] {
	return workbook.chartParts.map((chart, chartIndex) => ({
		partPath: chart.partPath,
		...(chart.sheetName ? { sheetName: chart.sheetName } : {}),
		...(chart.chartType ? { chartType: chart.chartType } : {}),
		...(chart.title ? { title: chart.title } : {}),
		chartIndex,
		series: chart.series.map((series, seriesIndex) => ({
			seriesIndex,
			sourceRefs: chartSourceRefs(workbook, series, chart.sheetName),
		})),
	}))
}

function collectDrawingModelSummary(
	workbook: Workbook,
	parts: readonly XlsxPackageGraph['parts'][number][],
): Readonly<Record<string, unknown>> {
	return {
		sheets: workbook.sheets
			.filter(
				(sheet) =>
					sheet.drawingRefs.hasDrawing ||
					sheet.drawingRefs.hasLegacyDrawing ||
					sheet.imageRefs.length > 0 ||
					sheet.drawingObjectRefs.length > 0,
			)
			.map((sheet) => {
				const drawingMlObjectRefs = sheet.drawingObjectRefs.filter(
					(object) => object.source !== 'vml',
				)
				const vmlObjectRefs = sheet.drawingObjectRefs.filter((object) => object.source === 'vml')
				return {
					sheetName: sheet.name,
					hasDrawingMl: sheet.drawingRefs.hasDrawing,
					hasVml: sheet.drawingRefs.hasLegacyDrawing,
					imageCount: sheet.imageRefs.length,
					drawingMlObjectCount: drawingMlObjectRefs.length,
					vmlObjectCount: vmlObjectRefs.length,
					drawingMlPartPaths: uniqueStrings([
						...sheet.imageRefs.map((image) => image.drawingPartPath),
						...drawingMlObjectRefs.map((object) => object.drawingPartPath),
					]),
					vmlPartPaths: uniqueStrings(vmlObjectRefs.map((object) => object.drawingPartPath)),
				}
			}),
		packagePartCounts: {
			drawingMl: parts.filter((part) => part.featureFamily === 'preservedDrawing').length,
			media: parts.filter((part) => part.featureFamily === 'preservedMedia').length,
			chart: parts.filter((part) => part.featureFamily === 'preservedChart').length,
			chartSidecar: parts.filter(
				(part) =>
					part.featureFamily === 'preservedChartStyle' ||
					part.featureFamily === 'preservedChartColor',
			).length,
			vml: parts.filter((part) => part.featureFamily === 'preservedVml').length,
		},
		distinction:
			'DrawingML drawings, images, and charts use drawing/chart/media relationships; VML drawings are separate legacy shape/comment layout parts.',
	}
}

function chartSourceRefs(
	workbook: Workbook,
	series: ChartPartInfo['series'][number],
	defaultSheetName?: string,
): ChartSourceRefAudit[] {
	return [
		chartSourceRef(workbook, 'nameRef', series.nameRef, defaultSheetName),
		chartSourceRef(workbook, 'categoryRef', series.categoryRef, defaultSheetName),
		chartSourceRef(workbook, 'valueRef', series.valueRef, defaultSheetName),
	].filter((entry): entry is ChartSourceRefAudit => entry !== undefined)
}

function chartSourceRef(
	workbook: Workbook,
	sourceKind: ChartSourceRefAudit['sourceKind'],
	ref: string | undefined,
	defaultSheetName?: string,
): ChartSourceRefAudit | undefined {
	if (!ref) return undefined
	return {
		sourceKind,
		ref,
		referencedSheets: chartSourceReferencedSheets(workbook, ref, defaultSheetName),
	}
}

function chartSourceReferencedSheets(
	workbook: Workbook,
	ref: string,
	defaultSheetName?: string,
): string[] {
	const parsed = parseFormula(normalizeFormulaInput(ref))
	if (parsed.ok) {
		const sheets = flattenFormulaReferences(collectFormulaReferences(parsed.value)).flatMap(
			(reference) => {
				if (reference.kind === 'structured') {
					return chartStructuredReferenceSheetNames(workbook, reference, defaultSheetName)
				}
				if (reference.scope?.kind === 'sheet') return [reference.scope.sheet]
				if (reference.scope?.kind === 'sheetSpan') {
					return [reference.scope.startSheet, reference.scope.endSheet]
				}
				if (reference.scope?.kind === 'local' && defaultSheetName) return [defaultSheetName]
				return []
			},
		)
		if (sheets.length > 0) return uniqueStrings(sheets)
	}
	const regexMatch = /^'([^']|'')+'!|^([^'!:]+)!/.exec(ref)
	if (regexMatch?.[0]) {
		return [regexMatch[0].slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'")]
	}
	return defaultSheetName ? [defaultSheetName] : []
}

function chartStructuredReferenceSheetNames(
	workbook: Workbook,
	reference: Extract<FormulaReferenceInfo, { kind: 'structured' }>,
	defaultSheetName?: string,
): string[] {
	const tableName = reference.table.trim()
	if (tableName.length === 0) return defaultSheetName ? [defaultSheetName] : []
	const normalizedTableName = tableName.toLowerCase()
	return uniqueStrings(
		workbook.sheets.flatMap((sheet) =>
			sheet.tables.some((table) => table.name.toLowerCase() === normalizedTableName)
				? [sheet.name]
				: [],
		),
	)
}

function chartSourceRelatedOperations(
	operations: readonly Operation[],
	referencedSheets: ReadonlySet<string>,
): ChartSourceRelatedOperation[] {
	const related: ChartSourceRelatedOperation[] = []
	operations.forEach((operation, operationIndex) => {
		switch (operation.op) {
			case 'insertRows':
			case 'deleteRows':
				if (!referencedSheets.has(operation.sheet)) return
				related.push({
					operationIndex,
					op: operation.op,
					sheetName: operation.sheet,
					rangeImpact: 'rows',
					at: operation.at,
					count: operation.count,
				})
				break
			case 'insertCols':
			case 'deleteCols':
				if (!referencedSheets.has(operation.sheet)) return
				related.push({
					operationIndex,
					op: operation.op,
					sheetName: operation.sheet,
					rangeImpact: 'columns',
					at: operation.at,
					count: operation.count,
				})
				break
			case 'renameSheet':
				if (!referencedSheets.has(operation.sheet)) return
				related.push({
					operationIndex,
					op: operation.op,
					sheetName: operation.sheet,
					rangeImpact: 'sheet',
					newName: operation.newName,
				})
				break
			case 'deleteSheet':
				if (!referencedSheets.has(operation.sheet)) return
				related.push({
					operationIndex,
					op: operation.op,
					sheetName: operation.sheet,
					rangeImpact: 'sheet',
				})
				break
		}
	})
	return related
}

function chartRelatedPartPaths(
	chartPartPath: string,
	parts: readonly XlsxPackageGraph['parts'][number][],
): string[] {
	return uniqueStrings(
		parts
			.filter(
				(part) =>
					part.path === chartPartPath ||
					part.featureFamily === 'preservedChartStyle' ||
					part.featureFamily === 'preservedChartColor',
			)
			.map((part) => part.path),
	)
}

function compactRecord(
	value: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
	return Object.fromEntries(
		Object.entries(value).filter(([, entryValue]) => entryValue !== undefined),
	)
}

interface ExternalLinkRisk {
	readonly linkGroups: readonly ExternalLinkRiskGroup[]
	readonly relatedOperations: readonly ExternalLinkRelatedOperation[]
}

interface ExternalLinkRiskGroup {
	readonly key: string
	readonly workbook?: string | undefined
	readonly sheet?: string | undefined
	readonly references: readonly string[]
	readonly externalReference?: ExternalReferenceInfo | undefined
	readonly bindingRisk?: ExternalLinkBindingRisk | undefined
	readonly relatedOperations: readonly ExternalLinkRelatedOperation[]
}

interface ExternalLinkBindingRisk {
	readonly status: NonNullable<ExternalReferenceInfo['linkBindingStatus']>
	readonly externalBookRelId?: string | undefined
	readonly linkRelId?: string | undefined
	readonly linkRelationshipPart?: string | undefined
	readonly linkRelationshipKind?: ExternalReferenceInfo['linkRelationshipKind'] | undefined
	readonly linkRelationshipRawTarget?: string | undefined
	readonly target?: string | undefined
	readonly targetMode?: string | undefined
}

interface ExternalLinkRelatedOperation {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly sourceKind:
		| 'cellFormula'
		| 'fillFormula'
		| 'definedName'
		| 'deletedDefinedName'
		| 'externalLinkRewrite'
		| 'dataValidation'
		| 'conditionalFormat'
		| 'tableColumnFormula'
		| 'tableTotalsRowFormula'
		| 'chartSeriesName'
		| 'chartSeriesCategory'
		| 'chartSeriesValue'
		| 'sparklineGroupRange'
	readonly sheetName?: string | undefined
	readonly sourceRef?: string | undefined
	readonly range?: string | undefined
	readonly name?: string | undefined
	readonly scope?: string | undefined
	readonly formula?: string | undefined
	readonly workbook?: string | undefined
	readonly sheet?: string | undefined
	readonly references?: readonly string[] | undefined
	readonly externalReference?: ExternalReferenceInfo | undefined
	readonly selector?: ExternalLinkRewriteSelector | undefined
	readonly newTarget?: string | undefined
	readonly targetMode?: string | undefined
}

interface ExternalLinkRewriteSelector {
	readonly partPath?: string | undefined
	readonly relId?: string | undefined
	readonly linkRelId?: string | undefined
	readonly target?: string | undefined
}

interface ExternalFormulaReferenceGroup {
	readonly workbook: string
	readonly sheet?: string | undefined
	readonly references: readonly string[]
	readonly externalReference?: ExternalReferenceInfo | undefined
}

function buildExternalLinkRisk(
	workbook: Workbook,
	operations: readonly Operation[],
	externalReferences: readonly ExternalReferenceInfo[],
): ExternalLinkRisk {
	const relatedOperations = collectExternalLinkRelatedOperations(
		workbook,
		operations,
		externalReferences,
	)
	const groupByKey = new Map<string, ExternalLinkRiskGroup>()
	for (const reference of externalReferences) {
		groupByKey.set(externalLinkGroupKey(reference), {
			key: externalLinkGroupKey(reference),
			externalReference: copyExternalReferenceInfo(reference),
			...(externalLinkBindingRisk(reference)
				? { bindingRisk: externalLinkBindingRisk(reference) }
				: {}),
			references: [],
			relatedOperations: [],
		})
	}
	for (const operation of relatedOperations) {
		const key = operation.externalReference
			? externalLinkGroupKey(operation.externalReference)
			: `token:${operation.workbook ?? operation.selector?.target ?? operation.operationIndex}`
		const existing = groupByKey.get(key)
		const references = uniqueStrings([
			...(existing?.references ?? []),
			...(operation.references ?? []),
		])
		groupByKey.set(key, {
			key,
			...((operation.workbook ?? existing?.workbook)
				? { workbook: operation.workbook ?? existing?.workbook }
				: {}),
			...((operation.sheet ?? existing?.sheet)
				? { sheet: operation.sheet ?? existing?.sheet }
				: {}),
			references,
			...((operation.externalReference ?? existing?.externalReference)
				? {
						externalReference: copyExternalReferenceInfo(
							operation.externalReference ?? existing?.externalReference,
						),
					}
				: {}),
			...(externalLinkBindingRisk(operation.externalReference ?? existing?.externalReference)
				? {
						bindingRisk: externalLinkBindingRisk(
							operation.externalReference ?? existing?.externalReference,
						),
					}
				: {}),
			relatedOperations: [...(existing?.relatedOperations ?? []), operation],
		})
	}
	return {
		linkGroups: [...groupByKey.values()],
		relatedOperations,
	}
}

function collectExternalLinkRelatedOperations(
	workbook: Workbook,
	operations: readonly Operation[],
	externalReferences: readonly ExternalReferenceInfo[],
): ExternalLinkRelatedOperation[] {
	const related: ExternalLinkRelatedOperation[] = []
	operations.forEach((operation, operationIndex) => {
		switch (operation.op) {
			case 'setFormula':
				pushExternalLinkFormulaOperations(related, externalReferences, operationIndex, operation, [
					{
						sourceKind: 'cellFormula',
						sheetName: operation.sheet,
						sourceRef: `${operation.sheet}!${operation.ref}`,
						formula: operation.formula,
					},
				])
				break
			case 'fillFormula':
				pushExternalLinkFormulaOperations(related, externalReferences, operationIndex, operation, [
					{
						sourceKind: 'fillFormula',
						sheetName: operation.sheet,
						range: operation.range,
						formula: operation.formula,
					},
				])
				break
			case 'setDefinedName':
				pushExternalLinkFormulaOperations(related, externalReferences, operationIndex, operation, [
					{
						sourceKind: 'definedName',
						name: operation.name,
						...(operation.scope ? { scope: operation.scope } : {}),
						formula: operation.ref,
					},
				])
				break
			case 'deleteDefinedName':
				for (const name of matchingDefinedNames(workbook, operation)) {
					pushExternalLinkFormulaOperations(
						related,
						externalReferences,
						operationIndex,
						operation,
						[
							{
								sourceKind: 'deletedDefinedName',
								name: name.name,
								...(definedNameScopeName(workbook, name.scope)
									? { scope: definedNameScopeName(workbook, name.scope) }
									: {}),
								formula: name.formula,
							},
						],
					)
				}
				break
			case 'setDataValidation':
				pushExternalLinkFormulaOperations(
					related,
					externalReferences,
					operationIndex,
					operation,
					[operation.rule.formula1, operation.rule.formula2]
						.filter((formula): formula is string => formula !== undefined)
						.map((formula) => ({
							sourceKind: 'dataValidation',
							sheetName: operation.sheet,
							sourceRef: `${operation.sheet}!${operation.range}`,
							range: operation.range,
							formula,
						})),
				)
				break
			case 'setConditionalFormat':
				pushExternalLinkFormulaOperations(
					related,
					externalReferences,
					operationIndex,
					operation,
					conditionalFormatOperationFormulaTexts(operation.rule).map((formula) => ({
						sourceKind: 'conditionalFormat',
						sheetName: operation.sheet,
						sourceRef: `${operation.sheet}!${operation.range}`,
						range: operation.range,
						formula,
					})),
				)
				break
			case 'setTableColumn': {
				const sourceRef = `${operation.table}[${String(operation.column)}]`
				const sources: ExternalLinkFormulaOperationSource[] = []
				if (operation.formula) {
					sources.push({
						sourceKind: 'tableColumnFormula',
						sourceRef,
						name: operation.table,
						formula: operation.formula,
					})
				}
				if (operation.totalsRowFormula) {
					sources.push({
						sourceKind: 'tableTotalsRowFormula',
						sourceRef,
						name: operation.table,
						formula: operation.totalsRowFormula,
					})
				}
				pushExternalLinkFormulaOperations(
					related,
					externalReferences,
					operationIndex,
					operation,
					sources,
				)
				break
			}
			case 'setChartSeriesSource': {
				const sourceRef = chartOperationSourceRef(operation)
				const sources: ExternalLinkFormulaOperationSource[] = []
				if (operation.nameRef) {
					sources.push({
						sourceKind: 'chartSeriesName',
						sourceRef,
						formula: operation.nameRef,
					})
				}
				if (operation.categoryRef) {
					sources.push({
						sourceKind: 'chartSeriesCategory',
						sourceRef,
						formula: operation.categoryRef,
					})
				}
				if (operation.valueRef) {
					sources.push({
						sourceKind: 'chartSeriesValue',
						sourceRef,
						formula: operation.valueRef,
					})
				}
				pushExternalLinkFormulaOperations(
					related,
					externalReferences,
					operationIndex,
					operation,
					sources,
				)
				break
			}
			case 'setSparklineGroup':
				pushExternalLinkFormulaOperations(
					related,
					externalReferences,
					operationIndex,
					operation,
					operation.range
						? [
								{
									sourceKind: 'sparklineGroupRange',
									sheetName: operation.sheet,
									sourceRef: `${operation.sheet}!sparklineGroup${operation.groupIndex}`,
									range: operation.range,
									formula: operation.range,
								},
							]
						: [],
				)
				break
			case 'rewriteExternalLink':
				for (const reference of externalReferences.filter((entry) =>
					rewriteExternalLinkMatches(operation, entry),
				)) {
					related.push({
						operationIndex,
						op: operation.op,
						sourceKind: 'externalLinkRewrite',
						externalReference: copyExternalReferenceInfo(reference),
						selector: externalLinkRewriteSelector(operation),
						newTarget: operation.newTarget,
						...(operation.targetMode ? { targetMode: operation.targetMode } : {}),
					})
				}
				break
		}
	})
	return related
}

type ExternalLinkFormulaOperationSource = Pick<
	ExternalLinkRelatedOperation,
	'sourceKind' | 'sheetName' | 'sourceRef' | 'range' | 'name' | 'scope' | 'formula'
> & {
	readonly formula: string
}

function pushExternalLinkFormulaOperations(
	related: ExternalLinkRelatedOperation[],
	externalReferences: readonly ExternalReferenceInfo[],
	operationIndex: number,
	operation: Operation,
	sources: readonly ExternalLinkFormulaOperationSource[],
): void {
	for (const source of sources) {
		for (const group of externalFormulaReferenceGroups(source.formula, externalReferences)) {
			related.push({
				operationIndex,
				op: operation.op,
				...source,
				workbook: group.workbook,
				...(group.sheet ? { sheet: group.sheet } : {}),
				references: group.references,
				...(group.externalReference
					? { externalReference: copyExternalReferenceInfo(group.externalReference) }
					: {}),
			})
		}
	}
}

function conditionalFormatOperationFormulaTexts(
	rule: Extract<Operation, { op: 'setConditionalFormat' }>['rule'],
): string[] {
	return [
		...[rule.formula, rule.formula2].filter((formula): formula is string => formula !== undefined),
		...conditionalFormatValueObjectFormulas(rule.colorScale?.cfvo),
		...conditionalFormatValueObjectFormulas(rule.dataBar?.cfvo),
		...conditionalFormatValueObjectFormulas(rule.iconSet?.cfvo),
	]
}

function conditionalFormatValueObjectFormulas(
	values:
		| readonly {
				readonly type?: string
				readonly value?: string
		  }[]
		| undefined,
): string[] {
	return (
		values?.flatMap((value) => (value.type === 'formula' && value.value ? [value.value] : [])) ?? []
	)
}

function chartOperationSourceRef(
	operation: Extract<Operation, { op: 'setChartSeriesSource' }>,
): string {
	if (operation.partPath) return `${operation.partPath}#series${operation.seriesIndex}`
	if (operation.sheet)
		return `${operation.sheet}!chart${operation.chartIndex ?? 0}#series${operation.seriesIndex}`
	return `chart${operation.chartIndex ?? 0}#series${operation.seriesIndex}`
}

function externalFormulaReferenceGroups(
	formula: string,
	externalReferences: readonly ExternalReferenceInfo[],
): ExternalFormulaReferenceGroup[] {
	const parsed = parseFormula(normalizeFormulaInput(formula))
	if (!parsed.ok) return []
	const groups = new Map<string, ExternalFormulaReferenceGroup>()
	for (const reference of flattenFormulaReferences(collectFormulaReferences(parsed.value))) {
		if (reference.scope?.kind !== 'external') continue
		const key = `${reference.scope.workbook}\u0000${reference.scope.sheet}`
		const existing = groups.get(key)
		const references = uniqueStrings([...(existing?.references ?? []), reference.text])
		groups.set(key, {
			workbook: reference.scope.workbook,
			sheet: reference.scope.sheet,
			references,
			...(existing?.externalReference
				? { externalReference: existing.externalReference }
				: {
						...(resolveExternalReference(externalReferences, reference.scope.workbook)
							? {
									externalReference: resolveExternalReference(
										externalReferences,
										reference.scope.workbook,
									),
								}
							: {}),
					}),
		})
	}
	return [...groups.values()]
}

function flattenFormulaReferences(
	references: readonly FormulaReferenceInfo[],
): FormulaReferenceInfo[] {
	const flattened: FormulaReferenceInfo[] = []
	for (const reference of references) {
		if (reference.kind === 'union' || reference.kind === 'intersection') {
			flattened.push(...flattenFormulaReferences(reference.members))
			continue
		}
		flattened.push(reference)
	}
	return flattened
}

function resolveExternalReference(
	details: readonly ExternalReferenceInfo[],
	workbookToken: string,
): ExternalReferenceInfo | undefined {
	const numericIndex = parseExternalReferenceIndex(workbookToken)
	if (numericIndex !== undefined) return copyExternalReferenceInfo(details[numericIndex])
	const tokenName = externalReferenceBasename(workbookToken)
	const matches = details.filter((entry) => {
		if (entry.target === workbookToken || entry.partPath === workbookToken) return true
		return entry.target !== undefined && externalReferenceBasename(entry.target) === tokenName
	})
	return matches.length === 1 ? copyExternalReferenceInfo(matches[0]) : undefined
}

function parseExternalReferenceIndex(workbookToken: string): number | undefined {
	if (!/^\d+$/.test(workbookToken)) return undefined
	const index = Number(workbookToken)
	return Number.isSafeInteger(index) && index > 0 ? index - 1 : undefined
}

function externalReferenceBasename(path: string): string {
	const normalized = path.replace(/\\/g, '/')
	return normalized.slice(normalized.lastIndexOf('/') + 1)
}

function matchingDefinedNames(
	workbook: Workbook,
	operation: Extract<Operation, { op: 'deleteDefinedName' }>,
): ReturnType<Workbook['definedNames']['list']> {
	return workbook.definedNames.list().filter((name) => {
		if (name.name !== operation.name) return false
		if (!operation.scope) return name.scope.kind === 'workbook'
		return definedNameScopeName(workbook, name.scope) === operation.scope
	})
}

function definedNameScopeName(
	workbook: Workbook,
	scope: ReturnType<Workbook['definedNames']['list']>[number]['scope'],
): string | undefined {
	if (scope.kind === 'workbook') return undefined
	return workbook.sheets.find((sheet) => sheet.id === scope.sheetId)?.name
}

function rewriteExternalLinkMatches(
	operation: Extract<Operation, { op: 'rewriteExternalLink' }>,
	reference: ExternalReferenceInfo,
): boolean {
	if (operation.partPath !== undefined && reference.partPath !== operation.partPath) return false
	if (operation.relId !== undefined && reference.relId !== operation.relId) return false
	if (operation.linkRelId !== undefined && reference.linkRelId !== operation.linkRelId) return false
	if (operation.target !== undefined && reference.target !== operation.target) return false
	return (
		operation.partPath !== undefined ||
		operation.relId !== undefined ||
		operation.linkRelId !== undefined ||
		operation.target !== undefined
	)
}

function externalLinkRewriteSelector(
	operation: Extract<Operation, { op: 'rewriteExternalLink' }>,
): ExternalLinkRewriteSelector {
	return {
		...(operation.partPath ? { partPath: operation.partPath } : {}),
		...(operation.relId ? { relId: operation.relId } : {}),
		...(operation.linkRelId ? { linkRelId: operation.linkRelId } : {}),
		...(operation.target ? { target: operation.target } : {}),
	}
}

function externalLinkGroupKey(reference: ExternalReferenceInfo): string {
	return `part:${reference.partPath}`
}

function externalLinkBindingRisk(
	reference: ExternalReferenceInfo | undefined,
): ExternalLinkBindingRisk | undefined {
	if (!reference?.linkBindingStatus || reference.linkBindingStatus === 'externalBookRelId') {
		return undefined
	}
	return {
		status: reference.linkBindingStatus,
		...(reference.externalBookRelId ? { externalBookRelId: reference.externalBookRelId } : {}),
		...(reference.linkRelId ? { linkRelId: reference.linkRelId } : {}),
		...(reference.linkRelationshipPart
			? { linkRelationshipPart: reference.linkRelationshipPart }
			: {}),
		...(reference.linkRelationshipKind
			? { linkRelationshipKind: reference.linkRelationshipKind }
			: {}),
		...(reference.linkRelationshipRawTarget
			? { linkRelationshipRawTarget: reference.linkRelationshipRawTarget }
			: {}),
		...(reference.target ? { target: reference.target } : {}),
		...(reference.targetMode ? { targetMode: reference.targetMode } : {}),
	}
}

function externalLinkRewriteRecommendation(group: ExternalLinkRiskGroup): {
	readonly op: 'rewriteExternalLink'
	readonly partPath?: string | undefined
	readonly linkRelId?: string | undefined
	readonly target?: string | undefined
	readonly newTarget: '<new-target>'
	readonly targetMode?: string | undefined
} {
	const reference = group.externalReference
	return {
		op: 'rewriteExternalLink',
		...(reference?.partPath ? { partPath: reference.partPath } : {}),
		...(reference?.linkRelId ? { linkRelId: reference.linkRelId } : {}),
		...(reference?.target ? { target: reference.target } : {}),
		newTarget: '<new-target>',
		...(reference?.targetMode ? { targetMode: reference.targetMode } : {}),
	}
}

function copyExternalReferenceInfo(
	reference: ExternalReferenceInfo | undefined,
): ExternalReferenceInfo | undefined {
	return reference ? { ...reference } : undefined
}

interface LegacyCommentLocation {
	readonly sheetName: string
	readonly ref: string
	readonly location: string
	readonly author?: string
	readonly hasLegacyDrawing: boolean
	readonly shapeId?: string
	readonly vmlTarget?: { readonly row?: number; readonly column?: number }
}

interface ThreadedCommentLocation {
	readonly sheetName: string
	readonly ref: string
	readonly location: string
	readonly commentIndex: number
	readonly partPath?: string
	readonly id?: string
	readonly parentId?: string
	readonly personId?: string
	readonly author?: string
}

interface TableLocation {
	readonly sheetName: string
	readonly tableName: string
	readonly location: string
	readonly partPath?: string
	readonly tableId?: string
	readonly queryTablePartPath?: string
	readonly queryTableRelationshipId?: string
}

interface TableRelatedOperation {
	readonly operationIndex: number
	readonly op: string
	readonly sheetName?: string
	readonly tableName?: string
	readonly ref?: string
	readonly at?: number
	readonly count?: number
	readonly axis?: 'row' | 'column'
}

interface CommentRelatedOperation {
	readonly operationIndex: number
	readonly op: string
	readonly sheetName?: string
	readonly ref?: string
	readonly source?: string
	readonly target?: string
	readonly targetSheet?: string
	readonly mode?: string
	readonly at?: number
	readonly count?: number
	readonly axis?: 'row' | 'column'
	readonly range?: string
	readonly affectedLocations?: readonly string[]
	readonly partPath?: string
	readonly threadedCommentId?: string
	readonly commentIndex?: number
}

function collectCommentIntegrityIssues(checkIssues: readonly CheckIssue[]): {
	readonly legacy: readonly CheckIssue[]
	readonly threaded: readonly CheckIssue[]
} {
	return {
		legacy: checkIssues.filter(isLegacyCommentIntegrityIssue),
		threaded: checkIssues.filter(isThreadedCommentIntegrityIssue),
	}
}

function collectTableIntegrityIssues(checkIssues: readonly CheckIssue[]): readonly CheckIssue[] {
	return checkIssues.filter(
		(issue) =>
			issue.rule === 'table-integrity' ||
			issue.rule === 'table-package-integrity' ||
			issue.rule === 'table-query-integrity' ||
			(issue.rule === 'package-graph-integrity' &&
				(checkIssueFeatureFamilies(issue).some(isTableFeatureFamily) ||
					checkIssuePackagePartPaths(issue).some(isTablePackagePartPath))),
	)
}

function collectChartIntegrityIssues(checkIssues: readonly CheckIssue[]): readonly CheckIssue[] {
	return checkIssues.filter((issue) => issue.rule === 'chart-series-integrity')
}

function chartSourceRefRiskMessage(driftCount: number, integrityIssueCount: number): string {
	if (driftCount > 0 && integrityIssueCount > 0) {
		return `${driftCount} chart source reference group(s) may drift under planned topology edits and ${integrityIssueCount} chart source verify issue(s) require repair.`
	}
	if (integrityIssueCount > 0) {
		return `${integrityIssueCount} chart source verify issue(s) require repair before write.`
	}
	return `${driftCount} chart source reference group(s) may drift under planned sheet, row, or column topology edits.`
}

function collectExternalLinkIntegrityIssues(
	checkIssues: readonly CheckIssue[],
): readonly CheckIssue[] {
	return checkIssues.filter(isExternalLinkIntegrityIssue)
}

function collectExternalLinkPartPaths(
	parts: readonly XlsxPackageGraph['parts'][number][],
	externalReferences: readonly ExternalReferenceInfo[],
	issues: readonly CheckIssue[] = [],
	packageIssuePartPaths: readonly string[] = [],
): string[] {
	return uniqueStrings([
		...parts
			.filter((part) => isExternalLinkFeatureFamily(part.featureFamily))
			.map((part) => part.path),
		...externalReferences.flatMap((reference) => [
			reference.partPath,
			...(reference.linkRelationshipPart ? [reference.linkRelationshipPart] : []),
		]),
		...issues.flatMap((issue) =>
			checkIssuePackagePartPaths(issue).filter(isExternalLinkPackagePartPath),
		),
		...packageIssuePartPaths.filter(isExternalLinkPackagePartPath),
	])
}

function collectTableLocations(workbook: Workbook): TableLocation[] {
	return workbook.sheets.flatMap((sheet) =>
		sheet.tables.map((table) => {
			const location = `${sheet.name}!${rangeRefToA1(table.ref)}`
			return {
				sheetName: sheet.name,
				tableName: table.name,
				location,
				...(table.partPath ? { partPath: table.partPath } : {}),
				tableId: String(table.id),
				...(table.queryTable?.partPath ? { queryTablePartPath: table.queryTable.partPath } : {}),
				...(table.queryTable?.relationshipId
					? { queryTableRelationshipId: table.queryTable.relationshipId }
					: {}),
			}
		}),
	)
}

function collectTablePartPaths(
	parts: readonly XlsxPackageGraph['parts'][number][],
	locations: readonly TableLocation[],
	issues: readonly CheckIssue[] = [],
	packageIssuePartPaths: readonly string[] = [],
): string[] {
	return uniqueStrings([
		...parts.filter((part) => isTableFeatureFamily(part.featureFamily)).map((part) => part.path),
		...locations.flatMap((location) => [
			...(location.partPath ? [location.partPath] : []),
			...(location.queryTablePartPath ? [location.queryTablePartPath] : []),
		]),
		...issues.flatMap((issue) => checkIssuePackagePartPaths(issue).filter(isTablePackagePartPath)),
		...packageIssuePartPaths.filter(isTablePackagePartPath),
	])
}

function collectTableRelatedOperations(
	workbook: Workbook,
	operations: readonly Operation[],
): TableRelatedOperation[] {
	const tableSheets = new Set(
		workbook.sheets.filter((sheet) => sheet.tables.length > 0).map((sheet) => sheet.name),
	)
	const related: TableRelatedOperation[] = []
	for (const [operationIndex, op] of operations.entries()) {
		switch (op.op) {
			case 'createTable':
				related.push({
					operationIndex,
					op: op.op,
					sheetName: op.sheet,
					tableName: op.name,
					ref: op.ref,
				})
				break
			case 'appendRows':
			case 'deleteTable':
			case 'renameTable':
			case 'resizeTable':
			case 'setTableColumn':
			case 'setTableStyle':
				related.push({
					operationIndex,
					op: op.op,
					tableName: op.table,
					...('ref' in op ? { ref: op.ref } : {}),
				})
				break
			case 'insertRows':
			case 'deleteRows':
				if (!tableSheets.has(op.sheet)) break
				related.push({
					operationIndex,
					op: op.op,
					sheetName: op.sheet,
					at: op.at,
					count: op.count,
					axis: 'row',
				})
				break
			case 'insertCols':
			case 'deleteCols':
				if (!tableSheets.has(op.sheet)) break
				related.push({
					operationIndex,
					op: op.op,
					sheetName: op.sheet,
					at: op.at,
					count: op.count,
					axis: 'column',
				})
				break
		}
	}
	return related
}

function collectLegacyCommentLocations(workbook: Workbook): LegacyCommentLocation[] {
	return workbook.sheets.flatMap((sheet) =>
		[...sheet.comments.entries()].map(([ref, comment]) => ({
			sheetName: sheet.name,
			ref,
			location: `${sheet.name}!${ref}`,
			...(comment.author ? { author: comment.author } : {}),
			hasLegacyDrawing: comment.legacyDrawing !== undefined,
			...(comment.legacyDrawing?.shapeId ? { shapeId: comment.legacyDrawing.shapeId } : {}),
			...(comment.legacyDrawing
				? {
						vmlTarget: {
							...(comment.legacyDrawing.row !== undefined
								? { row: comment.legacyDrawing.row }
								: {}),
							...(comment.legacyDrawing.column !== undefined
								? { column: comment.legacyDrawing.column }
								: {}),
						},
					}
				: {}),
		})),
	)
}

function collectThreadedCommentLocations(workbook: Workbook): ThreadedCommentLocation[] {
	return workbook.sheets.flatMap((sheet) =>
		sheet.threadedComments.map((comment, commentIndex) => ({
			sheetName: sheet.name,
			ref: comment.ref,
			location: `${sheet.name}!${comment.ref}`,
			commentIndex,
			...(comment.partPath ? { partPath: comment.partPath } : {}),
			...(comment.id ? { id: comment.id } : {}),
			...(comment.parentId ? { parentId: comment.parentId } : {}),
			...(comment.personId ? { personId: comment.personId } : {}),
			...(comment.author ? { author: comment.author } : {}),
		})),
	)
}

function collectLegacyCommentPartPaths(
	parts: readonly XlsxPackageGraph['parts'][number][],
	locations: readonly LegacyCommentLocation[],
	issues: readonly CheckIssue[] = [],
	packageIssuePartPaths: readonly string[] = [],
): string[] {
	const hasVmlLayout = locations.some((location) => location.hasLegacyDrawing)
	return uniqueStrings([
		...parts
			.filter(
				(part) =>
					part.featureFamily === 'preservedComments' ||
					(hasVmlLayout && part.featureFamily === 'preservedVml'),
			)
			.map((part) => part.path),
		...issues.flatMap((issue) =>
			checkIssuePackagePartPaths(issue).filter(isLegacyCommentPackagePartPath),
		),
		...packageIssuePartPaths.filter(isLegacyCommentPackagePartPath),
	])
}

function collectThreadedCommentPartPaths(
	parts: readonly XlsxPackageGraph['parts'][number][],
	locations: readonly ThreadedCommentLocation[],
	issues: readonly CheckIssue[] = [],
	packageIssuePartPaths: readonly string[] = [],
): string[] {
	return uniqueStrings([
		...parts
			.filter((part) => part.featureFamily === 'preservedThreadedComments')
			.map((part) => part.path),
		...locations.flatMap((location) => (location.partPath ? [location.partPath] : [])),
		...issues.flatMap((issue) =>
			checkIssuePackagePartPaths(issue).filter(isThreadedCommentPackagePartPath),
		),
		...packageIssuePartPaths.filter(isThreadedCommentPackagePartPath),
	])
}

function isLegacyCommentIntegrityIssue(issue: CheckIssue): boolean {
	if (issue.rule === 'legacy-comment-drawing-integrity') return true
	if (issue.rule === 'drawing-integrity' && issue.details?.kind === 'orphan-vml-drawing-part') {
		return checkIssuePackagePartPaths(issue).some(isLegacyCommentPackagePartPath)
	}
	if (issue.rule !== 'package-graph-integrity') return false
	if (checkIssueFeatureFamilies(issue).some(isThreadedCommentFeatureFamily)) return false
	if (checkIssuePackagePartPaths(issue).some(isThreadedCommentPackagePartPath)) return false
	if (checkIssueFeatureFamilies(issue).some(isLegacyCommentFeatureFamily)) return true
	return (
		checkIssuePackagePartPaths(issue).some(isLegacyCommentPackagePartPath) ||
		checkIssueContentTypes(issue).some(isLegacyCommentContentType)
	)
}

function isThreadedCommentIntegrityIssue(issue: CheckIssue): boolean {
	if (issue.rule === 'threaded-comment-integrity') return true
	if (issue.rule !== 'package-graph-integrity') return false
	if (checkIssueFeatureFamilies(issue).some(isThreadedCommentFeatureFamily)) return true
	return (
		checkIssuePackagePartPaths(issue).some(isThreadedCommentPackagePartPath) ||
		checkIssueContentTypes(issue).some(isThreadedCommentContentType)
	)
}

function isExternalLinkIntegrityIssue(issue: CheckIssue): boolean {
	if (issue.rule === 'external-link-integrity') return true
	if (issue.rule !== 'package-graph-integrity') return false
	if (checkIssueFeatureFamilies(issue).some(isExternalLinkFeatureFamily)) return true
	return checkIssuePackagePartPaths(issue).some(isExternalLinkPackagePartPath)
}

function checkIssuePackagePartPaths(issue: CheckIssue): string[] {
	const details = issue.details ?? {}
	const values = [
		...(issue.refs ?? []).filter(isPackagePartPath),
		stringDetail(details, 'partPath'),
		stringDetail(details, 'drawingPartPath'),
		stringDetail(details, 'firstPartPath'),
		stringDetail(details, 'duplicatePartPath'),
		stringDetail(details, 'relationshipPartPath'),
	].filter((value): value is string => value !== undefined && isPackagePartPath(value))
	return uniqueStrings(values)
}

function checkIssueContentTypes(issue: CheckIssue): string[] {
	const details = issue.details ?? {}
	return [
		stringDetail(details, 'contentType'),
		stringDetail(details, 'expected'),
		stringDetail(details, 'actual'),
	].filter((value): value is string => typeof value === 'string' && value.includes('/'))
}

function checkIssueFeatureFamilies(issue: CheckIssue): string[] {
	const featureFamily = stringDetail(issue.details ?? {}, 'featureFamily')
	return featureFamily ? [featureFamily] : []
}

function stringDetail(details: Readonly<Record<string, unknown>>, key: string): string | undefined {
	const value = details[key]
	return typeof value === 'string' ? value : undefined
}

function isPackagePartPath(value: string): boolean {
	return value === '[Content_Types].xml' || /^[^!#]+\/[^!#]+\.[A-Za-z0-9]+$/i.test(value)
}

function isAnalyticalPackagePartPath(path: string): boolean {
	return (
		/^xl\/pivotTables\/[^/]+\.xml$/i.test(path) ||
		/^xl\/pivotCache\/[^/]+\.xml$/i.test(path) ||
		/^xl\/slicerCaches\/[^/]+\.xml$/i.test(path) ||
		/^xl\/slicers\/[^/]+\.xml$/i.test(path) ||
		/^xl\/timelineCaches\/[^/]+\.xml$/i.test(path) ||
		/^xl\/timelines\/[^/]+\.xml$/i.test(path)
	)
}

function isLegacyCommentPackagePartPath(path: string): boolean {
	return /^xl\/comments\d+\.xml$/i.test(path) || /^xl\/drawings\/[^/]+\.vml$/i.test(path)
}

function isThreadedCommentPackagePartPath(path: string): boolean {
	return /^xl\/threadedComments\/[^/]+\.xml$/i.test(path) || /^xl\/persons\/[^/]+\.xml$/i.test(path)
}

function isTablePackagePartPath(path: string): boolean {
	return /^xl\/tables\/[^/]+\.xml$/i.test(path) || /^xl\/queryTables\/[^/]+\.xml$/i.test(path)
}

function isExternalLinkPackagePartPath(path: string): boolean {
	return /^xl\/externalLinks\/[^/]+\.xml$/i.test(path)
}

function isLegacyCommentFeatureFamily(featureFamily: string): boolean {
	return featureFamily === 'preservedComments' || featureFamily === 'preservedVml'
}

function isThreadedCommentFeatureFamily(featureFamily: string): boolean {
	return featureFamily === 'preservedThreadedComments'
}

function isTableFeatureFamily(featureFamily: string): boolean {
	return featureFamily === 'preservedTable' || featureFamily === 'preservedQueryTable'
}

function isExternalLinkFeatureFamily(featureFamily: string): boolean {
	return featureFamily === 'preservedExternalLink'
}

function isAnalyticalPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isAnalyticalSidecar(issue.featureFamily)) return true
	return (
		packageGraphIssuePackagePaths(issue).some(isAnalyticalPackagePartPath) ||
		packageGraphIssueContentTypes(issue).some(isAnalyticalContentType)
	)
}

function isTablePackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isTableFeatureFamily(issue.featureFamily)) return true
	return packageGraphIssuePackagePaths(issue).some(isTablePackagePartPath)
}

function isExternalLinkPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isExternalLinkFeatureFamily(issue.featureFamily)) return true
	return packageGraphIssuePackagePaths(issue).some(isExternalLinkPackagePartPath)
}

function isLegacyCommentPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isLegacyCommentFeatureFamily(issue.featureFamily)) return true
	return (
		packageGraphIssuePackagePaths(issue).some(isLegacyCommentPackagePartPath) ||
		packageGraphIssueContentTypes(issue).some(isLegacyCommentContentType)
	)
}

function isThreadedCommentPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isThreadedCommentFeatureFamily(issue.featureFamily)) return true
	return (
		packageGraphIssuePackagePaths(issue).some(isThreadedCommentPackagePartPath) ||
		packageGraphIssueContentTypes(issue).some(isThreadedCommentContentType)
	)
}

function isVisualPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (issue.featureFamily && isVisualSidecar(issue.featureFamily)) return true
	return (
		packageGraphIssuePackagePaths(issue).some(isVisualPackagePartPath) ||
		packageGraphIssueContentTypes(issue).some(isVisualContentType)
	)
}

function isRoutedPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): boolean {
	if (isLegacyCommentPackageGraphIssue(issue) || isThreadedCommentPackageGraphIssue(issue)) {
		return true
	}
	if (isVisualPackageGraphIssue(issue)) return true
	if (isAnalyticalPackageGraphIssue(issue)) return true
	if (!issue.featureFamily) return false
	return (
		isTableFeatureFamily(issue.featureFamily) || isExternalLinkFeatureFamily(issue.featureFamily)
	)
}

function isLegacyCommentContentType(contentType: string): boolean {
	return (
		contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml' ||
		contentType === 'application/vnd.openxmlformats-officedocument.vmlDrawing'
	)
}

function isThreadedCommentContentType(contentType: string): boolean {
	return (
		contentType === 'application/vnd.ms-excel.threadedcomments+xml' ||
		contentType === 'application/vnd.ms-excel.person+xml'
	)
}

function isAnalyticalContentType(contentType: string): boolean {
	return (
		contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml' ||
		contentType ===
			'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml' ||
		contentType ===
			'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml' ||
		contentType === 'application/vnd.ms-excel.slicerCache+xml' ||
		contentType === 'application/vnd.ms-excel.slicer+xml' ||
		contentType === 'application/vnd.ms-excel.timelineCache+xml' ||
		contentType === 'application/vnd.ms-excel.timeline+xml'
	)
}

function isVisualContentType(contentType: string): boolean {
	return (
		contentType === 'application/vnd.openxmlformats-officedocument.drawing+xml' ||
		contentType === 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml' ||
		contentType === 'application/vnd.ms-office.chartstyle+xml' ||
		contentType === 'application/vnd.ms-office.chartcolorstyle+xml' ||
		contentType === 'application/vnd.openxmlformats-officedocument.vmlDrawing' ||
		contentType.startsWith('image/')
	)
}

function packageGraphIssuePackagePaths(issue: XlsxPackageGraphFidelityIssue): string[] {
	const values = [
		issue.partPath,
		issue.sourcePartPath,
		issue.relationshipPartPath,
		typeof issue.actual === 'string' ? issue.actual : undefined,
		typeof issue.expected === 'string' ? issue.expected : undefined,
	].filter((value): value is string => value !== undefined && isPackagePartPath(value))
	return uniqueStrings(values)
}

function packageGraphIssueContentTypes(issue: XlsxPackageGraphFidelityIssue): string[] {
	return [
		issue.contentType,
		typeof issue.expected === 'string' ? issue.expected : undefined,
	].filter((value): value is string => typeof value === 'string' && value.includes('/'))
}

function packageGraphIssuesForParts(
	audit: PackageGraphAudit,
	partPaths: readonly string[],
	featureFamilies: readonly string[] = [],
): readonly XlsxPackageGraphFidelityIssue[] {
	if (partPaths.length === 0 && featureFamilies.length === 0) return []
	const pathSet = new Set(partPaths)
	const featureFamilySet = new Set(featureFamilies)
	return audit.issues.filter((issue) => {
		if (issue.partPath !== undefined && pathSet.has(issue.partPath)) return true
		if (issue.sourcePartPath !== undefined && pathSet.has(issue.sourcePartPath)) return true
		if (issue.relationshipPartPath !== undefined && pathSet.has(issue.relationshipPartPath)) {
			return true
		}
		if (typeof issue.actual === 'string' && pathSet.has(issue.actual)) return true
		return issue.featureFamily !== undefined && featureFamilySet.has(issue.featureFamily)
	})
}

function collectLegacyCommentRelatedOperations(
	locations: readonly LegacyCommentLocation[],
	operations: readonly Operation[],
): CommentRelatedOperation[] {
	const related: CommentRelatedOperation[] = []
	for (const [operationIndex, op] of operations.entries()) {
		if (op.op === 'setComment' || op.op === 'deleteComment') {
			related.push({
				operationIndex,
				op: op.op,
				sheetName: op.sheet,
				ref: op.ref,
			})
			continue
		}
		if ((op.op === 'copyRange' || op.op === 'moveRange') && copiesComments(op.mode)) {
			related.push({
				operationIndex,
				op: op.op,
				sheetName: op.sheet,
				source: op.source,
				target: op.target,
				...(op.targetSheet ? { targetSheet: op.targetSheet } : {}),
				...(op.mode ? { mode: op.mode } : {}),
			})
			continue
		}
		const topology = collectCommentTopologyRelatedOperation(locations, op, operationIndex)
		if (topology) related.push(topology)
	}
	return related
}

function collectThreadedCommentRelatedOperations(
	locations: readonly ThreadedCommentLocation[],
	operations: readonly Operation[],
): CommentRelatedOperation[] {
	const related: CommentRelatedOperation[] = []
	for (const [operationIndex, op] of operations.entries()) {
		if (op.op === 'setThreadedComment') {
			related.push({
				operationIndex,
				op: op.op,
				sheetName: op.sheet,
				...(op.ref ? { ref: op.ref } : {}),
				...(op.partPath ? { partPath: op.partPath } : {}),
				...(op.threadedCommentId ? { threadedCommentId: op.threadedCommentId } : {}),
				...(op.commentIndex !== undefined ? { commentIndex: op.commentIndex } : {}),
			})
			continue
		}
		const topology = collectCommentTopologyRelatedOperation(locations, op, operationIndex)
		if (topology) related.push(topology)
	}
	return related
}

function copiesComments(mode: string | undefined): boolean {
	return mode === undefined || mode === 'all' || mode === 'comments'
}

function collectCommentTopologyRelatedOperation(
	locations: readonly (LegacyCommentLocation | ThreadedCommentLocation)[],
	operation: Operation,
	operationIndex: number,
): CommentRelatedOperation | undefined {
	if (!isRowColumnTopologyOperation(operation)) return undefined
	const affectedLocations = locations
		.filter(
			(location) =>
				location.sheetName === operation.sheet &&
				commentLocationAffectedByTopologyOperation(location, operation),
		)
		.map((location) => location.location)
	if (affectedLocations.length === 0) return undefined
	return {
		operationIndex,
		op: operation.op,
		sheetName: operation.sheet,
		at: operation.at,
		count: operation.count,
		axis: operation.op === 'insertRows' || operation.op === 'deleteRows' ? 'row' : 'column',
		range: topologyOperationRange(operation),
		affectedLocations: uniqueStrings(affectedLocations),
	}
}

function commentLocationAffectedByTopologyOperation(
	location: LegacyCommentLocation | ThreadedCommentLocation,
	operation: Extract<Operation, { op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols' }>,
): boolean {
	let range: RangeRef
	try {
		range = parseRange(location.ref)
	} catch {
		return true
	}
	const axis = operation.op === 'insertRows' || operation.op === 'deleteRows' ? 'row' : 'col'
	return topologyOperationAffectsRange(operation, axis, range)
}

interface X14ConditionalFormatExtensionPayload {
	readonly sheetName: string
	readonly sheetPartPath?: string
	readonly source: 'x14:conditionalFormatting'
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
			if (format.deleted) continue
			const preservedAttributeNames = Object.keys(format.preservedRuleAttributes ?? {})
				.filter((name) => !X14_CONDITIONAL_FORMAT_MODELED_ATTRIBUTES.has(name))
				.sort()
			const preservedChildElements = uniqueStrings(
				(format.preservedRuleChildXml ?? []).map(preservedChildElementName),
			).sort()
			if (preservedAttributeNames.length === 0 && preservedChildElements.length === 0) {
				continue
			}
			payloads.push({
				sheetName: sheet.name,
				...(sheet.preservedXml?.partPath ? { sheetPartPath: sheet.preservedXml.partPath } : {}),
				source: 'x14:conditionalFormatting',
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

interface X14DataValidationExtensionPayload {
	readonly sheetName: string
	readonly sheetPartPath?: string
	readonly source: 'x14:dataValidations'
	readonly sqref: string
	readonly index: number
	readonly type?: string
	readonly operator?: string
	readonly hasFormula1: boolean
	readonly hasFormula2: boolean
	readonly preservedAttributeNames: readonly string[]
	readonly preservedChildElements: readonly string[]
}

function collectX14DataValidationExtensionPayloads(
	workbook: Workbook,
): X14DataValidationExtensionPayload[] {
	const payloads: X14DataValidationExtensionPayload[] = []
	for (const sheet of workbook.sheets) {
		for (const validation of sheet.x14DataValidations) {
			if (validation.deleted) continue
			const preservedAttributeNames = Object.keys(validation.preservedAttributes ?? {})
				.filter((name) => !X14_DATA_VALIDATION_MODELED_ATTRIBUTES.has(name))
				.sort()
			const preservedChildElements = uniqueStrings(
				(validation.preservedChildXml ?? []).map(preservedChildElementName),
			).sort()
			if (preservedAttributeNames.length === 0 && preservedChildElements.length === 0) {
				continue
			}
			payloads.push({
				sheetName: sheet.name,
				...(sheet.preservedXml?.partPath ? { sheetPartPath: sheet.preservedXml.partPath } : {}),
				source: 'x14:dataValidations',
				sqref: validation.sqref,
				index: validation.index,
				...(validation.type ? { type: validation.type } : {}),
				...(validation.operator ? { operator: validation.operator } : {}),
				hasFormula1: validation.formula1 !== undefined,
				hasFormula2: validation.formula2 !== undefined,
				preservedAttributeNames,
				preservedChildElements,
			})
		}
	}
	return payloads
}

const X14_CONDITIONAL_FORMAT_MODELED_ATTRIBUTES = new Set(['type', 'priority', 'id'])

const X14_DATA_VALIDATION_MODELED_ATTRIBUTES = new Set([
	'sqref',
	'type',
	'operator',
	'errorStyle',
	'imeMode',
	'allowBlank',
	'showInputMessage',
	'showErrorMessage',
	'showDropDown',
	'promptTitle',
	'prompt',
	'errorTitle',
	'error',
])

interface X14SemanticEdit {
	readonly operationIndex: number
	readonly op: Operation['op']
	readonly sheet?: string
	readonly sheetName: string
	readonly range?: string
	readonly table?: string
	readonly column?: string | number
	readonly newName?: string
	readonly at?: number
	readonly count?: number
	readonly priority?: number
	readonly ruleIndex?: number
	readonly affectedPayloads: readonly X14AffectedPayload[]
	readonly reason: string
}

interface X14AffectedPayload {
	readonly sheetName: string
	readonly sheetPartPath?: string
	readonly source:
		| X14ConditionalFormatExtensionPayload['source']
		| X14DataValidationExtensionPayload['source']
	readonly sqref: string
	readonly index: number
	readonly priority?: number
}

function collectX14ConditionalFormatSemanticEdits(
	workbook: Workbook,
	operations: readonly Operation[],
	payloads: readonly X14ConditionalFormatExtensionPayload[],
): X14SemanticEdit[] {
	const edits: X14SemanticEdit[] = []
	operations.forEach((operation, operationIndex) => {
		if (operation.op === 'setConditionalFormat' || operation.op === 'deleteConditionalFormat') {
			const affectedPayloads = payloads.filter((payload) =>
				conditionalFormatOperationAffectsPayload(operation, payload),
			)
			if (affectedPayloads.length === 0) return
			edits.push({
				operationIndex,
				op: operation.op,
				sheet: operation.sheet,
				sheetName: operation.sheet,
				...(operation.range ? { range: operation.range } : {}),
				...(operation.op === 'deleteConditionalFormat' && operation.priority !== undefined
					? { priority: operation.priority }
					: {}),
				...(operation.op === 'deleteConditionalFormat' && operation.ruleIndex !== undefined
					? { ruleIndex: operation.ruleIndex }
					: {}),
				affectedPayloads: affectedPayloads.map(x14AffectedPayload),
				reason:
					'conditional-format operation may change rule order, priority, or range semantics around preserved x14 extension XML',
			})
			return
		}
		const topologyEdit = x14TopologyEdit(operation, operationIndex, payloads)
		if (topologyEdit) edits.push(topologyEdit)
		const tableEdit = x14ConditionalFormatTableReferenceEdit(
			workbook,
			operation,
			operationIndex,
			payloads,
		)
		if (tableEdit) edits.push(tableEdit)
	})
	return edits
}

function collectX14DataValidationSemanticEdits(
	workbook: Workbook,
	operations: readonly Operation[],
	payloads: readonly X14DataValidationExtensionPayload[],
): X14SemanticEdit[] {
	const edits: X14SemanticEdit[] = []
	operations.forEach((operation, operationIndex) => {
		if (operation.op === 'setDataValidation' || operation.op === 'deleteDataValidation') {
			const affectedPayloads = payloads.filter((payload) =>
				dataValidationOperationAffectsPayload(operation, payload),
			)
			if (affectedPayloads.length === 0) return
			edits.push({
				operationIndex,
				op: operation.op,
				sheet: operation.sheet,
				sheetName: operation.sheet,
				range: operation.range,
				affectedPayloads: affectedPayloads.map(x14AffectedPayload),
				reason:
					'data-validation operation overlaps preserved x14 extension XML for the same worksheet range',
			})
			return
		}
		const topologyEdit = x14TopologyEdit(operation, operationIndex, payloads)
		if (topologyEdit) edits.push(topologyEdit)
		const tableEdit = x14DataValidationTableReferenceEdit(
			workbook,
			operation,
			operationIndex,
			payloads,
		)
		if (tableEdit) edits.push(tableEdit)
	})
	return edits
}

function x14TopologyEdit(
	operation: Operation,
	operationIndex: number,
	payloads: readonly (X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload)[],
): X14SemanticEdit | undefined {
	if (!isRowColumnTopologyOperation(operation)) return undefined
	const affectedPayloads = payloads.filter((payload) =>
		topologyOperationAffectsPayloadSqref(operation, payload),
	)
	if (affectedPayloads.length === 0) return undefined
	return {
		operationIndex,
		op: operation.op,
		sheet: operation.sheet,
		sheetName: operation.sheet,
		range: topologyOperationRange(operation),
		at: operation.at,
		count: operation.count,
		affectedPayloads: affectedPayloads.map(x14AffectedPayload),
		reason:
			'row or column topology operation can rewrite shifted x14 sqref metadata or tombstone fully deleted x14 entries',
	}
}

function x14ConditionalFormatTableReferenceEdit(
	workbook: Workbook,
	operation: Operation,
	operationIndex: number,
	payloads: readonly X14ConditionalFormatExtensionPayload[],
): X14SemanticEdit | undefined {
	if (!isTableReferenceRewriteOperation(operation)) return undefined
	const affectedPayloads = payloads.filter((payload) =>
		x14ConditionalFormatPayloadReferencesTable(workbook, payload, operation),
	)
	if (affectedPayloads.length === 0) return undefined
	return x14TableReferenceEdit(operation, operationIndex, affectedPayloads)
}

function x14DataValidationTableReferenceEdit(
	workbook: Workbook,
	operation: Operation,
	operationIndex: number,
	payloads: readonly X14DataValidationExtensionPayload[],
): X14SemanticEdit | undefined {
	if (!isTableReferenceRewriteOperation(operation)) return undefined
	const affectedPayloads = payloads.filter((payload) =>
		x14DataValidationPayloadReferencesTable(workbook, payload, operation),
	)
	if (affectedPayloads.length === 0) return undefined
	return x14TableReferenceEdit(operation, operationIndex, affectedPayloads)
}

function x14TableReferenceEdit(
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
	operationIndex: number,
	affectedPayloads: readonly (
		| X14ConditionalFormatExtensionPayload
		| X14DataValidationExtensionPayload
	)[],
): X14SemanticEdit {
	const sheetName = affectedPayloads[0]?.sheetName ?? ''
	return {
		operationIndex,
		op: operation.op,
		...(sheetName ? { sheet: sheetName, sheetName } : { sheetName }),
		table: operation.table,
		...(operation.op === 'setTableColumn' ? { column: operation.column } : {}),
		...(operation.newName ? { newName: operation.newName } : {}),
		affectedPayloads: affectedPayloads.map(x14AffectedPayload),
		reason:
			'table rename or column rename rewrites structured references in x14 formulas that reference the table',
	}
}

function isRowColumnTopologyOperation(
	operation: Operation,
): operation is Extract<
	Operation,
	{ op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols' }
> {
	return (
		operation.op === 'insertRows' ||
		operation.op === 'deleteRows' ||
		operation.op === 'insertCols' ||
		operation.op === 'deleteCols'
	)
}

function topologyOperationAffectsPayloadSqref(
	operation: Extract<Operation, { op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols' }>,
	payload: X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload,
): boolean {
	if (operation.sheet !== payload.sheetName) return false
	const ranges = parseSqrefRanges(payload.sqref)
	if (ranges.length === 0) return true
	const axis = operation.op === 'insertRows' || operation.op === 'deleteRows' ? 'row' : 'col'
	return ranges.some((range) => topologyOperationAffectsRange(operation, axis, range))
}

function topologyOperationAffectsRange(
	operation: Extract<Operation, { op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols' }>,
	axis: 'row' | 'col',
	range: RangeRef,
): boolean {
	const end =
		axis === 'row'
			? Math.max(range.start.row, range.end.row)
			: Math.max(range.start.col, range.end.col)
	const opStart = operation.at
	const opEnd = operation.at + operation.count - 1
	if (operation.op === 'insertRows' || operation.op === 'insertCols') {
		return opStart <= end
	}
	return opStart <= end && opEnd >= 0 && operation.count > 0
}

function topologyOperationRange(
	operation: Extract<Operation, { op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols' }>,
): string {
	const start = operation.at
	const end = operation.at + operation.count - 1
	if (operation.op === 'insertRows' || operation.op === 'deleteRows') {
		return `${start + 1}:${end + 1}`
	}
	return `${indexToColumn(start)}:${indexToColumn(end)}`
}

function isTableReferenceRewriteOperation(
	operation: Operation,
): operation is Extract<Operation, { op: 'renameTable' | 'setTableColumn' }> {
	return (
		operation.op === 'renameTable' ||
		(operation.op === 'setTableColumn' && operation.newName !== undefined)
	)
}

function x14ConditionalFormatPayloadReferencesTable(
	workbook: Workbook,
	payload: X14ConditionalFormatExtensionPayload,
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
): boolean {
	return formulaTextsReferenceTable(
		x14ConditionalFormatPayloadFormulaTexts(workbook, payload),
		workbook,
		payload,
		operation,
	)
}

function x14DataValidationPayloadReferencesTable(
	workbook: Workbook,
	payload: X14DataValidationExtensionPayload,
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
): boolean {
	return formulaTextsReferenceTable(
		x14DataValidationPayloadFormulaTexts(workbook, payload),
		workbook,
		payload,
		operation,
	)
}

function x14ConditionalFormatPayloadFormulaTexts(
	workbook: Workbook,
	payload: X14ConditionalFormatExtensionPayload,
): string[] {
	const format = workbook.sheets
		.find((sheet) => sheet.name === payload.sheetName)
		?.x14ConditionalFormats.find((entry) => entry.index === payload.index)
	return uniqueStrings([
		...(format?.formulas ?? []),
		...conditionalFormatValueObjectFormulas(format?.dataBar?.cfvo),
		...conditionalFormatValueObjectFormulas(format?.iconSet?.cfvo),
	])
}

function x14DataValidationPayloadFormulaTexts(
	workbook: Workbook,
	payload: X14DataValidationExtensionPayload,
): string[] {
	const validation = workbook.sheets
		.find((sheet) => sheet.name === payload.sheetName)
		?.x14DataValidations.find((entry) => entry.index === payload.index)
	return [validation?.formula1, validation?.formula2].filter(
		(formula): formula is string => formula !== undefined,
	)
}

function formulaTextsReferenceTable(
	formulas: readonly string[],
	workbook: Workbook,
	payload: X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload,
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
): boolean {
	if (formulas.length === 0) return false
	return formulas.some((formula) =>
		formulaStructuredReferences(formula).some((reference) =>
			structuredReferenceMatchesTableOperation(workbook, payload, reference, operation),
		),
	)
}

function formulaStructuredReferences(
	formula: string,
): Extract<FormulaReferenceInfo, { kind: 'structured' }>[] {
	const parsed = parseFormula(normalizeFormulaInput(formula))
	if (!parsed.ok) return []
	return flattenFormulaReferences(collectFormulaReferences(parsed.value)).filter(
		(reference): reference is Extract<FormulaReferenceInfo, { kind: 'structured' }> =>
			reference.kind === 'structured',
	)
}

function structuredReferenceMatchesTableOperation(
	workbook: Workbook,
	payload: X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload,
	reference: Extract<FormulaReferenceInfo, { kind: 'structured' }>,
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
): boolean {
	const tableName = reference.table || tableNameForLocalStructuredReference(workbook, payload)
	if (!tableName || !tableNameMatchesTableOperation(tableName, operation)) return false
	if (operation.op === 'renameTable') return true
	const columns = [String(operation.column), operation.newName].flatMap((column) =>
		column === undefined ? [] : [column.toLowerCase()],
	)
	if (!reference.column && !reference.endColumn) return true
	return (
		(reference.column !== undefined && columns.includes(reference.column.toLowerCase())) ||
		(reference.endColumn !== undefined && columns.includes(reference.endColumn.toLowerCase()))
	)
}

function tableNameMatchesTableOperation(
	tableName: string,
	operation: Extract<Operation, { op: 'renameTable' | 'setTableColumn' }>,
): boolean {
	const normalized = tableName.toLowerCase()
	if (normalized === operation.table.toLowerCase()) return true
	return operation.op === 'renameTable' && normalized === operation.newName.toLowerCase()
}

function tableNameForLocalStructuredReference(
	workbook: Workbook,
	payload: X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload,
): string | undefined {
	const sheet = workbook.sheets.find((entry) => entry.name === payload.sheetName)
	if (!sheet) return undefined
	return sheet.tables.find((table) => rangesMayOverlap(toRangeStringSafe(table.ref), payload.sqref))
		?.name
}

function toRangeStringSafe(range: RangeRef): string {
	return `${indexToColumn(Math.min(range.start.col, range.end.col))}${Math.min(range.start.row, range.end.row) + 1}:${indexToColumn(Math.max(range.start.col, range.end.col))}${Math.max(range.start.row, range.end.row) + 1}`
}

function conditionalFormatOperationAffectsPayload(
	operation: Extract<Operation, { op: 'setConditionalFormat' | 'deleteConditionalFormat' }>,
	payload: X14ConditionalFormatExtensionPayload,
): boolean {
	if (operation.sheet !== payload.sheetName) return false
	if (operation.op === 'deleteConditionalFormat') {
		if (operation.ruleIndex !== undefined) return operation.ruleIndex === payload.index
		if (operation.priority !== undefined) return operation.priority === payload.priority
	}
	if (!operation.range) return true
	return rangesMayOverlap(operation.range, payload.sqref)
}

function dataValidationOperationAffectsPayload(
	operation: Extract<Operation, { op: 'setDataValidation' | 'deleteDataValidation' }>,
	payload: X14DataValidationExtensionPayload,
): boolean {
	if (operation.sheet !== payload.sheetName) return false
	return rangesMayOverlap(operation.range, payload.sqref)
}

function rangesMayOverlap(left: string, right: string): boolean {
	const leftRanges = parseSqrefRanges(left)
	const rightRanges = parseSqrefRanges(right)
	if (leftRanges.length === 0 || rightRanges.length === 0) return true
	return leftRanges.some((leftRange) =>
		rightRanges.some((rightRange) => rangesOverlap(leftRange, rightRange)),
	)
}

function parseSqrefRanges(sqref: string): RangeRef[] {
	const ranges: RangeRef[] = []
	for (const part of sqref.trim().split(/\s+/)) {
		if (!part) continue
		try {
			ranges.push(parseRange(part))
		} catch {
			return []
		}
	}
	return ranges
}

function rangesOverlap(left: RangeRef, right: RangeRef): boolean {
	const leftStartRow = Math.min(left.start.row, left.end.row)
	const leftEndRow = Math.max(left.start.row, left.end.row)
	const leftStartCol = Math.min(left.start.col, left.end.col)
	const leftEndCol = Math.max(left.start.col, left.end.col)
	const rightStartRow = Math.min(right.start.row, right.end.row)
	const rightEndRow = Math.max(right.start.row, right.end.row)
	const rightStartCol = Math.min(right.start.col, right.end.col)
	const rightEndCol = Math.max(right.start.col, right.end.col)
	return (
		leftStartRow <= rightEndRow &&
		leftEndRow >= rightStartRow &&
		leftStartCol <= rightEndCol &&
		leftEndCol >= rightStartCol
	)
}

function x14AffectedPayload(
	payload: X14ConditionalFormatExtensionPayload | X14DataValidationExtensionPayload,
): X14AffectedPayload {
	return {
		sheetName: payload.sheetName,
		...(payload.sheetPartPath ? { sheetPartPath: payload.sheetPartPath } : {}),
		source: payload.source,
		sqref: payload.sqref,
		index: payload.index,
		...('priority' in payload && payload.priority !== undefined
			? { priority: payload.priority }
			: {}),
	}
}

function x14ConditionalFormatSuggestedAction(hasSemanticEditRisk: boolean): string {
	if (hasSemanticEditRisk) {
		return 'Inspect inspectSheet(sheet).x14ConditionalFormats entries by sheetName, sheetPartPath, index, priority, and sqref before committing conditional-format edits; verify postWrite.packageGraphAudit after write.'
	}
	return 'No action is required for unrelated cell edits; keep inspectSheet(sheet).x14ConditionalFormats provenance with the edit record if auditing opaque x14 payload preservation.'
}

function x14DataValidationSuggestedAction(hasSemanticEditRisk: boolean): string {
	if (hasSemanticEditRisk) {
		return 'Inspect inspectSheet(sheet).x14DataValidations entries by sheetName, sheetPartPath, index, and sqref before committing data-validation edits; verify postWrite.packageGraphAudit after write.'
	}
	return 'No action is required for unrelated cell edits; keep inspectSheet(sheet).x14DataValidations provenance with the edit record if auditing opaque x14 payload preservation.'
}

function x14ConditionalFormatRiskMessage(
	payloadCount: number,
	relatedOperations: readonly X14SemanticEdit[],
): string {
	if (
		relatedOperations.every(
			(operation) =>
				operation.op === 'setConditionalFormat' || operation.op === 'deleteConditionalFormat',
		)
	) {
		return `${payloadCount} preserved x14 conditional-format extension payload(s) share a sheet with planned conditional-format semantic edits.`
	}
	return `${payloadCount} preserved x14 conditional-format extension payload(s) intersect planned conditional-format, topology, or table-reference edits.`
}

function x14DataValidationRiskMessage(
	payloadCount: number,
	relatedOperations: readonly X14SemanticEdit[],
): string {
	if (
		relatedOperations.every(
			(operation) =>
				operation.op === 'setDataValidation' || operation.op === 'deleteDataValidation',
		)
	) {
		return `${payloadCount} preserved x14 data-validation extension payload(s) overlap planned data-validation semantic edits.`
	}
	return `${payloadCount} preserved x14 data-validation extension payload(s) intersect planned data-validation, topology, or table-reference edits.`
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

function rangeRefToA1(range: RangeRef): string {
	return `${indexToColumn(range.start.col)}${range.start.row + 1}:${indexToColumn(range.end.col)}${range.end.row + 1}`
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

function isVisualSidecar(featureFamily: string): boolean {
	return (
		featureFamily === 'preservedDrawing' ||
		featureFamily === 'preservedChart' ||
		featureFamily === 'preservedChartSheet' ||
		featureFamily === 'preservedChartStyle' ||
		featureFamily === 'preservedChartColor' ||
		featureFamily === 'preservedMedia' ||
		featureFamily === 'preservedVml'
	)
}

function isDrawingMlVisualPath(path: string): boolean {
	return (
		/^xl\/drawings\/[^/]+\.xml$/i.test(path) ||
		/^xl\/drawings\/_rels\/[^/]+\.xml\.rels$/i.test(path) ||
		/^xl\/media\/[^/]+$/i.test(path) ||
		/^xl\/charts\/[^/]+\.xml$/i.test(path) ||
		/^xl\/charts\/_rels\/[^/]+\.xml\.rels$/i.test(path)
	)
}

function isVisualPackagePartPath(path: string): boolean {
	return isDrawingMlVisualPath(path) || /^xl\/drawings\/[^/]+\.vml$/i.test(path)
}

function isAnalyticalSidecar(featureFamily: string): boolean {
	return (
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
