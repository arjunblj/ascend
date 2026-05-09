import { createHash } from 'node:crypto'
import { copyFile, readFile, rename, writeFile } from 'node:fs/promises'
import { dirname, extname, join } from 'node:path'
import { AscendException, ascendError, type FeatureReport, type Operation } from '@ascend/schema'
import { listCapabilities, summarizeCapabilities } from './capabilities.ts'
import { AscendWorkbook } from './workbook.ts'

export interface AgentPlanResult {
	readonly file: string
	readonly inputSha256: string
	readonly operationCount: number
	readonly planDigest: string
	readonly trace: AgentWorkflowTrace
	readonly modelOutput: AgentModelOutput
	readonly preview: ReturnType<AscendWorkbook['preview']>
	readonly check: ReturnType<AscendWorkbook['check']>
	readonly lint: ReturnType<AscendWorkbook['lint']>
	readonly preservation: ReturnType<AscendWorkbook['writePlanSummary']>
	readonly unsupportedFeatures: readonly unknown[]
	readonly lossAudit: LossAudit
	readonly capabilities: ReturnType<typeof summarizeCapabilities>
}

export interface AgentCommitOptions {
	readonly output?: string
	readonly inPlace?: boolean
	readonly backup?: string
	readonly expectSha256?: string
	readonly allowLoss?: readonly string[] | 'all'
}

export interface AgentCommitResult {
	readonly file: string
	readonly output: string
	readonly backup?: string
	readonly inputSha256: string
	readonly outputSha256: string
	readonly planDigest: string
	readonly operationCount: number
	readonly trace: AgentWorkflowTrace
	readonly modelOutput: AgentModelOutput
	readonly apply: ReturnType<AscendWorkbook['apply']>
	readonly recalc: ReturnType<AscendWorkbook['recalc']> | null
	readonly check: ReturnType<AscendWorkbook['check']>
	readonly lint: ReturnType<AscendWorkbook['lint']>
	readonly preservation: ReturnType<AscendWorkbook['writePlanSummary']>
	readonly lossAudit: LossAudit
}

export interface LossAudit {
	readonly ok: boolean
	readonly blockedFeatures: readonly FeatureReport[]
	readonly allowedLoss: readonly string[] | 'all'
	readonly policy: 'block-preserved-and-unsupported'
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
		readonly preservationParts?: number
	}
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
): Promise<AgentPlanResult> {
	const inputSha256 = await fileSha256(file)
	const wb = await AscendWorkbook.open(file)
	const preview = wb.preview(ops)
	const lossAudit = auditLossPolicy(wb.report.features)
	const check = wb.check()
	const lint = wb.lint()
	const preservation = wb.writePlanSummary()
	const planDigest = digestPlan(inputSha256, ops)
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
			preservationPhase(preservation),
		],
		artifacts: [
			artifact('ops', ops, `${ops.length} operation(s)`),
			artifact('preview', preview, `${preview.changedCells.length} changed cell(s)`),
			artifact('lossAudit', lossAudit, `${lossAudit.blockedFeatures.length} blocked feature(s)`),
			artifact('preservation', preservation, `${preservation.totalParts} package part(s)`),
		],
	})
	return {
		file,
		inputSha256,
		operationCount: ops.length,
		planDigest,
		trace,
		modelOutput: modelOutputFromTrace(trace),
		preview,
		check,
		lint,
		preservation,
		unsupportedFeatures: wb.report.features,
		lossAudit,
		capabilities: summarizeCapabilities(listCapabilities({ gapsOnly: true })),
	}
}

export async function commitAgentPlan(
	file: string,
	ops: readonly Operation[],
	options: AgentCommitOptions = {},
): Promise<AgentCommitResult> {
	const inputSha256 = await fileSha256(file)
	if (options.expectSha256 && options.expectSha256 !== inputSha256) {
		throw new AscendException(
			ascendError('VALIDATION_ERROR', 'Input workbook hash does not match --expect-sha256', {
				details: { expected: options.expectSha256, actual: inputSha256 },
				suggestedFix: 'Re-run ascend plan and commit with the new inputSha256.',
			}),
		)
	}

	const output = options.inPlace ? file : options.output
	if (!output) {
		throw new AscendException(
			ascendError('INVALID_ARGUMENT', 'Commit requires --output unless --in-place is set', {
				suggestedFix: 'Pass --output out.xlsx or --in-place with an optional --backup path.',
			}),
		)
	}
	if (!options.inPlace && options.backup) {
		throw new AscendException(
			ascendError('INVALID_ARGUMENT', '--backup is only valid with --in-place commits', {
				suggestedFix:
					'Use --output for non-destructive writes, or add --in-place when creating a backup.',
			}),
		)
	}

	const wb = await AscendWorkbook.open(file)
	const lossAudit = auditLossPolicy(wb.report.features, options.allowLoss)
	if (!lossAudit.ok) {
		throw new AscendException(
			ascendError('VALIDATION_ERROR', 'Workbook contains preserved or unsupported features', {
				details: { lossAudit },
				suggestedFix:
					'Inspect the plan lossAudit. If the write is intentional, pass --allow-loss <feature> for every blocked feature or --allow-loss all.',
			}),
		)
	}
	const apply = wb.apply(ops, { transaction: true })
	if (apply.errors.length > 0)
		throw new AscendException(apply.errors[0] ?? ascendError('VALIDATION_ERROR', 'Apply failed'))

	let recalc: ReturnType<AscendWorkbook['recalc']> | null = null
	if (apply.recalcRequired) {
		recalc = wb.recalc()
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
	}

	if (options.inPlace && options.backup) await copyFile(file, options.backup)
	const preservation = wb.writePlanSummary()
	await writeWorkbookAtomically(wb, output)
	const outputSha256 = await fileSha256(output)
	const check = wb.check()
	const lint = wb.lint()
	const planDigest = digestPlan(inputSha256, ops)
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
			applyPhase(apply),
			recalcPhase(recalc),
			preservationPhase(preservation),
			okPhase('write', `Workbook written to ${output}.`),
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
			artifact('preservation', preservation, `${preservation.totalParts} package part(s)`),
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
		trace,
		modelOutput: modelOutputFromTrace(trace),
		apply,
		recalc,
		check,
		lint,
		preservation,
		lossAudit,
	}
	return result
}

export function auditLossPolicy(
	features: readonly FeatureReport[],
	allowLoss: readonly string[] | 'all' = [],
): LossAudit {
	const allowed = allowLoss === 'all' ? 'all' : allowLoss.map((entry) => entry.toLowerCase())
	const blockedFeatures = features.filter((feature) => {
		if (feature.tier !== 'preserved' && feature.tier !== 'unsupported') return false
		if (allowed === 'all') return false
		return !allowed.includes(feature.feature.toLowerCase()) && !allowed.includes(feature.tier)
	})
	return {
		ok: blockedFeatures.length === 0,
		blockedFeatures,
		allowedLoss: allowLoss,
		policy: 'block-preserved-and-unsupported',
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
	return sha256Bytes(await readFile(file))
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
			lossAudit.blockedFeatures,
		)
	return okPhase('loss-audit', 'No unapproved preserved or unsupported feature loss detected.', 0)
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
	setCount(counts, 'preservationParts', phaseCount(trace, 'preservation-audit'))
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
	if (trace.phases.some((phase) => phase.phase === 'loss-audit' && phase.status === 'blocked')) {
		return [
			'Inspect lossAudit.blockedFeatures.',
			'Commit only with explicit allowLoss entries if the write is intentionally lossy.',
		]
	}
	if (trace.phases.some((phase) => phase.status === 'failed')) {
		return ['Inspect failed trace phases and run repair-plan before committing.']
	}
	if (trace.kind === 'plan') return ['Commit with this planDigest and inputSha256 when ready.']
	if (trace.kind === 'commit')
		return ['Verify the output workbook hash and retain the traceDigest.']
	return ['Follow the suggested repair actions in order.']
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
