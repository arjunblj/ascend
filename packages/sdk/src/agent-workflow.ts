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

export interface RepairPlanResult {
	readonly file: string
	readonly inputSha256: string
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
	return {
		file,
		inputSha256,
		operationCount: ops.length,
		planDigest: digestPlan(inputSha256, ops),
		preview,
		check: wb.check(),
		lint: wb.lint(),
		preservation: wb.writePlanSummary(),
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

	const result: AgentCommitResult = {
		file,
		output,
		...(options.inPlace && options.backup ? { backup: options.backup } : {}),
		inputSha256,
		outputSha256,
		planDigest: digestPlan(inputSha256, ops),
		operationCount: ops.length,
		apply,
		recalc,
		check: wb.check(),
		lint: wb.lint(),
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
	return { file, inputSha256, check, lint, unsupportedFeatures, actions }
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
