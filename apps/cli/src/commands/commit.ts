import { readFile } from 'node:fs/promises'
import { ascendError, type Operation } from '@ascend/schema'
import {
	type AgentCommitOptions,
	type AgentCommitResult,
	commitAgentPlan,
	compactAgentCommitResult,
	createAgentCommitPackageActionProof,
	operationValidationDetails,
	parseOperations,
} from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'
import { createAgentProgressReporter } from '../progress.ts'
import { printWritePolicySummary } from './agent-workflow-output.ts'

export const usage = `Usage: ascend commit <file> --ops <file.json> --output <out.xlsx> [flags]
       ascend commit <file> --ops <file.json> --in-place [--backup <backup.xlsx>]

  Apply validated operations and write atomically with an optional input hash guard.

Arguments:
  <file>                    Path to the workbook file

Flags:
  --ops <file.json>         Operations JSON file
  --output <out.xlsx>       Non-destructive output path
  --in-place                Replace the input file atomically
  --backup <backup.xlsx>    Backup path for --in-place
  --password <value>        Password for encrypted XLSX/XLSM workbooks
  --expect-sha256 <hash>    Reject commit if the input changed since plan
  --allow-loss <feature>    Allow preserved/unsupported feature loss by feature, tier, or "all"
  --approval <id>           Approve an explicit plan approval id, comma-separated list, or "all"
  --progress jsonl          Emit machine-readable progress events to stderr
  --compact                 Return compact JSON verification counts instead of full trace artifacts
  --package-actions         Include package action proof in JSON output
  --proof                   Include compact workflow proof bundle in JSON output
  --json                    Output as JSON
`

export async function commitCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const opsFile = flags.get('ops')
	if (!file || !opsFile) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required commit input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'commit',
					required: ['file', 'ops'],
					missing: [...(!file ? ['file'] : []), ...(!opsFile ? ['ops'] : [])],
					workflow: ['inspect', 'plan', 'commit', 'reopen', 'verify'],
				},
				suggestedFix:
					'Run ascend commit <file> --ops <file.json> --output out.xlsx --expect-sha256 <inputSha256> --json.',
			}),
			flags,
		)
		return 1
	}

	const ops = await readOpsFile(opsFile, flags)
	if (!ops) return 1
	const onProgress = createAgentProgressReporter(flags)
	const options: AgentCommitOptions = {
		...(flags.get('output') ? { output: flags.get('output') as string } : {}),
		...(flags.has('in-place') ? { inPlace: true } : {}),
		...(flags.get('backup') ? { backup: flags.get('backup') as string } : {}),
		...(flags.get('password') !== undefined ? { password: flags.get('password') as string } : {}),
		...(flags.get('expect-sha256') ? { expectSha256: flags.get('expect-sha256') as string } : {}),
		...(flags.get('allow-loss')
			? { allowLoss: parseAllowLoss(flags.get('allow-loss') as string) }
			: {}),
		...(flags.get('approval')
			? { approvals: parseApprovalFlags(flags.get('approval') as string) }
			: {}),
		...(onProgress ? { onProgress } : {}),
	}
	const result = await commitAgentPlan(file, ops, options)
	if (flags.has('json')) {
		const payload = flags.has('compact') ? compactAgentCommitResult(result) : result
		console.log(jsonOut(withProofBundle(withPackageActions(payload, result, flags), result, flags)))
		return 0
	}
	console.log(heading(`Committed: ${file}`))
	console.log(bullet('Output', result.output))
	if (result.backup) console.log(bullet('Backup', result.backup))
	console.log(bullet('Operations', result.operationCount))
	console.log(bullet('Input SHA-256', result.inputSha256))
	console.log(bullet('Output SHA-256', result.outputSha256))
	console.log(bullet('Plan digest', result.planDigest))
	console.log(bullet('Package graph issues', result.packageGraphAudit.issues.length))
	printWritePolicySummary(result.writePolicy)
	console.log(
		bullet('Post-write package graph issues', result.postWrite.packageGraphAudit.issues.length),
	)
	return 0
}

function withPackageActions<T>(
	payload: T,
	result: Awaited<ReturnType<typeof commitAgentPlan>>,
	flags: Map<string, string>,
): T | (T & { readonly packageActions: ReturnType<typeof createAgentCommitPackageActionProof> }) {
	if (!flags.has('package-actions')) return payload
	return {
		...payload,
		packageActions: createAgentCommitPackageActionProof(result),
	}
}

function withProofBundle<T>(
	payload: T,
	result: AgentCommitResult,
	flags: Map<string, string>,
): T | (T & { readonly proofBundle: ReturnType<typeof createCommitProofBundle> }) {
	if (!flags.has('proof')) return payload
	return {
		...payload,
		proofBundle: createCommitProofBundle(result, flags),
	}
}

function createCommitProofBundle(result: AgentCommitResult, flags: Map<string, string>) {
	const expectedSha256 = flags.get('expect-sha256')
	const whatChanged = result.apply.affectedCells.map((ref) => ({ ref }))
	const whySafe = [
		{
			gate: 'input-guard',
			ok: expectedSha256 !== undefined && expectedSha256 === result.inputSha256,
			evidence: {
				guard: expectedSha256 === undefined ? null : 'expect-sha256',
				expectedSha256: expectedSha256 ?? null,
				inputSha256: result.inputSha256,
			},
		},
		{
			gate: 'approval',
			ok: result.approvals.length === 0,
			evidence: {
				approvalCount: result.approvals.length,
				approvalIds: result.approvals.map((approval) => approval.id),
			},
		},
		{
			gate: 'write-policy',
			ok: result.writePolicy.ok,
			evidence: {
				diagnosticCount: result.writePolicy.diagnostics.length,
				blockerCount: result.writePolicy.diagnostics.filter(
					(diagnostic) => diagnostic.severity === 'blocker',
				).length,
			},
		},
		{
			gate: 'commit',
			ok: result.postWrite.valid && result.postWrite.auditsPassed,
			evidence: {
				outputSha256: result.outputSha256,
				postWriteValid: result.postWrite.valid,
				auditsPassed: result.postWrite.auditsPassed,
			},
		},
		{
			gate: 'reopen-verify',
			ok: result.postWrite.reopened && result.postWrite.check.valid && result.postWrite.lint.clean,
			evidence: {
				reopened: result.postWrite.reopened,
				checkValid: result.postWrite.check.valid,
				checkIssueCount: result.postWrite.check.issues.length,
				lintClean: result.postWrite.lint.clean,
				lintWarningCount: result.postWrite.lint.warnings.length,
			},
		},
		{
			gate: 'package-graph',
			ok:
				result.postWrite.packageGraphAudit.ok &&
				result.postWrite.unresolvedPackageGraphIssueCount === 0,
			evidence: {
				packageGraphAuditOk: result.postWrite.packageGraphAudit.ok,
				expectedPackageGraphIssueCount: result.postWrite.expectedPackageGraphIssueCount,
				unresolvedPackageGraphIssueCount: result.postWrite.unresolvedPackageGraphIssueCount,
			},
		},
	] as const
	return {
		safeToUse: whySafe.every((gate) => gate.ok),
		whatChanged,
		whySafe,
		evidence: {
			inputSha256: result.inputSha256,
			planDigest: result.planDigest,
			outputSha256: result.outputSha256,
			operationCount: result.operationCount,
			affectedCellCount: result.apply.affectedCells.length,
			postWriteValid: result.postWrite.valid,
			auditsPassed: result.postWrite.auditsPassed,
			reopened: result.postWrite.reopened,
			checkValid: result.postWrite.check.valid,
			lintClean: result.postWrite.lint.clean,
			writePolicyOk: result.writePolicy.ok,
			packageGraphAuditOk: result.postWrite.packageGraphAudit.ok,
		},
	}
}

function parseAllowLoss(value: string): readonly string[] | 'all' {
	return parseListOrAll(value)
}

function parseApprovalFlags(value: string): readonly string[] | 'all' {
	return parseListOrAll(value)
}

function parseListOrAll(value: string): readonly string[] | 'all' {
	const entries = value
		.split(',')
		.map((entry) => entry.trim())
		.filter((entry) => entry.length > 0)
	if (entries.some((entry) => entry.toLowerCase() === 'all')) return 'all'
	return entries
}

async function readOpsFile(
	opsFile: string,
	flags: Map<string, string>,
): Promise<readonly Operation[] | null> {
	const raw = await readFile(opsFile, 'utf-8')
	const parsed = parseOperations(JSON.parse(raw))
	if (!parsed.ok) {
		cliError(
			ascendError('VALIDATION_ERROR', parsed.error, {
				details: operationValidationDetails(parsed),
				suggestedFix: 'Run ascend ops --json for canonical operation schemas and examples.',
			}),
			flags,
		)
		return null
	}
	return parsed.value
}
