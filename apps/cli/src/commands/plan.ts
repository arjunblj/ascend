import { readFile } from 'node:fs/promises'
import { ascendError, type Operation } from '@ascend/schema'
import {
	createAgentPlan,
	createPackageActionProof,
	operationValidationDetails,
	parseOperations,
} from '@ascend/sdk'
import { cliError, jsonErr, jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'
import { createAgentProgressReporter } from '../progress.ts'
import { printWritePolicySummary } from './agent-workflow-output.ts'

export const usage = `Usage: ascend plan <file> --ops <file.json> [flags]

  Validate, preview, recalc-audit, package-graph-audit, and preservation-audit operations without saving.

Arguments:
  <file>              Path to the workbook file

Flags:
  --ops <file.json>   Operations JSON file
  --password <value>  Password for encrypted XLSX/XLSM workbooks
  --package-actions   Include package action proof in JSON output
  --progress jsonl    Emit machine-readable progress events to stderr
  --json              Output as JSON
`

export async function planCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const opsFile = flags.get('ops')
	if (!file || !opsFile) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required plan input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'plan',
					required: ['file', 'ops'],
					missing: [...(!file ? ['file'] : []), ...(!opsFile ? ['ops'] : [])],
					workflow: ['inspect', 'plan', 'commit', 'reopen', 'verify'],
				},
				suggestedFix:
					'Run ascend plan <file> --ops <file.json> --json after creating an operations JSON file.',
			}),
			flags,
		)
		return 1
	}

	const ops = await readOpsFile(opsFile, flags)
	if (!ops) return 1
	const onProgress = createAgentProgressReporter(flags)
	const password = flags.get('password')
	const result = await createAgentPlan(file, ops, {
		...(password !== undefined ? { password } : {}),
		...(onProgress ? { onProgress } : {}),
	})
	if (flags.has('json')) {
		if (result.preview.errors.length > 0) {
			const first = result.preview.errors[0]
			console.log(
				jsonErr(
					first
						? { ...first, details: { ...(first.details ?? {}), plan: result } }
						: ascendError('VALIDATION_ERROR', 'Plan failed', { details: { plan: result } }),
				),
			)
		} else {
			console.log(jsonOut(withPackageActions(result, flags)))
		}
		return result.preview.errors.length === 0 ? 0 : 1
	}

	console.log(heading(`Plan: ${file}`))
	console.log(bullet('Operations', result.operationCount))
	console.log(bullet('Input SHA-256', result.inputSha256))
	console.log(bullet('Plan digest', result.planDigest))
	console.log(bullet('Would succeed', result.preview.wouldSucceed ? 'yes' : 'no'))
	console.log(bullet('Cell changes', result.preview.cellChanges.length))
	console.log(bullet('Package graph issues', result.packageGraphAudit.issues.length))
	console.log(bullet('Write parts', result.preservation.totalParts))
	printWritePolicySummary(result.writePolicy)
	console.log(bullet('Approval required', result.needsApproval ? 'yes' : 'no'))
	for (const approval of result.approvals) {
		console.log(bullet(`Approval ${approval.id}`, approval.title))
	}
	if (result.preview.errors.length > 0) {
		console.log('')
		for (const error of result.preview.errors) console.log(bullet('Error', error.message))
	}
	return result.preview.errors.length === 0 ? 0 : 1
}

function withPackageActions<T extends Awaited<ReturnType<typeof createAgentPlan>>>(
	result: T,
	flags: Map<string, string>,
): T | (T & { readonly packageActions: ReturnType<typeof createPackageActionProof> }) {
	if (!flags.has('package-actions')) return result
	return {
		...result,
		packageActions: createPackageActionProof(result.preservation, {
			writePolicy: result.writePolicy,
			packageGraphAudit: result.packageGraphAudit,
		}),
	}
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
