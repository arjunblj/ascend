import { readFile } from 'node:fs/promises'
import { ascendError, type Operation } from '@ascend/schema'
import { createAgentPlan, parseOperations } from '@ascend/sdk'
import { cliError, jsonErr, jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'
import { createAgentProgressReporter } from '../progress.ts'

export const usage = `Usage: ascend plan <file> --ops <file.json> [flags]

  Validate, preview, recalc-audit, and preservation-audit operations without saving.

Arguments:
  <file>              Path to the workbook file

Flags:
  --ops <file.json>   Operations JSON file
  --progress jsonl    Emit machine-readable progress events to stderr
  --json              Output as JSON
`

export async function planCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const opsFile = flags.get('ops')
	if (!file || !opsFile) {
		cliError('Usage: ascend plan <file> --ops <file.json>', flags)
		return 1
	}

	const ops = await readOpsFile(opsFile, flags)
	if (!ops) return 1
	const onProgress = createAgentProgressReporter(flags)
	const result = await createAgentPlan(file, ops, onProgress ? { onProgress } : {})
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
			console.log(jsonOut(result))
		}
		return result.preview.errors.length === 0 ? 0 : 1
	}

	console.log(heading(`Plan: ${file}`))
	console.log(bullet('Operations', result.operationCount))
	console.log(bullet('Input SHA-256', result.inputSha256))
	console.log(bullet('Plan digest', result.planDigest))
	console.log(bullet('Would succeed', result.preview.wouldSucceed ? 'yes' : 'no'))
	console.log(bullet('Cell changes', result.preview.cellChanges.length))
	console.log(bullet('Write parts', result.preservation.totalParts))
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

async function readOpsFile(
	opsFile: string,
	flags: Map<string, string>,
): Promise<readonly Operation[] | null> {
	const raw = await readFile(opsFile, 'utf-8')
	const parsed = parseOperations(JSON.parse(raw))
	if (!parsed.ok) {
		cliError(
			ascendError('VALIDATION_ERROR', parsed.error, {
				details: { issues: parsed.issues },
				suggestedFix: 'Run ascend ops --json for canonical operation schemas and examples.',
			}),
			flags,
		)
		return null
	}
	return parsed.value
}
