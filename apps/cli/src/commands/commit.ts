import { readFile } from 'node:fs/promises'
import { ascendError, type Operation } from '@ascend/schema'
import { type AgentCommitOptions, commitAgentPlan, parseOperations } from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'
import { createAgentProgressReporter } from '../progress.ts'

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
  --expect-sha256 <hash>    Reject commit if the input changed since plan
  --allow-loss <feature>    Allow preserved/unsupported feature loss by feature, tier, or "all"
  --approval <id>           Approve an explicit plan approval id, comma-separated list, or "all"
  --progress jsonl          Emit machine-readable progress events to stderr
  --json                    Output as JSON
`

export async function commitCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const opsFile = flags.get('ops')
	if (!file || !opsFile) {
		cliError('Usage: ascend commit <file> --ops <file.json> --output <out.xlsx>', flags)
		return 1
	}

	const ops = await readOpsFile(opsFile, flags)
	if (!ops) return 1
	const onProgress = createAgentProgressReporter(flags)
	const options: AgentCommitOptions = {
		...(flags.get('output') ? { output: flags.get('output') as string } : {}),
		...(flags.has('in-place') ? { inPlace: true } : {}),
		...(flags.get('backup') ? { backup: flags.get('backup') as string } : {}),
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
		console.log(jsonOut(result))
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
	return 0
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
				details: { issues: parsed.issues },
				suggestedFix: 'Run ascend ops --json for canonical operation schemas and examples.',
			}),
			flags,
		)
		return null
	}
	return parsed.value
}
