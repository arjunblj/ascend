import { ascendError } from '@ascend/schema'
import { createRepairPlan } from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { bullet, heading, table } from '../output/pretty.ts'

export const usage = `Usage: ascend repair-plan <file> [flags]

  Suggest safe next actions when check, lint, or unsupported-feature audits need attention.

Arguments:
  <file>          Path to the workbook file

Flags:
  --json          Output as JSON
`

export async function repairPlanCommand(
	args: string[],
	flags: Map<string, string>,
): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required repair-plan input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'repair-plan',
					required: ['file'],
					missing: ['file'],
					workflow: ['inspect', 'repair-plan', 'plan'],
				},
				suggestedFix:
					'Run ascend repair-plan <file> --json after check, lint, or unsupported-feature reports need recovery actions.',
			}),
			flags,
		)
		return 1
	}
	const result = await createRepairPlan(file)
	if (flags.has('json')) {
		console.log(jsonOut(result))
		return 0
	}
	console.log(heading(`Repair Plan: ${file}`))
	console.log(bullet('Input SHA-256', result.inputSha256))
	console.log(bullet('Actions', result.actions.length))
	console.log('')
	console.log(
		table(
			['Code', 'Action', 'Command'],
			result.actions.map((action) => [action.code, action.title, action.command ?? '']),
		),
	)
	return 0
}
