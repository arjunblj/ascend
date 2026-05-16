import { ascendError } from '@ascend/schema'
import { getOperationsSchema, listOperations } from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { table } from '../output/pretty.ts'

export const usage = `Usage: ascend ops [flags]

  List canonical operation schemas and examples for agent writes.

Flags:
  --op <name>     Show one operation
  --json          Output as JSON
`

export async function opsCommand(_args: string[], flags: Map<string, string>): Promise<number> {
	const op = flags.get('op')
	const operations = op ? listOperations().filter((entry) => entry.op === op) : listOperations()
	const schemas = op
		? getOperationsSchema().filter((entry) => entry.op === op)
		: getOperationsSchema()
	if (op && operations.length === 0) {
		cliError(unknownOperationError(op), flags)
		return 1
	}
	const data = { operations, schemas }
	if (flags.has('json')) {
		console.log(jsonOut(data))
		return 0
	}
	console.log(
		table(
			['Operation', 'Required', 'Optional', 'Description'],
			operations.map((entry) => [
				entry.op,
				entry.requiredFields.join(', '),
				entry.optionalFields?.join(', ') ?? '',
				entry.description,
			]),
		),
	)
	return 0
}

function unknownOperationError(op: string) {
	return ascendError('INVALID_ARGUMENT', `Unknown operation: ${op}`, {
		retryable: true,
		retryStrategy: 'modified',
		details: {
			command: 'ops',
			op,
			availableOperations: listOperations()
				.map((entry) => entry.op)
				.sort(),
			workflow: ['ops', 'plan', 'commit'],
		},
		suggestedFix:
			'Run ascend ops --json to list canonical operation schemas, then retry with a supported --op value.',
	})
}
