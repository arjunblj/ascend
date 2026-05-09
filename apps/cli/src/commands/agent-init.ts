import { jsonOut } from '../output/json.ts'
import { bullet, heading } from '../output/pretty.ts'

export const usage = `Usage: ascend agent-init [flags]

  Print the recommended headless spreadsheet workflow for coding agents.

Flags:
  --json    Output as JSON
`

interface AgentInitInfo {
	readonly workflow: readonly string[]
	readonly commands: Record<string, string>
	readonly mcpResources: readonly string[]
	readonly safetyDefaults: readonly string[]
	readonly progress: {
		readonly mode: 'jsonl'
		readonly stderr: true
		readonly commands: readonly string[]
	}
}

const AGENT_INIT: AgentInitInfo = {
	workflow: [
		'inspect workbook structure',
		'locate only the needed ranges, tables, formulas, visuals, or metadata',
		'build operations from the published operation schemas',
		'run plan before every write',
		'commit with output paths, input hash guards, approvals, and allow-loss only when explicit',
		'verify the result with check, lint, diff, trace, or export',
		'use repair-plan when validation, recalc, or unsupported-feature audits need recovery actions',
	],
	commands: {
		inspect: 'ascend inspect <file> --json --verbose',
		read: 'ascend read <file> <range> --json',
		docs: 'ascend docs <query> --json',
		operations: 'ascend ops --json',
		capabilities: 'ascend capabilities --json',
		plan: 'ascend plan <file> --ops ops.json --progress jsonl --json',
		commit:
			'ascend commit <file> --ops ops.json --output out.xlsx --expect-sha256 <hash> --progress jsonl --json',
		check: 'ascend check <file> --progress jsonl --json',
		repair: 'ascend repair-plan <file> --json',
	},
	mcpResources: [
		'ascend://llms.txt',
		'ascend://llms-full.txt',
		'ascend://docs/agent-api.md',
		'ascend://capabilities',
		'ascend://operations',
		'ascend://agent-workflow',
	],
	safetyDefaults: [
		'Prefer non-destructive --output writes over --in-place.',
		'Use --expect-sha256 from plan output to reject stale inputs.',
		'Pass --approval only for approval ids emitted by plan.',
		'Pass --allow-loss only for feature loss explicitly accepted by the user.',
		'Keep stdout machine-readable and read --progress jsonl from stderr for long workflows.',
	],
	progress: {
		mode: 'jsonl',
		stderr: true,
		commands: ['plan', 'commit', 'check'],
	},
}

export async function agentInitCommand(
	_args: string[],
	flags: Map<string, string>,
): Promise<number> {
	if (flags.has('json')) {
		console.log(jsonOut(AGENT_INIT))
		return 0
	}

	console.log(heading('Ascend Agent Workflow'))
	for (const step of AGENT_INIT.workflow) console.log(bullet('Step', step))
	console.log('')
	for (const [name, command] of Object.entries(AGENT_INIT.commands)) {
		console.log(bullet(name, command))
	}
	console.log('')
	for (const safety of AGENT_INIT.safetyDefaults) console.log(bullet('Safety', safety))
	return 0
}
