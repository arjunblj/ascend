import {
	type CapabilityFilters,
	type CapabilityPriority,
	type CapabilityStatus,
	listCapabilities,
	summarizeCapabilities,
} from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'
import { bullet, heading, table } from '../output/pretty.ts'

export const usage = `Usage: ascend capabilities [flags]

  Show Ascend's canonical Excel capability matrix.

Flags:
  --feature <text>      Filter by feature id, label, or family
  --family <name>       Filter by exact capability family
  --priority <P0-P3>    Filter by priority
  --status <status>     Filter by capability status
  --gaps                Show non-editable/non-equivalent gaps only
  --json                Output as JSON
`

export async function capabilitiesCommand(
	_args: string[],
	flags: Map<string, string>,
): Promise<number> {
	const filters: CapabilityFilters = {
		...(flags.get('feature') ? { feature: flags.get('feature') as string } : {}),
		...(flags.get('family') ? { family: flags.get('family') as string } : {}),
		...(flags.get('priority') ? { priority: flags.get('priority') as CapabilityPriority } : {}),
		...(flags.get('status') ? { status: flags.get('status') as CapabilityStatus } : {}),
		...(flags.has('gaps') ? { gapsOnly: true } : {}),
	}
	const capabilities = listCapabilities(filters)
	const summary = summarizeCapabilities(capabilities)
	const data = { summary, capabilities }
	if (flags.has('json')) {
		console.log(jsonOut(data))
		return 0
	}
	console.log(heading('Excel Capabilities'))
	console.log(bullet('Total', summary.total))
	console.log(bullet('Gaps', summary.gaps))
	console.log('')
	console.log(
		table(
			['ID', 'Priority', 'Status', 'Family', 'Next milestone'],
			capabilities.map((capability) => [
				capability.id,
				capability.priority,
				capability.status,
				capability.family,
				capability.nextMilestone,
			]),
		),
	)
	return 0
}
