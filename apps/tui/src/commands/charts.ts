import { commandsForGroup } from './registry.ts'

export const CHART_COMMANDS = commandsForGroup('insert').filter((command) =>
	command.id.includes('chart'),
)
