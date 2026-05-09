import { commandsForGroup } from './registry.ts'

export const PIVOT_COMMANDS = commandsForGroup('data').filter((command) =>
	command.id.includes('pivot'),
)
