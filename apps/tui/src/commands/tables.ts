import { commandsForGroup } from './registry.ts'

export const TABLE_COMMANDS = commandsForGroup('insert').filter((command) =>
	command.id.includes('table'),
)
