import { commandsForGroup } from './registry.ts'

export const FORMAT_COMMANDS = commandsForGroup('home').filter((command) =>
	command.id.includes('format'),
)
