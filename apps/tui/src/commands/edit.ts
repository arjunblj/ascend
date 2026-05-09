import { commandsForGroup } from './registry.ts'

export const EDIT_COMMANDS = commandsForGroup('home').filter((command) =>
	[
		'home.copy',
		'home.cut',
		'home.paste',
		'home.pasteSpecial',
		'home.clear',
		'home.findReplace',
	].includes(command.id),
)
