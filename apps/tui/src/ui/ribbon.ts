import { commandsForGroup } from '../commands/registry.ts'
import { fitAnsi } from '../render/ansi-text.ts'
import { THEME } from '../render/styles.ts'

export function renderRibbon(width: number): string {
	const tabs =
		width < 56
			? ' Alt/F10  File  Home  Data  View '
			: width < 72
				? ' Alt/F10  File  Home  Insert  Formulas  Data  View '
				: ' Alt/F10  File  Home  Insert  Page Layout  Formulas  Data  Review  View '
	const quick = quickActions(width)
	return `${THEME.ribbon}${fitAnsi(`${tabs}${quick}`, width)}${THEME.reset}`
}

function quickActions(width: number): string {
	if (width < 72) return '  / Search'
	const actions = [
		...commandsForGroup('file').slice(0, 3),
		...commandsForGroup('home').slice(0, 4),
		...commandsForGroup('data').slice(0, 2),
	]
		.map((command) => {
			const key = command.excelKeys[0] ?? command.fallbackKeys[0] ?? command.id
			return `${command.title} ${key}`
		})
		.join('   ')
	return `  |  ${actions}   / Search`
}
