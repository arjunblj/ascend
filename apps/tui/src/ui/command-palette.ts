import fuzzysort from 'fuzzysort'
import { listCommands } from '../commands/registry.ts'
import { fitAnsi } from '../render/ansi-text.ts'
import { THEME } from '../render/styles.ts'
import type { CommandDescriptor, CommandPaletteState } from '../runtime/types.ts'

export function renderCommandPalette(
	state: CommandPaletteState,
	width: number,
	height: number,
): readonly string[] {
	const query = state.query
	const maxResults = Math.max(0, height - 2)
	const commands = commandPaletteResults(query, maxResults)
	const lines = [fitAnsi(`Search commands: ${query}`, width)]
	for (const [index, command] of commands.entries()) {
		const marker = index === state.selectedIndex ? `${THEME.active}>${THEME.reset}` : ' '
		const excel = command.excelKeys.length > 0 ? `  ${command.excelKeys.join(', ')}` : ''
		const fallback = command.fallbackKeys.length > 0 ? `  ${command.fallbackKeys.join(', ')}` : ''
		const hint = command.dialogId ? '  dialog' : ''
		lines.push(fitAnsi(`${marker} ${command.title}${excel}${fallback}${hint}`, width))
	}
	if (commands.length === 0) lines.push(fitAnsi('  No matching commands', width))
	lines.push(fitAnsi('Up/Down Select  Enter Run  Esc Close', width))
	return lines
}

export function commandPaletteResults(
	query: string,
	limit = Number.POSITIVE_INFINITY,
): readonly CommandDescriptor[] {
	const maxResults = Number.isFinite(limit) ? Math.max(0, limit) : undefined
	if (query.trim() === '') return listCommands().slice(0, maxResults)
	return fuzzysort
		.go(query, listCommands(), {
			keys: ['title', 'id', (command) => command.fallbackKeys.join(' ')],
			...(maxResults !== undefined ? { limit: maxResults } : {}),
			all: true,
		})
		.map((result) => result.obj)
}
