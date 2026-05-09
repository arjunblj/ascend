import { describe, expect, test } from 'bun:test'
import { CHART_COMMANDS } from './charts.ts'
import { DATA_COMMANDS } from './data.ts'
import { EDIT_COMMANDS } from './edit.ts'
import { FILE_COMMANDS } from './file.ts'
import { FORMAT_COMMANDS } from './format.ts'
import { FORMULA_COMMANDS } from './formula.ts'
import { PIVOT_COMMANDS } from './pivots.ts'
import { commandsForGroup, findCommand } from './registry.ts'
import { REVIEW_COMMANDS } from './review.ts'
import { TABLE_COMMANDS } from './tables.ts'
import { VIEW_COMMANDS } from './view.ts'

describe('command registry', () => {
	test('finds commands by Excel-like fallback text', () => {
		expect(findCommand('table create')?.id).toBe('insert.table')
		expect(findCommand('comment')?.id).toBe('review.comment')
		expect(findCommand('chart')?.id).toBe('insert.chart')
		expect(findCommand('pivot')?.id).toBe('data.pivotFields')
		expect(findCommand('print')?.id).toBe('file.printPreview')
		expect(findCommand('find')?.id).toBe('home.findReplace')
		expect(findCommand('replace')?.id).toBe('home.findReplace')
		expect(findCommand('show formulas')?.id).toBe('view.showFormulas')
		expect(findCommand('objects')?.id).toBe('view.objects')
		expect(findCommand('charts')?.id).toBe('view.objects')
		expect(findCommand('pivots')?.id).toBe('view.objects')
	})

	test('scopes command modules to their ribbon groups', () => {
		expect(FILE_COMMANDS.every((command) => command.group === 'file')).toBe(true)
		expect(DATA_COMMANDS.every((command) => command.group === 'data')).toBe(true)
		expect(FORMULA_COMMANDS.every((command) => command.group === 'formulas')).toBe(true)
		expect(REVIEW_COMMANDS.every((command) => command.group === 'review')).toBe(true)
		expect(VIEW_COMMANDS.every((command) => command.group === 'view')).toBe(true)
		expect(EDIT_COMMANDS.some((command) => command.id === 'home.copy')).toBe(true)
		expect(EDIT_COMMANDS.some((command) => command.id === 'home.findReplace')).toBe(true)
		expect(FORMAT_COMMANDS.map((command) => command.id)).toEqual(['home.formatCells'])
		expect(TABLE_COMMANDS.map((command) => command.id)).toEqual(['insert.table'])
		expect(CHART_COMMANDS.map((command) => command.id)).toEqual(['insert.chart'])
		expect(PIVOT_COMMANDS.map((command) => command.id)).toEqual(['data.pivotFields'])
	})

	test('commandsForGroup returns only the requested group', () => {
		expect(commandsForGroup('home').every((command) => command.group === 'home')).toBe(true)
	})
})
