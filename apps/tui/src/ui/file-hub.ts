import { fitAnsi, sanitizeTerminalText } from '../render/ansi-text.ts'
import { THEME } from '../render/styles.ts'
import type { FileHubEntry, FileHubState } from '../runtime/types.ts'

export function renderFileHub(
	state: FileHubState,
	width: number,
	height: number,
): readonly string[] {
	const lines: string[] = []
	lines.push(`${THEME.ribbon}${fitAnsi(' File backstage', width)}${THEME.reset}`)
	lines.push(fitAnsi(sectionTabs(state.section), width))
	lines.push(fitAnsi(sectionTitle(state.section), width))
	const entries = entriesForSection(state)
	for (const [index, entry] of entries.slice(0, Math.max(0, height - 7)).entries()) {
		const active = index === clampSelection(state.selectedIndex, entries.length)
		const marker = active ? `${THEME.active}>${THEME.reset} ` : '  '
		const pin = entry.pinned ? '* ' : '  '
		const missing = entry.missing ? ' missing' : ''
		const detail = entry.detail ? `  ${entry.detail}` : ''
		lines.push(
			fitAnsi(
				`${marker}${pin}${sanitizeTerminalText(entry.label)}${sanitizeTerminalText(detail)}${missing}`,
				width,
			),
		)
	}
	lines.push(''.padEnd(width))
	lines.push(fitAnsi(sectionFooter(state.section), width))
	return lines.slice(0, height)
}

const DEFAULT_ENTRIES: readonly FileHubEntry[] = [
	{ label: 'Open a workbook...', detail: 'Ctrl+O' },
	{ label: 'Create a blank workbook', detail: 'New' },
	{ label: 'Recover workbooks', detail: 'AutoRecover snapshots' },
]

function sectionTabs(active: FileHubState['section']): string {
	const tabs: readonly FileHubState['section'][] = [
		'recent',
		'open',
		'new',
		'saveAs',
		'export',
		'recover',
		'info',
	]
	return tabs.map((tab) => (tab === active ? `[${tabLabel(tab)}]` : ` ${tabLabel(tab)} `)).join(' ')
}

function sectionTitle(section: FileHubState['section']): string {
	switch (section) {
		case 'open':
			return 'Open: choose a recent file, paste a path, or run :open <path>.'
		case 'new':
			return 'New: start a blank workbook without leaving the TUI.'
		case 'saveAs':
			return 'Save As: write this workbook to a new .xlsx path.'
		case 'export':
			return 'Export: write CSV, TSV, or JSON while preserving the workbook in memory.'
		case 'recover':
			return 'Recover: inspect AutoRecover snapshots before replacing current work.'
		case 'info':
			return 'Info: workbook metadata, protection, compatibility, and preservation status.'
		default:
			return 'Recent Workbooks'
	}
}

function entriesForSection(state: FileHubState): readonly FileHubEntry[] {
	if (state.section === 'recent') return state.entries.length > 0 ? state.entries : DEFAULT_ENTRIES
	if (state.section === 'open') {
		return [
			{ label: 'Open workbook path', detail: ':open <path> or Ctrl+O' },
			{ label: 'Browse recent workbooks', detail: 'Recent section' },
			{ label: 'Open without discarding dirty work', detail: 'Save first, or use :open! <path>' },
		]
	}
	if (state.section === 'new') {
		return [
			{ label: 'Blank workbook', detail: ':new' },
			{ label: 'Discard and create blank workbook', detail: ':new!' },
			{ label: 'Template gallery', detail: 'planned' },
		]
	}
	if (state.section === 'saveAs') {
		return [
			{ label: 'Save as Excel workbook', detail: ':save-as /path/book.xlsx' },
			{ label: 'Atomic write with source guard', detail: 'preserves existing workbook state' },
			{ label: 'Unsaved state remains visible', detail: 'status bar shows Saved or Unsaved' },
		]
	}
	if (state.section === 'export') {
		return [
			{ label: 'Export CSV', detail: ':export /path/sheet.csv' },
			{ label: 'Export TSV', detail: ':export /path/sheet.tsv' },
			{ label: 'Export JSON', detail: ':export {"path":"out.json","format":"json"}' },
		]
	}
	if (state.section === 'recover') {
		return [
			{ label: 'Recover snapshots', detail: ':recover' },
			{ label: 'Review before restore', detail: 'planned durable snapshot browser' },
			{ label: 'Dirty-work warning', detail: 'save or force before replacement' },
		]
	}
	return [
		{ label: 'Workbook info', detail: ':objects for charts, pivots, drawings, comments' },
		{ label: 'Compatibility', detail: 'preservation capsules and protected review' },
		{ label: 'Performance', detail: ':perf and --telemetry-json' },
	]
}

function sectionFooter(section: FileHubState['section']): string {
	const shared = 'Esc Back   Ctrl+P Commands   F1 Help'
	switch (section) {
		case 'saveAs':
			return `Type :save-as <path>   F12 Save As   ${shared}`
		case 'export':
			return `Type :export <path>   Ctrl+Shift+S Export   ${shared}`
		case 'open':
			return `Type :open <path>   Ctrl+O Open   ${shared}`
		default:
			return `Up/Down Select   Enter Choose   ${shared}`
	}
}

function tabLabel(section: FileHubState['section']): string {
	return section === 'saveAs' ? 'Save As' : section[0]?.toUpperCase() + section.slice(1)
}

function clampSelection(index: number, length: number): number {
	if (length <= 0) return 0
	return Math.min(length - 1, Math.max(0, index))
}
