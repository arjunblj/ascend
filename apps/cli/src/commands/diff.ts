import { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'
import { heading, table } from '../output/pretty.ts'

export async function diffCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const fileA = args[0]
	const fileB = args[1]
	if (!fileA || !fileB) {
		console.error('Usage: ascend diff <file-a> <file-b>')
		return 1
	}

	const [a, b] = await Promise.all([AscendWorkbook.open(fileA), AscendWorkbook.open(fileB)])
	const result = a.diff(b)

	if (flags.has('json')) {
		console.log(jsonOut(result))
		return 0
	}

	const totalChanges =
		result.namesAdded.length +
		result.namesRemoved.length +
		result.namesChanged.length +
		result.sheets.reduce((n, s) => n + s.cellsChanged.length, 0)

	if (totalChanges === 0) {
		console.log('No differences found')
		return 0
	}

	console.log(heading('Workbook Diff'))

	if (result.namesAdded.length > 0) console.log(`  Names added: ${result.namesAdded.join(', ')}`)
	if (result.namesRemoved.length > 0)
		console.log(`  Names removed: ${result.namesRemoved.join(', ')}`)

	for (const sheet of result.sheets) {
		if (sheet.cellsChanged.length === 0) continue
		console.log(`\n  Sheet "${sheet.name}": ${sheet.cellsChanged.length} cell(s) changed`)
		console.log(
			table(
				['Cell', 'Before', 'After'],
				sheet.cellsChanged.map((c) => [c.ref, summarizeValue(c.before), summarizeValue(c.after)]),
			),
		)
	}

	return 0
}

function summarizeValue(v: unknown): string {
	if (v === null || v === undefined) return '(empty)'
	if (typeof v === 'object' && 'kind' in v) {
		const cv = v as { kind: string; value?: unknown; serial?: number }
		if (cv.kind === 'empty') return '(empty)'
		if (cv.value !== undefined) return String(cv.value)
		if (cv.serial !== undefined) return `[date:${cv.serial}]`
	}
	return String(v)
}
