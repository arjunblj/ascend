import { ascendError, type CellValue } from '@ascend/schema'
import { cliError, jsonOut } from '../output/json.ts'
import { formatCellValue } from '../output/pretty.ts'
import { openWorkbookWithProgress } from '../progress.ts'

export const usage = `Usage: ascend find <file> <query> [flags]

  Search for cells matching a value.

Arguments:
  <file>              Path to the workbook file
  <query>             Value to search for (string, number, or true/false)

Flags:
  --sheet <name>      Sheet name (defaults to first sheet)
  --match <mode>      Match mode: exact, contains, startsWith, endsWith (default: exact)
  --json              Output as JSON
`

type MatchMode = 'exact' | 'contains' | 'startsWith' | 'endsWith'

function parseMatchMode(value: string | undefined): MatchMode | null {
	if (!value) return 'exact'
	switch (value) {
		case 'exact':
		case 'contains':
		case 'startsWith':
		case 'endsWith':
			return value
		default:
			return null
	}
}

function parseQuery(query: string): string | number | boolean {
	const lower = query.toLowerCase()
	if (lower === 'true') return true
	if (lower === 'false') return false
	const num = Number(query)
	if (!Number.isNaN(num) && String(num) === query) return num
	return query
}

export async function findCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	const queryArg = args[1]
	if (!file || !queryArg) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing required find input', {
				retryable: true,
				retryStrategy: 'modified',
				details: {
					command: 'find',
					required: ['file', 'query'],
					missing: [...(!file ? ['file'] : []), ...(!queryArg ? ['query'] : [])],
					workflow: ['inspect', 'read'],
				},
				suggestedFix:
					'Run ascend find <file> <query> --sheet <name> --json after inspecting workbook sheets.',
			}),
			flags,
		)
		return 1
	}

	const matchMode = parseMatchMode(flags.get('match'))
	if (flags.has('match') && matchMode === null) {
		cliError('Invalid --match. Use one of: exact, contains, startsWith, endsWith', flags)
		return 1
	}

	const sheetFlag = flags.get('sheet')
	const { workbook: wb } = await openWorkbookWithProgress(file, {
		mode: 'values',
		...(sheetFlag ? { sheets: [sheetFlag] } : {}),
	})

	const sheetName = flags.get('sheet') ?? wb.sheets[0]
	if (!sheetName) {
		cliError('No sheets in workbook', flags)
		return 1
	}

	const query = parseQuery(queryArg)
	const matches = wb.find(sheetName, {
		value: query,
		match: matchMode ?? 'exact',
	})

	if (flags.has('json')) {
		console.log(
			jsonOut({
				sheet: sheetName,
				query: queryArg,
				match: matchMode ?? 'exact',
				count: matches.length,
				matches: matches.map((m) => ({ ref: m.ref, value: cellValueToJson(m.value) })),
			}),
		)
		return 0
	}

	if (matches.length === 0) {
		console.log(`No matches for "${queryArg}" in ${sheetName}`)
		return 0
	}

	for (const m of matches) {
		const display = formatCellValue(m.value, { display: true })
		console.log(`${sheetName}!${m.ref}\t${display}`)
	}
	console.log(`\n${matches.length} match(es) found`)
	return 0
}

function cellValueToJson(v: CellValue): unknown {
	switch (v.kind) {
		case 'empty':
			return null
		case 'number':
		case 'string':
		case 'boolean':
		case 'error':
			return v.value
		case 'date':
			return { kind: 'date', serial: v.serial }
		case 'richText':
			return v.runs.map((r) => r.text).join('')
		default:
			return null
	}
}
