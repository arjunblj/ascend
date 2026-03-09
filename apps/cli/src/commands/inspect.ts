import type { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'
import { bullet, heading, table } from '../output/pretty.ts'
import { openWorkbookWithProgress } from '../progress.ts'

export const usage = `Usage: ascend inspect <file> [sheet] [flags]

  Inspect workbook structure and sheet details.

Arguments:
  <file>          Path to the workbook file
  [sheet]         Optional sheet name to inspect

Flags:
  --sheet <name>  Sheet name (alternative to positional argument)
  --detail <type> Show detail for: cf, dv, hyperlinks, tables, comments, compatibility
  --mode <mode>   Load mode: metadata, values, or full
  --json          Output as JSON
  --verbose       Show compatibility report and timing
`

export async function inspectCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		console.error('Usage: ascend inspect <file> [sheet]')
		return 1
	}

	const sheetArg = args[1] ?? flags.get('sheet')
	const detail = flags.get('detail')
	const verbose = flags.has('verbose')
	const workbookDetail = detail === 'compatibility'
	const parsedMode = parseInspectMode(flags.get('mode'))
	if (flags.has('mode') && parsedMode === null) {
		console.error('Invalid --mode. Use one of: metadata, values, full')
		return 1
	}
	const explicitMode = parsedMode ?? undefined
	if (detail && sheetArg && explicitMode && explicitMode !== 'full') {
		console.error('Sheet detail views require --mode full')
		return 1
	}

	const openOptions =
		explicitMode !== undefined
			? {
					mode: explicitMode,
					...(sheetArg ? { sheets: [sheetArg] } : {}),
				}
			: workbookDetail
				? { mode: 'metadata-only' as const }
				: detail && sheetArg
					? { sheets: [sheetArg] }
					: sheetArg
						? { mode: 'values' as const, sheets: [sheetArg] }
						: { mode: 'metadata-only' as const }
	const { workbook: wb, durationMs: openMs } = await openWorkbookWithProgress(file, openOptions)
	const info = wb.inspect()

	if (detail === 'compatibility') {
		return printCompatibilityDetail(wb, flags.has('json'))
	}

	if (detail && sheetArg) {
		return printSheetDetail(wb, sheetArg, detail, flags.has('json'))
	}

	if (detail) {
		console.error(
			'Sheet detail requires a sheet name. Use: ascend inspect <file> <sheet> --detail <type>',
		)
		return 1
	}

	if (sheetArg) {
		const sheet = info.sheets.find((s) => s.name === sheetArg)
		if (!sheet) {
			console.error(`Sheet "${sheetArg}" not found`)
			return 1
		}
		if (flags.has('json')) {
			console.log(jsonOut(sheet))
		} else {
			console.log(heading(`Sheet: ${sheet.name}`))
			console.log(bullet('Cell data loaded', sheet.cellDataLoaded ? 'yes' : 'no'))
			console.log(bullet('Rows', formatCount(sheet.rowCount)))
			console.log(bullet('Columns', formatCount(sheet.colCount)))
			console.log(bullet('Cells', formatCount(sheet.cellCount)))
			console.log(bullet('Tables', formatCount(sheet.tableCount)))
			console.log(bullet('Comments', formatCount(sheet.commentCount)))
			console.log(bullet('Conditional formats', formatCount(sheet.conditionalFormatCount)))
			console.log(bullet('Data validations', formatCount(sheet.dataValidationCount)))
			console.log(bullet('Images', formatCount(sheet.imageCount)))
			console.log(bullet('Frozen panes', formatLoadedBool(sheet.hasFrozenPanes)))
			console.log(bullet('Column widths', formatCount(sheet.colWidthCount)))
			console.log(bullet('Row heights', formatCount(sheet.rowHeightCount)))
			console.log(bullet('Hyperlinks', formatCount(sheet.hyperlinkCount)))
			console.log(bullet('Ignored errors', formatCount(sheet.ignoredErrorCount)))
			console.log(bullet('Auto filter', formatLoadedBool(sheet.hasAutoFilter)))
			console.log(bullet('Drawing refs', formatLoadedBool(sheet.hasDrawingRefs)))
			console.log(bullet('Page metadata', formatLoadedBool(sheet.hasPageMetadata)))
			console.log(bullet('Protection', formatLoadedBool(sheet.hasProtection)))
		}
		return 0
	}

	if (flags.has('json')) {
		const out = verbose ? { ...info, timing: { openMs: Math.round(openMs * 100) / 100 } } : info
		console.log(jsonOut(out))
		return 0
	}

	console.log(heading(`Workbook: ${file}`))
	console.log(bullet('Format', info.sourceFormat))
	console.log(bullet('Source sheets', info.sheetCount))
	console.log(bullet('Loaded sheets', info.loadedSheetCount))
	console.log(bullet('Load mode', info.load.mode))
	console.log(bullet('Partial view', info.load.isPartial ? 'yes' : 'no'))
	console.log(bullet('Cell data loaded', info.load.cellsHydrated ? 'yes' : 'no'))
	console.log(bullet('Total cells', formatCount(info.cellCount)))
	console.log(bullet('Comments', formatCount(info.commentCount)))
	console.log(bullet('Conditional formats', formatCount(info.conditionalFormatCount)))
	console.log(bullet('Data validations', formatCount(info.dataValidationCount)))
	console.log(bullet('Images', formatCount(info.imageCount)))
	console.log(bullet('Pivot tables', info.pivotTableCount))
	console.log(bullet('Pivot caches', info.pivotCacheCount))
	console.log(bullet('Slicers', info.slicerCount))
	console.log(bullet('Slicer caches', info.slicerCacheCount))
	console.log(bullet('Workbook views', info.workbookViewCount))
	console.log(bullet('External references', info.externalReferenceCount))
	console.log(bullet('Workbook protection', info.hasWorkbookProtection ? 'yes' : 'no'))
	console.log(bullet('Cell styles', info.styleSummary.cellXfCount))
	console.log(bullet('Diff styles', info.styleSummary.dxfCount))
	console.log(bullet('Theme part', info.themeSummary.hasThemePart ? 'yes' : 'no'))
	console.log(bullet('Theme colors', info.themeSummary.colorCount))
	if (info.themeSummary.name) {
		console.log(bullet('Theme name', info.themeSummary.name))
	}
	if (info.themeSummary.majorFontLatin || info.themeSummary.minorFontLatin) {
		console.log(
			bullet(
				'Theme fonts',
				[
					info.themeSummary.majorFontLatin
						? `major=${info.themeSummary.majorFontLatin}`
						: undefined,
					info.themeSummary.minorFontLatin
						? `minor=${info.themeSummary.minorFontLatin}`
						: undefined,
				]
					.filter(Boolean)
					.join(', '),
			),
		)
	}
	console.log(bullet('Compatibility', info.compatibility.status))
	if (info.definedNames.length > 0) {
		console.log(bullet('Defined names', info.definedNames.join(', ')))
	}

	if (info.sheets.length > 0) {
		console.log('')
		console.log(
			table(
				['Sheet', 'Loaded', 'Rows', 'Cols', 'Cells', 'Tables', 'CF', 'DV', 'Images', 'Filter'],
				info.sheets.map((s) => [
					s.name,
					s.cellDataLoaded ? 'yes' : 'no',
					formatCount(s.rowCount),
					formatCount(s.colCount),
					formatCount(s.cellCount),
					formatCount(s.tableCount),
					formatCount(s.conditionalFormatCount),
					formatCount(s.dataValidationCount),
					formatCount(s.imageCount),
					formatLoadedBool(s.hasAutoFilter),
				]),
			),
		)
	}

	if (verbose) {
		console.log('')
		console.log(heading('Compatibility Details'))
		for (const f of info.compatibility.features) {
			console.log(bullet(`${f.feature} (${f.tier})`, `${f.count} location(s)`))
			if (f.note) console.log(`    ${f.note}`)
		}
		console.log('')
		console.log(heading('Timing'))
		console.log(bullet('Open', `${openMs.toFixed(1)}ms`))
	}

	return 0
}

function parseInspectMode(
	mode: string | undefined,
): 'metadata-only' | 'values' | 'full' | undefined | null {
	if (mode === undefined || mode === '') return undefined
	switch (mode) {
		case 'metadata':
			return 'metadata-only'
		case 'values':
		case 'full':
			return mode
		default:
			return null
	}
}

function printSheetDetail(
	wb: AscendWorkbook,
	sheetName: string,
	detail: string,
	json: boolean,
): number {
	const handle = wb.sheet(sheetName)
	if (!handle) {
		console.error(`Sheet "${sheetName}" not found`)
		return 1
	}
	switch (detail) {
		case 'cf': {
			const cfs = handle.conditionalFormats
			if (json) {
				console.log(jsonOut(cfs))
				return 0
			}
			console.log(heading(`Conditional Formats: ${sheetName}`))
			for (const cf of cfs) {
				console.log(bullet('Range', cf.sqref))
				for (const rule of cf.rules) {
					console.log(
						`    type=${rule.type ?? '(none)'} priority=${rule.priority} operator=${rule.operator ?? ''}`,
					)
					if (rule.formulas.length > 0) console.log(`    formulas: ${rule.formulas.join(', ')}`)
				}
			}
			if (cfs.length === 0) console.log('  (none)')
			return 0
		}
		case 'dv': {
			const dvs = handle.dataValidations
			if (json) {
				console.log(jsonOut(dvs))
				return 0
			}
			console.log(heading(`Data Validations: ${sheetName}`))
			for (const dv of dvs) {
				console.log(bullet('Range', dv.sqref ?? ''))
				console.log(`    type=${dv.type ?? '(any)'} allowBlank=${dv.allowBlank ?? false}`)
				if (dv.formula1) console.log(`    formula1: ${dv.formula1}`)
				if (dv.formula2) console.log(`    formula2: ${dv.formula2}`)
			}
			if (dvs.length === 0) console.log('  (none)')
			return 0
		}
		case 'hyperlinks': {
			const links = handle.hyperlinks()
			if (json) {
				console.log(jsonOut(Object.fromEntries(links)))
				return 0
			}
			console.log(heading(`Hyperlinks: ${sheetName}`))
			for (const [ref, link] of links) {
				console.log(bullet(ref, link.target ?? link.location ?? ''))
			}
			if (links.size === 0) console.log('  (none)')
			return 0
		}
		case 'comments': {
			const comments = handle.comments()
			if (json) {
				console.log(jsonOut(Object.fromEntries(comments)))
				return 0
			}
			console.log(heading(`Comments: ${sheetName}`))
			for (const [ref, comment] of comments) {
				console.log(bullet(ref, `${comment.author ?? ''}: ${comment.text}`))
			}
			if (comments.size === 0) console.log('  (none)')
			return 0
		}
		case 'tables': {
			const tables = wb.inspect().sheets.find((s) => s.name === sheetName)
			if (json) {
				console.log(jsonOut({ tableCount: tables?.tableCount ?? 0 }))
				return 0
			}
			console.log(heading(`Tables: ${sheetName}`))
			console.log(bullet('Count', formatCount(tables?.tableCount ?? null)))
			return 0
		}
		case 'compatibility': {
			return printCompatibilityDetail(wb, json)
		}
		default:
			console.error(
				`Unknown detail type: ${detail}. Options: cf, dv, hyperlinks, comments, tables, compatibility`,
			)
			return 1
	}
}

function printCompatibilityDetail(wb: AscendWorkbook, json: boolean): number {
	const report = wb.report
	if (json) {
		console.log(jsonOut(report))
		return 0
	}
	console.log(heading('Compatibility Report'))
	console.log(bullet('Status', report.status))
	for (const f of report.features) {
		console.log(bullet(`${f.feature} (${f.tier})`, `${f.count} location(s)`))
		if (f.note) console.log(`    ${f.note}`)
	}
	return 0
}

function formatCount(value: number | null): string {
	return value === null ? 'unknown (not hydrated)' : String(value)
}

function formatLoadedBool(value: boolean | null): string {
	if (value === null) return 'unknown (not hydrated)'
	return value ? 'yes' : 'no'
}
