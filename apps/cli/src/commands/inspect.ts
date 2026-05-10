import type { WorkbookDocument } from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { bullet, formatCellValue, heading, table } from '../output/pretty.ts'
import { openWorkbookDocumentWithProgress } from '../progress.ts'

export const usage = `Usage: ascend inspect <file> [sheet] [flags]

  Inspect workbook structure and sheet details.

Arguments:
  <file>          Path to the workbook file
  [sheet]         Optional sheet name to inspect

Flags:
  --sheet <name>  Sheet name (alternative to positional argument)
  --detail <type> Show detail for: cf, dv, hyperlinks, tables, comments, drawings, images, compatibility, visuals, pivots, slicers, names, external-refs, views
  --mode <mode>   Load mode: metadata, values, or full
  --json          Output as JSON
  --verbose       Show compatibility report and timing
`

export async function inspectCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		cliError('Usage: ascend inspect <file> [sheet]', flags)
		return 1
	}

	const sheetArg = args[1] ?? flags.get('sheet')
	const detail = flags.get('detail')
	const verbose = flags.has('verbose')
	const workbookDetail = detail === 'compatibility'
	const workbookStructureDetail =
		detail === 'pivots' ||
		detail === 'slicers' ||
		detail === 'names' ||
		detail === 'external-refs' ||
		detail === 'views' ||
		detail === 'visuals'
	const parsedMode = parseInspectMode(flags.get('mode'))
	if (flags.has('mode') && parsedMode === null) {
		cliError('Invalid --mode. Use one of: metadata, values, full', flags)
		return 1
	}
	const explicitMode = parsedMode ?? undefined
	if (detail && sheetArg && explicitMode && explicitMode !== 'full') {
		cliError('Sheet detail views require --mode full', flags)
		return 1
	}

	const openOptions =
		explicitMode !== undefined
			? {
					mode: explicitMode,
					...(sheetArg ? { sheets: [sheetArg] } : {}),
				}
			: workbookDetail
				? { mode: 'full' as const }
				: workbookStructureDetail
					? { mode: 'full' as const }
					: detail && sheetArg
						? { sheets: [sheetArg] }
						: sheetArg
							? { mode: 'values' as const, sheets: [sheetArg] }
							: { mode: 'metadata-only' as const }
	const { document: wb, durationMs: openMs } = await openWorkbookDocumentWithProgress(
		file,
		openOptions,
	)

	if (detail === 'compatibility') {
		return printCompatibilityDetail(wb, flags.has('json'))
	}

	if (detail === 'pivots') {
		return printPivotDetail(wb, flags.has('json'))
	}

	if (detail === 'slicers') {
		return printSlicerDetail(wb, flags.has('json'))
	}

	if (detail === 'names') {
		return printNamesDetail(wb, flags.has('json'))
	}

	if (detail === 'external-refs') {
		return printExternalRefsDetail(wb, flags.has('json'))
	}

	if (detail === 'views') {
		return printWorkbookViewsDetail(wb, flags.has('json'))
	}

	if (detail === 'visuals') {
		return printVisualInventoryDetail(wb, flags.has('json'))
	}

	if (detail && sheetArg) {
		return printSheetDetail(wb, sheetArg, detail, flags.has('json'), flags)
	}

	if (detail) {
		cliError(
			'Sheet detail requires a sheet name. Use: ascend inspect <file> <sheet> --detail <type>',
			flags,
		)
		return 1
	}

	if (sheetArg) {
		const sheet = wb.inspectSheet(sheetArg)
		if (!sheet) {
			cliError(`Sheet "${sheetArg}" not found`, flags)
			return 1
		}
		if (flags.has('json')) {
			console.log(jsonOut(sheet))
		} else {
			console.log(heading(`Sheet: ${sheet.name}`))
			console.log(bullet('Cell data loaded', sheet.cellDataLoaded ? 'yes' : 'no'))
			console.log(
				bullet(
					'Rich sheet metadata loaded',
					sheet.commentCount !== null ||
						sheet.conditionalFormatCount !== null ||
						sheet.dataValidationCount !== null ||
						sheet.imageCount !== null ||
						sheet.hyperlinkCount !== null
						? 'yes'
						: 'no',
				),
			)
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
			console.log(bullet('Auto filter ref', sheet.autoFilter?.ref ?? 'none'))
			console.log(bullet('Drawing refs', formatLoadedBool(sheet.hasDrawingRefs)))
			console.log(
				bullet(
					'Drawing detail',
					sheet.drawingRefs
						? `drawing=${sheet.drawingRefs.hasDrawing ? 'yes' : 'no'}, legacy=${sheet.drawingRefs.hasLegacyDrawing ? 'yes' : 'no'}`
						: 'unknown (not hydrated)',
				),
			)
			console.log(bullet('Page metadata', formatLoadedBool(sheet.hasPageMetadata)))
			console.log(
				bullet(
					'Page detail',
					sheet.cellDataLoaded
						? [
								sheet.pageMargins ? 'margins' : undefined,
								sheet.pageSetup ? 'setup' : undefined,
								sheet.printOptions ? 'print-options' : undefined,
								sheet.headerFooter ? 'header-footer' : undefined,
							]
								.filter(Boolean)
								.join(', ') || 'none'
						: 'unknown (not hydrated)',
				),
			)
			console.log(bullet('Protection', formatLoadedBool(sheet.hasProtection)))
			console.log(
				bullet('Merges', sheet.merges ? String(sheet.merges.length) : 'unknown (not hydrated)'),
			)
		}
		return 0
	}

	if (flags.has('json')) {
		const info = wb.inspect()
		const out = verbose ? { ...info, timing: { openMs: Math.round(openMs * 100) / 100 } } : info
		console.log(jsonOut(out))
		return 0
	}

	const info = wb.inspect()
	console.log(heading(`Workbook: ${file}`))
	console.log(bullet('Format', info.sourceFormat))
	console.log(bullet('Source sheets', info.sheetCount))
	console.log(bullet('Loaded sheets', info.loadedSheetCount))
	console.log(bullet('Load mode', info.load.mode))
	console.log(bullet('Partial view', info.load.isPartial ? 'yes' : 'no'))
	console.log(bullet('Cell data loaded', info.load.cellsHydrated ? 'yes' : 'no'))
	console.log(
		bullet('Rich sheet metadata loaded', info.load.richSheetMetadataHydrated ? 'yes' : 'no'),
	)
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
	wb: WorkbookDocument,
	sheetName: string,
	detail: string,
	json: boolean,
	flags: Map<string, string>,
): number {
	const handle = wb.sheet(sheetName)
	if (!handle) {
		cliError(`Sheet "${sheetName}" not found`, flags)
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
		case 'drawings': {
			const detailInfo = wb.inspectSheet(sheetName)
			if (json) {
				console.log(
					jsonOut({
						drawingRefs: detailInfo?.drawingRefs ?? null,
					}),
				)
				return 0
			}
			console.log(heading(`Drawing Refs: ${sheetName}`))
			if (!detailInfo?.drawingRefs) {
				console.log('  unknown (not hydrated)')
				return 0
			}
			console.log(bullet('Drawing', detailInfo.drawingRefs.hasDrawing ? 'yes' : 'no'))
			console.log(bullet('Legacy drawing', detailInfo.drawingRefs.hasLegacyDrawing ? 'yes' : 'no'))
			return 0
		}
		case 'images': {
			const images = handle.imageRefs
			if (json) {
				console.log(jsonOut(images))
				return 0
			}
			console.log(heading(`Images: ${sheetName}`))
			for (const image of images) {
				console.log(
					bullet(
						image.name ?? image.targetPath,
						[
							image.targetPath,
							image.description,
							image.anchor?.kind ? `anchor=${image.anchor.kind}` : undefined,
						]
							.filter(Boolean)
							.join(' | '),
					),
				)
			}
			if (images.length === 0) console.log('  (none)')
			return 0
		}
		case 'tables': {
			const sheetInfo = wb.inspectSheet(sheetName)
			const tables = sheetInfo?.tables ?? null
			if (json) {
				console.log(jsonOut({ tableCount: tables?.length ?? null, tables }))
				return 0
			}
			console.log(heading(`Tables: ${sheetName}`))
			console.log(
				bullet('Count', tables === null ? 'unknown (not hydrated)' : String(tables.length)),
			)
			for (const tableInfo of tables ?? []) {
				console.log(
					bullet(
						tableInfo.name,
						`${formatRange(tableInfo.ref)} rows=${tableInfo.rowCount} headers=${tableInfo.hasHeaders ? 'yes' : 'no'} totals=${tableInfo.hasTotals ? 'yes' : 'no'}`,
					),
				)
				if (tableInfo.styleInfo?.name) {
					console.log(`    style: ${tableInfo.styleInfo.name}`)
				}
				if (tableInfo.autoFilter?.ref) {
					console.log(`    filter: ${tableInfo.autoFilter.ref}`)
				}
				if (tableInfo.sortState?.ref) {
					console.log(`    sort: ${tableInfo.sortState.ref}`)
				}
				if (tableInfo.headerRow) {
					console.log(
						`    headers: ${tableInfo.headerRow.map((value) => formatCellValue(value)).join(' | ')}`,
					)
				}
				if (tableInfo.totalsRow) {
					console.log(
						`    totals: ${tableInfo.totalsRow.map((value) => formatCellValue(value)).join(' | ')}`,
					)
				}
				console.log(
					`    columns: ${tableInfo.columnDefs.map((column) => column.name).join(', ') || '(none)'}`,
				)
			}
			return 0
		}
		case 'compatibility': {
			return printCompatibilityDetail(wb, json)
		}
		default:
			cliError(
				`Unknown detail type: ${detail}. Options: cf, dv, hyperlinks, comments, drawings, images, tables, compatibility, visuals, pivots, slicers, names, external-refs, views`,
				flags,
			)
			return 1
	}
}

function printCompatibilityDetail(wb: WorkbookDocument, json: boolean): number {
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

function printVisualInventoryDetail(wb: WorkbookDocument, json: boolean): number {
	const inventory = wb.visualInventory()
	if (json) {
		console.log(jsonOut(inventory))
		return 0
	}
	console.log(heading('Visual Inventory'))
	console.log(bullet('Sheet images', formatCount(inventory.sheetImageCount)))
	console.log(bullet('Package chart features', String(inventory.packageChartFeatureCount)))
	console.log(bullet('Package drawing features', String(inventory.packageDrawingFeatureCount)))
	console.log(bullet('Package media features', String(inventory.packageMediaFeatureCount)))
	if (inventory.packageFeatures.length > 0) {
		console.log('')
		console.log(heading('Package Features'))
		for (const feature of inventory.packageFeatures) {
			console.log(
				bullet(
					`${feature.feature} (${feature.category}, ${feature.tier})`,
					`${feature.count} location(s)`,
				),
			)
		}
	}
	if (inventory.sheets.length > 0) {
		console.log('')
		console.log(
			table(
				['Sheet', 'Drawing', 'Legacy', 'Images'],
				inventory.sheets.map((sheet) => [
					sheet.sheet,
					formatLoadedBool(sheet.hasDrawing),
					formatLoadedBool(sheet.hasLegacyDrawing),
					formatCount(sheet.imageCount),
				]),
			),
		)
	}
	if (inventory.notes.length > 0) {
		console.log('')
		console.log(heading('Notes'))
		for (const note of inventory.notes) console.log(bullet('-', note))
	}
	return 0
}

function printPivotDetail(wb: WorkbookDocument, json: boolean): number {
	const workbookInfo = wb.inspect()
	if (json) {
		console.log(
			jsonOut({
				pivotTables: wb.pivotTables(),
				pivotCaches: wb.pivotCaches(),
				pivotRefreshPlans: wb.pivotRefreshPlans(),
			}),
		)
		return 0
	}
	console.log(heading('Pivot Tables'))
	console.log(bullet('Count', String(workbookInfo.pivotTableCount)))
	for (const pivot of wb.pivotTables()) {
		console.log(
			bullet(
				pivot.name ?? pivot.partPath,
				[
					pivot.sheetName,
					pivot.locationRef,
					pivot.cacheId !== undefined ? `cache=${pivot.cacheId}` : undefined,
				]
					.filter(Boolean)
					.join(' | '),
			),
		)
	}
	if (workbookInfo.pivotTableCount === 0) console.log('  (none)')
	if (wb.pivotCaches().length > 0) {
		console.log('')
		console.log(heading('Pivot Caches'))
		for (const cache of wb.pivotCaches()) {
			console.log(
				bullet(
					cache.partPath,
					[
						cache.sourceSheet ? `sheet=${cache.sourceSheet}` : undefined,
						cache.sourceRef ? `ref=${cache.sourceRef}` : undefined,
						cache.cacheId !== undefined ? `cache=${cache.cacheId}` : undefined,
					]
						.filter(Boolean)
						.join(' | '),
				),
			)
		}
	}
	const refreshPlans = wb.pivotRefreshPlans()
	if (refreshPlans.length > 0) {
		console.log('')
		console.log(heading('Pivot Refresh Plans'))
		for (const plan of refreshPlans) {
			console.log(
				bullet(
					plan.partPath,
					[
						plan.outputState,
						plan.requiresExternalRefresh ? 'external refresh required' : undefined,
						plan.pivotTables.length > 0
							? `pivots=${plan.pivotTables.map((pivot) => pivot.name ?? pivot.partPath).join(', ')}`
							: undefined,
					]
						.filter(Boolean)
						.join(' | '),
				),
			)
			for (const warning of plan.warnings) console.log(`    warning: ${warning}`)
		}
	}
	return 0
}

function printSlicerDetail(wb: WorkbookDocument, json: boolean): number {
	const workbookInfo = wb.inspect()
	if (json) {
		console.log(
			jsonOut({
				slicerCaches: wb.slicerCaches(),
				slicers: wb.slicers(),
				timelineCaches: wb.timelineCaches(),
				timelines: wb.timelines(),
			}),
		)
		return 0
	}
	console.log(heading('Slicers'))
	console.log(bullet('Count', String(workbookInfo.slicerCount)))
	for (const slicer of wb.slicers()) {
		console.log(
			bullet(
				slicer.name ?? slicer.partPath,
				[slicer.cacheName, slicer.caption].filter(Boolean).join(' | '),
			),
		)
	}
	if (workbookInfo.slicerCount === 0) console.log('  (none)')
	if (wb.slicerCaches().length > 0) {
		console.log('')
		console.log(heading('Slicer Caches'))
		for (const cache of wb.slicerCaches()) {
			console.log(
				bullet(
					cache.name ?? cache.partPath,
					[
						cache.sourceName,
						cache.pivotCacheId !== undefined ? `pivotCache=${cache.pivotCacheId}` : undefined,
						cache.pivotTableNames.length > 0
							? `pivots=${cache.pivotTableNames.join(', ')}`
							: undefined,
					]
						.filter(Boolean)
						.join(' | '),
				),
			)
		}
	}
	if (wb.timelines().length > 0) {
		console.log('')
		console.log(heading('Timelines'))
		for (const timeline of wb.timelines()) {
			console.log(
				bullet(
					timeline.name ?? timeline.partPath,
					[timeline.cacheName, timeline.caption].filter(Boolean).join(' | '),
				),
			)
		}
	}
	if (wb.timelineCaches().length > 0) {
		console.log('')
		console.log(heading('Timeline Caches'))
		for (const cache of wb.timelineCaches()) {
			console.log(
				bullet(
					cache.name ?? cache.partPath,
					[
						cache.sourceName,
						cache.pivotCacheId !== undefined ? `pivotCache=${cache.pivotCacheId}` : undefined,
						cache.pivotTableNames.length > 0
							? `pivots=${cache.pivotTableNames.join(', ')}`
							: undefined,
					]
						.filter(Boolean)
						.join(' | '),
				),
			)
		}
	}
	return 0
}

function printNamesDetail(wb: WorkbookDocument, json: boolean): number {
	const names = wb.definedNames()
	if (json) {
		console.log(jsonOut(names))
		return 0
	}
	console.log(heading('Defined Names'))
	for (const entry of names) {
		const scope = entry.scope === 'sheet' ? `sheet=${entry.sheet ?? '(unknown)'}` : 'workbook'
		console.log(bullet(entry.name, `${scope} | ${entry.normalizedFormula}`))
		if (entry.references.length > 0) {
			console.log(`    refs: ${entry.references.map((reference) => reference.text).join(', ')}`)
		}
		if (entry.functions.length > 0) {
			console.log(`    functions: ${entry.functions.join(', ')}`)
		}
		if (entry.parseError) {
			console.log(`    parse-error: ${entry.parseError}`)
		}
	}
	if (names.length === 0) console.log('  (none)')
	return 0
}

function printExternalRefsDetail(wb: WorkbookDocument, json: boolean): number {
	const refs = wb.externalReferences()
	const info = wb.inspect()
	const usages = wb.externalReferenceUsages()
	if (json) {
		console.log(
			jsonOut({
				references: refs,
				details: info.externalReferenceDetails,
				usages,
			}),
		)
		return 0
	}
	console.log(heading('External References'))
	for (const detail of info.externalReferenceDetails) {
		console.log(
			bullet(
				detail.partPath,
				[
					detail.target,
					detail.relId ? `rel=${detail.relId}` : undefined,
					detail.targetMode ? `mode=${detail.targetMode}` : undefined,
				]
					.filter(Boolean)
					.join(' | '),
			),
		)
	}
	for (const ref of refs.filter(
		(ref) => !info.externalReferenceDetails.some((d) => d.partPath === ref),
	)) {
		console.log(bullet(ref, ref))
	}
	if (refs.length === 0) console.log('  (none)')
	if (usages.length > 0) {
		console.log('')
		console.log(heading('Formula Usages'))
		for (const usage of usages) {
			console.log(
				bullet(
					usage.sourceRef ?? usage.name ?? usage.sourceKind,
					`${usage.workbook}${usage.sheet ? `:${usage.sheet}` : ''} | ${usage.references.join(', ')}`,
				),
			)
		}
	}
	return 0
}

function printWorkbookViewsDetail(wb: WorkbookDocument, json: boolean): number {
	const views = wb.workbookViews()
	if (json) {
		console.log(jsonOut(views))
		return 0
	}
	console.log(heading('Workbook Views'))
	for (const [index, view] of views.entries()) {
		console.log(
			bullet(
				`View ${index + 1}`,
				[
					view.activeTab !== undefined ? `activeTab=${view.activeTab}` : undefined,
					view.firstSheet !== undefined ? `firstSheet=${view.firstSheet}` : undefined,
					view.visibility,
					view.tabRatio !== undefined ? `tabRatio=${view.tabRatio}` : undefined,
				]
					.filter(Boolean)
					.join(' | '),
			),
		)
	}
	if (views.length === 0) console.log('  (none)')
	return 0
}

function formatCount(value: number | null): string {
	return value === null ? 'unknown (not hydrated)' : String(value)
}

function formatLoadedBool(value: boolean | null): string {
	if (value === null) return 'unknown (not hydrated)'
	return value ? 'yes' : 'no'
}

function formatRange(range: {
	start: { row: number; col: number }
	end: { row: number; col: number }
}): string {
	return `${columnLabel(range.start.col)}${range.start.row + 1}:${columnLabel(range.end.col)}${range.end.row + 1}`
}

function columnLabel(col: number): string {
	let n = col
	let label = ''
	while (n >= 0) {
		label = String.fromCharCode(65 + (n % 26)) + label
		n = Math.floor(n / 26) - 1
	}
	return label
}
