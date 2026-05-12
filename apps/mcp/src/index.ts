import { AscendException, ascendError, type CellValue, EMPTY } from '@ascend/schema'
import {
	type AgentCommitOptions,
	Ascend,
	type CapabilityFilters,
	type CompactRangeWindowInfo,
	commitAgentPlan,
	createAgentPlan,
	createRepairPlan,
	ensureOutputExtension,
	escapeDelimitedCell,
	formatDisplayCellValue,
	indexToColumn,
	inferExportFormat,
	listCapabilities,
	normalizeExportFormat,
	type PivotOutputMaterializeMode,
	parseA1,
	parseOperations,
	type RangeObjectsInfo,
	type RangeRowsInfo,
	type RangeWindowInfo,
	readAgentDoc,
	searchAgentDocs,
	summarizeCapabilities,
	toA1Ref,
	WorkbookDocument,
} from '@ascend/sdk'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { z } from 'zod'
import { errorResponse, okResponse } from './response.ts'

export function createServer(): McpServer {
	const server = new McpServer({
		name: 'ascend',
		version: '0.0.0',
	})

	registerAgentResources(server)
	registerAgentPrompts(server)

	server.tool(
		'ascend.search_docs',
		'Search Ascend machine-readable docs, workflow guidance, API references, and release notes for agent recovery',
		{
			query: z.string().min(1).describe('Search query, topic, command, tool, or workflow'),
			limit: z.number().int().positive().max(20).optional().describe('Maximum results to return'),
			tokens: z
				.number()
				.int()
				.positive()
				.max(8000)
				.optional()
				.describe('Approximate maximum tokens to include per result snippet'),
		},
		async ({ query, limit, tokens }) => {
			const results = await searchAgentDocs({
				query,
				...(limit !== undefined ? { limit } : {}),
				...(tokens !== undefined ? { tokens } : {}),
			})
			return okResponse(
				{ query, results },
				`Found ${results.length} Ascend documentation result(s) for "${query}"`,
			)
		},
	)

	server.tool(
		'ascend.search_examples',
		'Search Ascend examples and MCP setup snippets for concrete CLI, SDK, and MCP usage patterns',
		{
			query: z.string().min(1).describe('Search query, task, operation, or integration target'),
			limit: z.number().int().positive().max(20).optional().describe('Maximum examples to return'),
			tokens: z
				.number()
				.int()
				.positive()
				.max(8000)
				.optional()
				.describe('Approximate maximum tokens to include per result snippet'),
		},
		async ({ query, limit, tokens }) => {
			const results = await searchAgentDocs({
				query,
				kind: 'example',
				...(limit !== undefined ? { limit } : {}),
				...(tokens !== undefined ? { tokens } : {}),
			})
			return okResponse(
				{ query, results },
				`Found ${results.length} Ascend example result(s) for "${query}"`,
			)
		},
	)

	server.tool(
		'ascend.inspect',
		'Inspect workbook or sheet metadata',
		{
			file: z.string().describe('Path to workbook file'),
			sheet: z.string().optional().describe('Sheet name to inspect'),
		},
		async ({ file, sheet }) => {
			try {
				const wb = await WorkbookDocument.open(
					file,
					sheet ? { mode: 'values', sheets: [sheet] } : { mode: 'metadata-only' },
				)
				if (sheet) {
					const data = wb.inspectSheet(sheet)
					if (!data) {
						return errorResponse(
							sheetNotFoundError(sheet, await loadAvailableSheets(file, wb.sheets)),
						)
					}
					return {
						...okResponse(data, `Inspected sheet "${sheet}"`),
					}
				}
				const data = wb.inspect()
				return okResponse(data, `Inspected workbook "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.list_sheets',
		'List all sheet names and tables in a workbook (lightweight alternative to inspect)',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'metadata-only' })
				const info = wb.inspect()
				const sheets = info.sheets.map((s) => ({
					name: s.name,
					rows: s.rowCount,
					cols: s.colCount,
					tableCount: s.tableCount,
				}))
				return okResponse({ sheets }, `${sheets.length} sheet(s) in "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.active_content',
		'Inspect macros, Excel 4 macro sheets, ActiveX/form controls, Custom UI callbacks, signatures, and related preserve-first edit risks',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'metadata-only' })
				const info = wb.inspect()
				const activeFeatureFamilies = new Set([
					'preservedMacro',
					'preservedMacroSheet',
					'preservedActiveContent',
					'preservedSignature',
					'preservedCustomUi',
				])
				const compatibilityFeatures = info.compatibility.features.filter(
					(feature) =>
						activeFeatureFamilies.has(feature.feature) ||
						feature.locations.some((location) =>
							/(vba|macro|activex|ctrlprops|customui|_xmlsignatures|signature)/i.test(location),
						),
				)
				return okResponse(
					{
						activeContentCount: info.activeContentCount,
						macroSheetCount: info.macroSheetCount,
						activeContent: info.activeContent,
						macroSheets: info.macroSheets,
						compatibilityFeatures,
						capabilityWarnings: info.capabilityWarnings.filter(
							(warning) => warning.family === 'active content',
						),
					},
					`Inspected ${info.activeContentCount} active-content part(s) in "${file}"`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.package_graph',
		'Inspect XLSX OPC package parts, content types, relationship ids, raw/resolved targets, ownership, feature families, and preservation policy',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'metadata-only' })
				const graph = await wb.packageGraph()
				return okResponse(
					graph,
					`Inspected ${graph.parts.length} package part(s) and ${graph.relationships.length} relationship(s) in "${file}"`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.visuals',
		'Inspect workbook visual inventory: charts, drawings, media, image anchors, drawing object links, and preserve-first visual gaps',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'full' })
				const inventory = wb.visualInventory()
				return okResponse(inventory, `Inspected visual inventory for "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.pivots',
		'Inspect PivotTables, cache records, saved-output audits, refresh plans, and planned setCells ops for supported output materialization',
		{
			file: z.string().describe('Path to workbook file'),
			pivotTable: z.string().optional().describe('Optional PivotTable name filter'),
			partPath: z.string().optional().describe('Optional PivotTable part path filter'),
			mode: z
				.enum(['missing', 'mismatches', 'all'])
				.optional()
				.describe(
					'Materialization planning mode: write only missing cells, mismatched cells, or all supported output cells',
				),
		},
		async ({ file, pivotTable, partPath, mode }) => {
			try {
				const wb = await WorkbookDocument.open(file, {
					mode: 'full',
					pivotCacheRecordMaterializeLimit: 'all',
				})
				const materializeOptions = {
					...(pivotTable ? { pivotTable } : {}),
					...(partPath ? { partPath } : {}),
					...(mode ? { mode: mode as PivotOutputMaterializeMode } : {}),
				}
				const plan = wb.pivotOutputMaterializeOps(materializeOptions)
				return okResponse(
					{
						pivotTables: wb.pivotTables(),
						pivotCaches: wb.pivotCaches(),
						pivotOutputAudits: wb.pivotOutputAudits(),
						pivotRefreshPlans: wb.pivotRefreshPlans(),
						pivotOutputMaterializePlan: plan,
					},
					`Planned ${plan.plannedCellCount} PivotTable output cell(s) for "${file}"`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.read',
		'Read cell values from a range',
		{
			file: z.string().describe('Path to workbook file'),
			range: z.string().describe('Cell range (e.g. "A1:C10")'),
			sheet: z.string().optional().describe('Sheet name (defaults to first sheet)'),
			rowOffset: z.number().int().nonnegative().optional().describe('Row offset within the range'),
			rowLimit: z.number().int().positive().optional().describe('Maximum rows to return'),
			changedSince: z
				.string()
				.optional()
				.describe('For compact reads, return only cells changed since this prior changeToken'),
			format: z
				.enum(['cells', 'rows', 'objects', 'tsv', 'compact'])
				.optional()
				.describe(
					'Read format: cell records, row arrays, object rows, TSV text, or sparse compact cells',
				),
			display: z
				.boolean()
				.optional()
				.describe('Return display strings instead of raw typed values'),
			headers: z
				.array(z.string())
				.optional()
				.describe('Explicit headers for object mode; defaults to first-row headers'),
			cols: z
				.array(z.string())
				.optional()
				.describe(
					'Columns to return by absolute letter (A, C), 1-based position in the requested range, or object header',
				),
		},
		async ({
			file,
			range,
			sheet,
			rowOffset,
			rowLimit,
			changedSince,
			format,
			display,
			headers,
			cols,
		}) => {
			try {
				const wb = await WorkbookDocument.open(
					file,
					sheet ? { mode: 'values', sheets: [sheet] } : { mode: 'values' },
				)
				const sheetName = sheet ?? wb.sheets[0]
				if (!sheetName) {
					return errorResponse('No sheets in workbook')
				}
				const handle = wb.sheet(sheetName)
				if (!handle) {
					return errorResponse(
						sheetNotFoundError(sheetName, await loadAvailableSheets(file, wb.sheets)),
					)
				}
				const readOpts = {
					...(rowOffset !== undefined ? { rowOffset } : {}),
					...(rowLimit !== undefined ? { rowLimit } : {}),
				}
				const mode = format ?? 'cells'
				const info =
					mode === 'compact'
						? buildCompactReadResult(
								handle.readWindowCompact(range, {
									...readOpts,
									includeRefs: false,
									omitEmpty: true,
									flatValues: true,
									changedSince: changedSince ?? '',
								}),
								cols,
							)
						: mode === 'tsv'
							? buildTsvReadResult(handle.readRows(range, readOpts), cols)
							: mode === 'rows'
								? pruneRowsInfo(handle.readRows(range, readOpts), cols)
								: mode === 'objects'
									? pruneObjectsInfo(
											handle.readObjects(range, {
												...readOpts,
												headers: headers && headers.length > 0 ? headers : 'first-row',
											}),
											cols,
										)
									: pruneWindowInfo(handle.readWindow(range, readOpts), cols)
				return okResponse(
					mode === 'tsv' || mode === 'compact'
						? info
						: display
							? displayReadResult(mode, info)
							: info,
					`Read range ${range} from "${sheetName}"`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.find',
		'Find cells by value or formula text',
		{
			file: z.string().describe('Path to workbook file'),
			query: z.string().min(1).describe('Text to search for'),
			sheet: z.string().optional().describe('Sheet name (defaults to first sheet)'),
			in: z
				.enum(['value', 'formula', 'both'])
				.optional()
				.describe('Search values, formulas, or both'),
			caseSensitive: z.boolean().optional().describe('Match case exactly'),
			limit: z.number().int().positive().max(500).optional().describe('Maximum matches to return'),
		},
		async ({ file, query, sheet, in: searchIn, caseSensitive, limit }) => {
			try {
				const wb = await WorkbookDocument.open(
					file,
					sheet ? { mode: 'formula', sheets: [sheet] } : { mode: 'formula' },
				)
				const sheetName = sheet ?? wb.sheets[0]
				if (!sheetName) return errorResponse('No sheets in workbook')
				const handle = wb.sheet(sheetName)
				if (!handle) {
					return errorResponse(
						sheetNotFoundError(sheetName, await loadAvailableSheets(file, wb.sheets)),
					)
				}
				const usedRange = handle.usedRange()
				if (!usedRange) {
					return okResponse(
						{ sheet: sheetName, query, in: searchIn ?? 'both', matches: [], truncated: false },
						`No populated cells to search in "${sheetName}"`,
					)
				}
				const cells = handle.rangeCompact(rangeRefToString(usedRange), {
					includeRefs: true,
				}).cells
				const searchMode = searchIn ?? 'both'
				const maxMatches = limit ?? 100
				const matches: Array<{
					ref: string
					row: number
					col: number
					value: string
					formula: string | null
					matchedOn: 'value' | 'formula'
				}> = []
				for (const cell of cells) {
					const ref = cell.ref
					if (!ref) continue
					const valueText = formatDisplayCellValue(cell.value)
					const formulaText = cell.formula ?? ''
					const matchValue =
						(searchMode === 'value' || searchMode === 'both') &&
						includesQuery(valueText, query, caseSensitive ?? false)
					const matchFormula =
						(searchMode === 'formula' || searchMode === 'both') &&
						cell.formula !== null &&
						includesQuery(formulaText, query, caseSensitive ?? false)
					if (!matchValue && !matchFormula) continue
					matches.push({
						ref,
						row: cell.row,
						col: cell.col,
						value: valueText,
						formula: cell.formula,
						matchedOn: matchFormula ? 'formula' : 'value',
					})
					if (matches.length >= maxMatches) break
				}
				return okResponse(
					{
						sheet: sheetName,
						query,
						in: searchMode,
						caseSensitive: caseSensitive ?? false,
						limit: maxMatches,
						truncated: matches.length >= maxMatches && cells.length > matches.length,
						matches,
					},
					`Found ${matches.length} matching cells in "${sheetName}"`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.read_table',
		'Read a structured Excel table by name without manually resolving its range',
		{
			file: z.string().describe('Path to workbook file'),
			table: z.string().describe('Excel table name'),
			rowOffset: z.number().int().nonnegative().optional().describe('Data row offset'),
			rowLimit: z.number().int().positive().optional().describe('Maximum data rows to return'),
			display: z
				.boolean()
				.optional()
				.describe('Return display strings instead of raw typed values'),
		},
		async ({ file, table, rowOffset, rowLimit, display }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'full' })
				const handle = wb.table(table)
				if (!handle) return errorResponse(`Table "${table}" not found`)
				const page = handle.readRows({
					...(rowOffset !== undefined ? { offset: rowOffset } : {}),
					...(rowLimit !== undefined ? { limit: rowLimit } : {}),
				})
				const rows = display
					? page.rows.map((row) => ({
							...row,
							values: Object.fromEntries(
								Object.entries(row.values).map(([key, value]) => [
									key,
									formatDisplayCellValue(value),
								]),
							),
						}))
					: page.rows
				return okResponse(
					{
						name: handle.name,
						columns: handle.columns,
						ref: rangeRefToString(handle.ref),
						rowCount: handle.rowCount,
						hasHeaders: handle.hasHeaders,
						hasTotals: handle.hasTotals,
						headerRow: display
							? handle.headerRow()?.map((cell) => formatDisplayCellValue(cell))
							: handle.headerRow(),
						totalsRow: display
							? handle.totalsRow()?.map((cell) => formatDisplayCellValue(cell))
							: handle.totalsRow(),
						sortState: handle.sortState,
						autoFilter: handle.autoFilter,
						page: { ...page, rows },
						rows: rows.map((row) => row.values),
					},
					`Read table "${table}"`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.agent_view',
		'Read a compressed semantic summary for a worksheet range',
		{
			file: z.string().describe('Path to workbook file'),
			range: z.string().describe('Cell range (e.g. "A1:Z200")'),
			sheet: z.string().optional().describe('Sheet name (defaults to first sheet)'),
			rowChunkSize: z.number().int().positive().optional().describe('Rows per streamed chunk'),
			sampleRowLimit: z.number().int().positive().optional().describe('Maximum sample rows'),
			sampleValueLimit: z
				.number()
				.int()
				.positive()
				.optional()
				.describe('Maximum sample values per column'),
		},
		async ({ file, range, sheet, rowChunkSize, sampleRowLimit, sampleValueLimit }) => {
			try {
				const wb = await WorkbookDocument.open(
					file,
					sheet ? { mode: 'formula', sheets: [sheet] } : { mode: 'formula' },
				)
				const sheetName = sheet ?? wb.sheets[0]
				if (!sheetName) return errorResponse('No sheets in workbook')
				const view = wb.agentView(sheetName, range, {
					...(rowChunkSize !== undefined ? { rowChunkSize } : {}),
					...(sampleRowLimit !== undefined ? { sampleRowLimit } : {}),
					...(sampleValueLimit !== undefined ? { sampleValueLimit } : {}),
				})
				if (!view) {
					return errorResponse(
						sheetNotFoundError(sheetName, await loadAvailableSheets(file, wb.sheets)),
					)
				}
				return okResponse(view, `Generated agent view for ${range} on "${sheetName}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.preview',
		'Preview operations without saving the workbook',
		{
			file: z.string().describe('Path to workbook file'),
			ops: z.array(z.record(z.string(), z.unknown())).describe('Operations to preview'),
		},
		async ({ file, ops }) => {
			const parsed = parseOperations(ops)
			if (!parsed.ok) {
				return errorResponse(
					ascendError('VALIDATION_ERROR', parsed.error, {
						details: { issues: parsed.issues },
						retryStrategy: 'modified',
						suggestedFix:
							'Use ascend.list_operations for canonical operation schemas and examples.',
					}),
				)
			}
			try {
				const wb = await Ascend.open(file)
				const result = wb.preview(parsed.value)
				if (result.errors.length > 0) {
					const first = result.errors[0]
					return errorResponse(
						first
							? {
									...first,
									details: { ...(first.details ?? {}), preview: result },
								}
							: ascendError('VALIDATION_ERROR', 'Preview failed', { details: { preview: result } }),
					)
				}
				return okResponse(result, `Previewed ${ops.length} operation(s)`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.write',
		'Apply operations to a workbook',
		{
			file: z.string().describe('Path to workbook file'),
			ops: z.array(z.record(z.string(), z.unknown())).describe('Operations to apply'),
		},
		async ({ file, ops }) => {
			const parsed = parseOperations(ops)
			if (!parsed.ok) {
				return errorResponse(
					ascendError('VALIDATION_ERROR', parsed.error, {
						details: { issues: parsed.issues },
						retryStrategy: 'modified',
						suggestedFix:
							'Use ascend.list_operations for canonical operation schemas and examples.',
					}),
				)
			}
			try {
				const wb = await Ascend.open(file)
				const result = wb.apply(parsed.value)
				if (result.errors.length > 0) {
					const first = result.errors[0]
					return errorResponse(
						first ?? ascendError('VALIDATION_ERROR', 'Failed to apply operations'),
					)
				}
				if (result.recalcRequired) {
					const recalc = wb.recalc()
					if (recalc.errors.length > 0) {
						const first = recalc.errors[0]
						return errorResponse(
							first
								? ascendError('FORMULA_EVAL_ERROR', `${first.ref}: ${first.error.message}`, {
										refs: [first.ref],
										details: { evalError: first.error },
									})
								: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed'),
						)
					}
				}
				await wb.save(file)
				return okResponse(result, `Applied ${ops.length} operation(s)`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.calc',
		'Recalculate all formulas in a workbook',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await Ascend.open(file)
				const result = wb.recalc()
				if (result.errors.length > 0) {
					const first = result.errors[0]
					return errorResponse(
						first
							? ascendError('FORMULA_EVAL_ERROR', `${first.ref}: ${first.error.message}`, {
									refs: [first.ref],
									details: { evalError: first.error },
								})
							: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed'),
					)
				}
				await wb.save(file)
				return okResponse(result, `Recalculated workbook "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.eval',
		'Evaluate a formula against a workbook without writing a scratch cell',
		{
			file: z.string().describe('Path to workbook file'),
			formula: z.string().describe('Formula to evaluate, with or without a leading ='),
			display: z.boolean().optional().describe('Return a display string alongside the typed value'),
		},
		async ({ file, formula, display }) => {
			try {
				const wb = await Ascend.open(file)
				const value = wb.eval(formula)
				return okResponse(
					{
						formula,
						value,
						...(display ? { display: formatDisplayCellValue(value) } : {}),
					},
					`Evaluated formula "${formula}"`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.list_operations',
		'List all available spreadsheet operations with their parameters and JSON Schema for LLM tool use',
		{},
		async () => {
			const ops = Ascend.listOperations()
			const schemas = Ascend.getOperationsSchema()
			return okResponse({ operations: ops, schemas }, `${ops.length} operations available`)
		},
	)

	server.tool(
		'ascend.capabilities',
		'List Ascend Excel capability coverage with OSS baseline gaps and next milestones',
		{
			feature: z.string().optional().describe('Filter by capability id, label, or family'),
			family: z.string().optional().describe('Filter by exact capability family'),
			priority: z.enum(['P0', 'P1', 'P2', 'P3']).optional().describe('Priority filter'),
			status: z
				.enum([
					'excel-equivalent',
					'editable',
					'inspectable',
					'preserved',
					'unsafe-blocked',
					'unsupported',
				])
				.optional()
				.describe('Status filter'),
			gapsOnly: z.boolean().optional().describe('Only return non-editable/non-equivalent gaps'),
		},
		async ({ feature, family, priority, status, gapsOnly }) => {
			const filters: CapabilityFilters = {
				...(feature ? { feature } : {}),
				...(family ? { family } : {}),
				...(priority ? { priority } : {}),
				...(status ? { status } : {}),
				...(gapsOnly ? { gapsOnly: true } : {}),
			}
			const capabilities = listCapabilities(filters)
			return okResponse(
				{ summary: summarizeCapabilities(capabilities), capabilities },
				`${capabilities.length} capabilities returned`,
			)
		},
	)

	server.tool(
		'ascend.plan',
		'Agent-safe edit planning: validate, preview, recalc-audit, package-graph-audit, and preservation-audit operations without saving',
		{
			file: z.string().describe('Path to workbook file'),
			ops: z.array(z.record(z.string(), z.unknown())).describe('Operations to plan'),
		},
		async ({ file, ops }) => {
			const parsed = parseOperations(ops)
			if (!parsed.ok) {
				return errorResponse(
					ascendError('VALIDATION_ERROR', parsed.error, {
						details: { issues: parsed.issues },
						retryStrategy: 'modified',
						suggestedFix:
							'Use ascend.list_operations for canonical operation schemas and examples.',
					}),
				)
			}
			try {
				const result = await createAgentPlan(file, parsed.value)
				if (result.preview.errors.length > 0) {
					const first = result.preview.errors[0]
					return errorResponse(
						first
							? { ...first, details: { ...(first.details ?? {}), plan: result } }
							: ascendError('VALIDATION_ERROR', 'Plan failed', { details: { plan: result } }),
					)
				}
				return okResponse(result, `Planned ${ops.length} operation(s)`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.commit',
		'Commit an agent edit plan atomically with optional input hash guard',
		{
			file: z.string().describe('Path to workbook file'),
			ops: z.array(z.record(z.string(), z.unknown())).describe('Operations to commit'),
			output: z.string().optional().describe('Non-destructive output path'),
			inPlace: z.boolean().optional().describe('Replace the input file atomically'),
			backup: z.string().optional().describe('Backup path for in-place commits'),
			expectSha256: z
				.string()
				.optional()
				.describe('Reject if the input hash has changed since plan'),
			allowLoss: z
				.union([z.string(), z.array(z.string())])
				.optional()
				.describe('Allow preserved/unsupported feature loss by feature, tier, or "all"'),
			approvals: z
				.union([z.string(), z.array(z.string())])
				.optional()
				.describe('Approve explicit plan approval ids, aliases, or "all"'),
		},
		async ({ file, ops, output, inPlace, backup, expectSha256, allowLoss, approvals }) => {
			const parsed = parseOperations(ops)
			if (!parsed.ok) {
				return errorResponse(
					ascendError('VALIDATION_ERROR', parsed.error, {
						details: { issues: parsed.issues },
						retryStrategy: 'modified',
						suggestedFix:
							'Use ascend.list_operations for canonical operation schemas and examples.',
					}),
				)
			}
			try {
				const options: AgentCommitOptions = {
					...(output ? { output } : {}),
					...(inPlace ? { inPlace: true } : {}),
					...(backup ? { backup } : {}),
					...(expectSha256 ? { expectSha256 } : {}),
					...(allowLoss ? { allowLoss: parseAllowLoss(allowLoss) } : {}),
					...(approvals ? { approvals: parseStringListOrAll(approvals) } : {}),
				}
				const result = await commitAgentPlan(file, parsed.value, options)
				return okResponse(result, `Committed ${ops.length} operation(s)`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.repair_plan',
		'Suggest safe next actions when check, lint, or unsupported-feature audits need attention',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const result = await createRepairPlan(file)
				return okResponse(result, `${result.actions.length} repair action(s)`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.check',
		'Run structural checks on a workbook',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'formula' })
				const result = wb.check()
				if (!result.valid) {
					const summary = `${result.issues.length} issue(s) found`
					return errorResponse(
						ascendError('VALIDATION_ERROR', summary, {
							details: { check: result },
						}),
					)
				}
				return okResponse(result, `Checked workbook "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.lint',
		'Lint formulas for common issues',
		{
			file: z.string().describe('Path to workbook file'),
		},
		async ({ file }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'formula' })
				const result = wb.lint()
				return okResponse(result, `Linted workbook "${file}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.trace',
		'Trace cell dependencies (precedents and dependents)',
		{
			file: z.string().describe('Path to workbook file'),
			cell: z.string().describe('Cell reference (e.g. "Sheet1!A1" or "A1")'),
		},
		async ({ file, cell }) => {
			try {
				const wb = await WorkbookDocument.open(file, { mode: 'formula' })
				const result = wb.trace(cell)
				if (!result) {
					return errorResponse(`Cannot trace "${cell}"`)
				}
				return okResponse(result, `Traced cell "${cell}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.diff',
		'Compute semantic diff between two workbooks',
		{
			fileA: z.string().describe('Path to first workbook'),
			fileB: z.string().describe('Path to second workbook'),
		},
		async ({ fileA, fileB }) => {
			try {
				const [a, b] = await Promise.all([Ascend.open(fileA), Ascend.open(fileB)])
				const result = a.diff(b)
				return okResponse(result, `Diffed "${fileA}" against "${fileB}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.export',
		'Export workbook to another format',
		{
			file: z.string().describe('Path to source workbook'),
			output: z.string().describe('Output file path'),
			format: z.string().optional().describe('Output format (inferred from extension if omitted)'),
		},
		async ({ file, output, format }) => {
			try {
				const wb = await Ascend.open(file)
				const normalized = format ? normalizeExportFormat(format) : inferExportFormat(output)
				if (!normalized) return errorResponse(`Unsupported format: ${format ?? output}`)
				const target = ensureOutputExtension(output, normalized)
				if (normalized === 'json') {
					await Bun.write(target, JSON.stringify(wb.toJSON(), null, 2))
				} else if (normalized === 'csv' || normalized === 'tsv') {
					const text = wb.toCsv(normalized === 'tsv' ? { dialect: { delimiter: '\t' } } : undefined)
					await Bun.write(target, text)
				} else {
					await wb.save(target)
				}
				return okResponse({ exported: target }, `Exported workbook to "${target}"`)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	return server
}

function registerAgentResources(server: McpServer): void {
	server.registerResource(
		'ascend.llms',
		'ascend://llms.txt',
		{
			title: 'Ascend llms.txt',
			description: 'Short machine-readable map for agents using Ascend.',
			mimeType: 'text/plain',
		},
		async (uri) => textResource(uri, (await readAgentDoc('llms.txt')) ?? ''),
	)

	server.registerResource(
		'ascend.llms_full',
		'ascend://llms-full.txt',
		{
			title: 'Ascend llms-full.txt',
			description: 'Expanded machine-readable Ascend documentation bundle for agents.',
			mimeType: 'text/plain',
		},
		async (uri) => textResource(uri, (await readAgentDoc('llms-full.txt')) ?? ''),
	)

	server.registerResource(
		'ascend.agent_api',
		'ascend://docs/agent-api.md',
		{
			title: 'Ascend Agent API Markdown Reference',
			description: 'Markdown reference for CLI, MCP, SDK, operation schemas, and safe workflows.',
			mimeType: 'text/markdown',
		},
		async (uri) => textResource(uri, (await readAgentDoc('docs/AGENT_API.md')) ?? ''),
	)

	server.registerResource(
		'ascend.capabilities',
		'ascend://capabilities',
		{
			title: 'Ascend Excel Capability Matrix',
			description:
				'Canonical coverage registry with statuses, priorities, OSS baselines, and gaps.',
			mimeType: 'application/json',
		},
		(uri) => {
			const capabilities = listCapabilities()
			return jsonResource(uri, {
				summary: summarizeCapabilities(capabilities),
				capabilities,
			})
		},
	)

	server.registerResource(
		'ascend.operations',
		'ascend://operations',
		{
			title: 'Ascend Operation Schemas',
			description:
				'Operation catalog with schemas, examples, invalid examples, recovery actions, and approvals.',
			mimeType: 'application/json',
		},
		(uri) =>
			jsonResource(uri, {
				operations: Ascend.listOperations(),
				schemas: Ascend.getOperationsSchema(),
			}),
	)

	server.registerResource(
		'ascend.agent_workflow',
		'ascend://agent-workflow',
		{
			title: 'Ascend Agent Workflow',
			description:
				'Recommended headless spreadsheet workflow for inspect, read, plan, commit, verify, and repair.',
			mimeType: 'text/markdown',
		},
		(uri) => textResource(uri, buildAgentWorkflowGuide()),
	)
}

function registerAgentPrompts(server: McpServer): void {
	server.registerPrompt(
		'ascend.agent_workflow',
		{
			title: 'Ascend Spreadsheet Agent Workflow',
			description: 'Prime an agent to use Ascend safely for headless Excel edits.',
			argsSchema: {
				file: z.string().optional().describe('Workbook path to operate on'),
				task: z.string().optional().describe('Spreadsheet task or edit intent'),
			},
		},
		({ file, task }) => ({
			description: 'Use Ascend MCP tools for a safe spreadsheet edit workflow.',
			messages: [
				{
					role: 'user',
					content: {
						type: 'text',
						text: buildAgentWorkflowPrompt(file, task),
					},
				},
			],
		}),
	)
}

function jsonResource(uri: URL, data: unknown) {
	return {
		contents: [
			{
				uri: uri.href,
				mimeType: 'application/json',
				text: JSON.stringify(data, null, 2),
			},
		],
	}
}

function textResource(uri: URL, text: string) {
	return {
		contents: [
			{
				uri: uri.href,
				mimeType: 'text/markdown',
				text,
			},
		],
	}
}

function buildAgentWorkflowGuide(): string {
	return [
		'# Ascend Agent Workflow',
		'',
		'1. Inspect workbook structure with ascend.inspect or ascend.list_sheets.',
		'2. Audit package fidelity with ascend.package_graph when sidecars, relationships, content types, or preservation policy can affect the edit.',
		'3. Audit high-risk workbook content with ascend.active_content before editing macro-enabled, signed, ActiveX, Custom UI, or Excel 4 macro-sheet files.',
		'4. Locate data with ascend.find, ascend.read, ascend.read_table, ascend.visuals, and ascend.pivots for PivotTable inventory/audits/materialization ops.',
		'5. Use ascend.search_docs or ascend.search_examples when you need command, schema, workflow, or example recovery context.',
		'6. Fetch operation schemas from ascend.list_operations or ascend://operations.',
		'7. Preview edits with ascend.plan before writing.',
		'8. Commit with ascend.commit using output paths, input hash guards, approvals, and allow-loss only when explicit.',
		'9. Verify with ascend.check, ascend.lint, ascend.trace, ascend.diff, and ascend.export as needed.',
		'10. Use ascend.repair_plan when checks, lints, approvals, or unsupported-feature audits need recovery actions.',
	].join('\n')
}

function buildAgentWorkflowPrompt(file?: string, task?: string): string {
	const target = file
		? `Workbook: ${file}`
		: 'Workbook: ask for or infer the workbook path before editing.'
	const intent = task ? `Task: ${task}` : 'Task: determine the requested spreadsheet change.'
	return [
		target,
		intent,
		'',
		'Use Ascend as the source of truth for spreadsheet structure and edit safety.',
		'If you need recovery context, call ascend.search_docs or ascend.search_examples before guessing.',
		'Start with ascend.inspect or ascend.list_sheets; call ascend.package_graph when package sidecars, relationship identity, or preservation policy matter; call ascend.active_content before editing macro-enabled, signed, ActiveX, Custom UI, or Excel 4 macro-sheet workbooks.',
		'Then use ascend.read, ascend.read_table, ascend.find, ascend.visuals, and ascend.pivots to gather only the necessary workbook context.',
		'Before modifying anything, read ascend://operations or call ascend.list_operations and build operations that match the published schemas.',
		'Always run ascend.plan and inspect approvals, unsupported features, preview diffs, recalc status, and modelOutput before commit.',
		'Use ascend.commit with a non-destructive output path by default, pass expectSha256 when available, and only pass approvals or allowLoss values emitted by the plan.',
		'After commit, verify with ascend.check, ascend.lint, ascend.diff, or ascend.export depending on the task.',
	].join('\n')
}

function displayReadResult(mode: 'cells' | 'rows' | 'objects', info: unknown): unknown {
	if (mode === 'rows') {
		const rowsInfo = info as { rows: readonly (readonly CellValue[])[] }
		return {
			...rowsInfo,
			rows: rowsInfo.rows.map((row) => row.map((cell) => formatDisplayCellValue(cell))),
		}
	}
	if (mode === 'objects') {
		const objectInfo = info as {
			headers: readonly string[]
			rows: readonly Readonly<Record<string, CellValue>>[]
		}
		return {
			...objectInfo,
			rows: objectInfo.rows.map((row) =>
				Object.fromEntries(
					Object.entries(row).map(([key, value]) => [key, formatDisplayCellValue(value)]),
				),
			),
		}
	}
	const cellInfo = info as { cells: readonly { ref: string; value: CellValue }[] }
	return {
		...cellInfo,
		cells: cellInfo.cells.map((cell) => ({
			...cell,
			value: formatDisplayCellValue(cell.value),
		})),
	}
}

interface SelectedColumnInfo {
	readonly position: number
	readonly col: number
	readonly letter: string
	readonly header?: string
}

function resolveColumnSelection(
	cols: readonly string[] | undefined,
	startCol: number,
	colCount: number,
	headers?: readonly string[],
): readonly number[] | null {
	if (!cols || cols.length === 0) return null
	const selected: number[] = []
	const seen = new Set<number>()
	const headerLookup = headers
		? new Map(headers.map((header, index) => [header.trim().toLowerCase(), index] as const))
		: undefined
	for (const raw of cols) {
		const token = raw.trim()
		if (token.length === 0) continue
		const relative = resolveColumnToken(token, startCol, colCount, headerLookup)
		if (relative === undefined || seen.has(relative)) continue
		seen.add(relative)
		selected.push(relative)
	}
	return selected
}

function resolveColumnToken(
	token: string,
	startCol: number,
	colCount: number,
	headerLookup?: ReadonlyMap<string, number>,
): number | undefined {
	if (/^\d+$/.test(token)) {
		const relative = Number.parseInt(token, 10) - 1
		return relative >= 0 && relative < colCount ? relative : undefined
	}
	if (/^[A-Za-z]{1,3}$/.test(token)) {
		const parsed = parseA1(`${token.toUpperCase()}1`)
		const relative = parsed.col - startCol
		if (relative >= 0 && relative < colCount) return relative
	}
	const headerRelative = headerLookup?.get(token.toLowerCase())
	return headerRelative !== undefined && headerRelative >= 0 && headerRelative < colCount
		? headerRelative
		: undefined
}

function selectedColumns(
	selection: readonly number[],
	startCol: number,
	headers?: readonly string[],
): readonly SelectedColumnInfo[] {
	return selection.map((relative) => {
		const absoluteCol = startCol + relative
		return {
			position: relative + 1,
			col: absoluteCol,
			letter: indexToColumn(absoluteCol),
			...(headers?.[relative] !== undefined ? { header: headers[relative] } : {}),
		}
	})
}

function pruneRowsInfo(info: RangeRowsInfo, cols?: readonly string[]) {
	const selection = resolveColumnSelection(cols, info.ref.start.col, info.colCount)
	if (!selection) return info
	return {
		...info,
		colCount: selection.length,
		selectedColumns: selectedColumns(selection, info.ref.start.col),
		rows: info.rows.map((row) => selection.map((relative) => row[relative] ?? EMPTY)),
	}
}

function pruneObjectsInfo(info: RangeObjectsInfo, cols?: readonly string[]) {
	const selection = resolveColumnSelection(cols, info.ref.start.col, info.colCount, info.headers)
	if (!selection) return info
	const headers = selection.map((relative) => info.headers[relative] ?? '')
	return {
		...info,
		colCount: selection.length,
		headers,
		selectedColumns: selectedColumns(selection, info.ref.start.col, info.headers),
		rows: info.rows.map((row) =>
			Object.fromEntries(headers.map((header) => [header, row[header] ?? EMPTY])),
		),
	}
}

function pruneWindowInfo(info: RangeWindowInfo, cols?: readonly string[]) {
	const selection = resolveColumnSelection(cols, info.ref.start.col, info.colCount)
	if (!selection) return info
	const selected = new Set(selection)
	return {
		...info,
		colCount: selection.length,
		selectedColumns: selectedColumns(selection, info.ref.start.col),
		cells: info.cells.filter((cell) => selected.has(cell.col - info.ref.start.col)),
	}
}

function pruneCompactWindowInfo(info: CompactRangeWindowInfo, cols?: readonly string[]) {
	const selection = resolveColumnSelection(cols, info.ref.start.col, info.colCount)
	if (!selection) return info
	const selected = new Set(selection)
	return {
		...info,
		colCount: selection.length,
		selectedColumns: selectedColumns(selection, info.ref.start.col),
		cells: info.cells.filter((cell) => selected.has(cell.col - info.ref.start.col)),
	}
}

function buildTsvReadResult(sourceInfo: RangeRowsInfo, cols?: readonly string[]) {
	const info = pruneRowsInfo(sourceInfo, cols)
	const rows = info.rows.map((row) =>
		row.map((cell) => escapeDelimitedCell(formatDisplayCellValue(cell), '\t')),
	)
	return {
		requestedRef: info.requestedRef,
		ref: info.ref,
		rowCount: info.rowCount,
		colCount: info.colCount,
		rowOffset: info.rowOffset,
		rowLimit: info.rowLimit,
		hasMore: info.hasMore,
		...(info.nextRowOffset !== undefined ? { nextRowOffset: info.nextRowOffset } : {}),
		...('selectedColumns' in info ? { selectedColumns: info.selectedColumns } : {}),
		format: 'tsv' as const,
		tsv: rows.map((row) => row.join('\t')).join('\n'),
	}
}

function buildCompactReadResult(sourceInfo: CompactRangeWindowInfo, cols?: readonly string[]) {
	const info = pruneCompactWindowInfo(sourceInfo, cols)
	return {
		requestedRef: info.requestedRef,
		ref: info.ref,
		rowCount: info.rowCount,
		colCount: info.colCount,
		rowOffset: info.rowOffset,
		rowLimit: info.rowLimit,
		hasMore: info.hasMore,
		...(info.nextRowOffset !== undefined ? { nextRowOffset: info.nextRowOffset } : {}),
		...(info.changeToken !== undefined ? { changeToken: info.changeToken } : {}),
		...('selectedColumns' in info ? { selectedColumns: info.selectedColumns } : {}),
		format: 'compact' as const,
		cells: info.cells.map((cell) => [
			cell.row - info.ref.start.row,
			cell.col - info.ref.start.col,
			cell.value as unknown,
			...(cell.formula ? [cell.formula] : []),
		]),
	}
}

async function loadAvailableSheets(
	file: string,
	fallbackSheets: readonly string[],
): Promise<readonly string[]> {
	if (fallbackSheets.length > 0) return fallbackSheets
	try {
		const workbook = await WorkbookDocument.open(file, { mode: 'metadata-only' })
		return workbook.sheets
	} catch {
		return fallbackSheets
	}
}

function sheetNotFoundError(sheetName: string, availableSheets: readonly string[]) {
	return ascendError('SHEET_NOT_FOUND', `Sheet "${sheetName}" not found`, {
		details: { availableSheets },
		retryStrategy: availableSheets.length > 0 ? 'modified' : 'none',
		suggestedFix:
			availableSheets.length > 0
				? `Use one of the available sheets: ${availableSheets.join(', ')}`
				: 'Inspect the workbook first to list available sheets.',
	})
}

function includesQuery(value: string, query: string, caseSensitive: boolean): boolean {
	if (caseSensitive) return value.includes(query)
	return value.toLowerCase().includes(query.toLowerCase())
}

function rangeRefToString(ref: {
	start: { row: number; col: number }
	end: { row: number; col: number }
}): string {
	return `${toA1Ref(ref.start.row, ref.start.col)}:${toA1Ref(ref.end.row, ref.end.col)}`
}

function parseAllowLoss(value: string | string[]): readonly string[] | 'all' {
	return parseStringListOrAll(value)
}

function parseStringListOrAll(value: string | string[]): readonly string[] | 'all' {
	const entries = (Array.isArray(value) ? value : value.split(','))
		.map((entry) => entry.trim())
		.filter((entry) => entry.length > 0)
	if (entries.some((entry) => entry.toLowerCase() === 'all')) return 'all'
	return entries
}

if (import.meta.main) {
	const server = createServer()
	const transport = new StdioServerTransport()
	server.connect(transport).catch((err) => {
		process.stderr.write(`${err}\n`)
		process.exit(1)
	})
}
