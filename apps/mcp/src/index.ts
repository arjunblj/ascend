import { randomUUID } from 'node:crypto'
import {
	type AscendError,
	AscendException,
	ascendError,
	type CellValue,
	EMPTY,
	type Operation,
} from '@ascend/schema'
import {
	type AgentCommitOptions,
	Ascend,
	AscendWorkbook,
	type CapabilityFilters,
	type CompactRangeWindowInfo,
	commitAgentPlan,
	commitAgentPlanFromWorkbook,
	compactAgentCommitResult,
	compactAgentPlanResult,
	createAgentPlan,
	createAgentPlanFromWorkbook,
	createPreparedAgentPlan,
	createRepairPlan,
	ensureOutputExtension,
	escapeDelimitedCell,
	formatDisplayCellValue,
	indexToColumn,
	inferExportFormat,
	listCapabilities,
	normalizeExportFormat,
	operationValidationDetails,
	type PathMutation,
	type PathMutationResult,
	type PivotOutputMaterializeMode,
	type PreparedAgentPlan,
	parseA1,
	parseOperations,
	type RangeObjectsInfo,
	type RangeRowsInfo,
	type RangeWindowInfo,
	readAgentDoc,
	SUPPORTED_PATH_MUTATION_SHAPES,
	searchAgentDocs,
	sha256Bytes,
	summarizeCapabilities,
	toA1Ref,
	WorkbookDocument,
} from '@ascend/sdk'
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import { z } from 'zod'
import { errorResponse, okResponse } from './response.ts'

const DEFAULT_MCP_RAW_PART_MAX_BYTES = 64 * 1024
const MAX_MCP_RAW_PART_MAX_BYTES = 1024 * 1024
const DEFAULT_AGENT_PREVIEW_ROWS = 500

const pathMutationSchema = z
	.unknown()
	.describe(
		'Path-addressed mutation object with path as a string or string array and optional value',
	)

type ResolvedOperationInput =
	| {
			readonly ok: true
			readonly ops: readonly Operation[]
			readonly pathMutations?: PathMutationResult
	  }
	| { readonly ok: false; readonly error: AscendError }

interface PreparedPlanHandle {
	readonly file: string
	readonly inputSha256: string
	readonly planDigest: string
	readonly pathMutations?: PathMutationResult
	commit(options: AgentCommitOptions): Promise<Awaited<ReturnType<typeof commitAgentPlan>>>
}

export interface McpServerOptions {
	readonly preparedPlanMaxHandles?: number
	readonly preparedPlanTtlMs?: number
	readonly now?: () => number
}

interface PreparedPlanMetadata {
	readonly id: string
	readonly expiresAt: string
	readonly ttlMs: number
}

interface PreparedPlanRecord {
	readonly handle: PreparedPlanHandle
	readonly expiresAtMs: number
}

const DEFAULT_PREPARED_PLAN_MAX_HANDLES = 64
const DEFAULT_PREPARED_PLAN_TTL_MS = 5 * 60 * 1000

function positiveIntegerOption(value: number | undefined, fallback: number): number {
	if (value === undefined || !Number.isFinite(value)) return fallback
	return Math.max(1, Math.floor(value))
}

function agentPlanLoadOptionsError(options: readonly string[]): AscendError {
	return ascendError(
		'VALIDATION_ERROR',
		'Agent plans and commits require a full workbook load; partial or capped load options are not supported',
		{
			details: {
				unsupportedLoadOptions: options,
				requiredLoad: { mode: 'full', allSheets: true, maxRows: null },
			},
			suggestedFix:
				'Use read or agent-view for bounded inspection, then call plan without load options.',
		},
	)
}

class PreparedPlanStore {
	private readonly handles = new Map<string, PreparedPlanRecord>()
	private readonly maxHandles: number
	private readonly ttlMs: number
	private readonly now: () => number

	constructor(options: McpServerOptions = {}) {
		this.maxHandles = positiveIntegerOption(
			options.preparedPlanMaxHandles,
			DEFAULT_PREPARED_PLAN_MAX_HANDLES,
		)
		this.ttlMs = positiveIntegerOption(options.preparedPlanTtlMs, DEFAULT_PREPARED_PLAN_TTL_MS)
		this.now = options.now ?? Date.now
	}

	add(handle: PreparedPlanHandle): PreparedPlanMetadata {
		this.pruneExpired()
		while (this.handles.size >= this.maxHandles) {
			const oldest = this.handles.keys().next().value
			if (oldest === undefined) break
			this.handles.delete(oldest)
		}
		const id = randomUUID()
		const expiresAtMs = this.now() + this.ttlMs
		this.handles.set(id, { handle, expiresAtMs })
		return { id, expiresAt: new Date(expiresAtMs).toISOString(), ttlMs: this.ttlMs }
	}

	take(id: string): PreparedPlanHandle | null {
		this.pruneExpired()
		const record = this.handles.get(id)
		if (!record) return null
		this.handles.delete(id)
		return record.handle
	}

	private pruneExpired(): void {
		const now = this.now()
		for (const [id, record] of this.handles) {
			if (record.expiresAtMs <= now) this.handles.delete(id)
		}
	}
}

function resolveOperationInputForWorkbook(
	wb: AscendWorkbook,
	ops: readonly Record<string, unknown>[] | undefined,
	mutations: readonly unknown[] | undefined,
): ResolvedOperationInput {
	const shape = resolveOperationInputShape(ops, mutations)
	if (!shape.ok || !('mutations' in shape)) return shape
	return compilePathMutationInput(wb, shape.mutations)
}

type OperationInputShape =
	| { readonly ok: true; readonly ops: readonly Operation[] }
	| { readonly ok: true; readonly mutations: readonly PathMutation[] }
	| { readonly ok: false; readonly error: AscendError }

function resolveOperationInputShape(
	ops: readonly Record<string, unknown>[] | undefined,
	mutations: readonly unknown[] | undefined,
): OperationInputShape {
	const hasOpsKey = ops !== undefined
	const hasMutationsKey = mutations !== undefined
	if (hasOpsKey && hasMutationsKey) {
		return {
			ok: false,
			error: ascendError('VALIDATION_ERROR', 'Provide either ops or mutations, not both', {
				retryStrategy: 'modified',
				suggestedFix: 'Send canonical operations in ops or path-addressed mutations in mutations.',
			}),
		}
	}
	const hasOps = ops !== undefined && ops.length > 0
	const hasMutations = mutations !== undefined && mutations.length > 0
	if (!hasOps && !hasMutations) {
		return {
			ok: false,
			error: ascendError('VALIDATION_ERROR', 'Missing or invalid ops or mutations', {
				retryStrategy: 'modified',
				suggestedFix:
					'Send non-empty ops, or send mutations like {"path":"/sheets/Sheet1/cells/A1/value","value":123}.',
			}),
		}
	}
	if (hasOps) {
		const parsed = parseOperations(ops ?? [])
		if (!parsed.ok) {
			return {
				ok: false,
				error: ascendError('VALIDATION_ERROR', parsed.error, {
					details: operationValidationDetails(parsed),
					retryStrategy: 'modified',
					suggestedFix: 'Use ascend.list_operations for canonical operation schemas and examples.',
				}),
			}
		}
		return { ok: true, ops: parsed.value }
	}
	const parsedMutations = parsePathMutationBody(mutations ?? [])
	if (!parsedMutations.ok) return parsedMutations
	return { ok: true, mutations: parsedMutations.mutations }
}

function parsePathMutationBody(
	value: readonly unknown[],
):
	| { readonly ok: true; readonly mutations: readonly PathMutation[] }
	| { readonly ok: false; readonly error: AscendError } {
	const mutations: PathMutation[] = []
	for (const [index, entry] of value.entries()) {
		if (entry === null || typeof entry !== 'object' || Array.isArray(entry)) {
			return {
				ok: false,
				error: pathMutationShapeError(index, 'Mutation must be an object with path and value.'),
			}
		}
		const path = (entry as Record<string, unknown>).path
		if (
			typeof path !== 'string' &&
			(!Array.isArray(path) || !path.every((segment) => typeof segment === 'string'))
		) {
			return {
				ok: false,
				error: pathMutationShapeError(index, 'Mutation path must be a string or string array.'),
			}
		}
		mutations.push({
			path,
			...(Object.hasOwn(entry, 'value') ? { value: (entry as Record<string, unknown>).value } : {}),
		})
	}
	return { ok: true, mutations }
}

function pathMutationShapeError(index: number, message: string): AscendError {
	return ascendError('VALIDATION_ERROR', message, {
		details: {
			issueCount: 1,
			issues: [`mutations[${index}]: ${message}`],
			issueDetails: [
				{ code: 'invalid_path_mutation', mutationIndex: index, path: `mutations[${index}]` },
			],
		},
		retryStrategy: 'modified',
		suggestedFix: 'Use mutations shaped like {"path":"/sheets/Sheet1/cells/A1/value","value":123}.',
	})
}

function compilePathMutationInput(
	wb: AscendWorkbook,
	mutations: readonly PathMutation[],
): ResolvedOperationInput {
	const compiled = wb.compilePathMutations(mutations)
	if (!compiled.replayable) return { ok: false, error: pathMutationCompileError(compiled) }
	return { ok: true, ops: compiled.ops, pathMutations: compiled }
}

function pathMutationCompileError(result: PathMutationResult): AscendError {
	return ascendError('VALIDATION_ERROR', 'Path mutation compilation failed', {
		details: {
			mutationCount: result.mutationCount,
			issueCount: result.issueCount,
			issues: result.issues.map((issue) => issue.message),
			issueDetails: result.issues,
			compiledOps: result.ops,
			supportedPathShapes: SUPPORTED_PATH_MUTATION_SHAPES,
		},
		retryStrategy: 'modified',
		suggestedFix:
			'Use supported paths such as /sheets/{sheet}/cells/{A1}/value, /sheets/{sheet}/cells/{A1}/formula, /sheets/{sheet}/ranges/{A1:B2}/clear, /tables/{table}/rows/append, or /names/{name}/ref.',
	})
}

function withPathMutationResult<T extends object>(
	result: T,
	compiled: PathMutationResult | undefined,
): T | (T & { readonly pathMutations: PathMutationResult }) {
	return compiled ? { ...result, pathMutations: compiled } : result
}

function withPreparedPlanHandle<T extends object>(
	result: T,
	preparedPlan: PreparedPlanMetadata | undefined,
): T | (T & { readonly preparedPlan: PreparedPlanMetadata }) {
	return preparedPlan ? { ...result, preparedPlan } : result
}

function preparedPlanHandle(prepared: PreparedAgentPlan): PreparedPlanHandle {
	const commit = prepared.commit
	return {
		file: prepared.file,
		inputSha256: prepared.inputSha256,
		planDigest: prepared.planDigest,
		commit: (options) => commit(options),
	}
}

export function createServer(options: McpServerOptions = {}): McpServer {
	const server = new McpServer({
		name: 'ascend',
		version: '0.0.0',
	})
	const preparedPlans = new PreparedPlanStore(options)

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
		'ascend.raw_part',
		'Read a bounded raw XLSX OPC package part for diagnostics. This is raw package inspection, not semantic workbook truth, and it does not bypass write-risk checks.',
		{
			file: z.string().describe('Path to workbook file'),
			partPath: z
				.string()
				.describe('Exact case-sensitive package part path, for example xl/workbook.xml'),
			encoding: z
				.enum(['text', 'base64', 'none'])
				.optional()
				.describe('Return part preview as UTF-8 text, base64, or metadata only'),
			maxBytes: z
				.number()
				.int()
				.nonnegative()
				.optional()
				.describe('Maximum bytes to include in the text/base64 preview'),
			caseInsensitive: z
				.boolean()
				.optional()
				.describe(
					'Allow diagnostic case-insensitive fallback; result marks fallback matches explicitly',
				),
		},
		async ({ file, partPath, encoding, maxBytes, caseInsensitive }) => {
			try {
				if (maxBytes !== undefined && maxBytes > MAX_MCP_RAW_PART_MAX_BYTES) {
					return errorResponse(
						ascendError('VALIDATION_ERROR', 'Invalid raw part maxBytes', {
							details: {
								receivedMaxBytes: maxBytes,
								rule: `at most ${MAX_MCP_RAW_PART_MAX_BYTES} bytes`,
							},
							retryStrategy: 'modified',
							suggestedFix: `Omit maxBytes for the bounded default, or provide a nonnegative integer up to ${MAX_MCP_RAW_PART_MAX_BYTES}.`,
						}),
					)
				}
				const wb = await WorkbookDocument.open(file, { mode: 'metadata-only' })
				const result = await wb.rawPackagePart({
					partPath,
					...(encoding ? { encoding } : {}),
					maxBytes: maxBytes ?? DEFAULT_MCP_RAW_PART_MAX_BYTES,
					...(caseInsensitive === true ? { caseInsensitive: true } : {}),
				})
				if (!result.validPath) {
					return errorResponse(
						ascendError('VALIDATION_ERROR', `Invalid package part path: ${partPath}`, {
							details: { ...result },
							retryStrategy: 'modified',
							suggestedFix:
								'Use an exact OPC package part path with forward slashes and no dot or empty segments.',
						}),
					)
				}
				if (result.caseInsensitiveAmbiguous) {
					return errorResponse(
						ascendError('VALIDATION_ERROR', `Ambiguous package part path: ${partPath}`, {
							details: { ...result },
							retryStrategy: 'modified',
							suggestedFix:
								'Retry with an exact case-sensitive package part path from ascend.package_graph.',
						}),
					)
				}
				if (!result.found) {
					return errorResponse(
						ascendError('FILE_NOT_FOUND', `Package part not found: ${partPath}`, {
							details: { ...result },
							retryStrategy: 'modified',
							suggestedFix: 'Call ascend.package_graph and retry with an exact part path.',
						}),
					)
				}
				return okResponse(
					result,
					`Read ${result.byteLength ?? 0} byte(s) from package part ${result.partPath}`,
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
			maxRows: z
				.number()
				.int()
				.positive()
				.optional()
				.describe(
					'Maximum worksheet rows to hydrate for lazy preview reads; defaults to rowOffset + rowLimit when rowLimit is provided',
				),
			preview: z
				.boolean()
				.optional()
				.describe(
					`When true, default to the first ${DEFAULT_AGENT_PREVIEW_ROWS} rows if rowLimit is omitted`,
				),
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
			maxRows,
			preview,
			changedSince,
			format,
			display,
			headers,
			cols,
		}) => {
			try {
				const mode = format ?? 'cells'
				const firstWindow = preview === true || mode === 'compact'
				const effectiveRowLimit = firstWindowRowLimit(rowLimit, firstWindow)
				const wb = await WorkbookDocument.open(file, {
					mode: 'values',
					...(sheet ? { sheets: [sheet] } : {}),
					...readPreviewLoadOptions(maxRows, rowOffset, rowLimit, firstWindow),
				})
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
					...(effectiveRowLimit !== undefined ? { rowLimit: effectiveRowLimit } : {}),
				}
				const info =
					mode === 'compact'
						? withPartialLoadInfo(
								buildCompactReadResult(
									handle.readWindowCompact(range, {
										...readOpts,
										includeRefs: false,
										omitEmpty: true,
										flatValues: true,
										changedSince: changedSince ?? '',
									}),
									cols,
								),
								wb,
							)
						: mode === 'tsv'
							? withPartialLoadInfo(buildTsvReadResult(handle.readRows(range, readOpts), cols), wb)
							: mode === 'rows'
								? withPartialLoadInfo(pruneRowsInfo(handle.readRows(range, readOpts), cols), wb)
								: mode === 'objects'
									? withPartialLoadInfo(
											pruneObjectsInfo(
												handle.readObjects(range, {
													...readOpts,
													headers: headers && headers.length > 0 ? headers : 'first-row',
												}),
												cols,
											),
											wb,
										)
									: withPartialLoadInfo(
											pruneWindowInfo(handle.readWindow(range, readOpts), cols),
											wb,
										)
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
		'ascend.dump',
		'Dump supported workbook cells and formulas as deterministic operations that can be replayed through ascend.plan and ascend.commit',
		{
			file: z.string().describe('Path to workbook file'),
			sheet: z.string().optional().describe('Optional sheet name to dump'),
			valuesOnly: z.boolean().optional().describe('Omit formulas from the dumped batch'),
			formulasOnly: z.boolean().optional().describe('Omit literal values from the dumped batch'),
		},
		async ({ file, sheet, valuesOnly, formulasOnly }) => {
			if (valuesOnly && formulasOnly) {
				return errorResponse(
					ascendError('INVALID_ARGUMENT', 'Use either valuesOnly or formulasOnly, not both', {
						retryStrategy: 'modified',
						suggestedFix: 'Retry with only one dump mode flag.',
					}),
				)
			}
			try {
				const wb = await AscendWorkbook.open(file)
				const result = wb.dumpBatch({
					...(sheet ? { sheets: [sheet] } : {}),
					...(valuesOnly ? { includeFormulas: false } : {}),
					...(formulasOnly ? { includeValues: false } : {}),
				})
				return okResponse(
					result,
					`Dumped ${result.ops.length} replay operation(s) from ${result.sheetCount} sheet(s)`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.template_merge',
		'Compile {{key}} workbook template placeholders into deterministic operations that can be replayed through ascend.plan and ascend.commit',
		{
			file: z.string().describe('Path to workbook file'),
			data: z
				.record(z.string(), z.union([z.string(), z.number(), z.boolean(), z.null()]))
				.describe('Template values keyed by placeholder name; values must be JSON scalars'),
			sheet: z.string().optional().describe('Optional sheet name to scan'),
			valuesOnly: z.boolean().optional().describe('Omit formulas from the template merge scan'),
			formulasOnly: z
				.boolean()
				.optional()
				.describe('Omit literal values from the template merge scan'),
			open: z.string().optional().describe('Placeholder opening delimiter, default {{'),
			close: z.string().optional().describe('Placeholder closing delimiter, default }}'),
		},
		async ({ file, data, sheet, valuesOnly, formulasOnly, open, close }) => {
			if (valuesOnly && formulasOnly) {
				return errorResponse(
					ascendError('INVALID_ARGUMENT', 'Use either valuesOnly or formulasOnly, not both', {
						retryStrategy: 'modified',
						suggestedFix: 'Retry with only one template merge mode flag.',
					}),
				)
			}
			try {
				const wb = await AscendWorkbook.open(file)
				const result = wb.templateMerge(data, {
					...(sheet ? { sheets: [sheet] } : {}),
					...(valuesOnly ? { includeFormulas: false } : {}),
					...(formulasOnly ? { includeValues: false } : {}),
					...(open || close
						? {
								delimiters: {
									...(open ? { open } : {}),
									...(close ? { close } : {}),
								},
							}
						: {}),
				})
				return okResponse(
					result,
					`Compiled ${result.ops.length} template merge operation(s) from ${result.sheetCount} sheet(s)`,
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
			maxRows: z
				.number()
				.int()
				.positive()
				.optional()
				.describe('Maximum worksheet rows to hydrate; partial agent views include load metadata'),
			rowChunkSize: z.number().int().positive().optional().describe('Rows per streamed chunk'),
			sampleRowLimit: z.number().int().positive().optional().describe('Maximum sample rows'),
			sampleValueLimit: z
				.number()
				.int()
				.positive()
				.optional()
				.describe('Maximum sample values per column'),
		},
		async ({ file, range, sheet, maxRows, rowChunkSize, sampleRowLimit, sampleValueLimit }) => {
			try {
				const wb = await WorkbookDocument.open(
					file,
					sheet
						? {
								mode: 'formula',
								sheets: [sheet],
								...(maxRows !== undefined ? { maxRows } : {}),
							}
						: { mode: 'formula', ...(maxRows !== undefined ? { maxRows } : {}) },
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
				return okResponse(
					withPartialLoadInfo(view, wb),
					`Generated agent view for ${range} on "${sheetName}"`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.preview',
		'Preview canonical operations or path-addressed mutations without saving the workbook',
		{
			file: z.string().describe('Path to workbook file'),
			ops: z
				.array(z.record(z.string(), z.unknown()))
				.optional()
				.describe('Canonical operations to preview'),
			mutations: z
				.array(pathMutationSchema)
				.optional()
				.describe('Path-addressed mutations to compile and preview'),
			journal: z
				.boolean()
				.optional()
				.describe('Include reversible mutation journal metadata for supported operations'),
		},
		async ({ file, ops, mutations, journal }) => {
			try {
				const wb = await Ascend.open(file)
				const input = resolveOperationInputForWorkbook(wb, ops, mutations)
				if (!input.ok) return errorResponse(input.error)
				const result = wb.preview(input.ops, journal ? { journal: true } : undefined)
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
				return okResponse(
					withPathMutationResult(result, input.pathMutations),
					`Previewed ${input.ops.length} operation(s)`,
				)
			} catch (e) {
				return errorResponse(
					e instanceof AscendException ? e.ascendError : String(e instanceof Error ? e.message : e),
				)
			}
		},
	)

	server.tool(
		'ascend.write',
		'Apply canonical operations or path-addressed mutations to a workbook',
		{
			file: z.string().describe('Path to workbook file'),
			ops: z
				.array(z.record(z.string(), z.unknown()))
				.optional()
				.describe('Canonical operations to apply'),
			mutations: z
				.array(pathMutationSchema)
				.optional()
				.describe('Path-addressed mutations to compile and apply'),
			journal: z
				.boolean()
				.optional()
				.describe('Include reversible mutation journal metadata for supported operations'),
		},
		async ({ file, ops, mutations, journal }) => {
			try {
				const wb = await Ascend.open(file)
				const input = resolveOperationInputForWorkbook(wb, ops, mutations)
				if (!input.ok) return errorResponse(input.error)
				const result = wb.apply(input.ops, journal ? { journal: true } : undefined)
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
				return okResponse(
					withPathMutationResult(result, input.pathMutations),
					`Applied ${input.ops.length} operation(s)`,
				)
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
		'Agent-safe edit planning: validate, preview, recalc-audit, package-graph-audit, and preservation-audit operations or path mutations without saving',
		{
			file: z.string().describe('Path to workbook file'),
			ops: z
				.array(z.record(z.string(), z.unknown()))
				.optional()
				.describe('Canonical operations to plan'),
			mutations: z
				.array(pathMutationSchema)
				.optional()
				.describe('Path-addressed mutations to compile and plan'),
			compact: z
				.boolean()
				.optional()
				.describe('Return a compact plan payload with bounded changed-cell details'),
			prepare: z
				.boolean()
				.optional()
				.describe(
					'Return a one-shot prepared plan handle that can be passed to ascend.commit; defaults to true, set false to opt out',
				),
			maxChangedCells: z
				.number()
				.int()
				.nonnegative()
				.optional()
				.describe('Maximum preview changed cells to include when compact is true'),
			maxRows: z
				.number()
				.int()
				.positive()
				.optional()
				.describe(
					'Unsupported on plan; use ascend.read or ascend.agent_view for capped inspection',
				),
		},
		async ({ file, ops, mutations, compact, prepare, maxChangedCells, maxRows }) => {
			try {
				if (maxRows !== undefined) return errorResponse(agentPlanLoadOptionsError(['maxRows']))
				const inputShape = resolveOperationInputShape(ops, mutations)
				if (!inputShape.ok) return errorResponse(inputShape.error)
				let input: ResolvedOperationInput
				let pathMutations: PathMutationResult | undefined
				let result: Awaited<ReturnType<typeof createAgentPlan>> | null
				let preparedPlan: PreparedPlanMetadata | undefined
				const shouldPrepare = prepare !== false
				if ('mutations' in inputShape) {
					const opened = await AscendWorkbook.openSourceBytes(file)
					const inputSha256 = sha256Bytes(opened.sourceBytes)
					const wb = opened.workbook
					input = compilePathMutationInput(wb, inputShape.mutations)
					if (input.ok) pathMutations = input.pathMutations
					result = input.ok
						? await createAgentPlanFromWorkbook(file, inputSha256, wb, input.ops)
						: null
					if (input.ok && result && shouldPrepare) {
						const preparedOps = input.ops
						const planDigest = result.planDigest
						preparedPlan = preparedPlans.add({
							file,
							inputSha256,
							planDigest,
							...(pathMutations !== undefined ? { pathMutations } : {}),
							commit: async (options) => {
								const current = await Bun.file(file).bytes()
								const currentSha256 = sha256Bytes(current)
								if (currentSha256 !== inputSha256) {
									throw new AscendException(
										ascendError(
											'VALIDATION_ERROR',
											'Input workbook changed after agent plan was prepared',
											{
												details: {
													expected: inputSha256,
													actual: currentSha256,
													planDigest,
												},
												suggestedFix: 'Re-run ascend plan and commit with the new input workbook.',
											},
										),
									)
								}
								return commitAgentPlanFromWorkbook(file, inputSha256, wb, preparedOps, options, {
									sourceBytes: opened.sourceBytes,
								})
							},
						})
					}
				} else {
					input = inputShape
					if (shouldPrepare) {
						const prepared = await createPreparedAgentPlan(file, inputShape.ops)
						result = prepared.plan
						preparedPlan = preparedPlans.add(preparedPlanHandle(prepared))
					} else {
						result = await createAgentPlan(file, inputShape.ops)
					}
				}
				if (!input.ok) return errorResponse(input.error)
				if (!result) return errorResponse('Plan failed')
				if (result.preview.errors.length > 0) {
					const first = result.preview.errors[0]
					return errorResponse(
						first
							? { ...first, details: { ...(first.details ?? {}), plan: result } }
							: ascendError('VALIDATION_ERROR', 'Plan failed', { details: { plan: result } }),
					)
				}
				const payload = compact
					? compactAgentPlanResult(result, {
							...(maxChangedCells !== undefined ? { maxChangedCells } : {}),
						})
					: result
				return okResponse(
					withPreparedPlanHandle(withPathMutationResult(payload, pathMutations), preparedPlan),
					`Planned ${input.ops.length} operation(s)`,
				)
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
			file: z.string().optional().describe('Path to workbook file'),
			planHandle: z
				.string()
				.optional()
				.describe('Prepared plan handle returned by ascend.plan with prepare=true'),
			ops: z
				.array(z.record(z.string(), z.unknown()))
				.optional()
				.describe('Canonical operations to commit'),
			mutations: z
				.array(pathMutationSchema)
				.optional()
				.describe('Path-addressed mutations to compile and commit'),
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
				.describe('Allow preserved/unsupported feature loss by feature, feature:tier, or "all"'),
			approvals: z
				.union([z.string(), z.array(z.string())])
				.optional()
				.describe('Approve explicit plan approval ids or "all"'),
			compact: z
				.boolean()
				.optional()
				.describe('Return compact commit verification counts instead of full trace artifacts'),
			maxRows: z
				.number()
				.int()
				.positive()
				.optional()
				.describe(
					'Unsupported on commit; use ascend.read or ascend.agent_view for capped inspection',
				),
		},
		async ({
			file,
			planHandle,
			ops,
			mutations,
			output,
			inPlace,
			backup,
			expectSha256,
			allowLoss,
			approvals,
			compact,
			maxRows,
		}) => {
			try {
				if (maxRows !== undefined) return errorResponse(agentPlanLoadOptionsError(['maxRows']))
				const options: AgentCommitOptions = {
					...(output ? { output } : {}),
					...(inPlace ? { inPlace: true } : {}),
					...(backup ? { backup } : {}),
					...(expectSha256 ? { expectSha256 } : {}),
					...(allowLoss ? { allowLoss: parseAllowLoss(allowLoss) } : {}),
					...(approvals ? { approvals: parseStringListOrAll(approvals) } : {}),
				}
				if (planHandle) {
					const prepared = preparedPlans.take(planHandle)
					if (!prepared) return errorResponse('Prepared plan handle was not found')
					const result = await prepared.commit(options)
					const payload = compact ? compactAgentCommitResult(result) : result
					return okResponse(
						withPathMutationResult(payload, prepared.pathMutations),
						`Committed ${result.operationCount} operation(s)`,
					)
				}
				if (!file) return errorResponse('Missing file')
				const inputShape = resolveOperationInputShape(ops, mutations)
				if (!inputShape.ok) return errorResponse(inputShape.error)
				let input: ResolvedOperationInput
				let pathMutations: PathMutationResult | undefined
				let result: Awaited<ReturnType<typeof commitAgentPlan>>
				if ('mutations' in inputShape) {
					const opened = await AscendWorkbook.openSourceBytes(file)
					const inputSha256 = sha256Bytes(opened.sourceBytes)
					const wb = opened.workbook
					input = compilePathMutationInput(wb, inputShape.mutations)
					if (!input.ok) return errorResponse(input.error)
					pathMutations = input.pathMutations
					result = await commitAgentPlanFromWorkbook(file, inputSha256, wb, input.ops, options, {
						sourceBytes: opened.sourceBytes,
					})
				} else {
					input = inputShape
					result = await commitAgentPlan(file, input.ops, options)
				}
				const payload = compact ? compactAgentCommitResult(result) : result
				return okResponse(
					withPathMutationResult(payload, pathMutations),
					`Committed ${input.ops.length} operation(s)`,
				)
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
			maxRows: z
				.number()
				.int()
				.positive()
				.optional()
				.describe('Maximum worksheet rows to hydrate; capped checks return partial-load metadata'),
		},
		async ({ file, maxRows }) => {
			try {
				const wb = await WorkbookDocument.open(file, {
					...(maxRows !== undefined ? { maxRows } : {}),
				})
				const result = withPartialLoadInfo(wb.check(), wb)
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
			maxRows: z
				.number()
				.int()
				.positive()
				.optional()
				.describe(
					'Maximum worksheet rows to hydrate; capped lint results include partial-load warnings',
				),
		},
		async ({ file, maxRows }) => {
			try {
				const wb = await WorkbookDocument.open(file, {
					mode: 'formula',
					...(maxRows !== undefined ? { maxRows } : {}),
				})
				const result = withPartialLoadInfo(wb.lint(), wb)
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
			maxRows: z
				.number()
				.int()
				.positive()
				.optional()
				.describe(
					'Maximum worksheet rows to hydrate; partial traces return structured diagnostics',
				),
		},
		async ({ file, cell, maxRows }) => {
			try {
				const wb = await WorkbookDocument.open(file, {
					mode: 'formula',
					...(maxRows !== undefined ? { maxRows } : {}),
				})
				const traceIssue = wb.traceIssue(cell)
				if (traceIssue) {
					return errorResponse(
						ascendError('VALIDATION_ERROR', traceIssue.message, {
							details: {
								...(traceIssue.rule ? { rule: traceIssue.rule } : {}),
								...(traceIssue.ref ? { ref: traceIssue.ref } : {}),
								load: wb.inspect().load,
							},
							...(traceIssue.suggestedFix ? { suggestedFix: traceIssue.suggestedFix } : {}),
						}),
					)
				}
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

function readPreviewLoadOptions(
	explicitMaxRows: number | undefined,
	rowOffset: number | undefined,
	rowLimit: number | undefined,
	preview: boolean,
): { readonly maxRows?: number } {
	if (explicitMaxRows !== undefined) return { maxRows: explicitMaxRows }
	if (rowLimit === undefined && !preview) return {}
	const offset = Math.max(0, rowOffset ?? 0)
	return { maxRows: offset + (rowLimit ?? DEFAULT_AGENT_PREVIEW_ROWS) }
}

function firstWindowRowLimit(rowLimit: number | undefined, preview: boolean): number | undefined {
	return rowLimit ?? (preview ? DEFAULT_AGENT_PREVIEW_ROWS : undefined)
}

function withPartialLoadInfo<T extends object>(info: T, wb: WorkbookDocument): T {
	const load = wb.inspect().load
	return load.isPartial ? ({ ...info, load } as T) : info
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
		snapshot: info.snapshot,
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
		snapshot: info.snapshot,
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
