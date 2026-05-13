import { randomUUID } from 'node:crypto'
import {
	type AscendError,
	AscendException,
	ascendError,
	type CellValue,
	type Operation,
} from '@ascend/schema'
import {
	type AgentCommitOptions,
	AscendWorkbook,
	type CapabilityFilters,
	type CapabilityPriority,
	type CapabilityStatus,
	type CompactRangeWindowInfo,
	commitAgentPlan,
	commitAgentPlanFromWorkbook,
	compactAgentPlanResult,
	createAgentPlan,
	createAgentPlanFromWorkbook,
	createPreparedAgentPlan,
	createRepairPlan,
	formatDisplayCellValue,
	getOperationsSchema,
	listCapabilities,
	listOperations,
	normalizeExportFormat,
	operationValidationDetails,
	type PathMutation,
	type PathMutationResult,
	type PivotOutputMaterializeMode,
	type PivotOutputMaterializeOptions,
	type PreparedAgentPlan,
	parseOperations,
	SUPPORTED_PATH_MUTATION_SHAPES,
	sha256Bytes,
	summarizeCapabilities,
	WorkbookDocument,
} from '@ascend/sdk'
import { binaryResponse, jsonFailure, jsonFailureError, jsonSuccess } from './response.ts'

const DEFAULT_API_RAW_PART_MAX_BYTES = 64 * 1024
const MAX_API_RAW_PART_MAX_BYTES = 1024 * 1024
const DEFAULT_AGENT_PREVIEW_ROWS = 500

async function parseJson<T>(req: Request): Promise<T | null> {
	try {
		const body = await req.json()
		return body as T
	} catch {
		return null
	}
}

function requireString(obj: unknown, key: string): string | null {
	if (obj === null || typeof obj !== 'object') return null
	const v = (obj as Record<string, unknown>)[key]
	return typeof v === 'string' ? v : null
}

function requireArray(obj: unknown, key: string): unknown[] | null {
	if (obj === null || typeof obj !== 'object') return null
	const v = (obj as Record<string, unknown>)[key]
	return Array.isArray(v) ? v : null
}

function requireOptionalNumber(obj: unknown, key: string): number | undefined {
	if (obj === null || typeof obj !== 'object') return undefined
	const v = (obj as Record<string, unknown>)[key]
	return typeof v === 'number' && Number.isFinite(v) ? v : undefined
}

function firstWindowMaxRows(
	explicitMaxRows: number | undefined,
	rowOffset: number | undefined,
	rowLimit: number | undefined,
	preview: boolean,
): number | undefined {
	if (explicitMaxRows !== undefined) return Math.max(1, Math.floor(explicitMaxRows))
	if (rowLimit === undefined && !preview) return undefined
	const offset = Math.max(0, Math.floor(rowOffset ?? 0))
	const limit = rowLimit ?? DEFAULT_AGENT_PREVIEW_ROWS
	return offset + Math.max(1, Math.floor(limit))
}

function firstWindowRowLimit(rowLimit: number | undefined, preview: boolean): number | undefined {
	return rowLimit ?? (preview ? DEFAULT_AGENT_PREVIEW_ROWS : undefined)
}

function withPartialLoadInfo<T extends object>(info: T, wb: WorkbookDocument): T {
	const load = wb.inspect().load
	return load.isPartial ? ({ ...info, load } as T) : info
}

function rawPartEncoding(
	value: unknown,
): { readonly ok: true; readonly encoding?: 'text' | 'base64' | 'none' } | { readonly ok: false } {
	if (value === undefined) return { ok: true }
	return value === 'text' || value === 'base64' || value === 'none'
		? { ok: true, encoding: value }
		: { ok: false }
}

function rawPartMaxBytes(
	value: unknown,
):
	| { readonly ok: true; readonly maxBytes?: number }
	| { readonly ok: false; readonly rule: string } {
	if (value === undefined) return { ok: true }
	if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) {
		return { ok: false, rule: 'nonnegative integer' }
	}
	if (value > MAX_API_RAW_PART_MAX_BYTES) {
		return { ok: false, rule: `at most ${MAX_API_RAW_PART_MAX_BYTES} bytes` }
	}
	return { ok: true, maxBytes: value }
}

type OperationInput =
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

export interface ApiFetchOptions {
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

class PreparedPlanStore {
	private readonly handles = new Map<string, PreparedPlanRecord>()
	private readonly maxHandles: number
	private readonly ttlMs: number
	private readonly now: () => number

	constructor(options: ApiFetchOptions = {}) {
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
	body: Record<string, unknown> | null,
): OperationInput {
	const shape = resolveOperationInputShape(body)
	if (!shape.ok || !('mutations' in shape)) return shape
	return compilePathMutationInput(wb, shape.mutations)
}

type OperationInputShape =
	| { readonly ok: true; readonly ops: readonly Operation[] }
	| { readonly ok: true; readonly mutations: readonly PathMutation[] }
	| { readonly ok: false; readonly error: AscendError }

function resolveOperationInputShape(body: Record<string, unknown> | null): OperationInputShape {
	const hasOpsKey = body !== null && Object.hasOwn(body, 'ops')
	const hasMutationsKey = body !== null && Object.hasOwn(body, 'mutations')
	const opsArr = body ? requireArray(body, 'ops') : null
	const mutationArr = body ? requireArray(body, 'mutations') : null
	if (hasOpsKey && hasMutationsKey) {
		return {
			ok: false,
			error: ascendError('VALIDATION_ERROR', 'Provide either ops or mutations, not both', {
				retryStrategy: 'modified',
				suggestedFix: 'Send canonical operations in ops or path-addressed mutations in mutations.',
			}),
		}
	}
	const hasOps = opsArr !== null && opsArr.length > 0
	const hasMutations = mutationArr !== null && mutationArr.length > 0
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
		const parsed = parseOperations(opsArr)
		if (!parsed.ok) {
			return {
				ok: false,
				error: ascendError('VALIDATION_ERROR', parsed.error, {
					details: operationValidationDetails(parsed),
					retryStrategy: 'modified',
					suggestedFix: 'Call /operations for canonical operation schemas and examples.',
				}),
			}
		}
		return { ok: true, ops: parsed.value }
	}

	const parsedMutations = parsePathMutationBody(mutationArr ?? [])
	if (!parsedMutations.ok) return parsedMutations
	return { ok: true, mutations: parsedMutations.mutations }
}

function compilePathMutationInput(
	wb: AscendWorkbook,
	mutations: readonly PathMutation[],
): OperationInput {
	const compiled = wb.compilePathMutations(mutations)
	if (!compiled.replayable) {
		return { ok: false, error: pathMutationCompileError(compiled) }
	}
	return { ok: true, ops: compiled.ops, pathMutations: compiled }
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

function parsePivotOutputMaterializeOptions(
	body: Record<string, unknown> | null,
): PivotOutputMaterializeOptions | null {
	if (!body) return {}
	const mode = body.mode
	if (mode !== undefined && mode !== 'missing' && mode !== 'mismatches' && mode !== 'all') {
		return null
	}
	const pivotTable = typeof body.pivotTable === 'string' ? body.pivotTable : undefined
	const partPath = typeof body.partPath === 'string' ? body.partPath : undefined
	return {
		...(pivotTable ? { pivotTable } : {}),
		...(partPath ? { partPath } : {}),
		...(mode ? { mode: mode as PivotOutputMaterializeMode } : {}),
	}
}

function parseAllowLoss(value: unknown): readonly string[] | 'all' | undefined {
	return parseStringListOrAll(value)
}

function parseApprovals(value: unknown): readonly string[] | 'all' | undefined {
	return parseStringListOrAll(value)
}

function parseStringListOrAll(value: unknown): readonly string[] | 'all' | undefined {
	const entries =
		typeof value === 'string'
			? value.split(',')
			: Array.isArray(value)
				? value.filter((entry): entry is string => typeof entry === 'string')
				: []
	const normalized = entries.map((entry) => entry.trim()).filter((entry) => entry.length > 0)
	if (normalized.length === 0) return undefined
	if (normalized.some((entry) => entry.toLowerCase() === 'all')) return 'all'
	return normalized
}

function statusForError(ae: AscendError): number {
	if (
		ae.code === 'SHEET_NOT_FOUND' ||
		ae.code === 'NAME_NOT_FOUND' ||
		ae.code === 'TABLE_NOT_FOUND'
	)
		return 404
	if (
		ae.code === 'IMPORT_ERROR' ||
		ae.code === 'INVALID_REF' ||
		ae.code === 'INVALID_RANGE' ||
		ae.code === 'VALIDATION_ERROR' ||
		ae.code === 'INVALID_ARGUMENT'
	)
		return 400
	return 500
}

function isFileNotFoundError(e: unknown): boolean {
	if (
		typeof e === 'object' &&
		e !== null &&
		'code' in e &&
		(e as { code: string }).code === 'ENOENT'
	)
		return true
	if (typeof e === 'object' && e !== null && 'errno' in e && (e as { errno: number }).errno === -2)
		return true
	return false
}

function handleError(e: unknown, fileContext?: string): Response {
	if (e instanceof AscendException) {
		const status = statusForError(e.ascendError)
		return jsonFailureError(e.ascendError, status)
	}
	if (isFileNotFoundError(e))
		return jsonFailure(fileContext ? `File not found: ${fileContext}` : 'File not found', 404)
	const msg = e instanceof Error ? e.message : String(e)
	return jsonFailure(msg, 500)
}

function sheetNotFoundResponse(sheetName: string, wb: WorkbookDocument | AscendWorkbook): Response {
	const available = wb.inspect().sheets.map((s) => s.name)
	return jsonFailureError(
		ascendError('SHEET_NOT_FOUND', `Sheet "${sheetName}" not found`, {
			suggestedFix:
				available.length > 0
					? `Available sheets: ${available.join(', ')}`
					: 'Workbook has no sheets',
			details: { availableSheets: available },
		}),
		404,
	)
}

function activeContentPayload(wb: WorkbookDocument) {
	const info = wb.inspect()
	const activeFeatureFamilies = new Set([
		'preservedMacro',
		'preservedMacroSheet',
		'preservedActiveContent',
		'preservedSignature',
	])
	return {
		activeContentCount: info.activeContentCount,
		macroSheetCount: info.macroSheetCount,
		activeContent: info.activeContent,
		macroSheets: info.macroSheets,
		compatibilityFeatures: info.compatibility.features.filter(
			(feature) =>
				activeFeatureFamilies.has(feature.feature) ||
				feature.locations.some((location) =>
					/(vba|macro|activex|ctrlprops|_xmlsignatures|signature)/i.test(location),
				),
		),
		capabilityWarnings: info.capabilityWarnings.filter(
			(warning) => warning.family === 'active content',
		),
	}
}

const CORS_HEADERS: Record<string, string> = {
	'Access-Control-Allow-Origin': '*',
	'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
	'Access-Control-Allow-Headers': 'Content-Type',
}

function withCors(res: Response): Response {
	for (const [k, v] of Object.entries(CORS_HEADERS)) {
		res.headers.set(k, v)
	}
	return res
}

export function createApiFetch(options: ApiFetchOptions = {}) {
	const preparedPlans = new PreparedPlanStore(options)
	return async (req: Request) => {
		const url = new URL(req.url)
		const method = req.method
		const path = url.pathname

		if (method === 'OPTIONS') {
			return withCors(new Response(null, { status: 204 }))
		}

		try {
			if (method === 'GET' && path === '/health') {
				return withCors(jsonSuccess({ status: 'ok' }))
			}

			if (method === 'GET' && path === '/operations') {
				return withCors(
					jsonSuccess({ operations: listOperations(), schemas: getOperationsSchema() }),
				)
			}

			if (method === 'GET' && path === '/capabilities') {
				const feature = url.searchParams.get('feature')
				const family = url.searchParams.get('family')
				const priority = url.searchParams.get('priority')
				const status = url.searchParams.get('status')
				const filters: CapabilityFilters = {
					...(feature ? { feature } : {}),
					...(family ? { family } : {}),
					...(priority ? { priority: priority as CapabilityPriority } : {}),
					...(status ? { status: status as CapabilityStatus } : {}),
					...(url.searchParams.get('gaps') === 'true' ? { gapsOnly: true } : {}),
				}
				const capabilities = listCapabilities(filters)
				return withCors(jsonSuccess({ summary: summarizeCapabilities(capabilities), capabilities }))
			}

			if (method === 'POST' && path === '/inspect') {
				const body = await parseJson<{ file?: string; sheet?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				const sheetName = body ? requireString(body, 'sheet') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					const wb = await WorkbookDocument.open(
						file,
						sheetName ? { mode: 'values', sheets: [sheetName] } : { mode: 'metadata-only' },
					)
					if (!sheetName) return jsonSuccess(wb.inspect())
					const sheet = wb.inspectSheet(sheetName)
					if (!sheet) return sheetNotFoundResponse(sheetName, wb)
					return jsonSuccess(sheet)
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/active-content') {
				const body = await parseJson<{ file?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					const wb = await WorkbookDocument.open(file, { mode: 'metadata-only' })
					return jsonSuccess(activeContentPayload(wb))
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/package-graph') {
				const body = await parseJson<{ file?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					const wb = await WorkbookDocument.open(file, { mode: 'metadata-only' })
					return jsonSuccess(await wb.packageGraph())
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/raw-part') {
				const body = await parseJson<Record<string, unknown>>(req)
				const file = body ? requireString(body, 'file') : null
				const partPath = body ? requireString(body, 'partPath') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				if (!partPath) return jsonFailure('Missing or invalid partPath', 400)
				const encoding = rawPartEncoding(body?.encoding)
				if (!encoding.ok) {
					return jsonFailureError(
						ascendError('VALIDATION_ERROR', 'Invalid raw part encoding', {
							details: {
								allowedEncodings: ['text', 'base64', 'none'],
								receivedEncoding: body?.encoding,
							},
							suggestedFix: 'Use encoding "text", "base64", or "none".',
						}),
						400,
					)
				}
				const maxBytes = rawPartMaxBytes(body?.maxBytes)
				if (!maxBytes.ok) {
					return jsonFailureError(
						ascendError('VALIDATION_ERROR', 'Invalid raw part maxBytes', {
							details: { receivedMaxBytes: body?.maxBytes, rule: maxBytes.rule },
							suggestedFix: `Omit maxBytes for the bounded default, or provide a nonnegative integer up to ${MAX_API_RAW_PART_MAX_BYTES}.`,
						}),
						400,
					)
				}
				try {
					const wb = await WorkbookDocument.open(file, { mode: 'metadata-only' })
					const result = await wb.rawPackagePart({
						partPath,
						...(encoding.encoding ? { encoding: encoding.encoding } : {}),
						maxBytes: maxBytes.maxBytes ?? DEFAULT_API_RAW_PART_MAX_BYTES,
						...(body?.caseInsensitive === true ? { caseInsensitive: true } : {}),
					})
					if (!result.validPath) {
						return jsonFailureError(
							ascendError('VALIDATION_ERROR', `Invalid package part path: ${partPath}`, {
								details: { ...result },
								suggestedFix:
									'Use an exact OPC package part path with forward slashes and no dot or empty segments.',
							}),
							400,
						)
					}
					if (result.caseInsensitiveAmbiguous) {
						return jsonFailureError(
							ascendError('VALIDATION_ERROR', `Ambiguous package part path: ${partPath}`, {
								details: { ...result },
								suggestedFix:
									'Retry with an exact case-sensitive package part path from /package-graph.',
							}),
							400,
						)
					}
					if (!result.found) {
						return jsonFailureError(
							ascendError('FILE_NOT_FOUND', `Package part not found: ${partPath}`, {
								details: { ...result },
								suggestedFix: 'Call /package-graph and retry with an exact part path.',
							}),
							404,
						)
					}
					return jsonSuccess(result)
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/visuals') {
				const body = await parseJson<{ file?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					const wb = await WorkbookDocument.open(file, { mode: 'full' })
					return jsonSuccess(wb.visualInventory())
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/dump') {
				const body = await parseJson<{
					file?: string
					sheet?: string
					valuesOnly?: boolean
					formulasOnly?: boolean
				}>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				if (body?.valuesOnly === true && body?.formulasOnly === true) {
					return jsonFailure('Use either valuesOnly or formulasOnly, not both', 400)
				}
				try {
					const wb = await AscendWorkbook.open(file)
					return jsonSuccess(
						wb.dumpBatch({
							...(body?.sheet ? { sheets: [body.sheet] } : {}),
							...(body?.valuesOnly === true ? { includeFormulas: false } : {}),
							...(body?.formulasOnly === true ? { includeValues: false } : {}),
						}),
					)
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/template-merge') {
				const body = await parseJson<{
					file?: string
					data?: Record<string, unknown>
					sheet?: string
					valuesOnly?: boolean
					formulasOnly?: boolean
					delimiters?: { open?: string; close?: string }
				}>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				if (!body || !body.data || Array.isArray(body.data) || typeof body.data !== 'object') {
					return jsonFailure('Missing or invalid template data', 400)
				}
				if (body.valuesOnly === true && body.formulasOnly === true) {
					return jsonFailure('Use either valuesOnly or formulasOnly, not both', 400)
				}
				try {
					const wb = await AscendWorkbook.open(file)
					return jsonSuccess(
						wb.templateMerge(body.data, {
							...(body.sheet ? { sheets: [body.sheet] } : {}),
							...(body.valuesOnly === true ? { includeFormulas: false } : {}),
							...(body.formulasOnly === true ? { includeValues: false } : {}),
							...(body.delimiters &&
							(typeof body.delimiters.open === 'string' ||
								typeof body.delimiters.close === 'string')
								? {
										delimiters: {
											...(typeof body.delimiters.open === 'string'
												? { open: body.delimiters.open }
												: {}),
											...(typeof body.delimiters.close === 'string'
												? { close: body.delimiters.close }
												: {}),
										},
									}
								: {}),
						}),
					)
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/pivots') {
				const body = await parseJson<{
					file?: string
					pivotTable?: string
					partPath?: string
					mode?: PivotOutputMaterializeMode
				}>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				const options = parsePivotOutputMaterializeOptions(
					body && typeof body === 'object' ? (body as Record<string, unknown>) : null,
				)
				if (!options) {
					return jsonFailure('Invalid mode. Use one of: missing, mismatches, all', 400)
				}
				try {
					const wb = await WorkbookDocument.open(file, {
						mode: 'full',
						pivotCacheRecordMaterializeLimit: 'all',
					})
					return jsonSuccess({
						pivotTables: wb.pivotTables(),
						pivotCaches: wb.pivotCaches(),
						pivotOutputAudits: wb.pivotOutputAudits(),
						pivotRefreshPlans: wb.pivotRefreshPlans(),
						pivotOutputMaterializePlan: wb.pivotOutputMaterializeOps(options),
					})
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/read') {
				const body = await parseJson<{
					file?: string
					range?: string
					sheet?: string
					format?: string
					headers?: string[]
					display?: boolean
					maxRows?: number
					preview?: boolean
					changedSince?: string
				}>(req)
				const file = body ? requireString(body, 'file') : null
				const range = body ? requireString(body, 'range') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				if (!range) return jsonFailure('Missing or invalid range', 400)
				try {
					const sheetName = body ? requireString(body, 'sheet') : null
					const format = (body ? requireString(body, 'format') : null) ?? 'cells'
					const headers = body ? requireArray(body, 'headers') : null
					const changedSince = body ? requireString(body, 'changedSince') : null
					const rowOffset = body ? requireOptionalNumber(body, 'rowOffset') : undefined
					const explicitRowLimit = body ? requireOptionalNumber(body, 'rowLimit') : undefined
					const preview =
						body !== null &&
						typeof body === 'object' &&
						(body as Record<string, unknown>).preview === true
					const rowLimit = firstWindowRowLimit(explicitRowLimit, preview)
					const maxRows = firstWindowMaxRows(
						body ? requireOptionalNumber(body, 'maxRows') : undefined,
						rowOffset,
						explicitRowLimit,
						preview,
					)
					const display =
						body !== null &&
						typeof body === 'object' &&
						(body as Record<string, unknown>).display === true
					const wb = await WorkbookDocument.open(file, {
						mode: 'values',
						...(sheetName ? { sheets: [sheetName] } : {}),
						...(maxRows !== undefined ? { maxRows } : {}),
					})
					const sheet = sheetName ? wb.sheet(sheetName) : wb.sheet(wb.sheets[0] ?? '')
					if (!sheet) {
						const missing = sheetName ?? wb.sheets[0] ?? ''
						return sheetNotFoundResponse(missing || '(default)', wb)
					}
					if (format === 'rows') {
						const info = sheet.readRows(range, {
							...(rowOffset !== undefined ? { rowOffset } : {}),
							...(rowLimit !== undefined ? { rowLimit } : {}),
						})
						return jsonSuccess(withPartialLoadInfo(display ? displayRows(info) : info, wb))
					}
					if (format === 'objects') {
						const info = sheet.readObjects(range, {
							...(rowOffset !== undefined ? { rowOffset } : {}),
							...(rowLimit !== undefined ? { rowLimit } : {}),
							headers: headers?.every((entry) => typeof entry === 'string')
								? (headers as string[])
								: 'first-row',
						})
						return jsonSuccess(withPartialLoadInfo(display ? displayObjects(info) : info, wb))
					}
					if (format === 'compact') {
						const info = buildCompactReadResult(
							sheet.readWindowCompact(range, {
								...(rowOffset !== undefined ? { rowOffset } : {}),
								...(rowLimit !== undefined ? { rowLimit } : {}),
								includeRefs: false,
								omitEmpty: true,
								flatValues: true,
								changedSince: changedSince ?? '',
							}),
						)
						return jsonSuccess(withPartialLoadInfo(info, wb))
					}
					if (format !== 'cells') return jsonFailure('Invalid read format', 400)
					const info = sheet.readWindow(range, {
						...(rowOffset !== undefined ? { rowOffset } : {}),
						...(rowLimit !== undefined ? { rowLimit } : {}),
					})
					return jsonSuccess(withPartialLoadInfo(display ? displayCells(info) : info, wb))
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/agent-view') {
				const body = await parseJson<{
					file?: string
					range?: string
					sheet?: string
					rowChunkSize?: number
					sampleRowLimit?: number
					sampleValueLimit?: number
				}>(req)
				const file = body ? requireString(body, 'file') : null
				const range = body ? requireString(body, 'range') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				if (!range) return jsonFailure('Missing or invalid range', 400)
				try {
					const sheetName = body ? requireString(body, 'sheet') : null
					const rowChunkSize = body ? requireOptionalNumber(body, 'rowChunkSize') : undefined
					const sampleRowLimit = body ? requireOptionalNumber(body, 'sampleRowLimit') : undefined
					const sampleValueLimit = body
						? requireOptionalNumber(body, 'sampleValueLimit')
						: undefined
					const wb = await WorkbookDocument.open(
						file,
						sheetName ? { mode: 'formula', sheets: [sheetName] } : { mode: 'formula' },
					)
					const targetSheet = sheetName ?? wb.sheets[0]
					if (!targetSheet) return sheetNotFoundResponse(sheetName ?? '', wb)
					const info = wb.agentView(targetSheet, range, {
						...(rowChunkSize !== undefined ? { rowChunkSize } : {}),
						...(sampleRowLimit !== undefined ? { sampleRowLimit } : {}),
						...(sampleValueLimit !== undefined ? { sampleValueLimit } : {}),
					})
					if (!info) return sheetNotFoundResponse(targetSheet, wb)
					return jsonSuccess(info)
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/plan') {
				const body = await parseJson<Record<string, unknown>>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					const inputShape = resolveOperationInputShape(body)
					if (!inputShape.ok) return jsonFailureError(inputShape.error, 400)
					let input: OperationInput
					let pathMutations: PathMutationResult | undefined
					let result: Awaited<ReturnType<typeof createAgentPlan>> | null
					let preparedPlan: PreparedPlanMetadata | undefined
					if ('mutations' in inputShape) {
						const opened = await AscendWorkbook.openSourceBytes(file)
						const inputSha256 = sha256Bytes(opened.sourceBytes)
						const wb = opened.workbook
						input = compilePathMutationInput(wb, inputShape.mutations)
						if (input.ok) pathMutations = input.pathMutations
						result = input.ok
							? await createAgentPlanFromWorkbook(file, inputSha256, wb, input.ops)
							: null
						if (input.ok && result && body?.prepare === true) {
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
													suggestedFix:
														'Re-run ascend plan and commit with the new input workbook.',
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
						if (body?.prepare === true) {
							const prepared = await createPreparedAgentPlan(file, inputShape.ops)
							result = prepared.plan
							preparedPlan = preparedPlans.add(preparedPlanHandle(prepared))
						} else {
							result = await createAgentPlan(file, inputShape.ops)
						}
					}
					if (!input.ok) return jsonFailureError(input.error, 400)
					if (!result) {
						return jsonFailureError(ascendError('VALIDATION_ERROR', 'Plan failed'), 400)
					}
					if (result.preview.errors.length > 0) {
						const first = result.preview.errors[0]
						return jsonFailureError(
							first
								? { ...first, details: { ...(first.details ?? {}), plan: result } }
								: ascendError('VALIDATION_ERROR', 'Plan failed', { details: { plan: result } }),
							400,
						)
					}
					const maxChangedCells = body ? requireOptionalNumber(body, 'maxChangedCells') : undefined
					const payload =
						body?.compact === true
							? compactAgentPlanResult(result, {
									...(maxChangedCells !== undefined ? { maxChangedCells } : {}),
								})
							: result
					return jsonSuccess(
						withPreparedPlanHandle(withPathMutationResult(payload, pathMutations), preparedPlan),
					)
				} catch (e) {
					return handleError(e, file ?? undefined)
				}
			}

			if (method === 'POST' && path === '/commit') {
				const body = await parseJson<Record<string, unknown>>(req)
				const planHandle = body ? requireString(body, 'planHandle') : null
				const file = body ? requireString(body, 'file') : null
				if (!file && !planHandle) return jsonFailure('Missing or invalid file', 400)
				try {
					const output = body ? requireString(body, 'output') : null
					const backup = body ? requireString(body, 'backup') : null
					const expectSha256 = body ? requireString(body, 'expectSha256') : null
					const allowLoss = body
						? parseAllowLoss((body as Record<string, unknown>).allowLoss)
						: undefined
					const approvals = body
						? parseApprovals((body as Record<string, unknown>).approvals)
						: undefined
					const inPlace =
						body !== null &&
						typeof body === 'object' &&
						(body as Record<string, unknown>).inPlace === true
					const options: AgentCommitOptions = {
						...(output ? { output } : {}),
						...(inPlace ? { inPlace: true } : {}),
						...(backup ? { backup } : {}),
						...(expectSha256 ? { expectSha256 } : {}),
						...(allowLoss ? { allowLoss } : {}),
						...(approvals ? { approvals } : {}),
					}
					if (planHandle) {
						const prepared = preparedPlans.take(planHandle)
						if (!prepared) {
							return jsonFailureError(
								ascendError('VALIDATION_ERROR', 'Prepared plan handle was not found', {
									suggestedFix: 'Re-run ascend plan with prepare=true before committing.',
								}),
								400,
							)
						}
						const result = await prepared.commit(options)
						return jsonSuccess(withPathMutationResult(result, prepared.pathMutations))
					}
					const inputShape = resolveOperationInputShape(body)
					if (!inputShape.ok) return jsonFailureError(inputShape.error, 400)
					let input: OperationInput
					let pathMutations: PathMutationResult | undefined
					let result: Awaited<ReturnType<typeof commitAgentPlan>>
					if ('mutations' in inputShape) {
						if (!file) return jsonFailure('Missing or invalid file', 400)
						const opened = await AscendWorkbook.openSourceBytes(file)
						const inputSha256 = sha256Bytes(opened.sourceBytes)
						const wb = opened.workbook
						input = compilePathMutationInput(wb, inputShape.mutations)
						if (!input.ok) return jsonFailureError(input.error, 400)
						pathMutations = input.pathMutations
						result = await commitAgentPlanFromWorkbook(file, inputSha256, wb, input.ops, options, {
							sourceBytes: opened.sourceBytes,
						})
					} else {
						input = inputShape
						if (!file) return jsonFailure('Missing or invalid file', 400)
						result = await commitAgentPlan(file, input.ops, options)
					}
					return jsonSuccess(withPathMutationResult(result, pathMutations))
				} catch (e) {
					return handleError(e, file ?? undefined)
				}
			}

			if (method === 'POST' && path === '/repair-plan') {
				const body = await parseJson<{ file?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					return jsonSuccess(await createRepairPlan(file))
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/write') {
				const body = await parseJson<Record<string, unknown>>(req)
				const file = body ? requireString(body, 'file') : null
				const journal = body?.journal === true
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					const wb = await AscendWorkbook.open(file)
					const input = resolveOperationInputForWorkbook(wb, body)
					if (!input.ok) return jsonFailureError(input.error, 400)
					const result = wb.apply(input.ops, { journal })
					if (result.errors.length > 0) {
						const first = result.errors[0]
						return jsonFailureError(
							first ?? ascendError('VALIDATION_ERROR', 'Failed to apply operations'),
							400,
						)
					}
					if (result.recalcRequired) {
						const recalc = wb.recalc()
						if (recalc.errors.length > 0) {
							const first = recalc.errors[0]
							return jsonFailureError(
								first
									? ascendError('FORMULA_EVAL_ERROR', `${first.ref}: ${first.error.message}`, {
											refs: [first.ref],
											details: { evalError: first.error },
										})
									: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed'),
								400,
							)
						}
					}
					await wb.save(file)
					return jsonSuccess(withPathMutationResult(result, input.pathMutations))
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/preview') {
				const body = await parseJson<Record<string, unknown>>(req)
				const file = body ? requireString(body, 'file') : null
				const journal = body?.journal === true
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					const wb = await AscendWorkbook.open(file)
					const input = resolveOperationInputForWorkbook(wb, body)
					if (!input.ok) return jsonFailureError(input.error, 400)
					const result = wb.preview(input.ops, { journal })
					if (result.errors.length > 0) {
						const first = result.errors[0]
						return jsonFailureError(
							first
								? {
										...first,
										details: { ...(first.details ?? {}), preview: result },
									}
								: ascendError('VALIDATION_ERROR', 'Preview failed', {
										details: { preview: result },
									}),
							400,
						)
					}
					return jsonSuccess(withPathMutationResult(result, input.pathMutations))
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/calc') {
				const body = await parseJson<{ file?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					const wb = await AscendWorkbook.open(file)
					const result = wb.recalc()
					if (result.errors.length > 0) {
						const first = result.errors[0]
						return jsonFailureError(
							first
								? ascendError('FORMULA_EVAL_ERROR', `${first.ref}: ${first.error.message}`, {
										refs: [first.ref],
										details: { evalError: first.error },
									})
								: ascendError('FORMULA_EVAL_ERROR', 'Recalculation failed'),
							400,
						)
					}
					await wb.save(file)
					return jsonSuccess(result)
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/check') {
				const body = await parseJson<{ file?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					const wb = await WorkbookDocument.open(file)
					return jsonSuccess(wb.check())
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/lint') {
				const body = await parseJson<{ file?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				try {
					const wb = await WorkbookDocument.open(file, { mode: 'formula' })
					return jsonSuccess(wb.lint())
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/trace') {
				const body = await parseJson<{ file?: string; cell?: string; maxRows?: number }>(req)
				const file = body ? requireString(body, 'file') : null
				const cell = body ? requireString(body, 'cell') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				if (!cell) return jsonFailure('Missing or invalid cell', 400)
				try {
					const maxRows = body ? requireOptionalNumber(body, 'maxRows') : undefined
					const wb = await WorkbookDocument.open(file, {
						mode: 'formula',
						...(maxRows !== undefined ? { maxRows } : {}),
					})
					const traceIssue = wb.traceIssue(cell)
					if (traceIssue) {
						return jsonFailureError(
							ascendError('VALIDATION_ERROR', traceIssue.message, {
								details: {
									...(traceIssue.rule ? { rule: traceIssue.rule } : {}),
									...(traceIssue.ref ? { ref: traceIssue.ref } : {}),
									load: wb.inspect().load,
								},
								...(traceIssue.suggestedFix ? { suggestedFix: traceIssue.suggestedFix } : {}),
							}),
							400,
						)
					}
					const result = wb.trace(cell)
					if (!result) return jsonFailure('Cell not found', 400)
					return jsonSuccess(result)
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/diff') {
				const body = await parseJson<{ fileA?: string; fileB?: string }>(req)
				const fileA = body ? requireString(body, 'fileA') : null
				const fileB = body ? requireString(body, 'fileB') : null
				if (!fileA) return jsonFailure('Missing or invalid fileA', 400)
				if (!fileB) return jsonFailure('Missing or invalid fileB', 400)
				try {
					const wbA = await AscendWorkbook.open(fileA)
					const wbB = await AscendWorkbook.open(fileB)
					const diff = wbA.diff(wbB)
					return jsonSuccess(diff)
				} catch (e) {
					return handleError(e)
				}
			}

			if (method === 'POST' && path === '/export') {
				const body = await parseJson<{ file?: string; format?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				const format = body ? requireString(body, 'format') : null
				if (!file) return jsonFailure('Missing or invalid file', 400)
				if (!format) return jsonFailure('Missing or invalid format', 400)
				try {
					const wb = await AscendWorkbook.open(file)
					const fmt = normalizeExportFormat(format)
					if (!fmt) return jsonFailure(`Unsupported format: ${format}`, 400)
					if (fmt === 'xlsx' || fmt === 'xlsm') {
						const bytes = wb.toBytes()
						return binaryResponse(
							bytes,
							fmt === 'xlsm'
								? 'application/vnd.ms-excel.sheet.macroEnabled.12+xml'
								: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
						)
					}
					if (fmt === 'csv' || fmt === 'tsv') {
						const csv = wb.toCsv(fmt === 'tsv' ? { dialect: { delimiter: '\t' } } : undefined)
						const bytes = new TextEncoder().encode(csv)
						return binaryResponse(bytes, fmt === 'tsv' ? 'text/tab-separated-values' : 'text/csv')
					}
					if (fmt === 'json') {
						return jsonSuccess(wb.toJSON())
					}
					return jsonFailure(`Unsupported format: ${format}`, 400)
				} catch (e) {
					return handleError(e, file)
				}
			}

			return withCors(jsonFailure('Not Found', 404))
		} catch (e) {
			const msg = e instanceof Error ? e.message : String(e)
			return withCors(jsonFailure(msg, 500))
		}
	}
}

export function createServer(opts?: { port?: number } & ApiFetchOptions) {
	const fetch = createApiFetch(opts)
	const requestedPort = opts?.port ?? (Number(process.env.PORT) || 3000)
	if (requestedPort !== 0) return Bun.serve({ port: requestedPort, fetch })

	let lastError: unknown
	try {
		return Bun.serve({ port: 0, fetch })
	} catch (error) {
		lastError = error
	}

	for (let attempt = 0; attempt < 25; attempt++) {
		const port = 20_000 + Math.floor(Math.random() * 40_000)
		try {
			return Bun.serve({ port, fetch })
		} catch (error) {
			lastError = error
		}
	}
	throw lastError
}

function displayCells<T extends { cells: readonly { ref: string; value: CellValue }[] }>(
	info: T,
): Omit<T, 'cells'> & { cells: Array<{ ref: string; value: string }> } {
	return {
		...info,
		cells: info.cells.map((cell) => ({ ref: cell.ref, value: formatDisplayCellValue(cell.value) })),
	}
}

function displayRows<T extends { rows: readonly (readonly CellValue[])[] }>(
	info: T,
): Omit<T, 'rows'> & { rows: string[][] } {
	return {
		...info,
		rows: info.rows.map((row) => row.map((cell) => formatDisplayCellValue(cell))),
	}
}

function displayObjects<
	T extends {
		headers: readonly string[]
		rows: readonly Readonly<Record<string, CellValue>>[]
	},
>(info: T): Omit<T, 'rows'> & { rows: Array<Record<string, string>> } {
	return {
		...info,
		rows: info.rows.map((row) =>
			Object.fromEntries(
				Object.entries(row).map(([key, value]) => [key, formatDisplayCellValue(value)]),
			),
		),
	}
}

function buildCompactReadResult(info: CompactRangeWindowInfo) {
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
		format: 'compact' as const,
		cells: info.cells.map((cell) => [
			cell.row - info.ref.start.row,
			cell.col - info.ref.start.col,
			cell.value as unknown,
			...(cell.formula ? [cell.formula] : []),
		]),
	}
}
