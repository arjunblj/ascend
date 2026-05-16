import { readFileSync, statSync } from 'node:fs'
import { type AscendError, AscendException, ascendError, type CellValue } from '@ascend/schema'
import {
	type AgentCommitOptions,
	type AgentCommitResult,
	AscendWorkbook,
	type CapabilityFilters,
	type CapabilityPriority,
	type CapabilityStatus,
	type CompactRangeWindowInfo,
	commitAgentPlan,
	commitAgentPlanFromWorkbook,
	compactAgentCommitResult,
	compactAgentPlanResult,
	compilePathMutationInput,
	createAgentCommitPackageActionProof,
	createAgentPlan,
	createAgentPlanFromWorkbook,
	createPackageActionProof,
	createPreparedAgentPlan,
	createRepairPlan,
	formatDisplayCellValue,
	formulaAssist,
	getOperationsSchema,
	inspectWorkbookOpenPlan,
	listCapabilities,
	listOperations,
	normalizeExportFormat,
	type PathMutationResult,
	type PivotOutputMaterializeMode,
	type PivotOutputMaterializeOptions,
	type PreparedPlanMetadata,
	PreparedPlanStore,
	preparedPathMutationPlanHandle,
	preparedPlanHandle,
	type ResolvedOperationInput,
	resolveOperationInputShape,
	resolveOperationInputForWorkbook as resolveWorkbookOperationInput,
	sha256Bytes,
	summarizeCapabilities,
	WorkbookDocument,
	type WorkbookOpenIntent,
	withPathMutationResult,
	withPreparedPlanHandle,
} from '@ascend/sdk'
import { binaryResponse, jsonFailure, jsonFailureError, jsonSuccess } from './response.ts'

const DEFAULT_API_RAW_PART_MAX_BYTES = 64 * 1024
const MAX_API_RAW_PART_MAX_BYTES = 1024 * 1024
const DEFAULT_AGENT_PREVIEW_ROWS = 500
const DEFAULT_PATH_MUTATION_PLAN_CACHE_MAX_BYTES = 64 * 1024 * 1024

type PackageActionEvidence = Pick<
	Awaited<ReturnType<typeof createAgentPlan>>,
	'preservation' | 'writePolicy' | 'packageGraphAudit'
>

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

function withPackageActions<T, R extends PackageActionEvidence>(
	payload: T,
	result: R,
	includePackageActions: boolean,
): T | (T & { readonly packageActions: ReturnType<typeof createPackageActionProof> }) {
	if (!includePackageActions) return payload
	return {
		...payload,
		packageActions: isCommitPackageActionEvidence(result)
			? createAgentCommitPackageActionProof(result)
			: createPackageActionProof(result.preservation, {
					writePolicy: result.writePolicy,
					packageGraphAudit: result.packageGraphAudit,
				}),
	}
}

function withCommitProofBundle<T>(
	payload: T,
	result: AgentCommitResult,
	includeProofBundle: boolean,
	inputGuard: CommitProofInputGuard | undefined,
): T | (T & { readonly proofBundle: ReturnType<typeof createCommitProofBundle> }) {
	if (!includeProofBundle) return payload
	return {
		...payload,
		proofBundle: createCommitProofBundle(result, inputGuard),
	}
}

type CommitProofInputGuard =
	| { readonly kind: 'expect-sha256'; readonly expectedSha256: string }
	| { readonly kind: 'prepared-plan-handle' }

function createCommitProofBundle(
	result: AgentCommitResult,
	inputGuard: CommitProofInputGuard | undefined,
) {
	const whatChanged = result.apply.affectedCells.map((ref) => ({ ref }))
	const whySafe = [
		{
			gate: 'input-guard',
			ok:
				inputGuard?.kind === 'prepared-plan-handle' ||
				(inputGuard?.kind === 'expect-sha256' && inputGuard.expectedSha256 === result.inputSha256),
			evidence: {
				guard: inputGuard?.kind ?? null,
				expectedSha256: inputGuard?.kind === 'expect-sha256' ? inputGuard.expectedSha256 : null,
				inputSha256: result.inputSha256,
			},
		},
		{
			gate: 'approval',
			ok: result.approvals.length === 0,
			evidence: {
				approvalCount: result.approvals.length,
				approvalIds: result.approvals.map((approval) => approval.id),
			},
		},
		{
			gate: 'write-policy',
			ok: result.writePolicy.ok,
			evidence: {
				diagnosticCount: result.writePolicy.diagnostics.length,
				blockerCount: result.writePolicy.diagnostics.filter(
					(diagnostic) => diagnostic.severity === 'blocker',
				).length,
			},
		},
		{
			gate: 'commit',
			ok: result.postWrite.valid && result.postWrite.auditsPassed,
			evidence: {
				outputSha256: result.outputSha256,
				postWriteValid: result.postWrite.valid,
				auditsPassed: result.postWrite.auditsPassed,
			},
		},
		{
			gate: 'reopen-verify',
			ok: result.postWrite.reopened && result.postWrite.check.valid && result.postWrite.lint.clean,
			evidence: {
				reopened: result.postWrite.reopened,
				checkValid: result.postWrite.check.valid,
				checkIssueCount: result.postWrite.check.issues.length,
				lintClean: result.postWrite.lint.clean,
				lintWarningCount: result.postWrite.lint.warnings.length,
			},
		},
		{
			gate: 'package-graph',
			ok:
				result.postWrite.packageGraphAudit.ok &&
				result.postWrite.unresolvedPackageGraphIssueCount === 0,
			evidence: {
				packageGraphAuditOk: result.postWrite.packageGraphAudit.ok,
				expectedPackageGraphIssueCount: result.postWrite.expectedPackageGraphIssueCount,
				unresolvedPackageGraphIssueCount: result.postWrite.unresolvedPackageGraphIssueCount,
			},
		},
	] as const
	return {
		safeToUse: whySafe.every((gate) => gate.ok),
		whatChanged,
		whySafe,
		evidence: {
			inputSha256: result.inputSha256,
			planDigest: result.planDigest,
			outputSha256: result.outputSha256,
			operationCount: result.operationCount,
			affectedCellCount: result.apply.affectedCells.length,
			postWriteValid: result.postWrite.valid,
			auditsPassed: result.postWrite.auditsPassed,
			reopened: result.postWrite.reopened,
			checkValid: result.postWrite.check.valid,
			lintClean: result.postWrite.lint.clean,
			writePolicyOk: result.writePolicy.ok,
			packageGraphAuditOk: result.postWrite.packageGraphAudit.ok,
		},
	}
}

function isCommitPackageActionEvidence(result: PackageActionEvidence): result is AgentCommitResult {
	return 'outputSha256' in result && 'postWrite' in result
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

function openPlanIntent(
	value: unknown,
): { readonly ok: true; readonly intent?: WorkbookOpenIntent } | { readonly ok: false } {
	if (value === undefined) return { ok: true }
	return value === 'risk-inventory' ||
		value === 'read-values' ||
		value === 'formula-analysis' ||
		value === 'edit-plan'
		? { ok: true, intent: value }
		: { ok: false }
}

function unsupportedAgentPlanLoadOptions(body: Record<string, unknown> | null): readonly string[] {
	if (!body) return []
	return ['maxRows', 'mode', 'sheets'].filter((key) => body[key] !== undefined)
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

function optionalPasswordError(
	body: Record<string, unknown> | null,
	context: 'open-plan' | 'plan' | 'commit',
): AscendError | null {
	if (body?.password === undefined || typeof body.password === 'string') return null
	return ascendError('VALIDATION_ERROR', `Invalid ${context} password`, {
		details: { field: 'password', receivedType: typeof body.password },
		suggestedFix: 'Pass password as a string or omit it.',
	})
}

type AgentWorkflowFileContext =
	| 'open-plan'
	| 'inspect'
	| 'active-content'
	| 'trust-report'
	| 'package-graph'
	| 'raw-part'
	| 'visuals'
	| 'pivots'
	| 'dump'
	| 'template-merge'
	| 'read'
	| 'agent-view'
	| 'plan'
	| 'commit'
	| 'repair-plan'
	| 'check'
	| 'lint'
	| 'trace'
	| 'write'
	| 'preview'
	| 'calc'
	| 'export'

function missingAgentWorkflowFileError(context: AgentWorkflowFileContext): AscendError {
	const requirement = (() => {
		switch (context) {
			case 'open-plan':
				return 'Pass file so Ascend can inspect workbook risks before hydration.'
			case 'inspect':
				return 'Pass file so Ascend can inspect workbook structure before planning edits.'
			case 'active-content':
				return 'Pass file so Ascend can inspect active content risks before planning edits.'
			case 'trust-report':
				return 'Pass file so Ascend can build a trust report before planning edits.'
			case 'package-graph':
				return 'Pass file so Ascend can audit workbook package preservation before planning edits.'
			case 'raw-part':
				return 'Pass file so Ascend can inspect the requested raw package part safely.'
			case 'visuals':
				return 'Pass file so Ascend can inspect visual inventory before visual, chart, drawing, or image edits.'
			case 'pivots':
				return 'Pass file so Ascend can inspect PivotTables, cache metadata, and materialization plans before edits.'
			case 'dump':
				return 'Pass file so Ascend can dump replayable operations from a full workbook load.'
			case 'template-merge':
				return 'Pass file so Ascend can compile template placeholders into replayable operations.'
			case 'read':
				return 'Pass file so Ascend can read the requested workbook range before planning edits.'
			case 'agent-view':
				return 'Pass file so Ascend can build a bounded agent view before planning edits.'
			case 'commit':
				return 'Pass either file with ops/mutations for a direct commit, or planHandle from a prepared /plan response.'
			case 'plan':
				return 'Pass file with ops or mutations so Ascend can read the source workbook and build a safe plan.'
			case 'repair-plan':
				return 'Pass file so Ascend can inspect failed workflow state and suggest repair actions.'
			case 'check':
				return 'Pass file so Ascend can run structural verification on the workbook.'
			case 'lint':
				return 'Pass file so Ascend can run formula lint verification on the workbook.'
			case 'trace':
				return 'Pass file so Ascend can trace a cell dependency or issue.'
			case 'write':
				return 'Pass file so Ascend can apply operations to a workbook. Prefer /plan then /commit for agent-safe edits.'
			case 'preview':
				return 'Pass file so Ascend can preview operations before writing. Prefer /plan for full agent-safe edit audits.'
			case 'calc':
				return 'Pass file so Ascend can recalculate and save the workbook.'
			case 'export':
				return 'Pass file so Ascend can export the workbook in the requested format.'
		}
	})()
	return ascendError('VALIDATION_ERROR', `Missing or invalid ${context} workbook reference`, {
		retryable: true,
		retryStrategy: 'modified',
		details: {
			required: context === 'commit' ? ['file or planHandle'] : ['file'],
		},
		suggestedFix: requirement,
	})
}

function missingPackagePartPathError(): AscendError {
	return ascendError('VALIDATION_ERROR', 'Missing or invalid package part path', {
		retryable: true,
		retryStrategy: 'modified',
		details: { required: ['partPath'] },
		suggestedFix: 'Pass partPath from /package-graph using an exact package part path.',
	})
}

function missingReadRangeError(context: 'read' | 'agent-view'): AscendError {
	const suggestedFix =
		context === 'agent-view'
			? 'Pass range such as A1:D20 so Ascend can build a bounded agent view.'
			: 'Pass range such as A1:D20 so Ascend can read the requested workbook window.'
	return ascendError('VALIDATION_ERROR', `Missing or invalid ${context} range`, {
		retryable: true,
		retryStrategy: 'modified',
		details: { required: ['range'] },
		suggestedFix,
	})
}

function missingFormulaAssistFormulaError(): AscendError {
	return ascendError('VALIDATION_ERROR', 'Missing or invalid formula-assist formula', {
		retryable: true,
		retryStrategy: 'modified',
		details: { required: ['formula'] },
		suggestedFix: 'Pass formula text such as =SUM(A1:B2) for formula diagnostics and repair hints.',
	})
}

function missingTraceCellError(): AscendError {
	return ascendError('VALIDATION_ERROR', 'Missing or invalid trace cell', {
		retryable: true,
		retryStrategy: 'modified',
		details: { required: ['cell'] },
		suggestedFix: 'Pass cell such as Sheet1!A1 or A1 so Ascend can trace the requested dependency.',
	})
}

function traceCellNotFoundError(cell: string, wb: WorkbookDocument): AscendError {
	return ascendError('VALIDATION_ERROR', 'Trace cell not found', {
		retryable: true,
		retryStrategy: 'modified',
		details: {
			cell,
			workflow: ['inspect', 'read', 'trace'],
			load: wb.inspect().load,
		},
		suggestedFix:
			'Run /inspect or /read to confirm the target sheet and cell before retrying /trace.',
	})
}

function missingDiffWorkbookError(field: 'fileA' | 'fileB'): AscendError {
	return ascendError('VALIDATION_ERROR', `Missing or invalid diff ${field}`, {
		retryable: true,
		retryStrategy: 'modified',
		details: { required: [field] },
		suggestedFix: `Pass ${field} with an existing workbook path so Ascend can compare the two workbooks.`,
	})
}

function missingExportFormatError(): AscendError {
	return ascendError('VALIDATION_ERROR', 'Missing or invalid export format', {
		retryable: true,
		retryStrategy: 'modified',
		details: { required: ['format'], allowedFormats: ['csv', 'tsv', 'json', 'xlsx', 'xlsm'] },
		suggestedFix: 'Pass format as csv, tsv, json, xlsx, or xlsm.',
	})
}

function unsupportedExportFormatError(format: string): AscendError {
	return ascendError('VALIDATION_ERROR', `Unsupported export format: ${format}`, {
		retryable: true,
		retryStrategy: 'modified',
		details: {
			field: 'format',
			received: format,
			allowedFormats: ['csv', 'tsv', 'json', 'xlsx', 'xlsm'],
		},
		suggestedFix: 'Use format csv, tsv, json, xlsx, or xlsm.',
	})
}

function missingTemplateMergeDataError(): AscendError {
	return ascendError('VALIDATION_ERROR', 'Missing or invalid template data', {
		retryable: true,
		retryStrategy: 'modified',
		details: { required: ['data'] },
		suggestedFix:
			'Pass data as an object mapping template placeholder names to replacement values.',
	})
}

function conflictingReplayBatchFilterError(kind: 'dump' | 'template-merge'): AscendError {
	return ascendError('VALIDATION_ERROR', `Conflicting ${kind} replay filters`, {
		retryable: true,
		retryStrategy: 'modified',
		details: { mutuallyExclusive: ['valuesOnly', 'formulasOnly'] },
		suggestedFix: 'Use either valuesOnly or formulasOnly, not both.',
	})
}

function invalidPivotOutputModeError(): AscendError {
	return ascendError('VALIDATION_ERROR', 'Invalid pivot output materialize mode', {
		retryable: true,
		retryStrategy: 'modified',
		details: { field: 'mode', allowedModes: ['missing', 'mismatches', 'all'] },
		suggestedFix: 'Use mode missing, mismatches, or all.',
	})
}

function invalidReadFormatError(format: string): AscendError {
	return ascendError('VALIDATION_ERROR', `Invalid read format: ${format}`, {
		retryable: true,
		retryStrategy: 'modified',
		details: {
			field: 'format',
			received: format,
			allowedFormats: ['cells', 'rows', 'objects', 'compact'],
		},
		suggestedFix: 'Use read format cells, rows, objects, or compact.',
	})
}

function replayBatchLoadOptionsError(kind: string, options: readonly string[]): AscendError {
	return ascendError(
		'VALIDATION_ERROR',
		`${kind} replay batches require a full workbook load; partial or capped load options are not supported`,
		{
			details: {
				unsupportedLoadOptions: options,
				requiredLoad: { mode: 'full', allSheets: true, maxRows: null },
			},
			suggestedFix:
				'Use read or agent-view for bounded inspection, then call this replay-batch endpoint without load options.',
		},
	)
}

export interface ApiFetchOptions {
	readonly preparedPlanMaxHandles?: number
	readonly preparedPlanTtlMs?: number
	readonly pathMutationPlanCacheMaxBytes?: number
	readonly now?: () => number
}

interface PathMutationPlanOpenCacheEntry {
	readonly file: string
	readonly size: number
	readonly mtimeMs: number
	readonly ctimeMs: number
	readonly inputSha256: string
	readonly workbook: AscendWorkbook
	readonly sourceBytes: Uint8Array
}

interface RecentCheckCacheEntry {
	readonly file: string
	readonly size: number
	readonly mtimeMs: number
	readonly ctimeMs: number
	readonly sha256: string
	readonly check: AgentCommitResult['postWrite']['check']
}

function pathMutationPlanCacheLimit(value: number | undefined): number {
	if (value === undefined) return DEFAULT_PATH_MUTATION_PLAN_CACHE_MAX_BYTES
	return Number.isFinite(value) && value > 0 ? Math.floor(value) : 0
}

async function openPathMutationPlanWorkbook(
	file: string,
	cache: PathMutationPlanOpenCacheEntry | undefined,
	cacheMaxBytes: number,
	options: { readonly password?: string } = {},
): Promise<{
	readonly opened: PathMutationPlanOpenCacheEntry
	readonly cache: PathMutationPlanOpenCacheEntry | undefined
}> {
	const stat = statSync(file)
	if (
		options.password === undefined &&
		cache &&
		cache.file === file &&
		cache.size === stat.size &&
		cache.mtimeMs === stat.mtimeMs &&
		cache.ctimeMs === stat.ctimeMs
	) {
		return { opened: cache, cache }
	}
	const source = await AscendWorkbook.openSourceBytes(file, options)
	const opened: PathMutationPlanOpenCacheEntry = {
		file,
		size: stat.size,
		mtimeMs: stat.mtimeMs,
		ctimeMs: stat.ctimeMs,
		inputSha256: sha256Bytes(source.sourceBytes),
		workbook: source.workbook,
		sourceBytes: source.sourceBytes,
	}
	return {
		opened,
		cache:
			options.password === undefined && source.sourceBytes.byteLength <= cacheMaxBytes
				? opened
				: undefined,
	}
}

async function readRecentCheckCache(
	file: string,
	cache: RecentCheckCacheEntry | undefined,
): Promise<AgentCommitResult['postWrite']['check'] | undefined> {
	if (!cache || cache.file !== file) return undefined
	const stat = statSync(file)
	if (
		cache.size !== stat.size ||
		cache.mtimeMs !== stat.mtimeMs ||
		cache.ctimeMs !== stat.ctimeMs
	) {
		return undefined
	}
	const bytes = typeof Bun !== 'undefined' ? await Bun.file(file).bytes() : readFileSync(file)
	const currentSha256 = sha256Bytes(new Uint8Array(bytes))
	return currentSha256 === cache.sha256 ? cache.check : undefined
}

function recentCheckCacheEntry(
	file: string,
	result: AgentCommitResult,
): RecentCheckCacheEntry | undefined {
	try {
		const stat = statSync(file)
		return {
			file,
			size: stat.size,
			mtimeMs: stat.mtimeMs,
			ctimeMs: stat.ctimeMs,
			sha256: result.outputSha256,
			check: result.postWrite.check,
		}
	} catch {
		return undefined
	}
}

function resolveOperationInputForWorkbook(
	wb: AscendWorkbook,
	body: Record<string, unknown> | null,
): ResolvedOperationInput {
	return resolveWorkbookOperationInput(wb, operationInputSourceFromBody(body))
}

function operationInputSourceFromBody(body: Record<string, unknown> | null) {
	return {
		hasOpsKey: body !== null && Object.hasOwn(body, 'ops'),
		ops: body ? requireArray(body, 'ops') : null,
		hasMutationsKey: body !== null && Object.hasOwn(body, 'mutations'),
		mutations: body ? requireArray(body, 'mutations') : null,
		operationSchemaSuggestedFix: 'Call /operations for canonical operation schemas and examples.',
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
		ae.code === 'TABLE_NOT_FOUND' ||
		ae.code === 'FILE_NOT_FOUND'
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
	if (ae.code === 'EXPORT_ERROR') return 409
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

function fileNotFoundAscendError(fileContext?: string): AscendError {
	return ascendError(
		'FILE_NOT_FOUND',
		fileContext ? `File not found: ${fileContext}` : 'File not found',
		{
			retryable: true,
			retryStrategy: 'modified',
			...(fileContext ? { details: { file: fileContext } } : {}),
			suggestedFix:
				'Pass an existing workbook path that the API process can read, then retry the request.',
		},
	)
}

function handleError(e: unknown, fileContext?: string): Response {
	if (e instanceof AscendException) {
		const status = statusForError(e.ascendError)
		return jsonFailureError(e.ascendError, status)
	}
	if (isFileNotFoundError(e)) return jsonFailureError(fileNotFoundAscendError(fileContext), 404)
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
		'preservedCustomUi',
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
					/(vba|macro|activex|ctrlprops|customui|_xmlsignatures|signature)/i.test(location),
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

const SUPPORTED_API_ROUTES = [
	{ method: 'GET', path: '/health' },
	{ method: 'GET', path: '/operations' },
	{ method: 'GET', path: '/capabilities' },
	{ method: 'GET', path: '/agent-workflow' },
	{ method: 'POST', path: '/open-plan' },
	{ method: 'POST', path: '/inspect' },
	{ method: 'POST', path: '/active-content' },
	{ method: 'POST', path: '/trust-report' },
	{ method: 'POST', path: '/package-graph' },
	{ method: 'POST', path: '/raw-part' },
	{ method: 'POST', path: '/visuals' },
	{ method: 'POST', path: '/dump' },
	{ method: 'POST', path: '/template-merge' },
	{ method: 'POST', path: '/pivots' },
	{ method: 'POST', path: '/read' },
	{ method: 'POST', path: '/agent-view' },
	{ method: 'POST', path: '/formula-assist' },
	{ method: 'POST', path: '/plan' },
	{ method: 'POST', path: '/commit' },
	{ method: 'POST', path: '/repair-plan' },
	{ method: 'POST', path: '/write' },
	{ method: 'POST', path: '/preview' },
	{ method: 'POST', path: '/calc' },
	{ method: 'POST', path: '/check' },
	{ method: 'POST', path: '/lint' },
	{ method: 'POST', path: '/trace' },
	{ method: 'POST', path: '/diff' },
	{ method: 'POST', path: '/export' },
] as const

const API_AGENT_WORKFLOW = {
	workflow: [
		{
			step: 'open-plan',
			endpoint: 'POST /open-plan',
			proof: ['recommendedLoadOptions', 'reviewBeforeHydration', 'riskFeatures'],
		},
		{
			step: 'trust-preflight',
			endpoint: 'POST /trust-report',
			proof: ['trust', 'posture', 'findingCount', 'nextActions'],
		},
		{
			step: 'inspect',
			endpoint: 'POST /inspect',
			proof: ['sheetCount', 'compatibility', 'load'],
		},
		{
			step: 'read',
			endpoint: 'POST /read',
			proof: ['range', 'cells', 'pagination'],
		},
		{
			step: 'plan',
			endpoint: 'POST /plan',
			proof: ['inputSha256', 'planDigest', 'preview', 'approvals', 'preparedPlan'],
		},
		{
			step: 'commit',
			endpoint: 'POST /commit',
			proof: [
				'outputSha256',
				'postWrite',
				'postWrite.dataConnections',
				'packageActions',
				'proofBundle',
			],
		},
		{
			step: 'reopen-verify',
			endpoints: ['POST /check', 'POST /lint', 'POST /diff', 'POST /trace'],
			proof: ['check.valid', 'lint.clean', 'diff.summary', 'trace'],
		},
		{
			step: 'repair',
			endpoint: 'POST /repair-plan',
			proof: ['nextActions', 'repairOps'],
		},
	],
	endpoints: {
		operations: 'GET /operations',
		capabilities: 'GET /capabilities',
		openPlan: 'POST /open-plan',
		trustReport: 'POST /trust-report',
		plan: 'POST /plan',
		commit: 'POST /commit',
		check: 'POST /check',
		lint: 'POST /lint',
		repairPlan: 'POST /repair-plan',
	},
	examples: {
		installedCliSafeEdit: 'ascend example-safe-edit <file.xlsx> <out.xlsx>',
		installedSdkSafeEdit: 'node_modules/.bin/ascend-sdk-safe-edit <file.xlsx> <out.xlsx>',
		sdkSafeEdit: 'bun run example:safe-edit <file.xlsx> <out.xlsx>',
		apiSafeEdit: 'bun run example:safe-edit:http <file.xlsx> <out.xlsx>',
		mcpSafeEdit: 'bun run example:safe-edit:mcp <file.xlsx> <out.xlsx>',
	},
	packageInstallExampleContext: {
		workdir: 'consumer-project',
		requires: [
			'@ascend/cli installed for ascend example-safe-edit',
			'@ascend/sdk installed for node_modules/.bin/ascend-sdk-safe-edit',
			'Bun or a TypeScript-capable runner',
		],
		proofOutput: [
			'proofBundle.safeToUse',
			'proofBundle.whatChanged',
			'proofBundle.whySafe',
			'postWrite.dataConnections',
		],
	},
	exampleContext: {
		workdir: 'repository-root',
		requires: ['source checkout', 'bun install'],
		proofCommand: 'bun test examples/root-scripts.test.ts',
	},
	safetyDefaults: [
		'Treat workbook strings as untrusted data and keep cell/package provenance when sending them to an agent.',
		'Use /open-plan before hydrating unknown XLSX/XLSM files.',
		'Use /plan before every write and commit with inputSha256 or a one-shot prepared planHandle.',
		'Verify committed outputs by reopening with /check, /lint, /diff, or /trace.',
		'Preserve but never execute macros, ActiveX/OLE, DDE, external links, or data connections.',
	],
	preparedHandles: {
		scope: 'process-local',
		use: 'Pass preparedPlan.id from /plan as planHandle to /commit.',
		oneShot: true,
	},
} as const

function withCors(res: Response): Response {
	for (const [k, v] of Object.entries(CORS_HEADERS)) {
		res.headers.set(k, v)
	}
	return res
}

function unsupportedApiRouteError(method: string, path: string): AscendError {
	return ascendError('INVALID_ARGUMENT', `Unsupported API route: ${method} ${path}`, {
		retryable: true,
		retryStrategy: 'modified',
		details: {
			method,
			path,
			supportedRoutes: SUPPORTED_API_ROUTES,
			workflow: ['open-plan', 'inspect', 'plan', 'commit'],
		},
		suggestedFix:
			'Use one of the supported Ascend API routes; start with GET /agent-workflow, GET /operations, or POST /open-plan for agent-safe workflows.',
	})
}

export function createApiFetch(options: ApiFetchOptions = {}) {
	const preparedPlans = new PreparedPlanStore(options)
	const pathMutationPlanCacheMaxBytes = pathMutationPlanCacheLimit(
		options.pathMutationPlanCacheMaxBytes,
	)
	let pathMutationPlanCache: PathMutationPlanOpenCacheEntry | undefined
	let recentCheckCache: RecentCheckCacheEntry | undefined
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

			if (method === 'GET' && path === '/agent-workflow') {
				return withCors(jsonSuccess(API_AGENT_WORKFLOW))
			}

			if (method === 'POST' && path === '/inspect') {
				const body = await parseJson<{ file?: string; sheet?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				const sheetName = body ? requireString(body, 'sheet') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('inspect'), 400)
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

			if (method === 'POST' && path === '/open-plan') {
				const body = await parseJson<Record<string, unknown>>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('open-plan'), 400)
				const intent = openPlanIntent(body?.intent)
				if (!intent.ok) {
					return jsonFailureError(
						ascendError('VALIDATION_ERROR', 'Invalid open-plan intent', {
							details: {
								allowedIntents: ['risk-inventory', 'read-values', 'formula-analysis', 'edit-plan'],
								receivedIntent: body?.intent,
							},
							suggestedFix:
								'Use intent risk-inventory, read-values, formula-analysis, or edit-plan.',
						}),
						400,
					)
				}
				try {
					const passwordError = optionalPasswordError(body, 'open-plan')
					if (passwordError) return jsonFailureError(passwordError, 400)
					const password = body?.password as string | undefined
					return jsonSuccess(
						inspectWorkbookOpenPlan(new Uint8Array(readFileSync(file)), {
							...(intent.intent ? { intent: intent.intent } : {}),
							...(password !== undefined ? { password } : {}),
						}),
					)
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/active-content') {
				const body = await parseJson<{ file?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('active-content'), 400)
				try {
					const wb = await WorkbookDocument.open(file, { mode: 'metadata-only' })
					return jsonSuccess(activeContentPayload(wb))
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/trust-report') {
				const body = await parseJson<{ file?: string; maxFindings?: number }>(req)
				const file = body ? requireString(body, 'file') : null
				const maxFindings = body ? requireOptionalNumber(body, 'maxFindings') : undefined
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('trust-report'), 400)
				try {
					const wb = await WorkbookDocument.open(file, { mode: 'full' })
					return jsonSuccess(
						await wb.trustReport({
							...(maxFindings !== undefined ? { maxFindings } : {}),
						}),
					)
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/package-graph') {
				const body = await parseJson<{ file?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('package-graph'), 400)
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
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('raw-part'), 400)
				if (!partPath) return jsonFailureError(missingPackagePartPathError(), 400)
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
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('visuals'), 400)
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
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('dump'), 400)
				const unsupportedLoadOptions = unsupportedAgentPlanLoadOptions(body)
				if (unsupportedLoadOptions.length > 0) {
					return jsonFailureError(replayBatchLoadOptionsError('Dump', unsupportedLoadOptions), 400)
				}
				if (body?.valuesOnly === true && body?.formulasOnly === true) {
					return jsonFailureError(conflictingReplayBatchFilterError('dump'), 400)
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
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('template-merge'), 400)
				const unsupportedLoadOptions = unsupportedAgentPlanLoadOptions(body)
				if (unsupportedLoadOptions.length > 0) {
					return jsonFailureError(
						replayBatchLoadOptionsError('Template merge', unsupportedLoadOptions),
						400,
					)
				}
				if (!body || !body.data || Array.isArray(body.data) || typeof body.data !== 'object') {
					return jsonFailureError(missingTemplateMergeDataError(), 400)
				}
				if (body.valuesOnly === true && body.formulasOnly === true) {
					return jsonFailureError(conflictingReplayBatchFilterError('template-merge'), 400)
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
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('pivots'), 400)
				const options = parsePivotOutputMaterializeOptions(
					body && typeof body === 'object' ? (body as Record<string, unknown>) : null,
				)
				if (!options) {
					return jsonFailureError(invalidPivotOutputModeError(), 400)
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
					rowOffset?: number
					rowLimit?: number
					preview?: boolean
					changedSince?: string
				}>(req)
				const file = body ? requireString(body, 'file') : null
				const range = body ? requireString(body, 'range') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('read'), 400)
				if (!range) return jsonFailureError(missingReadRangeError('read'), 400)
				try {
					const sheetName = body ? requireString(body, 'sheet') : null
					const format = (body ? requireString(body, 'format') : null) ?? 'cells'
					if (
						format !== 'cells' &&
						format !== 'rows' &&
						format !== 'objects' &&
						format !== 'compact'
					) {
						return jsonFailureError(invalidReadFormatError(format), 400)
					}
					const headers = body ? requireArray(body, 'headers') : null
					const changedSince = body ? requireString(body, 'changedSince') : null
					const rowOffset = body ? requireOptionalNumber(body, 'rowOffset') : undefined
					const explicitRowLimit = body ? requireOptionalNumber(body, 'rowLimit') : undefined
					const preview =
						body !== null &&
						typeof body === 'object' &&
						(body as Record<string, unknown>).preview === true
					const firstWindow = preview || format === 'compact'
					const rowLimit = firstWindowRowLimit(explicitRowLimit, firstWindow)
					const maxRows = firstWindowMaxRows(
						body ? requireOptionalNumber(body, 'maxRows') : undefined,
						rowOffset,
						explicitRowLimit,
						firstWindow,
					)
					const display =
						body !== null &&
						typeof body === 'object' &&
						(body as Record<string, unknown>).display === true
					const wb = await WorkbookDocument.open(file, {
						mode: format === 'cells' ? 'formula' : 'values',
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
					maxRows?: number
					rowChunkSize?: number
					sampleRowLimit?: number
					sampleValueLimit?: number
					maxApproxTokens?: number
				}>(req)
				const file = body ? requireString(body, 'file') : null
				const range = body ? requireString(body, 'range') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('agent-view'), 400)
				if (!range) return jsonFailureError(missingReadRangeError('agent-view'), 400)
				try {
					const sheetName = body ? requireString(body, 'sheet') : null
					const maxRows = body ? requireOptionalNumber(body, 'maxRows') : undefined
					const rowChunkSize = body ? requireOptionalNumber(body, 'rowChunkSize') : undefined
					const sampleRowLimit = body ? requireOptionalNumber(body, 'sampleRowLimit') : undefined
					const sampleValueLimit = body
						? requireOptionalNumber(body, 'sampleValueLimit')
						: undefined
					const maxApproxTokens = body ? requireOptionalNumber(body, 'maxApproxTokens') : undefined
					const wb = await WorkbookDocument.open(
						file,
						sheetName
							? {
									mode: 'formula',
									sheets: [sheetName],
									...(maxRows !== undefined ? { maxRows } : {}),
								}
							: { mode: 'formula', ...(maxRows !== undefined ? { maxRows } : {}) },
					)
					const targetSheet = sheetName ?? wb.sheets[0]
					if (!targetSheet) return sheetNotFoundResponse(sheetName ?? '', wb)
					const info = wb.agentView(targetSheet, range, {
						...(rowChunkSize !== undefined ? { rowChunkSize } : {}),
						...(sampleRowLimit !== undefined ? { sampleRowLimit } : {}),
						...(sampleValueLimit !== undefined ? { sampleValueLimit } : {}),
						...(maxApproxTokens !== undefined ? { maxApproxTokens } : {}),
					})
					if (!info) return sheetNotFoundResponse(targetSheet, wb)
					return jsonSuccess(withPartialLoadInfo(info, wb))
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/formula-assist') {
				const body = await parseJson<Record<string, unknown>>(req)
				const formula = body ? requireString(body, 'formula') : null
				if (!formula) return jsonFailureError(missingFormulaAssistFormulaError(), 400)
				const cursor = body ? requireOptionalNumber(body, 'cursor') : undefined
				const completionLimit = body ? requireOptionalNumber(body, 'completionLimit') : undefined
				const prefix = body ? requireString(body, 'prefix') : null
				const functionName = body ? requireString(body, 'functionName') : null
				const reference = body ? requireString(body, 'reference') : null
				return jsonSuccess(
					formulaAssist(formula, {
						...(cursor !== undefined ? { cursor } : {}),
						...(prefix !== null ? { prefix } : {}),
						...(completionLimit !== undefined ? { completionLimit } : {}),
						...(functionName !== null ? { functionName } : {}),
						...(reference !== null ? { reference } : {}),
						replaceReferenceAtCursor: body?.replaceReferenceAtCursor === true,
						cycleReference: body?.cycleReference === true,
					}),
				)
			}

			if (method === 'POST' && path === '/plan') {
				const body = await parseJson<Record<string, unknown>>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('plan'), 400)
				try {
					const unsupportedLoadOptions = unsupportedAgentPlanLoadOptions(body)
					if (unsupportedLoadOptions.length > 0) {
						return jsonFailureError(agentPlanLoadOptionsError(unsupportedLoadOptions), 400)
					}
					const passwordError = optionalPasswordError(body, 'plan')
					if (passwordError) return jsonFailureError(passwordError, 400)
					const password = body ? requireString(body, 'password') : null
					const inputShape = resolveOperationInputShape(operationInputSourceFromBody(body))
					if (!inputShape.ok) return jsonFailureError(inputShape.error, 400)
					let input: ResolvedOperationInput
					let pathMutations: PathMutationResult | undefined
					let result: Awaited<ReturnType<typeof createAgentPlan>> | null
					let preparedPlan: PreparedPlanMetadata | undefined
					const prepare = body?.prepare !== false
					if ('mutations' in inputShape) {
						const openedResult = await openPathMutationPlanWorkbook(
							file,
							pathMutationPlanCache,
							pathMutationPlanCacheMaxBytes,
							password ? { password } : {},
						)
						pathMutationPlanCache = openedResult.cache
						const opened = openedResult.opened
						const inputSha256 = opened.inputSha256
						const wb = opened.workbook
						input = compilePathMutationInput(wb, inputShape.mutations)
						if (input.ok) pathMutations = input.pathMutations
						result = input.ok
							? await createAgentPlanFromWorkbook(file, inputSha256, wb, input.ops)
							: null
						if (input.ok && result && prepare) {
							preparedPlan = preparedPlans.add(
								preparedPathMutationPlanHandle({
									file,
									inputSha256,
									planDigest: result.planDigest,
									operationCount: result.operationCount,
									workbook: wb,
									ops: input.ops,
									sourceBytes: opened.sourceBytes,
									preparedCheck: result.check,
									...(pathMutations !== undefined ? { pathMutations } : {}),
								}),
							)
							pathMutationPlanCache = undefined
						}
					} else {
						input = inputShape
						if (prepare) {
							const prepared = await createPreparedAgentPlan(file, inputShape.ops, {
								...(password ? { password } : {}),
							})
							result = prepared.plan
							preparedPlan = preparedPlans.add(preparedPlanHandle(prepared))
						} else {
							result = await createAgentPlan(file, inputShape.ops, {
								...(password ? { password } : {}),
							})
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
					const responsePayload = withPackageActions(
						payload,
						result,
						body?.includePackageActions === true,
					)
					return jsonSuccess(
						withPreparedPlanHandle(
							withPathMutationResult(responsePayload, pathMutations),
							preparedPlan,
						),
					)
				} catch (e) {
					return handleError(e, file ?? undefined)
				}
			}

			if (method === 'POST' && path === '/commit') {
				const body = await parseJson<Record<string, unknown>>(req)
				const planHandle = body ? requireString(body, 'planHandle') : null
				const file = body ? requireString(body, 'file') : null
				if (!file && !planHandle) {
					return jsonFailureError(missingAgentWorkflowFileError('commit'), 400)
				}
				try {
					pathMutationPlanCache = undefined
					recentCheckCache = undefined
					const unsupportedLoadOptions = unsupportedAgentPlanLoadOptions(body)
					if (unsupportedLoadOptions.length > 0) {
						return jsonFailureError(agentPlanLoadOptionsError(unsupportedLoadOptions), 400)
					}
					const passwordError = optionalPasswordError(body, 'commit')
					if (passwordError) return jsonFailureError(passwordError, 400)
					const output = body ? requireString(body, 'output') : null
					const backup = body ? requireString(body, 'backup') : null
					const password = body ? requireString(body, 'password') : null
					const expectSha256 = body ? requireString(body, 'expectSha256') : null
					const allowLoss = body
						? parseAllowLoss((body as Record<string, unknown>).allowLoss)
						: undefined
					const approvals = body
						? parseApprovals((body as Record<string, unknown>).approvals)
						: undefined
					const compact = body?.compact === true
					const maxAffectedCells = body
						? requireOptionalNumber(body, 'maxAffectedCells')
						: undefined
					const includeProofBundle = body?.includeProofBundle === true
					const inPlace =
						body !== null &&
						typeof body === 'object' &&
						(body as Record<string, unknown>).inPlace === true
					const options: AgentCommitOptions = {
						...(output ? { output } : {}),
						...(inPlace ? { inPlace: true } : {}),
						...(backup ? { backup } : {}),
						...(password ? { password } : {}),
						...(expectSha256 ? { expectSha256 } : {}),
						...(allowLoss ? { allowLoss } : {}),
						...(approvals ? { approvals } : {}),
					}
					if (planHandle) {
						const prepared = preparedPlans.take(planHandle)
						if (!prepared.ok) return jsonFailureError(prepared.error, 400)
						let result: AgentCommitResult
						try {
							result = await prepared.handle.commit(options)
						} catch (error) {
							prepared.restore()
							throw error
						}
						recentCheckCache = recentCheckCacheEntry(result.output, result)
						const payload = compact
							? compactAgentCommitResult(result, {
									...(maxAffectedCells !== undefined ? { maxAffectedCells } : {}),
								})
							: result
						const responsePayload = withCommitProofBundle(
							withPackageActions(payload, result, body?.includePackageActions === true),
							result,
							includeProofBundle,
							{ kind: 'prepared-plan-handle' },
						)
						return jsonSuccess(
							withPathMutationResult(responsePayload, prepared.handle.pathMutations),
						)
					}
					const inputShape = resolveOperationInputShape(operationInputSourceFromBody(body))
					if (!inputShape.ok) return jsonFailureError(inputShape.error, 400)
					let input: ResolvedOperationInput
					let pathMutations: PathMutationResult | undefined
					let result: Awaited<ReturnType<typeof commitAgentPlan>>
					if ('mutations' in inputShape) {
						if (!file) return jsonFailure('Missing or invalid file', 400)
						const opened = await AscendWorkbook.openSourceBytes(file, {
							...(password ? { password } : {}),
						})
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
					recentCheckCache = recentCheckCacheEntry(result.output, result)
					const payload = compact
						? compactAgentCommitResult(result, {
								...(maxAffectedCells !== undefined ? { maxAffectedCells } : {}),
							})
						: result
					const responsePayload = withCommitProofBundle(
						withPackageActions(payload, result, body?.includePackageActions === true),
						result,
						includeProofBundle,
						expectSha256 ? { kind: 'expect-sha256', expectedSha256: expectSha256 } : undefined,
					)
					return jsonSuccess(withPathMutationResult(responsePayload, pathMutations))
				} catch (e) {
					return handleError(e, file ?? undefined)
				}
			}

			if (method === 'POST' && path === '/repair-plan') {
				const body = await parseJson<{ file?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('repair-plan'), 400)
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
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('write'), 400)
				try {
					pathMutationPlanCache = undefined
					recentCheckCache = undefined
					const wb = await AscendWorkbook.open(file)
					const input = resolveOperationInputForWorkbook(wb, body)
					if (!input.ok) return jsonFailureError(input.error, 400)
					const result = wb.apply(input.ops, { journal })
					if (result.errors.length > 0) {
						const first = result.errors[0]
						return jsonFailureError(
							first
								? {
										...first,
										details: { ...(first.details ?? {}), apply: result },
									}
								: ascendError('VALIDATION_ERROR', 'Failed to apply operations', {
										details: { apply: result },
									}),
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
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('preview'), 400)
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
				const body = await parseJson<{ file?: string; range?: string }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('calc'), 400)
				try {
					const wb = await AscendWorkbook.open(file)
					const result = wb.recalc(body?.range ? { range: body.range } : undefined)
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
				const body = await parseJson<{ file?: string; maxRows?: number }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('check'), 400)
				try {
					const maxRows = body ? requireOptionalNumber(body, 'maxRows') : undefined
					if (maxRows === undefined) {
						const cachedCheck = await readRecentCheckCache(file, recentCheckCache)
						if (cachedCheck) return jsonSuccess(cachedCheck)
					}
					const wb = await WorkbookDocument.open(file, {
						...(maxRows !== undefined ? { maxRows } : {}),
					})
					return jsonSuccess(withPartialLoadInfo(wb.check(), wb))
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/lint') {
				const body = await parseJson<{ file?: string; maxRows?: number }>(req)
				const file = body ? requireString(body, 'file') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('lint'), 400)
				try {
					const maxRows = body ? requireOptionalNumber(body, 'maxRows') : undefined
					const wb = await WorkbookDocument.open(file, {
						mode: 'formula',
						...(maxRows !== undefined ? { maxRows } : {}),
					})
					return jsonSuccess(withPartialLoadInfo(wb.lint(), wb))
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/trace') {
				const body = await parseJson<{ file?: string; cell?: string; maxRows?: number }>(req)
				const file = body ? requireString(body, 'file') : null
				const cell = body ? requireString(body, 'cell') : null
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('trace'), 400)
				if (!cell) return jsonFailureError(missingTraceCellError(), 400)
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
					if (!result) return jsonFailureError(traceCellNotFoundError(cell, wb), 400)
					return jsonSuccess(result)
				} catch (e) {
					return handleError(e, file)
				}
			}

			if (method === 'POST' && path === '/diff') {
				const body = await parseJson<{ fileA?: string; fileB?: string }>(req)
				const fileA = body ? requireString(body, 'fileA') : null
				const fileB = body ? requireString(body, 'fileB') : null
				if (!fileA) return jsonFailureError(missingDiffWorkbookError('fileA'), 400)
				if (!fileB) return jsonFailureError(missingDiffWorkbookError('fileB'), 400)
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
				if (!file) return jsonFailureError(missingAgentWorkflowFileError('export'), 400)
				if (!format) return jsonFailureError(missingExportFormatError(), 400)
				const fmt = normalizeExportFormat(format)
				if (!fmt) return jsonFailureError(unsupportedExportFormatError(format), 400)
				try {
					const wb = await AscendWorkbook.open(file)
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

			return withCors(jsonFailureError(unsupportedApiRouteError(method, path), 404))
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
): Omit<T, 'cells'> & {
	cells: Array<Omit<T['cells'][number], 'value'> & { value: string }>
} {
	const cells = info.cells.map((cell) => ({
		...cell,
		value: formatDisplayCellValue(cell.value),
	})) as Array<Omit<T['cells'][number], 'value'> & { value: string }>
	return {
		...info,
		cells,
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
		...(info.changeInvalidation !== undefined
			? { changeInvalidation: info.changeInvalidation }
			: {}),
		format: 'compact' as const,
		cells: info.cells.map((cell) => [
			cell.row - info.ref.start.row,
			cell.col - info.ref.start.col,
			cell.value as unknown,
			...(cell.formula ? [cell.formula] : []),
		]),
	}
}
