import {
	type AscendError,
	AscendException,
	ascendError,
	type CellValue,
	type Operation,
} from '@ascend/schema'
import {
	AscendWorkbook,
	formatDisplayCellValue,
	normalizeExportFormat,
	WorkbookDocument,
} from '@ascend/sdk'
import { binaryResponse, jsonFailure, jsonFailureError, jsonSuccess } from './response.ts'

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

export function createServer(opts?: { port?: number }) {
	return Bun.serve({
		port: opts?.port ?? (Number(process.env.PORT) || 3000),
		async fetch(req) {
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
						if (!sheet) return jsonFailure('Sheet not found', 400)
						return jsonSuccess(sheet)
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
					}>(req)
					const file = body ? requireString(body, 'file') : null
					const range = body ? requireString(body, 'range') : null
					if (!file) return jsonFailure('Missing or invalid file', 400)
					if (!range) return jsonFailure('Missing or invalid range', 400)
					try {
						const sheetName = body ? requireString(body, 'sheet') : null
						const format = (body ? requireString(body, 'format') : null) ?? 'cells'
						const headers = body ? requireArray(body, 'headers') : null
						const rowOffset = body ? requireOptionalNumber(body, 'rowOffset') : undefined
						const rowLimit = body ? requireOptionalNumber(body, 'rowLimit') : undefined
						const display =
							body !== null &&
							typeof body === 'object' &&
							(body as Record<string, unknown>).display === true
						const wb = await WorkbookDocument.open(
							file,
							sheetName ? { mode: 'values', sheets: [sheetName] } : { mode: 'values' },
						)
						const sheet = sheetName ? wb.sheet(sheetName) : wb.sheet(wb.sheets[0] ?? '')
						if (!sheet) return jsonFailure('Sheet not found', 400)
						if (format === 'rows') {
							const info = sheet.readRows(range, {
								...(rowOffset !== undefined ? { rowOffset } : {}),
								...(rowLimit !== undefined ? { rowLimit } : {}),
							})
							return jsonSuccess(display ? displayRows(info) : info)
						}
						if (format === 'objects') {
							const info = sheet.readObjects(range, {
								...(rowOffset !== undefined ? { rowOffset } : {}),
								...(rowLimit !== undefined ? { rowLimit } : {}),
								headers: headers?.every((entry) => typeof entry === 'string')
									? (headers as string[])
									: 'first-row',
							})
							return jsonSuccess(display ? displayObjects(info) : info)
						}
						if (format !== 'cells') return jsonFailure('Invalid read format', 400)
						const info = sheet.readWindow(range, {
							...(rowOffset !== undefined ? { rowOffset } : {}),
							...(rowLimit !== undefined ? { rowLimit } : {}),
						})
						return jsonSuccess(display ? displayCells(info) : info)
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
						if (!targetSheet) return jsonFailure('Sheet not found', 400)
						const info = wb.agentView(targetSheet, range, {
							...(rowChunkSize !== undefined ? { rowChunkSize } : {}),
							...(sampleRowLimit !== undefined ? { sampleRowLimit } : {}),
							...(sampleValueLimit !== undefined ? { sampleValueLimit } : {}),
						})
						if (!info) return jsonFailure('Sheet not found', 400)
						return jsonSuccess(info)
					} catch (e) {
						return handleError(e, file)
					}
				}

				if (method === 'POST' && path === '/write') {
					const body = await parseJson<{ file?: string; ops?: unknown[] }>(req)
					const file = body ? requireString(body, 'file') : null
					const opsArr = body ? requireArray(body, 'ops') : null
					if (!file) return jsonFailure('Missing or invalid file', 400)
					if (!opsArr || opsArr.length === 0) return jsonFailure('Missing or invalid ops', 400)
					const ops = opsArr as Operation[]
					try {
						const wb = await AscendWorkbook.open(file)
						const result = wb.apply(ops)
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
						return jsonSuccess(result)
					} catch (e) {
						return handleError(e, file)
					}
				}

				if (method === 'POST' && path === '/preview') {
					const body = await parseJson<{ file?: string; ops?: unknown[] }>(req)
					const file = body ? requireString(body, 'file') : null
					const opsArr = body ? requireArray(body, 'ops') : null
					if (!file) return jsonFailure('Missing or invalid file', 400)
					if (!opsArr || opsArr.length === 0) return jsonFailure('Missing or invalid ops', 400)
					const ops = opsArr as Operation[]
					try {
						const wb = await AscendWorkbook.open(file)
						const result = wb.preview(ops)
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
						return jsonSuccess(result)
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
						const wb = await WorkbookDocument.open(file, { mode: 'formula' })
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
					const body = await parseJson<{ file?: string; cell?: string }>(req)
					const file = body ? requireString(body, 'file') : null
					const cell = body ? requireString(body, 'cell') : null
					if (!file) return jsonFailure('Missing or invalid file', 400)
					if (!cell) return jsonFailure('Missing or invalid cell', 400)
					try {
						const wb = await WorkbookDocument.open(file, { mode: 'formula' })
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
		},
	})
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
