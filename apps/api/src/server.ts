import type { Operation } from '@ascend/schema'
import { AscendWorkbook, WorkbookSession } from '@ascend/sdk'
import { binaryResponse, jsonFailure, jsonSuccess } from './response.ts'

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

export function createServer(opts?: { port?: number }) {
	return Bun.serve({
		port: opts?.port ?? (Number(process.env.PORT) || 3000),
		async fetch(req) {
			const url = new URL(req.url)
			const method = req.method
			const path = url.pathname

			try {
				if (method === 'GET' && path === '/health') {
					return jsonSuccess({ status: 'ok' })
				}

				if (method === 'POST' && path === '/inspect') {
					const body = await parseJson<{ file?: string; sheet?: string }>(req)
					const file = body ? requireString(body, 'file') : null
					const sheetName = body ? requireString(body, 'sheet') : null
					if (!file) return jsonFailure('Missing or invalid file', 400)
					try {
						const wb = await WorkbookSession.open(
							file,
							sheetName ? { mode: 'values', sheets: [sheetName] } : { mode: 'metadata-only' },
						)
						if (!sheetName) return jsonSuccess(wb.inspect())
						const sheet = wb.inspectSheet(sheetName)
						if (!sheet) return jsonFailure('Sheet not found', 400)
						return jsonSuccess(sheet)
					} catch (e) {
						const msg = e instanceof Error ? e.message : String(e)
						if (msg.includes('ENOENT') || msg.includes('not found'))
							return jsonFailure(`File not found: ${file}`, 400)
						return jsonFailure(msg, 500)
					}
				}

				if (method === 'POST' && path === '/read') {
					const body = await parseJson<{ file?: string; range?: string; sheet?: string }>(req)
					const file = body ? requireString(body, 'file') : null
					const range = body ? requireString(body, 'range') : null
					if (!file) return jsonFailure('Missing or invalid file', 400)
					if (!range) return jsonFailure('Missing or invalid range', 400)
					try {
						const sheetName = body ? requireString(body, 'sheet') : null
						const wb = await WorkbookSession.open(
							file,
							sheetName ? { mode: 'values', sheets: [sheetName] } : { mode: 'values' },
						)
						const sheet = sheetName ? wb.sheet(sheetName) : wb.sheet(wb.sheets[0] ?? '')
						if (!sheet) return jsonFailure('Sheet not found', 400)
						const rangeInfo = sheet.range(range)
						return jsonSuccess(rangeInfo)
					} catch (e) {
						const msg = e instanceof Error ? e.message : String(e)
						if (msg.includes('ENOENT') || msg.includes('not found'))
							return jsonFailure(`File not found: ${file}`, 400)
						return jsonFailure(msg, 500)
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
							return jsonSuccess(result, 400)
						}
						if (result.recalcRequired) wb.recalc()
						await wb.save(file)
						return jsonSuccess(result)
					} catch (e) {
						const msg = e instanceof Error ? e.message : String(e)
						if (msg.includes('ENOENT') || msg.includes('not found'))
							return jsonFailure(`File not found: ${file}`, 400)
						return jsonFailure(msg, 500)
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
						return jsonSuccess(wb.preview(ops))
					} catch (e) {
						const msg = e instanceof Error ? e.message : String(e)
						if (msg.includes('ENOENT') || msg.includes('not found'))
							return jsonFailure(`File not found: ${file}`, 400)
						return jsonFailure(msg, 500)
					}
				}

				if (method === 'POST' && path === '/calc') {
					const body = await parseJson<{ file?: string }>(req)
					const file = body ? requireString(body, 'file') : null
					if (!file) return jsonFailure('Missing or invalid file', 400)
					try {
						const wb = await AscendWorkbook.open(file)
						const result = wb.recalc()
						await wb.save(file)
						return jsonSuccess(result)
					} catch (e) {
						const msg = e instanceof Error ? e.message : String(e)
						if (msg.includes('ENOENT') || msg.includes('not found'))
							return jsonFailure(`File not found: ${file}`, 400)
						return jsonFailure(msg, 500)
					}
				}

				if (method === 'POST' && path === '/check') {
					const body = await parseJson<{ file?: string }>(req)
					const file = body ? requireString(body, 'file') : null
					if (!file) return jsonFailure('Missing or invalid file', 400)
					try {
						const wb = await WorkbookSession.open(file, { mode: 'formula' })
						return jsonSuccess(wb.check())
					} catch (e) {
						const msg = e instanceof Error ? e.message : String(e)
						if (msg.includes('ENOENT') || msg.includes('not found'))
							return jsonFailure(`File not found: ${file}`, 400)
						return jsonFailure(msg, 500)
					}
				}

				if (method === 'POST' && path === '/lint') {
					const body = await parseJson<{ file?: string }>(req)
					const file = body ? requireString(body, 'file') : null
					if (!file) return jsonFailure('Missing or invalid file', 400)
					try {
						const wb = await WorkbookSession.open(file, { mode: 'formula' })
						return jsonSuccess(wb.lint())
					} catch (e) {
						const msg = e instanceof Error ? e.message : String(e)
						if (msg.includes('ENOENT') || msg.includes('not found'))
							return jsonFailure(`File not found: ${file}`, 400)
						return jsonFailure(msg, 500)
					}
				}

				if (method === 'POST' && path === '/trace') {
					const body = await parseJson<{ file?: string; cell?: string }>(req)
					const file = body ? requireString(body, 'file') : null
					const cell = body ? requireString(body, 'cell') : null
					if (!file) return jsonFailure('Missing or invalid file', 400)
					if (!cell) return jsonFailure('Missing or invalid cell', 400)
					try {
						const wb = await WorkbookSession.open(file, { mode: 'formula' })
						const result = wb.trace(cell)
						if (!result) return jsonFailure('Cell not found', 400)
						return jsonSuccess(result)
					} catch (e) {
						const msg = e instanceof Error ? e.message : String(e)
						if (msg.includes('ENOENT') || msg.includes('not found'))
							return jsonFailure(`File not found: ${file}`, 400)
						return jsonFailure(msg, 500)
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
						const msg = e instanceof Error ? e.message : String(e)
						if (msg.includes('ENOENT') || msg.includes('not found'))
							return jsonFailure('File not found', 400)
						return jsonFailure(msg, 500)
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
						const fmt = format.toLowerCase()
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
							const opts = fmt === 'tsv' ? ({ dialect: { delimiter: '\t' } } as never) : undefined
							const csv = wb.toCsv(opts)
							const bytes = new TextEncoder().encode(csv)
							return binaryResponse(bytes, fmt === 'tsv' ? 'text/tab-separated-values' : 'text/csv')
						}
						return jsonFailure(`Unsupported format: ${format}`, 400)
					} catch (e) {
						const msg = e instanceof Error ? e.message : String(e)
						if (msg.includes('ENOENT') || msg.includes('not found'))
							return jsonFailure(`File not found: ${file}`, 400)
						return jsonFailure(msg, 500)
					}
				}

				return jsonFailure('Not Found', 404)
			} catch (e) {
				const msg = e instanceof Error ? e.message : String(e)
				return jsonFailure(msg, 500)
			}
		},
	})
}
