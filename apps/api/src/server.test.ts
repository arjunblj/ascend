import { afterAll, describe, expect, test } from 'bun:test'
import { unlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendWorkbook } from '@ascend/sdk'
import { createZip, encode } from '../../../packages/io-xlsx/src/writer/zip.ts'
import { createApiFetch, createServer } from './server.ts'

const TEMP_FILE = join(
	tmpdir(),
	`ascend-api-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
)
const OUTPUT_FILE = join(
	tmpdir(),
	`ascend-api-out-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
)
const MACRO_FILE = join(
	tmpdir(),
	`ascend-api-macro-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
)
const MACRO_OUTPUT_FILE = join(
	tmpdir(),
	`ascend-api-macro-out-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
)

let server: ReturnType<typeof createServer> | undefined

interface ApiEnvelope {
	readonly ok: boolean
	readonly data?: {
		readonly approvals?: readonly { readonly id: string }[]
		readonly replayable?: boolean
		readonly formulaCount?: number
		readonly ops?: unknown[]
		readonly pathMutations?: {
			readonly replayable?: boolean
			readonly ops?: unknown[]
		}
		readonly preview?: {
			readonly changedCellCount?: number
			readonly emittedChangedCellCount?: number
			readonly changedCells?: unknown[]
			readonly wouldSucceed?: boolean
		}
		readonly preparedPlan?: {
			readonly id?: string
			readonly expiresAt?: string
			readonly ttlMs?: number
		}
		readonly journal?: { readonly supported?: boolean; readonly inverseOps?: unknown[] }
		readonly partPath?: string
		readonly featureFamily?: string
		readonly text?: string
		readonly base64?: string
		readonly origin?: string
		readonly semantics?: string
		readonly encoding?: string
		readonly previewByteLength?: number
		readonly truncated?: boolean
		readonly sha256?: string
		readonly rowCount?: number
		readonly cells?: unknown[]
		readonly format?: string
		readonly changeToken?: string
		readonly load?: {
			readonly mode?: string
			readonly isPartial?: boolean
			readonly maxRows?: number
			readonly cellsHydrated?: boolean
			readonly loadedSheets?: readonly string[]
		}
	}
	readonly error?: {
		readonly message?: string
		readonly code?: string
		readonly details?: {
			readonly issueCount?: number
			readonly issues?: readonly string[]
			readonly validPath?: boolean
			readonly rule?: string
			readonly load?: {
				readonly mode?: string
				readonly isPartial?: boolean
				readonly maxRows?: number
				readonly partialReasons?: readonly string[]
			}
			readonly supportedPathShapes?: readonly string[]
			readonly compiledOps?: readonly unknown[]
			readonly issueDetails?: readonly {
				readonly code?: string
				readonly opIndex?: number
				readonly path?: string
			}[]
		}
	}
}

afterAll(async () => {
	server?.stop(true)
	await unlink(TEMP_FILE).catch(() => {})
	await unlink(OUTPUT_FILE).catch(() => {})
	await unlink(MACRO_FILE).catch(() => {})
	await unlink(MACRO_OUTPUT_FILE).catch(() => {})
})

async function postJson(
	path: string,
	body: unknown,
): Promise<{ status: number; body: ApiEnvelope }> {
	server ??= createServer({ port: 0 })
	const response = await fetch(`http://127.0.0.1:${server.port}${path}`, {
		method: 'POST',
		headers: { 'Content-Type': 'application/json' },
		body: JSON.stringify(body),
	})
	return { status: response.status, body: (await response.json()) as ApiEnvelope }
}

async function postApiFetch(
	apiFetch: typeof fetch,
	path: string,
	body: unknown,
): Promise<{ status: number; body: ApiEnvelope }> {
	const response = await apiFetch(
		new Request(`http://ascend.local${path}`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify(body),
		}),
	)
	return { status: response.status, body: (await response.json()) as ApiEnvelope }
}

describe('Ascend API server', () => {
	test('dump emits replayable operation batches', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'B1', value: 'label' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B2', formula: 'A1*2' },
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/dump', { file: TEMP_FILE, sheet: 'Sheet1' })
		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.replayable).toBe(true)
		expect(result.body.data?.formulaCount).toBe(1)
		expect(result.body.data?.ops).toEqual([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 10 },
					{ ref: 'B1', value: 'label' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B2', formula: 'A1*2' },
		])
	})

	test('template-merge emits replayable operation batches and unresolved placeholders', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: '{{amount}}' },
					{ ref: 'A2', value: 'Missing {{client}}' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1+{{tax}}' },
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/template-merge', {
			file: TEMP_FILE,
			sheet: 'Sheet1',
			data: { amount: 10, tax: 2 },
		})
		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.replayable).toBe(false)
		expect(result.body.data?.unresolved).toEqual([
			{
				sheet: 'Sheet1',
				ref: 'A2',
				source: 'value',
				placeholder: '{{client}}',
				key: 'client',
			},
		])
		expect(result.body.data?.ops).toEqual([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'A1', value: 10 }],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1+2' },
		])
	})

	test('raw-part returns bounded package text and metadata', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const result = await postJson('/raw-part', {
			file: TEMP_FILE,
			partPath: 'xl/workbook.xml',
			maxBytes: 64,
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.partPath).toBe('xl/workbook.xml')
		expect(result.body.data?.origin).toBe('source')
		expect(result.body.data?.load?.mode).toBe('metadata-only')
		expect(result.body.data?.load?.isPartial).toBe(true)
		expect(result.body.data?.semantics).toBe('raw-package-bytes')
		expect(result.body.data?.featureFamily).toBe('workbook')
		expect(result.body.data?.text).toContain('<?xml')
		expect(result.body.data?.previewByteLength).toBe(64)
		expect(result.body.data?.truncated).toBe(true)
		expect(result.body.data?.sha256).toMatch(/^[a-f0-9]{64}$/)

		const metadataOnly = await postJson('/raw-part', {
			file: TEMP_FILE,
			partPath: '/xl/workbook.xml',
			encoding: 'none',
		})
		expect(metadataOnly.status).toBe(200)
		expect(metadataOnly.body.data?.encoding).toBe('none')
		expect(metadataOnly.body.data?.previewByteLength).toBe(0)
		expect(metadataOnly.body.data?.text).toBeUndefined()

		const missing = await postJson('/raw-part', { file: TEMP_FILE, partPath: 'xl/missing.xml' })
		expect(missing.status).toBe(404)
		expect(missing.body.ok).toBe(false)
		expect(missing.body.error?.code).toBe('FILE_NOT_FOUND')

		const invalid = await postJson('/raw-part', { file: TEMP_FILE, partPath: 'xl//workbook.xml' })
		expect(invalid.status).toBe(400)
		expect(invalid.body.error?.code).toBe('VALIDATION_ERROR')
		expect(invalid.body.error?.details?.validPath).toBe(false)

		const badMaxBytes = await postJson('/raw-part', {
			file: TEMP_FILE,
			partPath: 'xl/workbook.xml',
			maxBytes: -1,
		})
		expect(badMaxBytes.status).toBe(400)

		const badEncoding = await postJson('/raw-part', {
			file: TEMP_FILE,
			partPath: 'xl/workbook.xml',
			encoding: 'utf16',
		})
		expect(badEncoding.status).toBe(400)
	})

	test('read returns compact first-window data with partial load metadata', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: Array.from({ length: 20 }, (_, row) => [
					{ ref: `A${row + 1}`, value: row + 1 },
					{ ref: `B${row + 1}`, value: `row-${row + 1}` },
				]).flat(),
			},
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/read', {
			file: TEMP_FILE,
			range: 'A1:B20',
			format: 'compact',
			rowLimit: 3,
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.format).toBe('compact')
		expect(result.body.data?.rowCount).toBe(3)
		expect(result.body.data?.cells).toEqual([
			[0, 0, 1],
			[0, 1, 'row-1'],
			[1, 0, 2],
			[1, 1, 'row-2'],
			[2, 0, 3],
			[2, 1, 'row-3'],
		])
		expect(result.body.data?.changeToken).toBeDefined()
		expect(result.body.data?.load?.mode).toBe('values')
		expect(result.body.data?.load?.isPartial).toBe(true)
		expect(result.body.data?.load?.maxRows).toBe(3)
		expect(result.body.data?.load?.partialReasons).toContain(
			'only the first 3 row(s) are hydrated per loaded sheet',
		)
		expect(result.body.data?.load?.cellsHydrated).toBe(true)
		expect(result.body.data?.load?.loadedSheets).toEqual(['Sheet1'])

		const unchanged = await postJson('/read', {
			file: TEMP_FILE,
			range: 'A1:B20',
			format: 'compact',
			rowLimit: 3,
			changedSince: result.body.data?.changeToken,
		})
		expect(unchanged.status).toBe(200)
		expect(unchanged.body.data?.cells).toEqual([])
		expect(unchanged.body.data?.changeToken).toBeDefined()
	})

	test('trace returns structured partial-load diagnostics for capped formula views', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*2' },
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/trace', {
			file: TEMP_FILE,
			cell: 'Sheet1!A1',
			maxRows: 1,
		})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error?.code).toBe('VALIDATION_ERROR')
		expect(result.body.error?.details?.rule).toBe('partial-dependency-analysis')
		expect(result.body.error?.details?.load?.maxRows).toBe(1)
		expect(result.body.error?.details?.load?.partialReasons).toContain(
			'only the first 1 row(s) are hydrated per loaded sheet',
		)
	})

	test('preview accepts path-addressed mutations without saving', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])
		await wb.save(TEMP_FILE)

		const result = await postJson('/preview', {
			file: TEMP_FILE,
			journal: true,
			mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' }],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.pathMutations?.replayable).toBe(true)
		expect(result.body.data?.pathMutations?.ops).toEqual([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'new' }] },
		])
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.inverseOps).toEqual([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] },
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'old',
		})

		const ambiguous = await postJson('/preview', {
			file: TEMP_FILE,
			ops: [],
			mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' }],
		})
		expect(ambiguous.status).toBe(400)
		expect(ambiguous.body.error?.message).toBe('Provide either ops or mutations, not both')
	})

	test('preview, plan, write, and commit keep escaped path mutations canonical', async () => {
		const sheetName = "Q1.Forecast's Café Δ"
		const tableName = 'Sales.Δ'
		const columnName = 'Gross Profit/Δ~'
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'renameSheet', sheet: 'Sheet1', newName: sheetName },
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: columnName },
					{ ref: 'A2', value: 'North' },
					{ ref: 'B2', value: 10 },
				],
			},
			{ op: 'createTable', sheet: sheetName, ref: 'A1:B2', name: tableName, hasHeaders: true },
		])
		await wb.save(TEMP_FILE)
		const mutations = [
			{ path: `/sheets/${pointerSegment(sheetName)}/cells/A2/value`, value: 'pointer' },
			{ path: `sheets.${dotSegment(sheetName)}.cells.A3.value`, value: 'dot' },
			{ path: ['sheets', sheetName, 'cells', 'A4', 'value'], value: 'array' },
			{
				path: `tables.${dotSegment(tableName)}.columns.${dotSegment(columnName)}.formula`,
				value: 'SUM([Region])',
			},
			{
				path: ['tables', tableName, 'columns', columnName, 'name'],
				value: 'Net_Δ',
			},
		]
		const canonicalOps = [
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [
					{ ref: 'A2', value: 'pointer' },
					{ ref: 'A3', value: 'dot' },
					{ ref: 'A4', value: 'array' },
				],
			},
			{
				op: 'setTableColumn',
				table: tableName,
				column: columnName,
				formula: 'SUM([Region])',
			},
			{ op: 'setTableColumn', table: tableName, column: columnName, newName: 'Net_Δ' },
		]

		const preview = await postJson('/preview', { file: TEMP_FILE, mutations })
		expect(preview.status).toBe(200)
		expect(preview.body.ok).toBe(true)
		expect(preview.body.data?.pathMutations?.replayable).toBe(true)
		expect(preview.body.data?.pathMutations?.ops).toEqual(canonicalOps)

		const plan = await postJson('/plan', { file: TEMP_FILE, mutations })
		expect(plan.status).toBe(200)
		expect(plan.body.ok).toBe(true)
		expect(plan.body.data?.pathMutations?.ops).toEqual(canonicalOps)
		const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

		const writePath = `${TEMP_FILE}.escaped-write.xlsx`
		const commitInput = `${TEMP_FILE}.escaped-commit-input.xlsx`
		const commitOutput = `${OUTPUT_FILE}.escaped-commit-output.xlsx`
		try {
			await wb.save(writePath)
			const write = await postJson('/write', { file: writePath, mutations })
			expect(write.status).toBe(200)
			expect(write.body.ok).toBe(true)
			expect(write.body.data?.pathMutations?.ops).toEqual(canonicalOps)
			const writeReopened = await AscendWorkbook.open(writePath)
			expect(writeReopened.sheet(sheetName)?.cell('A3')?.value).toEqual({
				kind: 'string',
				value: 'dot',
			})
			expect(writeReopened.sheet(sheetName)?.cell('A4')?.value).toEqual({
				kind: 'string',
				value: 'array',
			})

			await wb.save(commitInput)
			const commit = await postJson('/commit', {
				file: commitInput,
				output: commitOutput,
				mutations,
				approvals: approvalIds,
			})
			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.pathMutations?.ops).toEqual(canonicalOps)
			const commitReopened = await AscendWorkbook.open(commitOutput)
			expect(commitReopened.sheet(sheetName)?.cell('A3')?.value).toEqual({
				kind: 'string',
				value: 'dot',
			})
			expect(commitReopened.sheet(sheetName)?.cell('A4')?.value).toEqual({
				kind: 'string',
				value: 'array',
			})
		} finally {
			await unlink(writePath).catch(() => {})
			await unlink(commitInput).catch(() => {})
			await unlink(commitOutput).catch(() => {})
		}
	})

	test('plan reports path mutation compiler errors as structured repair details', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const result = await postJson('/plan', {
			file: TEMP_FILE,
			mutations: [{ path: '/sheets/Missing/cells/A1/value', value: 1 }],
		})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error?.code).toBe('VALIDATION_ERROR')
		expect(result.body.error?.details?.issueCount).toBe(1)
		expect(result.body.error?.details?.issues).toEqual(['Sheet "Missing" not found.'])
		expect(result.body.error?.details?.issueDetails).toEqual([
			expect.objectContaining({
				code: 'sheet_not_found',
				path: '/sheets/Missing/cells/A1/value',
			}),
		])
		expect(result.body.error?.details?.supportedPathShapes).toEqual(
			expect.arrayContaining([
				'/sheets/{sheet}/ranges/{A1:B2}/conditionalFormat',
				'/tables/{table}/columns/{nameOrIndex}/totalsRowLabel',
				'/sheets/{sheet}/names/{name}/ref',
			]),
		)
	})

	test('plan reports malformed path syntax as structured repair details', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const result = await postJson('/plan', {
			file: TEMP_FILE,
			mutations: [
				{ path: '/sheets//cells/A1/value', value: 1 },
				{ path: '/sheets/%E0%A4%A/cells/A1/value', value: 1 },
				{ path: '/sheets/Sheet1~2/cells/A1/value', value: 1 },
			],
		})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error?.code).toBe('VALIDATION_ERROR')
		expect(result.body.error?.details?.issueCount).toBe(3)
		expect(result.body.error?.details?.issues).toEqual([
			'Path segment 1 must not be empty.',
			'Invalid percent encoding in path segment "%E0%A4%A".',
			'Invalid JSON Pointer escape in path segment "Sheet1~2".',
		])
		expect(result.body.error?.details?.issueDetails).toEqual([
			expect.objectContaining({ code: 'invalid_path', path: '/sheets//cells/A1/value' }),
			expect.objectContaining({ code: 'invalid_path', path: '/sheets/%E0%A4%A/cells/A1/value' }),
			expect.objectContaining({ code: 'invalid_path', path: '/sheets/Sheet1~2/cells/A1/value' }),
		])
	})

	test('malformed path mutations block preview, write, and commit consistently', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])
		await wb.save(TEMP_FILE)

		for (const endpoint of ['/preview', '/write', '/commit'] as const) {
			const result = await postJson(endpoint, {
				file: TEMP_FILE,
				output: OUTPUT_FILE,
				mutations: [{ path: '/sheets//cells/A1/value', value: 'new' }],
			})

			expect(result.status).toBe(400)
			expect(result.body.ok).toBe(false)
			expect(result.body.error?.code).toBe('VALIDATION_ERROR')
			expect(result.body.error?.details?.issueCount).toBe(1)
			expect(result.body.error?.details?.issues).toEqual(['Path segment 1 must not be empty.'])
			expect(result.body.error?.details?.issueDetails).toEqual([
				expect.objectContaining({
					code: 'invalid_path',
					path: '/sheets//cells/A1/value',
				}),
			])
		}

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'old',
		})
	})

	test('non-replayable path mutation batches do not expose or apply partial ops', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])
		await wb.save(TEMP_FILE)

		for (const endpoint of ['/plan', '/preview', '/write', '/commit'] as const) {
			const result = await postJson(endpoint, {
				file: TEMP_FILE,
				output: OUTPUT_FILE,
				mutations: [
					{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' },
					{ path: '/sheets/Missing/cells/A1/value', value: 1 },
				],
			})

			expect(result.status).toBe(400)
			expect(result.body.ok).toBe(false)
			expect(result.body.error?.code).toBe('VALIDATION_ERROR')
			expect(result.body.error?.details?.issueCount).toBe(1)
			expect(result.body.error?.details?.compiledOps).toEqual([])
			expect(result.body.error?.details?.issueDetails).toEqual([
				expect.objectContaining({
					code: 'sheet_not_found',
					path: '/sheets/Missing/cells/A1/value',
				}),
			])
		}

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'old',
		})
	})

	test('plan invalid ops return structured batch repair details', async () => {
		const ops = [
			{ op: 'insertRows', sheet: 'Sheet1', at: 0, count: '2' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: { nested: true } }] },
			{ op: 'missingOp', sheet: 'Sheet1' },
		]

		const result = await postJson('/plan', { file: TEMP_FILE, ops })
		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error?.code).toBe('VALIDATION_ERROR')
		expect(result.body.error?.details?.issueCount).toBe(3)
		expect(result.body.error?.details?.issues).toEqual(
			expect.arrayContaining([
				'ops[0].count must be a positive integer',
				'ops[1].updates[0].value must be a scalar value or null',
				'ops[2].op "missingOp" is not supported',
			]),
		)
		expect(result.body.error?.details?.issueDetails).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ code: 'invalid_type', opIndex: 0, path: 'ops[0].count' }),
				expect.objectContaining({ code: 'invalid_type', opIndex: 1 }),
				expect.objectContaining({ code: 'invalid_operation', opIndex: 2, path: 'ops[2].op' }),
			]),
		)
	})

	test('plan and commit require exact approval ids', async () => {
		const workbook = AscendWorkbook.create()
		workbook.apply([{ op: 'addSheet', name: 'Scratch' }])
		await workbook.save(TEMP_FILE)
		const ops = [{ op: 'deleteSheet', sheet: 'Scratch' }]

		const plan = await postJson('/plan', { file: TEMP_FILE, ops })
		expect(plan.status).toBe(200)
		expect(plan.body.ok).toBe(true)
		const approvalId = plan.body.data.approvals[0].id
		expect(approvalId).toBe('op:0:deletesheet')

		const aliasCommit = await postJson('/commit', {
			file: TEMP_FILE,
			ops,
			output: OUTPUT_FILE,
			approvals: ['deleteSheet'],
		})
		expect(aliasCommit.status).toBe(400)
		expect(aliasCommit.body.ok).toBe(false)
		expect(aliasCommit.body.error.message).toBe('Commit requires explicit approval')

		const exactCommit = await postJson('/commit', {
			file: TEMP_FILE,
			ops,
			output: OUTPUT_FILE,
			approvals: [approvalId],
		})
		expect(exactCommit.status).toBe(200)
		expect(exactCommit.body.ok).toBe(true)
		expect(exactCommit.body.data.approvals[0].id).toBe(approvalId)
	})

	test('commit requires exact approval ids for preserved lossy features', async () => {
		await Bun.write(MACRO_FILE, signedMacroWorkbook())
		const ops = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] }]
		const plan = await postJson('/plan', { file: MACRO_FILE, ops })
		expect(plan.status).toBe(200)
		expect(plan.body.ok).toBe(true)
		const approvalIds = plan.body.data.approvals.map((approval) => approval.id)
		expect(approvalIds).toEqual(
			expect.arrayContaining([
				expect.stringMatching(/^loss:preservedmacro:preserved:/),
				expect.stringMatching(/^loss:preservedsignature:preserved:/),
			]),
		)

		const aliasCommit = await postJson('/commit', {
			file: MACRO_FILE,
			ops,
			output: MACRO_OUTPUT_FILE,
			approvals: ['preservedMacro', 'preservedSignature'],
		})
		expect(aliasCommit.status).toBe(400)
		expect(aliasCommit.body.ok).toBe(false)
		expect(aliasCommit.body.error.message).toBe('Commit requires explicit approval')

		const exactCommit = await postJson('/commit', {
			file: MACRO_FILE,
			ops,
			output: MACRO_OUTPUT_FILE,
			approvals: approvalIds,
		})
		expect(exactCommit.status).toBe(200)
		expect(exactCommit.body.ok).toBe(true)
		expect(exactCommit.body.data.approvals.map((approval) => approval.id)).toEqual(approvalIds)
	})

	test('path mutation commit preserves canonical ops and exact approval ids', async () => {
		await Bun.write(MACRO_FILE, signedMacroWorkbook())
		const output = `${MACRO_OUTPUT_FILE}.path.xlsm`
		await unlink(output).catch(() => {})
		const mutations = [{ path: '/sheets/Sheet1/cells/A1/value', value: 11 }]
		const canonicalOps = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }]

		const plan = await postJson('/plan', { file: MACRO_FILE, mutations })
		expect(plan.status).toBe(200)
		expect(plan.body.ok).toBe(true)
		expect(plan.body.data?.pathMutations?.replayable).toBe(true)
		expect(plan.body.data?.pathMutations?.ops).toEqual(canonicalOps)
		const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []
		expect(approvalIds).toEqual(
			expect.arrayContaining([
				expect.stringMatching(/^loss:preservedmacro:preserved:/),
				expect.stringMatching(/^loss:preservedsignature:preserved:/),
			]),
		)

		const aliasCommit = await postJson('/commit', {
			file: MACRO_FILE,
			mutations,
			output,
			approvals: ['preservedMacro', 'preservedSignature'],
		})
		expect(aliasCommit.status).toBe(400)
		expect(aliasCommit.body.ok).toBe(false)
		expect(aliasCommit.body.error?.message).toBe('Commit requires explicit approval')

		const exactCommit = await postJson('/commit', {
			file: MACRO_FILE,
			mutations,
			output,
			approvals: approvalIds,
		})
		expect(exactCommit.status).toBe(200)
		expect(exactCommit.body.ok).toBe(true)
		expect(exactCommit.body.data?.pathMutations?.ops).toEqual(canonicalOps)
		expect(exactCommit.body.data?.approvals?.map((approval) => approval.id)).toEqual(approvalIds)
		const reopened = await AscendWorkbook.open(output)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 11 })
		await unlink(output).catch(() => {})
	})

	test('plan can return compact bounded preview details', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
					{ ref: 'A3', value: 3 },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/plan', {
			file: TEMP_FILE,
			compact: true,
			maxChangedCells: 1,
			ops: [
				{
					op: 'setCells',
					sheet: 'Sheet1',
					updates: [
						{ ref: 'A1', value: 10 },
						{ ref: 'A2', value: 20 },
						{ ref: 'A3', value: 30 },
					],
				},
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.preview?.wouldSucceed).toBe(true)
		expect(result.body.data?.preview?.changedCellCount).toBe(3)
		expect(result.body.data?.preview?.emittedChangedCellCount).toBe(1)
		expect(result.body.data?.preview?.changedCells).toHaveLength(1)
	})

	test('prepared plan handles commit without reopening operation input', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = `${OUTPUT_FILE}.prepared.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: TEMP_FILE,
				prepare: true,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 123 }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.preparedPlan?.id).toBeString()
			expect(plan.body.data?.preparedPlan?.expiresAt).toBeString()
			expect(plan.body.data?.preparedPlan?.ttlMs).toBeNumber()
			expect(plan.body.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 123 }] },
			])

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: [],
			})
			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 123 }] },
			])
			const reopened = await AscendWorkbook.open(output)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 123 })

			const reused = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output: `${output}.reuse.xlsx`,
				approvals: [],
			})
			expect(reused.status).toBe(400)
			expect(reused.body.error?.code).toBe('VALIDATION_ERROR')
		} finally {
			await unlink(output).catch(() => {})
			await unlink(`${output}.reuse.xlsx`).catch(() => {})
		}
	})

	test('prepared plan handles expire before commit', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		let now = 1_000
		const apiFetch = createApiFetch({
			preparedPlanTtlMs: 10,
			now: () => now,
		})

		const plan = await postApiFetch(apiFetch, '/plan', {
			file: TEMP_FILE,
			prepare: true,
			mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 456 }],
		})
		expect(plan.status).toBe(200)
		expect(plan.body.data?.preparedPlan?.id).toBeString()
		expect(plan.body.data?.preparedPlan?.ttlMs).toBe(10)

		now += 11
		const commit = await postApiFetch(apiFetch, '/commit', {
			planHandle: plan.body.data?.preparedPlan?.id,
			output: `${OUTPUT_FILE}.expired.xlsx`,
			approvals: [],
		})
		expect(commit.status).toBe(400)
		expect(commit.body.error?.code).toBe('VALIDATION_ERROR')
		await unlink(`${OUTPUT_FILE}.expired.xlsx`).catch(() => {})
	})
})

function signedMacroWorkbook(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>`),
				'xl/_rels/vbaProject.bin.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVbaSignature" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/vbaProject.bin': encode('macro-bytes'),
				'xl/vbaProjectSignature.bin': encode('signature-bytes'),
			}),
		),
	)
}

function pointerSegment(value: string): string {
	return encodeURIComponent(value.replace(/~/g, '~0').replace(/\//g, '~1'))
}

function dotSegment(value: string): string {
	return value.replace(/\\/g, '\\\\').replace(/\./g, '\\.')
}
