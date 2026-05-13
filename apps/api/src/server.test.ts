import { afterAll, describe, expect, test } from 'bun:test'
import { Buffer } from 'node:buffer'
import { createHash } from 'node:crypto'
import { unlink, writeFile } from 'node:fs/promises'
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
		readonly output?: string
		readonly backup?: string
		readonly outputSha256?: string
		readonly approvals?: readonly { readonly id: string }[]
		readonly replayable?: boolean
		readonly formulaCount?: number
		readonly ops?: unknown[]
		readonly changed?: readonly string[]
		readonly dirtyRegions?: readonly unknown[]
		readonly generations?: {
			readonly workbook?: number
			readonly formulas?: number
			readonly sheetMetadata?: number
			readonly styles?: number
		}
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
		readonly apply?: {
			readonly affectedCellCount?: number
		}
		readonly timings?: {
			readonly applyMs?: number
			readonly writePlanSummaryMs?: number
			readonly writePolicyCheckMs?: number
			readonly toBytesMs?: number
			readonly outputByteReadMs?: number
		}
		readonly trace?: {
			readonly artifactCount?: number
			readonly artifacts?: unknown[]
		}
		readonly postWrite?: {
			readonly valid?: boolean
			readonly outputSha256?: string
			readonly auditsPassed?: boolean
			readonly expectedPackageGraphIssueCount?: number
			readonly unresolvedPackageGraphIssueCount?: number
			readonly reopened?: boolean
			readonly timings?: {
				readonly reopenMs?: number
			}
			readonly check?: {
				readonly valid?: boolean
			}
			readonly lint?: {
				readonly clean?: boolean
				readonly warningCount?: number
				readonly errorCount?: number
				readonly parseErrorCount?: number
			}
			readonly packageGraphAudit?: {
				readonly ok?: boolean
				readonly issueCount?: number
			}
		}
		readonly modelOutput?: {
			readonly blocked?: boolean
			readonly nextActions?: readonly string[]
			readonly counts?: {
				readonly postWritePackageGraphIssues?: number
				readonly postWriteLintFailures?: number
			}
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
		readonly binaryLike?: boolean
		readonly textWarning?: string
		readonly rowCount?: number
		readonly cells?: unknown[]
		readonly format?: string
		readonly changeToken?: string
		readonly snapshot?: {
			readonly token?: string
			readonly generations?: {
				readonly workbook?: number
				readonly sheetMetadata?: number
				readonly formulas?: number
				readonly styles?: number
			}
			readonly load?: {
				readonly mode?: string
				readonly isPartial?: boolean
				readonly maxRows?: number
			}
		}
		readonly valid?: boolean
		readonly issues?: readonly {
			readonly rule?: string
			readonly message?: string
		}[]
		readonly clean?: boolean
		readonly warnings?: readonly {
			readonly rule?: string
			readonly message?: string
		}[]
		readonly load?: {
			readonly mode?: string
			readonly isPartial?: boolean
			readonly maxRows?: number
			readonly cellsHydrated?: boolean
			readonly loadedSheets?: readonly string[]
			readonly partialReasons?: readonly string[]
		}
	}
	readonly error?: {
		readonly message?: string
		readonly code?: string
		readonly details?: {
			readonly issueCount?: number
			readonly issues?: readonly string[]
			readonly found?: boolean
			readonly validPath?: boolean
			readonly semantics?: string
			readonly rule?: string
			readonly caseInsensitiveAmbiguous?: boolean
			readonly load?: {
				readonly mode?: string
				readonly isPartial?: boolean
				readonly maxRows?: number
				readonly partialReasons?: readonly string[]
			}
			readonly supportedPathShapes?: readonly string[]
			readonly planHandle?: string
			readonly reason?: string
			readonly unsupportedLoadOptions?: readonly string[]
			readonly requiredLoad?: {
				readonly mode?: string
				readonly allSheets?: boolean
				readonly maxRows?: null | number
			}
			readonly compiledOps?: readonly unknown[]
			readonly issueDetails?: readonly {
				readonly code?: string
				readonly mutationIndex?: number
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

		const replayInput = `${TEMP_FILE}.dump-replay-input.xlsx`
		const replayOutput = `${OUTPUT_FILE}.dump-replay-output.xlsx`
		try {
			await AscendWorkbook.create().save(replayInput)
			const plan = await postJson('/plan', { file: replayInput, ops: result.body.data?.ops })
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.preparedPlan?.id).toBeString()

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output: replayOutput,
				compact: true,
			})
			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)

			const replayed = await AscendWorkbook.open(replayOutput)
			expect(replayed.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'number',
				value: 10,
			})
			expect(replayed.sheet('Sheet1')?.cell('B1')?.value).toEqual({
				kind: 'string',
				value: 'label',
			})
			expect(replayed.sheet('Sheet1')?.cell('B2')?.formula).toBe('A1*2')
		} finally {
			await unlink(replayInput).catch(() => {})
			await unlink(replayOutput).catch(() => {})
		}
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

		const replayable = await postJson('/template-merge', {
			file: TEMP_FILE,
			sheet: 'Sheet1',
			data: { amount: 10, tax: 2, client: 'Acme' },
		})
		expect(replayable.status).toBe(200)
		expect(replayable.body.ok).toBe(true)
		expect(replayable.body.data?.replayable).toBe(true)
		expect(replayable.body.data?.unresolved).toEqual([])

		const replayOutput = `${OUTPUT_FILE}.template-replay-output.xlsx`
		try {
			const plan = await postJson('/plan', { file: TEMP_FILE, ops: replayable.body.data?.ops })
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.preparedPlan?.id).toBeString()

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output: replayOutput,
				compact: true,
			})
			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)

			const merged = await AscendWorkbook.open(replayOutput)
			expect(merged.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 10 })
			expect(merged.sheet('Sheet1')?.cell('A2')?.value).toEqual({
				kind: 'string',
				value: 'Missing Acme',
			})
			expect(merged.sheet('Sheet1')?.cell('B1')?.formula).toBe('A1+2')
		} finally {
			await unlink(replayOutput).catch(() => {})
		}
	})

	test('calc supports range-scoped recalc without clearing pending formulas outside the range', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'C1', value: 10 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'D1', formula: 'C1*2' },
		])
		wb.recalc()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 5 },
					{ ref: 'C1', value: 20 },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const ranged = await postJson('/calc', { file: TEMP_FILE, range: 'Sheet1!B1:B1' })
		expect(ranged.status).toBe(200)
		expect(ranged.body.ok).toBe(true)
		expect(ranged.body.data?.changed).toEqual(['Sheet1!B1'])
		expect(ranged.body.data?.dirtyRegions).toEqual([
			{ sheet: 'Sheet1', range: 'B1:B1', refs: ['Sheet1!B1'] },
		])
		expect(ranged.body.data?.generations?.formulas).toBeNumber()
		let reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 10 })
		expect(reopened.sheet('Sheet1')?.cell('D1')?.value).toEqual({ kind: 'number', value: 20 })

		const full = await postJson('/calc', { file: TEMP_FILE })
		expect(full.status).toBe(200)
		expect(full.body.ok).toBe(true)
		expect(full.body.data?.changed).toEqual(['Sheet1!D1'])
		reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('D1')?.value).toEqual({ kind: 'number', value: 40 })
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
		expect(missing.body.error?.details?.found).toBe(false)
		expect(missing.body.error?.details?.validPath).toBe(true)
		expect(missing.body.error?.details?.semantics).toBe('raw-package-bytes')

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
		expect(badMaxBytes.body.error?.code).toBe('VALIDATION_ERROR')
		expect(badMaxBytes.body.error?.details?.rule).toBe('nonnegative integer')

		const tooLargeMaxBytes = await postJson('/raw-part', {
			file: TEMP_FILE,
			partPath: 'xl/workbook.xml',
			maxBytes: 1024 * 1024 + 1,
		})
		expect(tooLargeMaxBytes.status).toBe(400)
		expect(tooLargeMaxBytes.body.error?.code).toBe('VALIDATION_ERROR')
		expect(tooLargeMaxBytes.body.error?.details?.rule).toContain('at most')

		const badEncoding = await postJson('/raw-part', {
			file: TEMP_FILE,
			partPath: 'xl/workbook.xml',
			encoding: 'utf16',
		})
		expect(badEncoding.status).toBe(400)
		expect(badEncoding.body.error?.code).toBe('VALIDATION_ERROR')
	})

	test('raw-part returns binary base64 previews with full-byte metadata', async () => {
		const binaryBytes = Uint8Array.from({ length: 70 * 1024 }, (_, index) => index % 251)
		const binaryFile = `${TEMP_FILE}.raw-binary.xlsx`
		await writeFile(binaryFile, binaryRawPartWorkbook(binaryBytes))
		try {
			const textPreview = await postJson('/raw-part', {
				file: binaryFile,
				partPath: 'xl/media/image1.png',
				encoding: 'text',
				maxBytes: 6,
			})
			expect(textPreview.status).toBe(200)
			expect(textPreview.body.data?.binaryLike).toBe(true)
			expect(textPreview.body.data?.textWarning).toContain('Part appears binary')

			const defaultBounded = await postJson('/raw-part', {
				file: binaryFile,
				partPath: 'xl/media/image1.png',
				encoding: 'base64',
			})
			expect(defaultBounded.status).toBe(200)
			expect(defaultBounded.body.data?.previewByteLength).toBe(64 * 1024)
			expect(defaultBounded.body.data?.truncated).toBe(true)
			expect(defaultBounded.body.data?.sha256).toBe(
				createHash('sha256').update(binaryBytes).digest('hex'),
			)

			const result = await postJson('/raw-part', {
				file: binaryFile,
				partPath: 'xl/media/image1.png',
				encoding: 'base64',
				maxBytes: 3,
			})

			expect(result.status).toBe(200)
			expect(result.body.ok).toBe(true)
			expect(result.body.data?.encoding).toBe('base64')
			expect(result.body.data?.base64).toBe(
				Buffer.from(binaryBytes.subarray(0, 3)).toString('base64'),
			)
			expect(result.body.data?.text).toBeUndefined()
			expect(result.body.data?.previewByteLength).toBe(3)
			expect(result.body.data?.truncated).toBe(true)
			expect(result.body.data?.sha256).toBe(createHash('sha256').update(binaryBytes).digest('hex'))

			const metadataOnly = await postJson('/raw-part', {
				file: binaryFile,
				partPath: 'xl/media/image1.png',
				encoding: 'none',
				maxBytes: 3,
			})
			expect(metadataOnly.status).toBe(200)
			expect(metadataOnly.body.data?.encoding).toBe('none')
			expect(metadataOnly.body.data?.base64).toBeUndefined()
			expect(metadataOnly.body.data?.text).toBeUndefined()
			expect(metadataOnly.body.data?.previewByteLength).toBe(0)
			expect(metadataOnly.body.data?.truncated).toBe(false)
			expect(metadataOnly.body.data?.sha256).toBe(result.body.data?.sha256)

			const ambiguous = await postJson('/raw-part', {
				file: binaryFile,
				partPath: 'Xl/Media/Case.Png',
				caseInsensitive: true,
			})
			expect(ambiguous.status).toBe(400)
			expect(ambiguous.body.error?.code).toBe('VALIDATION_ERROR')
			expect(ambiguous.body.error?.details?.caseInsensitiveAmbiguous).toBe(true)
		} finally {
			await unlink(binaryFile).catch(() => {})
		}
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
		expect(result.body.data?.snapshot?.token).toContain('partial')
		expect(result.body.data?.snapshot?.generations).toEqual({
			workbook: 0,
			sheetMetadata: 0,
			formulas: 0,
			styles: 0,
		})
		expect(result.body.data?.snapshot?.load).toMatchObject({
			mode: 'values',
			isPartial: true,
			maxRows: 3,
		})
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

	test('compact changedSince reads return a fresh window after source changes', async () => {
		const original = AscendWorkbook.create()
		original.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])
		await original.save(TEMP_FILE)

		const first = await postJson('/read', {
			file: TEMP_FILE,
			range: 'A1:A1',
			format: 'compact',
		})
		expect(first.status).toBe(200)
		expect(first.body.data?.cells).toEqual([[0, 0, 'old']])
		expect(first.body.data?.changeToken).toBeDefined()

		const changed = AscendWorkbook.create()
		changed.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'new' }] }])
		await changed.save(TEMP_FILE)

		const afterChange = await postJson('/read', {
			file: TEMP_FILE,
			range: 'A1:A1',
			format: 'compact',
			changedSince: first.body.data?.changeToken,
		})
		expect(afterChange.status).toBe(200)
		expect(afterChange.body.ok).toBe(true)
		expect(afterChange.body.data?.cells).toEqual([[0, 0, 'new']])
		expect(afterChange.body.data?.changeToken).toBeDefined()
	})

	test('read preview defaults compact reads to a bounded first window', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: Array.from({ length: 520 }, (_, row) => [
					{ ref: `A${row + 1}`, value: row + 1 },
					{ ref: `B${row + 1}`, value: `row-${row + 1}` },
				]).flat(),
			},
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/read', {
			file: TEMP_FILE,
			range: 'A1:B520',
			format: 'compact',
			preview: true,
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.format).toBe('compact')
		expect(result.body.data?.rowCount).toBe(500)
		expect(result.body.data?.cells).toHaveLength(1000)
		expect(result.body.data?.load?.mode).toBe('values')
		expect(result.body.data?.load?.isPartial).toBe(true)
		expect(result.body.data?.load?.maxRows).toBe(500)
		expect(result.body.data?.load?.partialReasons).toContain(
			'only the first 500 row(s) are hydrated per loaded sheet',
		)
	})

	test('compact reads default to a bounded first window', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: Array.from({ length: 520 }, (_, row) => [
					{ ref: `A${row + 1}`, value: row + 1 },
					{ ref: `B${row + 1}`, value: `row-${row + 1}` },
				]).flat(),
			},
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/read', {
			file: TEMP_FILE,
			range: 'A1:B520',
			format: 'compact',
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.format).toBe('compact')
		expect(result.body.data?.rowCount).toBe(500)
		expect(result.body.data?.cells).toHaveLength(1000)
		expect(result.body.data?.load?.mode).toBe('values')
		expect(result.body.data?.load?.isPartial).toBe(true)
		expect(result.body.data?.load?.maxRows).toBe(500)
		expect(result.body.data?.load?.partialReasons).toContain(
			'only the first 500 row(s) are hydrated per loaded sheet',
		)
	})

	test('agent-view exposes partial-load metadata for sheet-scoped capped formula views', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'addSheet', name: 'Data' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*2' },
			{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 'hidden' }] },
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/agent-view', {
			file: TEMP_FILE,
			sheet: 'Sheet1',
			range: 'A1:A3',
			maxRows: 1,
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.load?.isPartial).toBe(true)
		expect(result.body.data?.load?.maxRows).toBe(1)
		expect(result.body.data?.load?.partialReasons).toContain('only selected sheets are loaded')
		expect(result.body.data?.load?.partialReasons).toContain(
			'only the first 1 row(s) are hydrated per loaded sheet',
		)
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

	test('check and lint expose partial-load metadata for capped formula views', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*2' },
		])
		await wb.save(TEMP_FILE)

		const check = await postJson('/check', {
			file: TEMP_FILE,
			maxRows: 1,
		})

		expect(check.status).toBe(200)
		expect(check.body.ok).toBe(true)
		expect(check.body.data?.valid).toBe(false)
		expect(check.body.data?.issues?.[0]?.rule).toBe('partial-dependency-analysis')
		expect(check.body.data?.load?.isPartial).toBe(true)
		expect(check.body.data?.load?.maxRows).toBe(1)
		expect(check.body.data?.load?.partialReasons).toContain(
			'only the first 1 row(s) are hydrated per loaded sheet',
		)

		const lint = await postJson('/lint', {
			file: TEMP_FILE,
			maxRows: 1,
		})

		expect(lint.status).toBe(200)
		expect(lint.body.ok).toBe(true)
		expect(lint.body.data?.clean).toBe(false)
		expect(lint.body.data?.warnings?.[0]?.rule).toBe('partial-dependency-analysis')
		expect(lint.body.data?.load?.isPartial).toBe(true)
		expect(lint.body.data?.load?.maxRows).toBe(1)
		expect(lint.body.data?.load?.partialReasons).toContain(
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

	test('ops and path mutations are mutually exclusive across edit endpoints', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		for (const endpoint of ['/preview', '/plan', '/write', '/commit'] as const) {
			const result = await postJson(endpoint, {
				file: TEMP_FILE,
				output: OUTPUT_FILE,
				ops: [],
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' }],
			})

			expect(result.status).toBe(400)
			expect(result.body.ok).toBe(false)
			expect(result.body.error?.code).toBe('VALIDATION_ERROR')
			expect(result.body.error?.message).toBe('Provide either ops or mutations, not both')
		}
	})

	test('preview, plan, write, and commit keep escaped path mutations canonical', async () => {
		const sheetName = "Q1.Forecast's Café Δ"
		const tableName = 'Sales.Δ'
		const tablePathName = tableName.toLowerCase()
		const columnName = "Gross.Profit / Δ~'s"
		const columnPathName = columnName.toLowerCase()
		const workbookName = 'Global.Rate_Δ'
		const scopedName = 'Local.Rate_Δ'
		const definedNameRef = `'${sheetName.replace(/'/g, "''")}'!$B$2`
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
				path: `tables.${dotSegment(tablePathName)}.columns.${dotSegment(columnPathName)}.formula`,
				value: 'SUM([Region])',
			},
			{ path: `/names/${pointerSegment(workbookName)}/ref`, value: definedNameRef },
			{
				path: `sheets.${dotSegment(sheetName)}.names.${dotSegment(scopedName)}.ref`,
				value: definedNameRef,
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
			{ op: 'setDefinedName', name: workbookName, ref: definedNameRef },
			{ op: 'setDefinedName', name: scopedName, scope: sheetName, ref: definedNameRef },
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
			expect(writeReopened.definedName(workbookName)?.formula).toBe(definedNameRef)
			expect(writeReopened.definedName(scopedName, sheetName)?.formula).toBe(definedNameRef)

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
			expect(commitReopened.definedName(workbookName)?.formula).toBe(definedNameRef)
			expect(commitReopened.definedName(scopedName, sheetName)?.formula).toBe(definedNameRef)
		} finally {
			await unlink(writePath).catch(() => {})
			await unlink(commitInput).catch(() => {})
			await unlink(commitOutput).catch(() => {})
		}
	})

	test('preview defers path mutation renames after dependent edits', async () => {
		const sheetName = 'Q1.Forecast'
		const tableName = 'Sales.Δ'
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'renameSheet', sheet: 'Sheet1', newName: sheetName },
			{
				op: 'setCells',
				sheet: sheetName,
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Revenue' },
					{ ref: 'A2', value: 'North' },
					{ ref: 'B2', value: 10 },
				],
			},
			{ op: 'createTable', sheet: sheetName, ref: 'A1:B2', name: tableName, hasHeaders: true },
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/preview', {
			file: TEMP_FILE,
			mutations: [
				{ path: `/sheets/${pointerSegment(sheetName)}/name`, value: 'Summary' },
				{ path: `/sheets/${pointerSegment(sheetName)}/cells/C1/value`, value: 'safe order' },
				{ path: `/tables/${pointerSegment(tableName)}/name`, value: 'SalesData' },
				{
					path: `/tables/${pointerSegment(tableName)}/columns/Revenue/formula`,
					value: 'SUM([Revenue])',
				},
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.pathMutations?.replayable).toBe(true)
		expect(result.body.data?.pathMutations?.ops).toEqual([
			{ op: 'setCells', sheet: sheetName, updates: [{ ref: 'C1', value: 'safe order' }] },
			{
				op: 'setTableColumn',
				table: tableName,
				column: 'Revenue',
				formula: 'SUM([Revenue])',
			},
			{ op: 'renameSheet', sheet: sheetName, newName: 'Summary' },
			{ op: 'renameTable', table: tableName, newName: 'SalesData' },
		])
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

	test('invalid path mutation shapes return structured repair details consistently', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		for (const endpoint of ['/preview', '/plan', '/write', '/commit'] as const) {
			const result = await postJson(endpoint, {
				file: TEMP_FILE,
				output: OUTPUT_FILE,
				mutations: [{ path: 123, value: 'new' }],
			})

			expect(result.status).toBe(400)
			expect(result.body.error?.code).toBe('VALIDATION_ERROR')
			expect(result.body.error?.details?.issueCount).toBe(1)
			expect(result.body.error?.details?.issues).toEqual([
				'mutations[0]: Mutation path must be a string or string array.',
			])
			expect(result.body.error?.details?.issueDetails).toEqual([
				expect.objectContaining({
					code: 'invalid_path_mutation',
					mutationIndex: 0,
					path: 'mutations[0]',
				}),
			])
		}
	})

	test('non-replayable path mutation batches do not expose or apply partial ops', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'old' },
					{ ref: 'B1', value: 'Amount' },
					{ ref: 'B2', value: 1 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])
		await wb.save(TEMP_FILE)

		for (const endpoint of ['/plan', '/preview', '/write', '/commit'] as const) {
			const result = await postJson(endpoint, {
				file: TEMP_FILE,
				output: OUTPUT_FILE,
				mutations: [
					{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' },
					{ path: '/sheets/Sheet1/name', value: 'Bad/Name' },
					{ path: '/tables/Sales/name', value: 'Bad Name' },
				],
			})

			expect(result.status).toBe(400)
			expect(result.body.ok).toBe(false)
			expect(result.body.error?.code).toBe('VALIDATION_ERROR')
			expect(result.body.error?.details?.issueCount).toBe(2)
			expect(result.body.error?.details?.compiledOps).toEqual([])
			expect(result.body.error?.details?.issueDetails).toEqual([
				expect.objectContaining({
					code: 'invalid_value',
					path: '/sheets/Sheet1/name',
				}),
				expect.objectContaining({
					code: 'invalid_value',
					path: '/tables/Sales/name',
				}),
			])
		}

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'old',
		})
		expect(reopened.sheets).toContain('Sheet1')
		expect(reopened.table('Sales')?.name).toBe('Sales')

		const prepared = await postJson('/plan', {
			file: TEMP_FILE,
			prepare: true,
			mutations: [
				{ path: '/sheets/Sheet1/cells/A1/value', value: 'new' },
				{ path: '/sheets/Sheet1/name', value: 'Bad/Name' },
			],
		})
		expect(prepared.status).toBe(400)
		expect(prepared.body.ok).toBe(false)
		expect(prepared.body.data?.preparedPlan).toBeUndefined()
		expect(prepared.body.error?.details?.compiledOps).toEqual([])
		expect(prepared.body.error?.details?.issueDetails).toEqual([
			expect.objectContaining({
				code: 'invalid_value',
				path: '/sheets/Sheet1/name',
			}),
		])
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

	test('plan rejects capped load options instead of silently producing full plans', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const result = await postJson('/plan', {
			file: TEMP_FILE,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
			maxRows: 1,
		})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error?.code).toBe('VALIDATION_ERROR')
		expect(result.body.error?.details?.unsupportedLoadOptions).toEqual(['maxRows'])
		expect(result.body.error?.details?.requiredLoad).toEqual({
			mode: 'full',
			allSheets: true,
			maxRows: null,
		})
	})

	test('commit rejects capped load options instead of silently producing full commits', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const result = await postJson('/commit', {
			file: TEMP_FILE,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
			output: OUTPUT_FILE,
			maxRows: 1,
			mode: 'values',
			sheets: ['Sheet1'],
		})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error?.code).toBe('VALIDATION_ERROR')
		expect(result.body.error?.details?.unsupportedLoadOptions).toEqual([
			'maxRows',
			'mode',
			'sheets',
		])
		expect(result.body.error?.details?.requiredLoad).toEqual({
			mode: 'full',
			allSheets: true,
			maxRows: null,
		})
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
				compact: true,
			})
			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 123 }] },
			])
			expect(commit.body.data?.apply?.affectedCellCount).toBe(1)
			expect(commit.body.data?.timings?.applyMs).toBeNumber()
			expect(commit.body.data?.timings?.writePlanSummaryMs).toBeNumber()
			expect(commit.body.data?.timings?.writePolicyCheckMs).toBeNumber()
			expect(commit.body.data?.timings?.toBytesMs).toBeNumber()
			expect(commit.body.data?.timings?.outputByteReadMs).toBeNumber()
			expect(commit.body.data?.postWrite?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.expectedPackageGraphIssueCount).toBe(0)
			expect(commit.body.data?.postWrite?.unresolvedPackageGraphIssueCount).toBe(0)
			expect(commit.body.data?.postWrite?.reopened).toBe(true)
			expect(commit.body.data?.postWrite?.timings?.reopenMs).toBeNumber()
			expect(commit.body.data?.postWrite?.check?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.packageGraphAudit?.ok).toBe(true)
			expect(commit.body.data?.trace?.artifactCount).toBeNumber()
			expect(commit.body.data?.trace?.artifacts).toBeUndefined()
			const reopened = await AscendWorkbook.open(output)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 123 })

			const reused = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output: `${output}.reuse.xlsx`,
				approvals: [],
			})
			expect(reused.status).toBe(400)
			expect(reused.body.error?.code).toBe('VALIDATION_ERROR')
			expect(reused.body.error?.message).toBe('Prepared plan handle has already been used')
			expect(reused.body.error?.details).toMatchObject({
				rule: 'prepared-plan-handle-unavailable',
				reason: 'already-used',
				planHandle: plan.body.data?.preparedPlan?.id,
			})
		} finally {
			await unlink(output).catch(() => {})
			await unlink(`${output}.reuse.xlsx`).catch(() => {})
		}
	})

	test('direct path mutation commits preserve in-place backups and post-write truth', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'original' }] }])
		await wb.save(TEMP_FILE)
		const backup = `${OUTPUT_FILE}.direct-backup.xlsx`
		try {
			const commit = await postJson('/commit', {
				file: TEMP_FILE,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 'updated' }],
				inPlace: true,
				backup,
				approvals: [],
				compact: true,
			})
			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.output).toBe(TEMP_FILE)
			expect(commit.body.data?.backup).toBe(backup)
			expect(commit.body.data?.outputSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(commit.body.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'updated' }] },
			])
			expect(commit.body.data?.postWrite?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.reopened).toBe(true)
			expect(commit.body.data?.postWrite?.outputSha256).toBe(commit.body.data?.outputSha256)
			expect(commit.body.data?.postWrite?.check?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.packageGraphAudit?.ok).toBe(true)

			const reopenedInput = await AscendWorkbook.open(TEMP_FILE)
			expect(reopenedInput.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'updated',
			})
			const reopenedBackup = await AscendWorkbook.open(backup)
			expect(reopenedBackup.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'original',
			})
		} finally {
			await unlink(backup).catch(() => {})
		}
	})

	test('prepared path mutation handles preserve in-place backups and remain one-shot', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'original' }] }])
		await wb.save(TEMP_FILE)
		const backup = `${OUTPUT_FILE}.prepared-backup.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: TEMP_FILE,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 'updated' }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.data?.preparedPlan?.id).toBeString()

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				inPlace: true,
				backup,
				approvals: [],
			})
			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.output).toBe(TEMP_FILE)
			expect(commit.body.data?.backup).toBe(backup)
			expect(commit.body.data?.pathMutations?.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'updated' }] },
			])
			expect(commit.body.data?.postWrite?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.reopened).toBe(true)

			const reopenedInput = await AscendWorkbook.open(TEMP_FILE)
			expect(reopenedInput.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'updated',
			})
			const reopenedBackup = await AscendWorkbook.open(backup)
			expect(reopenedBackup.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'original',
			})

			const reused = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				inPlace: true,
				backup,
				approvals: [],
			})
			expect(reused.status).toBe(400)
			expect(reused.body.error?.message).toBe('Prepared plan handle has already been used')
		} finally {
			await unlink(backup).catch(() => {})
		}
	})

	test('prepared plan handles require exact destructive approval ids', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Scratch' }])
		await wb.save(TEMP_FILE)
		const output = `${OUTPUT_FILE}.prepared-approval.xlsx`
		const ops = [{ op: 'deleteSheet', sheet: 'Scratch' }]
		try {
			const plan = await postJson('/plan', { file: TEMP_FILE, ops })
			expect(plan.status).toBe(200)
			expect(plan.body.data?.preparedPlan?.id).toBeString()
			const approvalId = plan.body.data?.approvals?.[0]?.id
			expect(approvalId).toBe('op:0:deletesheet')

			const aliasCommit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: ['deleteSheet'],
			})
			expect(aliasCommit.status).toBe(400)
			expect(aliasCommit.body.ok).toBe(false)
			expect(aliasCommit.body.error?.message).toBe('Commit requires explicit approval')
			expect(await Bun.file(output).exists()).toBe(false)

			const retryPlan = await postJson('/plan', { file: TEMP_FILE, ops })
			const exactCommit = await postJson('/commit', {
				planHandle: retryPlan.body.data?.preparedPlan?.id,
				output,
				approvals: [approvalId],
			})
			expect(exactCommit.status).toBe(200)
			expect(exactCommit.body.ok).toBe(true)
			expect(exactCommit.body.data?.approvals?.[0]?.id).toBe(approvalId)
			const reopened = await AscendWorkbook.open(output)
			expect(reopened.sheets).not.toContain('Scratch')
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('prepared path mutation handles require exact preserved-loss approval ids', async () => {
		await Bun.write(MACRO_FILE, signedMacroWorkbook())
		const output = `${MACRO_OUTPUT_FILE}.prepared-path.xlsm`
		const mutations = [{ path: '/sheets/Sheet1/cells/A1/value', value: 17 }]
		const canonicalOps = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 17 }] }]
		try {
			const plan = await postJson('/plan', { file: MACRO_FILE, mutations })
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.preparedPlan?.id).toBeString()
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
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: ['preservedMacro', 'preservedSignature'],
			})
			expect(aliasCommit.status).toBe(400)
			expect(aliasCommit.body.ok).toBe(false)
			expect(aliasCommit.body.error?.message).toBe('Commit requires explicit approval')
			expect(await Bun.file(output).exists()).toBe(false)

			const retryPlan = await postJson('/plan', { file: MACRO_FILE, mutations })
			expect(retryPlan.body.data?.preparedPlan?.id).toBeString()
			expect(retryPlan.body.data?.pathMutations?.ops).toEqual(canonicalOps)
			const exactCommit = await postJson('/commit', {
				planHandle: retryPlan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
			})
			expect(exactCommit.status).toBe(200)
			expect(exactCommit.body.ok).toBe(true)
			expect(exactCommit.body.data?.pathMutations?.ops).toEqual(canonicalOps)
			expect(exactCommit.body.data?.approvals?.map((approval) => approval.id)).toEqual(approvalIds)
			expect(exactCommit.body.data?.postWrite?.valid).toBe(true)
			expect(exactCommit.body.data?.postWrite?.reopened).toBe(true)
			const reopened = await AscendWorkbook.open(output)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 17 })
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('prepared path mutation handles surface post-write audit failures as blocked output', async () => {
		await Bun.write(TEMP_FILE, preservedCustomWorkbook())
		const output = `${OUTPUT_FILE}.prepared-preserved.xlsx`
		const mutations = [{ path: '/sheets/Sheet1/cells/A1/value', value: 17 }]
		const canonicalOps = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 17 }] }]
		try {
			const plan = await postJson('/plan', { file: TEMP_FILE, mutations })
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.preparedPlan?.id).toBeString()
			expect(plan.body.data?.pathMutations?.ops).toEqual(canonicalOps)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual([expect.stringMatching(/^loss:preservedother:preserved:/)])

			const aliasCommit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: ['preservedOther'],
				compact: true,
			})
			expect(aliasCommit.status).toBe(400)
			expect(aliasCommit.body.ok).toBe(false)
			expect(aliasCommit.body.error?.message).toBe('Commit requires explicit approval')
			expect(await Bun.file(output).exists()).toBe(false)

			const retryPlan = await postJson('/plan', { file: TEMP_FILE, mutations })
			expect(retryPlan.body.data?.preparedPlan?.id).toBeString()
			expect(retryPlan.body.data?.pathMutations?.ops).toEqual(canonicalOps)
			const exactCommit = await postJson('/commit', {
				planHandle: retryPlan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})
			expect(exactCommit.status).toBe(200)
			expect(exactCommit.body.ok).toBe(true)
			expect(exactCommit.body.data?.pathMutations?.ops).toEqual(canonicalOps)
			expect(exactCommit.body.data?.approvals?.map((approval) => approval.id)).toEqual(approvalIds)
			expect(exactCommit.body.data?.postWrite?.valid).toBe(true)
			expect(exactCommit.body.data?.postWrite?.auditsPassed).toBe(false)
			expect(exactCommit.body.data?.postWrite?.outputSha256).toBe(
				exactCommit.body.data?.outputSha256,
			)
			expect(exactCommit.body.data?.postWrite?.packageGraphAudit?.ok).toBe(false)
			expect(exactCommit.body.data?.postWrite?.packageGraphAudit?.issueCount).toBeGreaterThan(0)
			expect(exactCommit.body.data?.postWrite?.expectedPackageGraphIssueCount).toBe(0)
			expect(exactCommit.body.data?.postWrite?.unresolvedPackageGraphIssueCount).toBeGreaterThan(0)
			expect(exactCommit.body.data?.modelOutput?.blocked).toBe(true)
			expect(
				exactCommit.body.data?.modelOutput?.counts?.postWritePackageGraphIssues,
			).toBeGreaterThan(0)
			expect(exactCommit.body.data?.modelOutput?.nextActions?.join('\n')).toContain(
				'postWrite.packageGraphAudit.issues',
			)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('prepared commits surface post-write formula lint failures as blocked output', async () => {
		const input = `${TEMP_FILE}.prepared-lint-source.xlsx`
		const output = `${OUTPUT_FILE}.prepared-lint-out.xlsx`
		const wb = AscendWorkbook.create()
		await wb.save(input)
		const complexFormula = `=${Array.from({ length: 26 }, () => '1').join('+')}`
		try {
			const plan = await postJson('/plan', {
				file: input,
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: complexFormula }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.preparedPlan?.id).toBeString()

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				compact: true,
			})
			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(false)
			expect(commit.body.data?.postWrite?.lint?.clean).toBe(false)
			expect(commit.body.data?.postWrite?.lint?.errorCount).toBeGreaterThan(0)
			expect(commit.body.data?.postWrite?.packageGraphAudit?.ok).toBe(true)
			expect(commit.body.data?.modelOutput?.blocked).toBe(true)
			expect(commit.body.data?.modelOutput?.counts?.postWriteLintFailures).toBeGreaterThan(0)
			expect(commit.body.data?.modelOutput?.nextActions?.join('\n')).toContain(
				'postWrite.lint.warnings',
			)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('direct commits surface post-write formula lint failures as blocked output', async () => {
		const input = `${TEMP_FILE}.direct-lint-source.xlsx`
		const output = `${OUTPUT_FILE}.direct-lint-out.xlsx`
		const wb = AscendWorkbook.create()
		await wb.save(input)
		const complexFormula = `=${Array.from({ length: 26 }, () => '1').join('+')}`
		try {
			const commit = await postJson('/commit', {
				file: input,
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: complexFormula }],
				output,
				compact: true,
			})
			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(false)
			expect(commit.body.data?.postWrite?.lint?.clean).toBe(false)
			expect(commit.body.data?.postWrite?.lint?.errorCount).toBeGreaterThan(0)
			expect(commit.body.data?.postWrite?.packageGraphAudit?.ok).toBe(true)
			expect(commit.body.data?.modelOutput?.blocked).toBe(true)
			expect(commit.body.data?.modelOutput?.counts?.postWriteLintFailures).toBeGreaterThan(0)
			expect(commit.body.data?.modelOutput?.nextActions?.join('\n')).toContain(
				'postWrite.lint.warnings',
			)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('plan can opt out of the default prepared handle', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const plan = await postJson('/plan', {
			file: TEMP_FILE,
			prepare: false,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
		})

		expect(plan.status).toBe(200)
		expect(plan.body.ok).toBe(true)
		expect(plan.body.data?.preparedPlan).toBeUndefined()
		expect(plan.body.data?.preview?.wouldSucceed).toBe(true)
	})

	test('prepared path mutation handles reject stale input before writing output', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = `${OUTPUT_FILE}.prepared-stale.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: TEMP_FILE,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 123 }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.data?.preparedPlan?.id).toBeString()

			const changed = AscendWorkbook.create()
			changed.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 9 }] }])
			await changed.save(TEMP_FILE)

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: [],
			})
			expect(commit.status).toBe(400)
			expect(commit.body.ok).toBe(false)
			expect(commit.body.error?.code).toBe('VALIDATION_ERROR')
			expect(commit.body.error?.message).toBe(
				'Input workbook changed after agent plan was prepared',
			)
			expect(commit.body.error?.details?.expected).toMatch(/^[a-f0-9]{64}$/)
			expect(commit.body.error?.details?.actual).toMatch(/^[a-f0-9]{64}$/)
			expect(commit.body.error?.details?.actual).not.toBe(commit.body.error?.details?.expected)
			expect(commit.body.error?.details?.planDigest).toMatch(/^[a-f0-9]{64}$/)
			expect(await Bun.file(output).exists()).toBe(false)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('plan and commit reject partial load options before preparing or writing', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = `${OUTPUT_FILE}.partial-load.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: TEMP_FILE,
				prepare: true,
				maxRows: 1,
				mode: 'values',
				sheets: ['Sheet1'],
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 123 }],
			})
			expect(plan.status).toBe(400)
			expect(plan.body.ok).toBe(false)
			expect(plan.body.data?.preparedPlan).toBeUndefined()
			expect(plan.body.error?.code).toBe('VALIDATION_ERROR')
			expect(plan.body.error?.details?.unsupportedLoadOptions).toEqual([
				'maxRows',
				'mode',
				'sheets',
			])
			expect(plan.body.error?.details?.requiredLoad).toEqual({
				mode: 'full',
				allSheets: true,
				maxRows: null,
			})

			const commit = await postJson('/commit', {
				file: TEMP_FILE,
				output,
				maxRows: 1,
				mode: 'values',
				sheets: ['Sheet1'],
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 123 }] }],
			})
			expect(commit.status).toBe(400)
			expect(commit.body.ok).toBe(false)
			expect(commit.body.error?.code).toBe('VALIDATION_ERROR')
			expect(commit.body.error?.details?.unsupportedLoadOptions).toEqual([
				'maxRows',
				'mode',
				'sheets',
			])
			expect(await Bun.file(output).exists()).toBe(false)
		} finally {
			await unlink(output).catch(() => {})
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
		expect(commit.body.error?.message).toBe('Prepared plan handle expired')
		expect(commit.body.error?.details).toMatchObject({
			rule: 'prepared-plan-handle-unavailable',
			reason: 'expired',
			planHandle: plan.body.data?.preparedPlan?.id,
		})
		await unlink(`${OUTPUT_FILE}.expired.xlsx`).catch(() => {})
	})

	test('prepared plan handle eviction is structured', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const apiFetch = createApiFetch({ preparedPlanMaxHandles: 1 })

		const first = await postApiFetch(apiFetch, '/plan', {
			file: TEMP_FILE,
			mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 1 }],
		})
		const second = await postApiFetch(apiFetch, '/plan', {
			file: TEMP_FILE,
			mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 2 }],
		})
		expect(first.body.data?.preparedPlan?.id).toBeString()
		expect(second.body.data?.preparedPlan?.id).toBeString()

		const evicted = await postApiFetch(apiFetch, '/commit', {
			planHandle: first.body.data?.preparedPlan?.id,
			output: `${OUTPUT_FILE}.evicted.xlsx`,
			approvals: [],
		})
		expect(evicted.status).toBe(400)
		expect(evicted.body.error?.message).toBe('Prepared plan handle was evicted')
		expect(evicted.body.error?.details).toMatchObject({
			rule: 'prepared-plan-handle-unavailable',
			reason: 'evicted',
			planHandle: first.body.data?.preparedPlan?.id,
		})
		await unlink(`${OUTPUT_FILE}.evicted.xlsx`).catch(() => {})
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

function preservedCustomWorkbook(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/custom/custom1.xml" ContentType="application/custom+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/custom/custom1.xml': encode('<custom>preserve me</custom>'),
			}),
		),
	)
}

function binaryRawPartWorkbook(binaryBytes: Uint8Array): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/media/image1.png': binaryBytes,
				'xl/media/case.png': new Uint8Array([1]),
				'XL/MEDIA/CASE.PNG': new Uint8Array([2]),
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
