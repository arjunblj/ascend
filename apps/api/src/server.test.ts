import { afterAll, describe, expect, test } from 'bun:test'
import { Buffer } from 'node:buffer'
import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { unlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
	AscendWorkbook,
	MUTATION_JOURNAL_ISSUE_SCHEMA,
	MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
	parseOperations,
	WorkbookDocument,
} from '@ascend/sdk'
import { createZip, encode } from '../../../packages/io-xlsx/src/writer/zip.ts'
import { makeXlsx } from '../../../packages/io-xlsx/test/helpers.ts'
import { createApiFetch, createServer } from './server.ts'

const TEMP_FILE = join(
	tmpdir(),
	`ascend-api-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
)
const OUTPUT_FILE = join(
	tmpdir(),
	`ascend-api-out-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
)
const CHARTSHEET_FIXTURE = join(import.meta.dir, '../../../fixtures/xlsx/exceljs/chart-sheet.xlsx')
const PIVOT_FIXTURE = join(
	import.meta.dir,
	'../../../fixtures/xlsx/libreoffice/PivotTable_CachedDefinitionAndDataInSync.xlsx',
)
const ENCRYPTED_FIXTURE = join(
	import.meta.dir,
	'../../../fixtures/xlsx/calamine/pass_protected.xlsx',
)
const MACRO_FILE = join(
	tmpdir(),
	`ascend-api-macro-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
)
const MACRO_OUTPUT_FILE = join(
	tmpdir(),
	`ascend-api-macro-out-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
)
const JOURNAL_V1_FIXTURE = JSON.parse(
	readFileSync(
		join(import.meta.dir, '../../../fixtures/journal/mutation-journal-v1.json'),
		'utf-8',
	),
) as {
	readonly scenario: {
		readonly ops: readonly Record<string, unknown>[]
		readonly journal: {
			readonly schemaVersion: number
			readonly schemaId: string
			readonly supported: boolean
			readonly exact: boolean
			readonly inverseOpCount: number
			readonly issueCount: number
			readonly issues: readonly unknown[]
		}
	}
}

function compactJournal(journal: {
	readonly schemaVersion?: number
	readonly schemaId?: string
	readonly supported?: boolean
	readonly exact?: boolean
	readonly inverseOps?: readonly unknown[]
	readonly issues?: readonly unknown[]
}): typeof JOURNAL_V1_FIXTURE.scenario.journal {
	const { schemaVersion, schemaId, supported, exact, inverseOps, issues } = journal
	if (
		schemaVersion === undefined ||
		schemaId === undefined ||
		supported === undefined ||
		exact === undefined ||
		inverseOps === undefined ||
		issues === undefined
	) {
		throw new Error('journal is missing required v1 fields')
	}
	return {
		schemaVersion,
		schemaId,
		supported,
		exact,
		inverseOpCount: inverseOps.length,
		issueCount: issues.length,
		issues,
	}
}

let server: ReturnType<typeof createServer> | undefined

interface ApiEnvelope {
	readonly ok: boolean
	readonly data?: {
		readonly output?: string
		readonly backup?: string
		readonly outputSha256?: string
		readonly trust?: string
		readonly posture?: string
		readonly includedInAgentContext?: {
			readonly activeContent?: boolean
			readonly hiddenSheets?: boolean
		}
		readonly executionPolicy?: {
			readonly macros?: string
			readonly externalLinks?: string
		}
		readonly findings?: readonly { readonly code?: string; readonly nextAction?: string }[]
		readonly nextActions?: readonly string[]
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
			readonly changedRanges?: readonly { readonly sheet?: string; readonly range?: string }[]
			readonly wouldSucceed?: boolean
			readonly journalSummary?: {
				readonly schemaVersion?: number
				readonly schemaId?: string
				readonly supported?: boolean
				readonly exact?: boolean
				readonly inverseOpCount?: number
				readonly issueCount?: number
				readonly issues?: unknown[]
			}
		}
		readonly preparedPlan?: {
			readonly id?: string
			readonly file?: string
			readonly inputSha256?: string
			readonly planDigest?: string
			readonly operationCount?: number
			readonly expiresAt?: string
			readonly ttlMs?: number
		}
		readonly apply?: {
			readonly affectedCellCount?: number
			readonly emittedAffectedCellCount?: number
			readonly affectedCellRefs?: readonly string[]
			readonly affectedRanges?: readonly { readonly sheet?: string; readonly range?: string }[]
			readonly journalSummary?: {
				readonly schemaVersion?: number
				readonly schemaId?: string
				readonly supported?: boolean
				readonly exact?: boolean
				readonly inverseOpCount?: number
				readonly issueCount?: number
				readonly issues?: unknown[]
			}
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
				readonly emittedIssueCount?: number
				readonly issues?: readonly {
					readonly code?: string
					readonly partPath?: string
					readonly preservationPolicy?: string
					readonly preservationMode?: string
				}[]
			}
			readonly opaquePayloads?: {
				readonly generatedWithOpaquePayloads?: number
				readonly x14ConditionalFormatExtensionPayloads?: number
				readonly x14DataValidationExtensionPayloads?: number
				readonly worksheetParts?: readonly string[]
				readonly preservationMode?: string
				readonly verification?: string
			}
			readonly comments?: {
				readonly legacyCommentLocations?: number
				readonly threadedCommentLocations?: number
				readonly legacyDrawingLocations?: number
				readonly locations?: readonly string[]
				readonly threadedCommentPartPaths?: readonly string[]
				readonly verification?: string
			}
			readonly tables?: {
				readonly tableLocations?: number
				readonly queryTableLocations?: number
				readonly tableAutoFilterLocations?: number
				readonly tableNames?: readonly string[]
				readonly locations?: readonly string[]
				readonly tablePartPaths?: readonly string[]
				readonly queryTablePartPaths?: readonly string[]
				readonly preservationMode?: string
				readonly verification?: string
			}
			readonly definedNames?: {
				readonly total?: number
				readonly workbookScoped?: number
				readonly sheetScoped?: number
				readonly hidden?: number
				readonly names?: readonly {
					readonly name?: string
					readonly formula?: string
					readonly scope?: string
					readonly sheet?: string
					readonly hidden?: boolean
				}[]
				readonly verification?: string
			}
			readonly externalReferences?: {
				readonly total?: number
				readonly boundByExternalBookRelId?: number
				readonly fallbackPathRelationships?: number
				readonly missingPathRelationships?: number
				readonly partPaths?: readonly string[]
				readonly targets?: readonly string[]
				readonly parts?: readonly {
					readonly partPath?: string
					readonly relId?: string
					readonly externalBookRelId?: string
					readonly linkRelId?: string
					readonly linkBindingStatus?: string
					readonly target?: string
					readonly targetMode?: string
				}[]
				readonly preservationMode?: string
				readonly verification?: string
			}
			readonly analytics?: {
				readonly pivotCaches?: number
				readonly pivotTables?: number
				readonly slicerCaches?: number
				readonly slicers?: number
				readonly timelineCaches?: number
				readonly timelines?: number
				readonly partPaths?: readonly string[]
				readonly requiresExternalRefresh?: boolean
				readonly preservationMode?: string
				readonly verification?: string
				readonly pivotCacheDetails?: readonly {
					readonly partPath?: string
					readonly cacheId?: number
					readonly sourceSheet?: string
					readonly sourceRef?: string
					readonly outputState?: string
					readonly requiresExternalRefresh?: boolean
				}[]
			}
			readonly activeContent?: {
				readonly total?: number
				readonly vbaProjects?: number
				readonly activeXControls?: number
				readonly formControls?: number
				readonly macroSheets?: number
				readonly vbaSignatures?: number
				readonly digitalSignatures?: number
				readonly customUi?: number
				readonly unknownActiveContent?: number
				readonly partPaths?: readonly string[]
				readonly executionPolicy?: string
				readonly preservationMode?: string
				readonly verification?: string
				readonly entries?: readonly {
					readonly kind?: string
					readonly partPath?: string
					readonly contentType?: string
					readonly anchor?: string
					readonly opaque?: boolean
					readonly executionPolicy?: string
				}[]
			}
			readonly visuals?: {
				readonly sheetsWithVisuals?: number
				readonly images?: number
				readonly drawingObjects?: number
				readonly drawingMlObjects?: number
				readonly vmlObjects?: number
				readonly chartParts?: number
				readonly chartSheets?: number
				readonly drawingPartPaths?: readonly string[]
				readonly mediaPartPaths?: readonly string[]
				readonly chartPartPaths?: readonly string[]
				readonly vmlPartPaths?: readonly string[]
				readonly preservationMode?: string
				readonly verification?: string
				readonly sheets?: readonly {
					readonly sheetName?: string
					readonly hasDrawingMl?: boolean
					readonly hasVml?: boolean
					readonly imageCount?: number
					readonly drawingPartPaths?: readonly string[]
					readonly mediaPartPaths?: readonly string[]
				}[]
			}
			readonly security?: {
				readonly workbookProtected?: boolean
				readonly workbookLocks?: readonly string[]
				readonly workbookPasswordProtected?: boolean
				readonly workbookRevisionPasswordProtected?: boolean
				readonly protectedSheets?: number
				readonly protectedSheetNames?: readonly string[]
				readonly sheetPasswordProtected?: number
				readonly sheetStrongHashProtected?: number
				readonly protectedRanges?: number
				readonly protectedRangeLocations?: readonly string[]
				readonly passwordHashVerification?: string
				readonly preservationMode?: string
				readonly verification?: string
				readonly sheets?: readonly {
					readonly sheetName?: string
					readonly protected?: boolean
					readonly passwordProtected?: boolean
					readonly strongHashProtected?: boolean
					readonly allowedActions?: readonly string[]
					readonly protectedRanges?: number
					readonly protectedRangeLocations?: readonly string[]
				}[]
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
		readonly writePolicy?: {
			readonly diagnostics?: readonly {
				readonly code?: string
				readonly featureFamily?: string
				readonly preservationMode?: string
				readonly packageParts?: readonly { readonly preservationMode?: string }[]
			}[]
			readonly summary?: {
				readonly calcChainPolicy?: string
				readonly preservationModes?: {
					readonly preserveExactParts?: number
					readonly generatedParts?: number
					readonly generatedWithOpaquePayloads?: number
					readonly invalidatedOnEditParts?: number
					readonly discardedForRecalcParts?: number
					readonly inspectOnlyParts?: number
					readonly reviewRequiredParts?: number
					readonly unsupportedFeatures?: number
					readonly lossyApprovalRequiredFeatures?: number
				}
			}
		}
		readonly journal?: {
			readonly schemaVersion?: number
			readonly schemaId?: string
			readonly supported?: boolean
			readonly exact?: boolean
			readonly inverseOps?: unknown[]
			readonly issues?: unknown[]
		}
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
		readonly retryable?: boolean
		readonly retryStrategy?: string
		readonly suggestedFix?: string
		readonly details?: {
			readonly file?: string
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
	test('/open-plan rejects missing workbook references with structured retry guidance', async () => {
		const result = await postJson('/open-plan', { intent: 'edit-plan' })

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error).toMatchObject({
			code: 'VALIDATION_ERROR',
			message: 'Missing or invalid open-plan workbook reference',
			retryable: true,
			retryStrategy: 'modified',
			details: { required: ['file'] },
			suggestedFix: expect.stringContaining('Pass file so Ascend can inspect workbook risks'),
		})
	})

	test('/inspect rejects missing workbook references with structured retry guidance', async () => {
		const result = await postJson('/inspect', { sheet: 'Sheet1' })

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error).toMatchObject({
			code: 'VALIDATION_ERROR',
			message: 'Missing or invalid inspect workbook reference',
			retryable: true,
			retryStrategy: 'modified',
			details: { required: ['file'] },
			suggestedFix: expect.stringContaining('Pass file so Ascend can inspect workbook structure'),
		})
	})

	test('/active-content rejects missing workbook references with structured retry guidance', async () => {
		const result = await postJson('/active-content', {})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error).toMatchObject({
			code: 'VALIDATION_ERROR',
			message: 'Missing or invalid active-content workbook reference',
			retryable: true,
			retryStrategy: 'modified',
			details: { required: ['file'] },
			suggestedFix: expect.stringContaining('Pass file so Ascend can inspect active content'),
		})
	})

	test('/trust-report rejects missing workbook references with structured retry guidance', async () => {
		const result = await postJson('/trust-report', {})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error).toMatchObject({
			code: 'VALIDATION_ERROR',
			message: 'Missing or invalid trust-report workbook reference',
			retryable: true,
			retryStrategy: 'modified',
			details: { required: ['file'] },
			suggestedFix: expect.stringContaining('Pass file so Ascend can build a trust report'),
		})
	})

	test('/read rejects missing workbook references with structured retry guidance', async () => {
		const result = await postJson('/read', { range: 'A1:A1' })

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error).toMatchObject({
			code: 'VALIDATION_ERROR',
			message: 'Missing or invalid read workbook reference',
			retryable: true,
			retryStrategy: 'modified',
			details: { required: ['file'] },
			suggestedFix: expect.stringContaining('Pass file so Ascend can read'),
		})
	})

	test('/agent-view rejects missing workbook references with structured retry guidance', async () => {
		const result = await postJson('/agent-view', { range: 'A1:A1' })

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error).toMatchObject({
			code: 'VALIDATION_ERROR',
			message: 'Missing or invalid agent-view workbook reference',
			retryable: true,
			retryStrategy: 'modified',
			details: { required: ['file'] },
			suggestedFix: expect.stringContaining('Pass file so Ascend can build a bounded agent view'),
		})
	})

	test('/plan rejects missing workbook references with structured retry guidance', async () => {
		const result = await postJson('/plan', {
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
		})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error).toMatchObject({
			code: 'VALIDATION_ERROR',
			message: 'Missing or invalid plan workbook reference',
			retryable: true,
			retryStrategy: 'modified',
			details: { required: ['file'] },
			suggestedFix: expect.stringContaining('Pass file with ops or mutations'),
		})
	})

	test('/commit rejects missing workbook references with structured retry guidance', async () => {
		const result = await postJson('/commit', {
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
			output: 'missing-reference-output.xlsx',
		})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error).toMatchObject({
			code: 'VALIDATION_ERROR',
			message: 'Missing or invalid commit workbook reference',
			retryable: true,
			retryStrategy: 'modified',
			details: { required: ['file or planHandle'] },
			suggestedFix: expect.stringContaining('Pass either file with ops/mutations'),
		})
	})

	test('/plan reports missing workbook files with structured retry guidance', async () => {
		const missing = join(tmpdir(), `ascend-api-missing-${Date.now()}.xlsx`)
		const result = await postJson('/plan', {
			file: missing,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
		})

		expect(result.status).toBe(404)
		expect(result.body.ok).toBe(false)
		expect(result.body.error).toMatchObject({
			code: 'FILE_NOT_FOUND',
			retryable: true,
			retryStrategy: 'modified',
			details: { file: missing },
			suggestedFix: expect.stringContaining('existing workbook path'),
		})
	})

	test('/commit reports missing workbook files without creating output artifacts', async () => {
		const missing = join(tmpdir(), `ascend-api-missing-commit-${Date.now()}.xlsx`)
		const output = `${missing}.out.xlsx`
		try {
			const result = await postJson('/commit', {
				file: missing,
				output,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
			})

			expect(result.status).toBe(404)
			expect(result.body.ok).toBe(false)
			expect(result.body.error).toMatchObject({
				code: 'FILE_NOT_FOUND',
				retryable: true,
				retryStrategy: 'modified',
				details: { file: missing },
				suggestedFix: expect.stringContaining('existing workbook path'),
			})
			expect(await Bun.file(output).exists()).toBe(false)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('/plan accepts encrypted workbook passwords and commit fails closed before decrypted export', async () => {
		const input = join(
			tmpdir(),
			`ascend-api-encrypted-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		const output = `${input}.out.xlsx`
		try {
			await Bun.write(input, readFileSync(ENCRYPTED_FIXTURE))
			const plan = await postJson('/plan', {
				file: input,
				password: '123',
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'Z10', value: 'blocked' }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.preparedPlan?.id).toBeString()
			expect(JSON.stringify(plan.body)).not.toContain('"123"')

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				compact: true,
			})
			expect(commit.status).toBe(409)
			expect(commit.body.ok).toBe(false)
			expect(commit.body.error).toMatchObject({
				code: 'EXPORT_ERROR',
				retryable: false,
				details: {
					sourceWasEncrypted: true,
					reEncryptionSupported: false,
					requestedExport: 'xlsx',
				},
			})
			expect(commit.body.error?.message).toContain(
				'Cannot export an edited encrypted workbook without re-encryption support',
			)
			expect(JSON.stringify(commit.body)).not.toContain('"123"')
			expect(await Bun.file(output).exists()).toBe(false)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('/plan accepts encrypted workbook passwords for path mutations and commit fails closed without source drift', async () => {
		const input = join(
			tmpdir(),
			`ascend-api-encrypted-path-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		const output = `${input}.out.xlsx`
		const sourceBytes = readFileSync(ENCRYPTED_FIXTURE)
		try {
			await Bun.write(input, sourceBytes)
			const plan = await postJson('/plan', {
				file: input,
				password: '123',
				mutations: [{ path: '/sheets/Sheet1/cells/Z10/value', value: 'blocked' }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.pathMutations).toMatchObject({
				replayable: true,
				ops: [
					{
						op: 'setCells',
						sheet: 'Sheet1',
						updates: [{ ref: 'Z10', value: 'blocked' }],
					},
				],
			})
			expect(plan.body.data?.preparedPlan?.id).toBeString()
			expect(JSON.stringify(plan.body)).not.toContain('"123"')

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				compact: true,
			})
			expect(commit.status).toBe(409)
			expect(commit.body.ok).toBe(false)
			expect(commit.body.error).toMatchObject({
				code: 'EXPORT_ERROR',
				retryable: false,
				details: {
					sourceWasEncrypted: true,
					reEncryptionSupported: false,
					requestedExport: 'xlsx',
				},
			})
			expect(commit.body.error?.message).toContain(
				'Cannot export an edited encrypted workbook without re-encryption support',
			)
			expect(JSON.stringify(commit.body)).not.toContain('"123"')
			expect(await Bun.file(output).exists()).toBe(false)
			expect(Buffer.from(readFileSync(input)).equals(sourceBytes)).toBe(true)
			await expect(
				AscendWorkbook.open(new Uint8Array(readFileSync(input)), { password: '123' }),
			).resolves.toBeDefined()
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('/trust-report exposes untrusted workbook boundaries and next actions', async () => {
		const trustFile = join(
			tmpdir(),
			`ascend-api-trust-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
		)
		await Bun.write(trustFile, signedMacroWorkbook())
		try {
			const result = await postJson('/trust-report', { file: trustFile, maxFindings: 10 })

			expect(result.status).toBe(200)
			expect(result.body.ok).toBe(true)
			expect(result.body.data?.trust).toBe('untrusted')
			expect(result.body.data?.posture).toBe('safe-parser-preserver')
			expect(result.body.data?.includedInAgentContext).toMatchObject({
				activeContent: false,
				hiddenSheets: false,
			})
			expect(result.body.data?.executionPolicy).toMatchObject({
				macros: 'preserve-only',
				externalLinks: 'do-not-refresh',
			})
			expect(result.body.data?.findings).toContainEqual(
				expect.objectContaining({ code: 'workbook.vbaProject' }),
			)
			expect(result.body.data?.nextActions).toContain(
				'Use visible workbook data as the default agent context; opt into hidden sheets, comments, names, and metadata only when the task requires them.',
			)
		} finally {
			await unlink(trustFile).catch(() => {})
		}
	})

	test('/active-content reports custom UI callbacks as active content for agents', async () => {
		const customUiFile = join(
			tmpdir(),
			`ascend-api-custom-ui-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
		)
		await Bun.write(customUiFile, customUiWorkbook())
		try {
			const result = await postJson('/active-content', { file: customUiFile })
			const data = result.body.data as
				| {
						activeContentCount?: number
						activeContent?: Array<{
							kind?: string
							partPath?: string
							contentType?: string
							relType?: string
							executionPolicy?: string
							customUi?: {
								callbacks?: readonly { readonly attribute?: string; readonly macro?: string }[]
							}
						}>
						compatibilityFeatures?: Array<{ feature?: string; locations?: readonly string[] }>
						capabilityWarnings?: Array<{ family?: string; message?: string }>
				  }
				| undefined

			expect(result.status).toBe(200)
			expect(result.body.ok).toBe(true)
			expect(data?.activeContentCount).toBe(1)
			const customUiEntry = data?.activeContent?.find(
				(entry) => entry.kind === 'customUi' && entry.partPath === 'customUI/customUI2.xml',
			)
			expect(customUiEntry).toMatchObject({
				contentType: 'application/vnd.ms-office.customUI+xml',
				relType: 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility',
				executionPolicy: 'blocked',
			})
			expect(customUiEntry?.customUi?.callbacks).toEqual(
				expect.arrayContaining([
					{ attribute: 'onLoad', macro: 'Ribbon.OnLoad' },
					{ attribute: 'loadImage', macro: 'Ribbon.LoadImage' },
					{ attribute: 'onAction', macro: 'Module1.RunReport' },
					{ attribute: 'getEnabled', macro: 'Ribbon.CanRun' },
				]),
			)
			expect(data?.compatibilityFeatures).toContainEqual(
				expect.objectContaining({
					feature: 'preservedCustomUi',
					locations: ['customUI/customUI2.xml'],
				}),
			)
			expect(data?.capabilityWarnings).toContainEqual(
				expect.objectContaining({
					family: 'active content',
					capabilityId: 'active.custom-ui',
					reason: expect.stringContaining('executionPolicy=blocked'),
				}),
			)
		} finally {
			await unlink(customUiFile).catch(() => {})
		}
	})

	test('formula-assist exposes diagnostics, completions, signature help, and reference edits', async () => {
		const result = await postJson('/formula-assist', {
			formula: '=SUM(A1:B2',
			cursor: 8,
			prefix: 'SU',
			completionLimit: 3,
			functionName: 'SUM',
			reference: 'C1',
			replaceReferenceAtCursor: true,
			cycleReference: true,
		})
		const data = result.body.data as {
			diagnostics?: { parseOk?: boolean; diagnostics?: Array<{ code?: string }> }
			tokens?: Array<{ text?: string; className?: string }>
			activeReference?: { text?: string; kind?: string }
			completions?: Array<{ name?: string }>
			signature?: { name?: string }
			signatureHelp?: { signature?: { name?: string }; activeParameter?: number }
			cycle?: { formula?: string; changed?: boolean }
			insertion?: { formula?: string; replaced?: { text?: string } }
			renameTarget?: { ok?: boolean; reason?: string; role?: { role?: string; text?: string } }
		}

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(data.diagnostics?.parseOk).toBe(false)
		expect(
			data.diagnostics?.diagnostics?.some((issue) => issue.code === 'formula-parse-error'),
		).toBe(true)
		expect(
			data.tokens?.some((token) => token.text === 'SUM' && token.className === 'function'),
		).toBe(true)
		expect(data.activeReference).toMatchObject({ text: 'A1:B2', kind: 'range' })
		expect(data.completions?.some((completion) => completion.name === 'SUM')).toBe(true)
		expect(data.signature?.name).toBe('SUM')
		expect(data.signatureHelp?.signature?.name).toBe('SUM')
		expect(data.cycle).toMatchObject({ formula: '=SUM(A1:$B$2', changed: true })
		expect(data.insertion).toMatchObject({ formula: '=SUM(C1', replaced: { text: 'A1:B2' } })

		const refusal = await postJson('/formula-assist', {
			formula: '=Budget+Sales[Amount]',
			cursor: 10,
		})
		const refusalData = refusal.body.data as {
			renameTarget?: { ok?: boolean; reason?: string; role?: { role?: string; text?: string } }
		}
		expect(refusal.status).toBe(200)
		expect(refusalData.renameTarget).toMatchObject({
			ok: false,
			reason: 'workbook-context-required',
			role: { role: 'table-name-use', text: 'Sales' },
		})
	})

	test('/plan accepts encrypted workbook passwords without echoing them', async () => {
		const result = await postJson('/plan', {
			file: ENCRYPTED_FIXTURE,
			password: '123',
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'planned' }] }],
		})
		const serialized = JSON.stringify(result.body)

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.preview?.wouldSucceed).toBe(true)
		expect(result.body.data?.preparedPlan?.id).toBeString()
		expect(serialized).not.toContain('"123"')
	})

	test('/plan rejects non-string encrypted workbook passwords with retry guidance', async () => {
		const result = await postJson('/plan', {
			file: ENCRYPTED_FIXTURE,
			password: 123,
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'planned' }] }],
		})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error).toMatchObject({
			code: 'VALIDATION_ERROR',
			message: 'Invalid plan password',
			details: { field: 'password', receivedType: 'number' },
			suggestedFix: 'Pass password as a string or omit it.',
		})
	})

	test('/commit rejects non-string encrypted workbook passwords before direct commit', async () => {
		const input = join(
			tmpdir(),
			`ascend-api-invalid-password-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		const output = `${input}.out.xlsx`
		try {
			const wb = AscendWorkbook.create()
			await wb.save(input)
			const result = await postJson('/commit', {
				file: input,
				password: 123,
				output,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'planned' }] }],
			})

			expect(result.status).toBe(400)
			expect(result.body.ok).toBe(false)
			expect(result.body.error).toMatchObject({
				code: 'VALIDATION_ERROR',
				message: 'Invalid commit password',
				details: { field: 'password', receivedType: 'number' },
				suggestedFix: 'Pass password as a string or omit it.',
			})
			expect(await Bun.file(output).exists()).toBe(false)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

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
			expect(commit.body.data?.outputSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(commit.body.data?.postWrite?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.unresolvedPackageGraphIssueCount).toBe(0)
			expect(commit.body.data?.postWrite?.reopened).toBe(true)
			expect(commit.body.data?.postWrite?.outputSha256).toBe(commit.body.data?.outputSha256)
			expect(commit.body.data?.postWrite?.check?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.packageGraphAudit?.ok).toBe(true)

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
			expect(commit.body.data?.outputSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(commit.body.data?.postWrite?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.unresolvedPackageGraphIssueCount).toBe(0)
			expect(commit.body.data?.postWrite?.reopened).toBe(true)
			expect(commit.body.data?.postWrite?.outputSha256).toBe(commit.body.data?.outputSha256)
			expect(commit.body.data?.postWrite?.check?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.packageGraphAudit?.ok).toBe(true)

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

	test('dump and template-merge reject capped load options instead of emitting replay ops', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'A1', value: '{{name}}' }],
			},
		])
		await wb.save(TEMP_FILE)

		const dump = await postJson('/dump', {
			file: TEMP_FILE,
			maxRows: 1,
		})
		expect(dump.status).toBe(400)
		expect(dump.body.ok).toBe(false)
		expect(dump.body.data?.ops).toBeUndefined()
		expect(dump.body.error?.code).toBe('VALIDATION_ERROR')
		expect(dump.body.error?.details?.unsupportedLoadOptions).toEqual(['maxRows'])
		expect(dump.body.error?.details?.requiredLoad).toEqual({
			mode: 'full',
			allSheets: true,
			maxRows: null,
		})

		const merge = await postJson('/template-merge', {
			file: TEMP_FILE,
			data: { name: 'Acme' },
			maxRows: 1,
			mode: 'values',
			sheets: ['Sheet1'],
		})
		expect(merge.status).toBe(400)
		expect(merge.body.ok).toBe(false)
		expect(merge.body.data?.ops).toBeUndefined()
		expect(merge.body.error?.code).toBe('VALIDATION_ERROR')
		expect(merge.body.error?.details?.unsupportedLoadOptions).toEqual(['maxRows', 'mode', 'sheets'])
		expect(merge.body.error?.details?.requiredLoad).toEqual({
			mode: 'full',
			allSheets: true,
			maxRows: null,
		})
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

	test('read exposes array and shared formula binding metadata', async () => {
		await Bun.write(TEMP_FILE, sharedFormulaWorkbook())

		const result = await postJson('/read', {
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'A1:D2',
			format: 'cells',
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		const cells = result.body.data?.cells as
			| Array<{
					readonly ref?: string
					readonly formula?: string
					readonly formulaBinding?: unknown
			  }>
			| undefined
		expect(cells?.map((cell) => [cell.ref, cell.formula, cell.formulaBinding ?? null])).toEqual([
			['A1', 'SUM(B1:B2)', { kind: 'array', ref: 'A1:A2' }],
			['B1', 'A1*2', { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'B1' }],
			['C1', 'SUM(Sales[[Revenue]:[Quantity]])', null],
			['D1', 'BudgetTotal*2', null],
			['B2', 'A2*2', { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'B1' }],
		])

		const inspect = await postJson('/inspect', { file: TEMP_FILE })
		expect(inspect.status).toBe(200)
		const definedNameDetails = (
			inspect.body.data as
				| {
						definedNameDetails?: readonly {
							readonly name?: string
							readonly formula?: string
							readonly normalizedFormula?: string
							readonly scope?: string
							readonly refs?: readonly string[]
						}[]
				  }
				| undefined
		)?.definedNameDetails
		expect(definedNameDetails).toContainEqual({
			name: 'BudgetTotal',
			formula: 'Calc!$A$1:$A$2',
			normalizedFormula: 'Calc!$A$1:$A$2',
			scope: 'workbook',
			references: [
				{ kind: 'range', text: 'Calc!$A$1:$A$2', scope: { kind: 'sheet', sheet: 'Calc' } },
			],
			refs: ['Calc!$A$1:$A$2'],
			functions: [],
			volatile: false,
		})
	})

	test('read exposes dynamic-array formulas and binding metadata', async () => {
		await Bun.write(TEMP_FILE, dynamicArrayWorkbook())

		const result = await postJson('/read', {
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'A1:C1',
			format: 'cells',
			display: true,
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		const cells = result.body.data?.cells as
			| Array<{
					readonly ref?: string
					readonly value?: string
					readonly formula?: string
					readonly formulaBinding?: unknown
			  }>
			| undefined
		expect(
			cells?.map((cell) => [cell.ref, cell.value, cell.formula, cell.formulaBinding ?? null]),
		).toEqual([
			['A1', '1', 'SEQUENCE(3)', { kind: 'dynamicArray', metadataIndex: 1, collapsed: false }],
			['B1', '6', 'SUM(A1#)', null],
			['C1', '1', '@A1', null],
		])
	})

	test('write blocks structural edits against imported formula bindings without changing read truth', async () => {
		await Bun.write(TEMP_FILE, sharedOnlyFormulaWorkbook())

		const sharedWrite = await postJson('/write', {
			file: TEMP_FILE,
			ops: [{ op: 'insertRows', sheet: 'Calc', at: 0, count: 1 }],
		})

		expect(sharedWrite.status).toBe(400)
		expect(sharedWrite.body.ok).toBe(false)
		expect(sharedWrite.body.error?.code).toBe('VALIDATION_ERROR')
		expect(sharedWrite.body.error?.message).toContain(
			'Calc!B1 contains imported shared formula metadata',
		)
		expect(sharedWrite.body.error?.refs).toEqual(['Calc!B1'])
		expect(sharedWrite.body.error?.suggestedFix).toContain('Materialize or rewrite')
		const sharedRead = await postJson('/read', {
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'B1:B2',
			format: 'cells',
		})
		expect(
			(
				sharedRead.body.data?.cells as
					| Array<{
							readonly ref?: string
							readonly formula?: string
							readonly formulaBinding?: unknown
					  }>
					| undefined
			)?.map((cell) => [cell.ref, cell.formula, cell.formulaBinding ?? null]),
		).toEqual([
			['B1', 'A1*2', { kind: 'shared', sharedIndex: '0', isMaster: true, masterRef: 'B1' }],
			['B2', 'A2*2', { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'B1' }],
		])

		await Bun.write(TEMP_FILE, dynamicArrayWorkbook())
		const dynamicWrite = await postJson('/write', {
			file: TEMP_FILE,
			ops: [{ op: 'insertCols', sheet: 'Calc', at: 0, count: 1 }],
		})

		expect(dynamicWrite.status).toBe(400)
		expect(dynamicWrite.body.ok).toBe(false)
		expect(dynamicWrite.body.error?.code).toBe('VALIDATION_ERROR')
		expect(dynamicWrite.body.error?.message).toContain(
			'Calc!A1 contains imported dynamicArray formula metadata',
		)
		expect(dynamicWrite.body.error?.refs).toEqual(['Calc!A1'])
		expect(dynamicWrite.body.error?.suggestedFix).toContain('Materialize or rewrite')
		const dynamicRead = await postJson('/read', {
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'A1:C1',
			format: 'cells',
			display: true,
		})
		expect(
			(
				dynamicRead.body.data?.cells as
					| Array<{
							readonly ref?: string
							readonly value?: string
							readonly formula?: string
							readonly formulaBinding?: unknown
					  }>
					| undefined
			)?.map((cell) => [cell.ref, cell.value, cell.formula, cell.formulaBinding ?? null]),
		).toEqual([
			['A1', '1', 'SEQUENCE(3)', { kind: 'dynamicArray', metadataIndex: 1, collapsed: false }],
			['B1', '6', 'SUM(A1#)', null],
			['C1', '1', '@A1', null],
		])
	})

	test('write can explicitly rewrite imported formula bindings without stale metadata', async () => {
		await Bun.write(TEMP_FILE, sharedOnlyFormulaWorkbook())
		const sharedWrite = await postJson('/write', {
			file: TEMP_FILE,
			ops: [
				{ op: 'setFormula', sheet: 'Calc', ref: 'B1', formula: '1+1' },
				{ op: 'setFormula', sheet: 'Calc', ref: 'B2', formula: '2+2' },
			],
		})

		expect(sharedWrite.status).toBe(200)
		expect(sharedWrite.body.ok).toBe(true)
		const sharedRead = await postJson('/read', {
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'B1:B2',
			format: 'cells',
			display: true,
		})
		expect(
			(
				sharedRead.body.data?.cells as
					| Array<{
							readonly ref?: string
							readonly value?: string
							readonly formula?: string
							readonly formulaBinding?: unknown
					  }>
					| undefined
			)?.map((cell) => [cell.ref, cell.value, cell.formula, cell.formulaBinding ?? null]),
		).toEqual([
			['B1', '2', '1+1', null],
			['B2', '4', '2+2', null],
		])

		await Bun.write(TEMP_FILE, dynamicArrayWorkbook())
		const dynamicWrite = await postJson('/write', {
			file: TEMP_FILE,
			ops: [
				{ op: 'setFormula', sheet: 'Calc', ref: 'A1', formula: '1+1' },
				{ op: 'setFormula', sheet: 'Calc', ref: 'B1', formula: 'A1+1' },
				{ op: 'setFormula', sheet: 'Calc', ref: 'C1', formula: 'A1+2' },
			],
		})

		expect(dynamicWrite.status).toBe(200)
		expect(dynamicWrite.body.ok).toBe(true)
		const dynamicRead = await postJson('/read', {
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'A1:C1',
			format: 'cells',
			display: true,
		})
		expect(
			(
				dynamicRead.body.data?.cells as
					| Array<{
							readonly ref?: string
							readonly value?: string
							readonly formula?: string
							readonly formulaBinding?: unknown
					  }>
					| undefined
			)?.map((cell) => [cell.ref, cell.value, cell.formula, cell.formulaBinding ?? null]),
		).toEqual([
			['A1', '2', '1+1', null],
			['B1', '3', 'A1+1', null],
			['C1', '4', 'A1+2', null],
		])
	})

	test('write reports materialized shared formula group when rewriting one member', async () => {
		await Bun.write(TEMP_FILE, sharedOnlyFormulaWorkbook())

		const write = await postJson('/write', {
			file: TEMP_FILE,
			ops: [{ op: 'setFormula', sheet: 'Calc', ref: 'B2', formula: '2+2' }],
		})

		expect(write.status).toBe(200)
		expect(write.body.ok).toBe(true)
		expect(
			(write.body.data as { affectedCells?: readonly string[] } | undefined)?.affectedCells,
		).toEqual(['B1', 'B2'])
		const read = await postJson('/read', {
			file: TEMP_FILE,
			sheet: 'Calc',
			range: 'B1:B2',
			format: 'cells',
		})
		expect(
			(
				read.body.data?.cells as
					| Array<{
							readonly ref?: string
							readonly formula?: string
							readonly formulaBinding?: unknown
					  }>
					| undefined
			)?.map((cell) => [cell.ref, cell.formula, cell.formulaBinding ?? null]),
		).toEqual([
			['B1', 'A1*2', null],
			['B2', '2+2', null],
		])
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

	test('compact changedSince reads invalidate when the requested window changes', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 1 },
					{ ref: 'A2', value: 2 },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const first = await postJson('/read', {
			file: TEMP_FILE,
			range: 'A1:A2',
			format: 'compact',
			rowLimit: 1,
		})
		expect(first.status).toBe(200)
		expect(first.body.data?.cells).toEqual([[0, 0, 1]])
		expect(first.body.data?.changeToken).toBeDefined()

		const widened = await postJson('/read', {
			file: TEMP_FILE,
			range: 'A1:A2',
			format: 'compact',
			rowLimit: 2,
			changedSince: first.body.data?.changeToken,
		})
		expect(widened.status).toBe(200)
		expect(widened.body.ok).toBe(true)
		expect(widened.body.data?.cells).toEqual([
			[0, 0, 1],
			[1, 0, 2],
		])
		expect(widened.body.data?.changeInvalidation).toEqual({
			baseToken: first.body.data?.changeToken,
			changeToken: widened.body.data?.changeToken,
			reason: 'base-snapshot-missing',
			requiredAction: 'use-returned-window',
		})

		const invalid = await postJson('/read', {
			file: TEMP_FILE,
			range: 'A1:A2',
			format: 'compact',
			rowLimit: 2,
			changedSince: 'not-a-token',
		})
		expect(invalid.status).toBe(200)
		expect(invalid.body.ok).toBe(true)
		expect(invalid.body.data?.cells).toEqual([
			[0, 0, 1],
			[1, 0, 2],
		])
		expect(invalid.body.data?.changeInvalidation).toEqual({
			baseToken: 'not-a-token',
			changeToken: invalid.body.data?.changeToken,
			reason: 'base-token-invalid',
			requiredAction: 'use-returned-window',
		})
	})

	test('compact changedSince reads return a fresh window after source changes', async () => {
		const input = `${TEMP_FILE}.changed-source.xlsx`
		try {
			const original = AscendWorkbook.create()
			original.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'old' }] }])
			await original.save(input)

			const first = await postJson('/read', {
				file: input,
				range: 'A1:A1',
				format: 'compact',
			})
			expect(first.status).toBe(200)
			expect(first.body.data?.cells).toEqual([[0, 0, 'old']])
			expect(first.body.data?.changeToken).toBeDefined()

			const changed = AscendWorkbook.create()
			changed.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'new' }] }])
			await changed.save(input)

			const afterChange = await postJson('/read', {
				file: input,
				range: 'A1:A1',
				format: 'compact',
				changedSince: first.body.data?.changeToken,
			})
			expect(afterChange.status).toBe(200)
			expect(afterChange.body.ok).toBe(true)
			expect(afterChange.body.data?.cells).toEqual([[0, 0, 'new']])
			expect(afterChange.body.data?.changeToken).toBeDefined()
			expect(afterChange.body.data?.changeInvalidation).toEqual({
				baseToken: first.body.data?.changeToken,
				changeToken: afterChange.body.data?.changeToken,
				reason: 'base-snapshot-missing',
				requiredAction: 'use-returned-window',
			})
		} finally {
			await unlink(input).catch(() => {})
		}
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

	test('agent-view exposes token budget metadata', async () => {
		const wb = AscendWorkbook.create()
		const updates = []
		for (let row = 1; row <= 20; row++) {
			for (let col = 0; col < 4; col++) {
				updates.push({
					ref: `${String.fromCharCode(65 + col)}${row}`,
					value: row === 1 ? `Header ${col + 1}` : `r${row}-c${col + 1}`,
				})
			}
		}
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates }])
		await wb.save(TEMP_FILE)

		const result = await postJson('/agent-view', {
			file: TEMP_FILE,
			range: 'A1:D20',
			maxApproxTokens: 384,
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.budget?.requestedApproxTokens).toBe(384)
		expect(result.body.data?.budget?.truncated).toBe(true)
		expect(result.body.data?.rowCount).toBe(20)
		expect(result.body.data?.colCount).toBe(4)
		expect(
			(result.body.data?.budget?.omittedSampleRows ?? 0) +
				(result.body.data?.budget?.omittedColumnSampleValues ?? 0),
		).toBeGreaterThan(0)
		expect(result.body.data?.budget?.omittedEvidence?.sampleRows?.count).toBe(
			result.body.data?.budget?.omittedSampleRows,
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

	test('check exposes blocked spill diagnostics for public agent repair', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=SEQUENCE(3)' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 'blocker' }] },
		])
		wb.recalc()
		await wb.save(TEMP_FILE)

		const result = await postJson('/check', { file: TEMP_FILE })

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.valid).toBe(false)
		const issue = (
			result.body.data as
				| {
						issues?: Array<{
							rule?: string
							ref?: string
							refs?: string[]
							details?: unknown
						}>
				  }
				| undefined
		)?.issues?.find((entry) => entry.rule === 'spill-diagnostics')
		expect(issue?.ref).toBe('Sheet1!A1')
		expect(issue?.refs).toEqual(['Sheet1!A1', 'Sheet1!A2'])
		expect(issue?.details).toEqual({
			error: '#SPILL!',
			cause: 'occupied-cell',
			spillRange: 'Sheet1!A1:A3',
			blockingRefs: ['Sheet1!A2'],
		})
	})

	test('check refreshes stale imported spill caches for public agent repair', async () => {
		await Bun.write(TEMP_FILE, staleSpillCacheWorkbook())

		const read = await postJson('/read', {
			file: TEMP_FILE,
			range: 'A1:A2',
			format: 'cells',
		})
		expect(read.status).toBe(200)
		expect(
			(
				read.body.data?.cells as
					| Array<{
							readonly ref?: string
							readonly formula?: string
							readonly formulaBinding?: unknown
					  }>
					| undefined
			)?.map((cell) => [cell.ref, cell.formula ?? null, cell.formulaBinding ?? null]),
		).toEqual([
			['A1', 'SEQUENCE(3)', null],
			['A2', null, null],
		])

		const result = await postJson('/check', { file: TEMP_FILE })

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		const issue = (
			result.body.data as
				| {
						issues?: Array<{
							rule?: string
							refs?: string[]
							details?: unknown
						}>
				  }
				| undefined
		)?.issues?.find((entry) => entry.rule === 'spill-diagnostics')
		expect(issue?.refs).toEqual(['Sheet1!A1', 'Sheet1!A2'])
		expect(issue?.details).toEqual({
			error: '#SPILL!',
			cause: 'occupied-cell',
			spillRange: 'Sheet1!A1:A3',
			blockingRefs: ['Sheet1!A2'],
		})
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

	test('preview and write return exact empty journals for no-op requests', async () => {
		const wb = AscendWorkbook.create()
		const file = join(
			tmpdir(),
			`ascend-api-noop-journal-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		await wb.save(file)

		const expectedJournal = {
			schemaVersion: MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
			schemaId: MUTATION_JOURNAL_ISSUE_SCHEMA.$id,
			supported: true,
			exact: true,
			entries: [],
			inverseOps: [],
			issues: [],
			undoPolicy: {
				undoable: true,
				exact: true,
				riskLevel: 'none',
				reason: 'exact',
				userMessage: 'Undo available.',
			},
		}
		try {
			const preview = await postJson('/preview', {
				file,
				journal: true,
				ops: [],
			})
			const write = await postJson('/write', {
				file,
				journal: true,
				ops: [],
			})
			const previewMutations = await postJson('/preview', {
				file,
				journal: true,
				mutations: [],
			})
			const writeMutations = await postJson('/write', {
				file,
				journal: true,
				mutations: [],
			})

			expect(preview.status).toBe(200)
			expect(preview.body.ok).toBe(true)
			expect(preview.body.data?.journal).toEqual(expectedJournal)
			expect(write.status).toBe(200)
			expect(write.body.ok).toBe(true)
			expect(write.body.data?.journal).toEqual(expectedJournal)
			expect(previewMutations.status).toBe(200)
			expect(previewMutations.body.ok).toBe(true)
			expect(previewMutations.body.data?.journal).toEqual(expectedJournal)
			expect(previewMutations.body.data?.pathMutations).toMatchObject({
				mutationCount: 0,
				issueCount: 0,
				issues: [],
				replayable: true,
			})
			expect(writeMutations.status).toBe(200)
			expect(writeMutations.body.ok).toBe(true)
			expect(writeMutations.body.data?.journal).toEqual(expectedJournal)
			expect(writeMutations.body.data?.pathMutations).toMatchObject({
				mutationCount: 0,
				issueCount: 0,
				issues: [],
				replayable: true,
			})
		} finally {
			await unlink(file).catch(() => {})
		}
	})

	test('preview preserves lossy journal issue metadata for agents', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const result = await postJson('/preview', {
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'groupRows', sheet: 'Sheet1', from: 1, to: 2, collapsed: true },
				{
					op: 'groupCols',
					sheet: 'Sheet1',
					from: 0,
					to: 1,
					collapsed: true,
					summaryRight: false,
				},
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.schemaVersion).toBe(MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION)
		expect(result.body.data?.journal?.schemaId).toBe(MUTATION_JOURNAL_ISSUE_SCHEMA.$id)
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.exact).toBe(false)
		expect(result.body.data?.journal?.inverseOps).toEqual([])
		expect(result.body.data?.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Grouped rows for Sheet1 cannot be restored with public operations',
				surface: 'row-layout',
				reason: 'row-layout-created',
				refs: [
					'Sheet1!2',
					'Sheet1!3',
					'Sheet1!4',
					'sheet:Sheet1:outlinePr:summaryBelow',
					'sheet:Sheet1:sheetFormatPr:outlineLevelRow',
				],
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Grouped columns for Sheet1 cannot be restored with public operations',
				surface: 'column-layout',
				reason: 'column-layout-created',
				refs: [
					'Sheet1!A',
					'Sheet1!B',
					'sheet:Sheet1:outlinePr:summaryRight',
					'sheet:Sheet1:sheetFormatPr:outlineLevelCol',
				],
			},
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.getWorkbookModel().getSheet('Sheet1')?.rowDefs.size).toBe(0)
		expect(reopened.getWorkbookModel().getSheet('Sheet1')?.colDefs).toEqual([])
	})

	test('preview and write preserve the public journal v1 golden issue payload', async () => {
		const file = join(
			tmpdir(),
			`ascend-api-journal-v1-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		const writeFilePath = `${file}.write.xlsx`
		const previewWorkbook = AscendWorkbook.create()
		const writeWorkbook = AscendWorkbook.create()
		await previewWorkbook.save(file)
		await writeWorkbook.save(writeFilePath)
		try {
			const preview = await postJson('/preview', {
				file,
				journal: true,
				ops: JOURNAL_V1_FIXTURE.scenario.ops,
			})
			const write = await postJson('/write', {
				file: writeFilePath,
				journal: true,
				ops: JOURNAL_V1_FIXTURE.scenario.ops,
			})

			expect(preview.status).toBe(200)
			expect(preview.body.ok).toBe(true)
			expect(compactJournal(preview.body.data?.journal ?? {})).toEqual(
				JOURNAL_V1_FIXTURE.scenario.journal,
			)
			expect(write.status).toBe(200)
			expect(write.body.ok).toBe(true)
			expect(compactJournal(write.body.data?.journal ?? {})).toEqual(
				JOURNAL_V1_FIXTURE.scenario.journal,
			)
		} finally {
			await unlink(file).catch(() => {})
			await unlink(writeFilePath).catch(() => {})
		}
	})

	test('preview marks new theme additions lossy with public inverse metadata', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const result = await postJson('/preview', {
			file: TEMP_FILE,
			journal: true,
			ops: [
				{
					op: 'setTheme',
					themeName: 'New Brand',
					themeColors: [{ slot: 'accent1', rgb: '123456' }],
				},
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.schemaVersion).toBe(MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION)
		expect(result.body.data?.journal?.schemaId).toBe(MUTATION_JOURNAL_ISSUE_SCHEMA.$id)
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.exact).toBe(false)
		expect(result.body.data?.journal?.inverseOps).toEqual([])
		expect(result.body.data?.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Theme metadata field themeName cannot be removed with public operations',
				surface: 'package-parts',
				reason: 'package-part-preservation',
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Theme color slot accent1 cannot be removed with public operations',
				surface: 'package-parts',
				reason: 'package-part-preservation',
			},
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.inspect().themeSummary.hasThemePart).toBe(false)
	})

	test('preview marks saved defined-name edits lossy for package-part proof', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const result = await postJson('/preview', {
			file: TEMP_FILE,
			journal: true,
			ops: [{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$B$1' }],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.schemaVersion).toBe(MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION)
		expect(result.body.data?.journal?.schemaId).toBe(MUTATION_JOURNAL_ISSUE_SCHEMA.$id)
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.exact).toBe(false)
		expect(result.body.data?.journal?.inverseOps).toEqual([
			{ op: 'deleteDefinedName', name: 'Budget' },
		])
		expect(result.body.data?.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message:
					'setDefinedName changes saved package state that public inverse operations cannot restore byte-for-byte',
				surface: 'package-parts',
				reason: 'package-part-preservation',
				refs: ['name:Budget'],
			},
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.definedName('Budget')).toBeUndefined()
	})

	test('preview exposes unsupported journal status with partial inverse ops', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 1 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/preview', {
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'appendRows', table: 'Sales', rows: [['East', 2]] },
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'D1', value: 'audit' }] },
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.supported).toBe(false)
		expect(result.body.data?.journal?.exact).toBe(false)
		expect(result.body.data?.journal?.inverseOps).toEqual([
			{ op: 'clearRange', sheet: 'Sheet1', range: 'D1', what: 'all' },
		])
		expect(result.body.data?.journal?.issues).toContainEqual({
			code: 'UNSUPPORTED_OPERATION',
			message: 'No reversible journal support for appendRows',
			reason: 'operation-unsupported',
			surface: 'tables',
		})
	})

	test('write preserves lossy journal issue metadata while saving', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const result = await postJson('/write', {
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'groupRows', sheet: 'Sheet1', from: 1, to: 2, collapsed: true },
				{
					op: 'groupCols',
					sheet: 'Sheet1',
					from: 0,
					to: 1,
					collapsed: true,
					summaryRight: false,
				},
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.exact).toBe(false)
		expect(result.body.data?.journal?.inverseOps).toEqual([])
		expect(result.body.data?.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Grouped rows for Sheet1 cannot be restored with public operations',
				surface: 'row-layout',
				reason: 'row-layout-created',
				refs: [
					'Sheet1!2',
					'Sheet1!3',
					'Sheet1!4',
					'sheet:Sheet1:outlinePr:summaryBelow',
					'sheet:Sheet1:sheetFormatPr:outlineLevelRow',
				],
			},
			{
				code: 'LOSSY_INVERSE',
				message: 'Grouped columns for Sheet1 cannot be restored with public operations',
				surface: 'column-layout',
				reason: 'column-layout-created',
				refs: [
					'Sheet1!A',
					'Sheet1!B',
					'sheet:Sheet1:outlinePr:summaryRight',
					'sheet:Sheet1:sheetFormatPr:outlineLevelCol',
				],
			},
		])

		const reopened = await AscendWorkbook.open(TEMP_FILE)
		const sheet = reopened.getWorkbookModel().getSheet('Sheet1')
		expect(sheet?.rowDefs.get(1)).toEqual({ hidden: true, outlineLevel: 1 })
		expect(sheet?.rowDefs.get(2)).toEqual({ hidden: true, outlineLevel: 1 })
		expect(sheet?.rowDefs.get(3)).toEqual({ collapsed: true })
		expect(sheet?.colDefs).toContainEqual({ min: 0, max: 0, hidden: true, outlineLevel: 1 })
		expect(sheet?.colDefs).toContainEqual({ min: 1, max: 1, hidden: true, outlineLevel: 1 })
	})

	test('write exposes unsupported journal status with partial inverse ops', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 1 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B2', name: 'Sales', hasHeaders: true },
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/write', {
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'appendRows', table: 'Sales', rows: [['East', 2]] },
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'D1', value: 'audit' }] },
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.supported).toBe(false)
		expect(result.body.data?.journal?.exact).toBe(false)
		expect(result.body.data?.journal?.inverseOps).toEqual([
			{ op: 'clearRange', sheet: 'Sheet1', range: 'D1', what: 'all' },
		])
		expect(result.body.data?.journal?.issues).toContainEqual({
			code: 'UNSUPPORTED_OPERATION',
			message: 'No reversible journal support for appendRows',
			reason: 'operation-unsupported',
			surface: 'tables',
		})
		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.sheet('Sheet1')?.cell('D1')?.value).toEqual({
			kind: 'string',
			value: 'audit',
		})
	})

	test('write errors preserve structured unavailable journals for agents', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)

		const result = await postJson('/write', {
			file: TEMP_FILE,
			journal: true,
			ops: [{ op: 'clearRange', sheet: 'Sheet1', range: 'A1:', what: 'all' }],
		})

		expect(result.status).toBe(400)
		expect(result.body.ok).toBe(false)
		expect(result.body.error?.details?.apply?.journal).toMatchObject({
			supported: false,
			exact: false,
			inverseOps: [],
			issues: [
				{
					code: 'JOURNAL_UNAVAILABLE',
					surface: 'package-parts',
					reason: 'journal-unavailable',
				},
			],
			undoPolicy: {
				undoable: false,
				exact: false,
				reason: 'unavailable',
				riskLevel: 'high',
			},
		})
	})

	test('write exact theme journal inverse ops restore saved theme truth after reopen', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setTheme',
				themeName: 'Office',
				colorSchemeName: 'Office Colors',
				majorFontLatin: 'Aptos Display',
				minorFontLatin: 'Aptos',
				themeColors: [
					{ slot: 'accent1', rgb: '4F81BD' },
					{ slot: 'lt1', systemColor: 'window', lastColor: 'FFFFFF' },
				],
			},
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/write', {
			file: TEMP_FILE,
			journal: true,
			ops: [
				{
					op: 'setTheme',
					themeName: 'Brand',
					colorSchemeName: 'Brand Colors',
					majorFontLatin: 'Inter Display',
					minorFontLatin: 'Inter',
					themeColors: [
						{ slot: 'accent1', rgb: '0F6CBD' },
						{ slot: 'lt1', systemColor: 'windowText', lastColor: '000000' },
					],
				},
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.exact).toBe(true)
		expect(result.body.data?.journal?.issues).toEqual([])
		const inverse = parseOperations(result.body.data?.journal?.inverseOps)
		expect(inverse.ok).toBe(true)
		if (!inverse.ok) throw new Error('Expected exact theme journal inverse ops to parse')

		const changed = await AscendWorkbook.open(TEMP_FILE)
		expect(changed.inspect().themeSummary).toMatchObject({
			hasThemePart: true,
			name: 'Brand',
			colorSchemeName: 'Brand Colors',
			majorFontLatin: 'Inter Display',
			minorFontLatin: 'Inter',
		})
		expect(changed.inspect().themeSummary.colors.find((color) => color.slot === 'accent1')).toEqual(
			{
				slot: 'accent1',
				rgb: '0F6CBD',
			},
		)
		expect(changed.inspect().themeSummary.colors.find((color) => color.slot === 'lt1')).toEqual({
			slot: 'lt1',
			systemColor: 'windowText',
			lastColor: '000000',
		})

		const rollback = changed.apply(inverse.value)
		expect(rollback.errors).toEqual([])
		await changed.save(TEMP_FILE)
		const restored = await AscendWorkbook.open(TEMP_FILE)
		expect(restored.inspect().themeSummary).toMatchObject({
			hasThemePart: true,
			name: 'Office',
			colorSchemeName: 'Office Colors',
			majorFontLatin: 'Aptos Display',
			minorFontLatin: 'Aptos',
		})
		expect(
			restored.inspect().themeSummary.colors.find((color) => color.slot === 'accent1'),
		).toEqual({
			slot: 'accent1',
			rgb: '4F81BD',
		})
		expect(restored.inspect().themeSummary.colors.find((color) => color.slot === 'lt1')).toEqual({
			slot: 'lt1',
			systemColor: 'window',
			lastColor: 'FFFFFF',
		})
		expect(restored.check().valid).toBe(true)
	})

	test('write exact chart journal inverse ops restore saved chart truth after reopen', async () => {
		await Bun.write(TEMP_FILE, Bun.file(CHARTSHEET_FIXTURE))

		const result = await postJson('/write', {
			file: TEMP_FILE,
			journal: true,
			ops: [
				{
					op: 'setChartSeriesSource',
					partPath: 'xl/charts/chart1.xml',
					seriesIndex: 0,
					nameRef: 'Sheet1!$B$1',
					categoryRef: 'Sheet1!$A$2:$A$6',
					valueRef: 'Sheet1!$B$2:$B$6',
				},
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.exact).toBe(true)
		expect(result.body.data?.journal?.issues).toEqual([])
		const inverse = parseOperations(result.body.data?.journal?.inverseOps)
		expect(inverse.ok).toBe(true)
		if (!inverse.ok) throw new Error('Expected exact chart journal inverse ops to parse')

		const changed = await AscendWorkbook.open(TEMP_FILE)
		const changedChart = changed.getWorkbookModel().chartParts[0]
		expect(changedChart).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Chart1',
			chartType: 'barChart',
			title: 'Wildlife Population',
		})
		expect(changedChart?.series).toHaveLength(3)
		expect(changedChart?.series[0]).toMatchObject({
			nameRef: 'Sheet1!$B$1',
			nameText: 'Bears',
			categoryRef: 'Sheet1!$A$2:$A$6',
			valueRef: 'Sheet1!$B$2:$B$6',
		})
		expect(changedChart?.series[1]).toMatchObject({
			nameRef: 'Sheet1!$C$1',
			nameText: 'Dolphins',
			categoryRef: 'Sheet1!$A$2:$A$7',
			valueRef: 'Sheet1!$C$2:$C$7',
		})

		const rollback = changed.apply(inverse.value)
		expect(rollback.errors).toEqual([])
		await changed.save(TEMP_FILE)
		const restored = await AscendWorkbook.open(TEMP_FILE)
		const restoredChart = restored.getWorkbookModel().chartParts[0]
		expect(restoredChart).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Chart1',
			chartType: 'barChart',
			title: 'Wildlife Population',
		})
		expect(restoredChart?.series).toHaveLength(3)
		expect(restoredChart?.series[0]).toMatchObject({
			nameRef: 'Sheet1!$B$1',
			nameText: 'Bears',
			categoryRef: 'Sheet1!$A$2:$A$7',
			valueRef: 'Sheet1!$B$2:$B$7',
		})
		expect(restoredChart?.series[1]).toMatchObject({
			nameRef: 'Sheet1!$C$1',
			nameText: 'Dolphins',
			categoryRef: 'Sheet1!$A$2:$A$7',
			valueRef: 'Sheet1!$C$2:$C$7',
		})
		expect(restored.check().valid).toBe(true)
	})

	test('write exact pivot journal inverse ops restore saved pivot cache truth after reopen', async () => {
		await Bun.write(TEMP_FILE, Bun.file(PIVOT_FIXTURE))

		const result = await postJson('/write', {
			file: TEMP_FILE,
			journal: true,
			ops: [
				{
					op: 'setPivotCache',
					pivotTable: 'PivotTable1',
					sourceSheet: 'Sheet1',
					sourceRef: 'A1:K4',
				},
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.exact).toBe(true)
		expect(result.body.data?.journal?.issues).toEqual([])
		const inverse = parseOperations(result.body.data?.journal?.inverseOps)
		expect(inverse.ok).toBe(true)
		if (!inverse.ok) throw new Error('Expected exact pivot journal inverse ops to parse')

		const changed = await AscendWorkbook.open(TEMP_FILE)
		const changedCache = changed.getWorkbookModel().pivotCaches[0]
		expect(changedCache).toMatchObject({
			cacheId: 37,
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			recordsPartPath: 'xl/pivotCache/pivotCacheRecords1.xml',
			recordCount: 4,
			sourceSheet: 'Sheet1',
			sourceRef: 'A1:K4',
		})
		expect(changedCache?.fields).toHaveLength(11)
		expect(changedCache?.records?.parsedCount).toBe(4)

		const rollback = changed.apply(inverse.value)
		expect(rollback.errors).toEqual([])
		await changed.save(TEMP_FILE)
		const restored = await AscendWorkbook.open(TEMP_FILE)
		const restoredCache = restored.getWorkbookModel().pivotCaches[0]
		expect(restoredCache).toMatchObject({
			cacheId: 37,
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			recordsPartPath: 'xl/pivotCache/pivotCacheRecords1.xml',
			recordCount: 4,
			sourceSheet: 'Sheet1',
			sourceRef: 'A1:K5',
		})
		expect(restoredCache?.fields).toHaveLength(11)
		expect(restoredCache?.records?.parsedCount).toBe(4)
		expect(restored.check().valid).toBe(true)
	})

	test('preview marks pivot cache public rollback gaps as lossy', async () => {
		await Bun.write(TEMP_FILE, Bun.file(PIVOT_FIXTURE))

		const result = await postJson('/preview', {
			file: TEMP_FILE,
			journal: true,
			ops: [{ op: 'setPivotCache', pivotTable: 'PivotTable1', refreshOnLoad: true }],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.exact).toBe(false)
		expect(result.body.data?.journal?.inverseOps).toEqual([])
		expect(result.body.data?.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Pivot cache selector cannot be restored exactly',
				surface: 'pivot-caches',
				reason: 'pivot-cache-unsettable',
			},
		])
		const reopened = await AscendWorkbook.open(TEMP_FILE)
		expect(reopened.getWorkbookModel().pivotCaches[0]?.refreshOnLoad).toBeUndefined()
	})

	test('write lossy journal inverse ops restore saved workbook truth after reopen', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Status' },
					{ ref: 'A2', value: 'Open' },
					{ ref: 'A3', value: 'Closed' },
					{ ref: 'B1', value: 2 },
					{ ref: 'C1', value: 3.14 },
					{ ref: 'J1', value: 'Product' },
					{ ref: 'K1', value: 'Qty' },
					{ ref: 'J2', value: 'Widget' },
					{ ref: 'K2', value: 5 },
					{ ref: 'J3', value: 'Bolt' },
					{ ref: 'K3', value: 6 },
					{ ref: 'L1', value: 'currency-style-anchor' },
					{ ref: 'L2', value: 'decimal-style-anchor' },
				],
			},
			{ op: 'setStyle', sheet: 'Sheet1', range: 'B1:B1', style: { numberFormat: '0.00' } },
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C1:C1', format: '0.0' },
			{ op: 'setStyle', sheet: 'Sheet1', range: 'L1:L1', style: { numberFormat: '$0.00' } },
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'L2:L2', format: '0.000' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'F1:F3',
				rule: { type: 'whole', operator: 'greaterThan', formula1: '0', allowBlank: true },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'G1:G3',
				rule: { type: 'cellIs', operator: 'greaterThan', formula: '5', priority: 1 },
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'J1:K3', name: 'Sales', hasHeaders: true },
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:A3', column: 0, values: ['Open'] },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'E5', url: 'https://example.com' },
			{ op: 'setWorkbookView', index: 0, view: { activeTab: 0, firstSheet: 0, tabRatio: 600 } },
			{
				op: 'setWorkbookProtection',
				protection: { lockStructure: true, workbookPassword: 'ABCD' },
			},
			{
				op: 'setSheetProtection',
				sheet: 'Sheet1',
				password: 'ABCD',
				options: { formatCells: false, autoFilter: true },
			},
			{ op: 'setTabColor', sheet: 'Sheet1', color: 'FF0000' },
			{ op: 'freezePane', sheet: 'Sheet1', row: 1, col: 1 },
			{
				op: 'setDocumentProperties',
				mode: 'replace',
				properties: {
					core: { title: 'Before' },
					app: { company: 'Ascend' },
					custom: [{ name: 'Reviewed', value: false, type: 'bool' }],
				},
			},
		])
		await wb.save(TEMP_FILE)

		const result = await postJson('/write', {
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 7 }] },
				{ op: 'setStyle', sheet: 'Sheet1', range: 'B1:B1', style: { numberFormat: '$0.00' } },
				{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C1:C1', format: '0.000' },
				{
					op: 'setDataValidation',
					sheet: 'Sheet1',
					range: 'F1:F3',
					rule: { type: 'whole', operator: 'greaterThan', formula1: '10' },
				},
				{
					op: 'setConditionalFormat',
					sheet: 'Sheet1',
					range: 'G1:G3',
					rule: { type: 'cellIs', operator: 'greaterThan', formula: '7', priority: 1 },
				},
				{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:A3', column: 0, values: ['Closed'] },
				{ op: 'renameTable', table: 'Sales', newName: 'Revenue' },
				{ op: 'setTableColumn', table: 'Revenue', column: 'Qty', newName: 'Units' },
				{ op: 'setTableStyle', table: 'Revenue', styleName: null },
				{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'E5', url: 'https://changed.example' },
				{
					op: 'setWorkbookView',
					index: 0,
					mode: 'replace',
					view: { activeTab: 0, firstSheet: 0, tabRatio: 720 },
				},
				{
					op: 'setWorkbookProtection',
					protection: { lockWindows: true, workbookPassword: 'DCBA' },
				},
				{
					op: 'setSheetProtection',
					sheet: 'Sheet1',
					password: 'DCBA',
					options: { insertRows: true, deleteRows: false },
				},
				{ op: 'setTabColor', sheet: 'Sheet1', color: '00FF00' },
				{ op: 'freezePane', sheet: 'Sheet1', row: 2, col: 0 },
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.exact).toBe(false)
		expect(result.body.data?.journal?.issues).toContainEqual(
			expect.objectContaining({
				surface: 'package-parts',
				reason: 'package-part-preservation',
			}),
		)
		const inverse = parseOperations(result.body.data?.journal?.inverseOps)
		expect(inverse.ok).toBe(true)
		if (!inverse.ok) throw new Error('Expected exact journal inverse ops to parse')

		const changed = await AscendWorkbook.open(TEMP_FILE)
		expect(changed.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 7 })
		expect(changed.cellStyle('Sheet1!B1')?.numberFormat).toBe('$0.00')
		expect(changed.cellStyle('Sheet1!C1')?.numberFormat).toBe('0.000')
		expect(changed.sheet('Sheet1')?.dataValidations[0]).toMatchObject({
			sqref: 'F1:F3',
			formula1: '10',
		})
		expect(changed.sheet('Sheet1')?.conditionalFormats[0]?.rules[0]?.formulas).toEqual(['7'])
		expect(changed.sheet('Sheet1')?.autoFilter).toMatchObject({
			ref: 'A1:A3',
			columns: [{ colId: 0, kind: 'filters', values: ['Closed'] }],
		})
		expect(changed.table('Sales')).toBeUndefined()
		expect(changed.table('Revenue')?.columns).toEqual(['Product', 'Units'])
		expect(changed.table('Revenue')?.columnDefs[1]?.formula).toBeUndefined()
		expect(changed.table('Revenue')?.styleInfo).toBeUndefined()
		expect(changed.inspectSheet('Sheet1')?.hyperlinks?.[0]?.target).toBe('https://changed.example')
		expect(changed.workbookViews()[0]).toMatchObject({ activeTab: 0, firstSheet: 0, tabRatio: 720 })
		expect(changed.getWorkbookModel().workbookProtection).toMatchObject({
			lockWindows: true,
			workbookPassword: 'DCBA',
		})
		expect(changed.sheet('Sheet1')?.protection).toMatchObject({
			password: 'DCBA',
			insertRows: true,
			deleteRows: false,
		})
		expect(changed.sheet('Sheet1')?.tabColor).toEqual({ rgb: '00FF00' })
		expect(changed.sheet('Sheet1')?.frozenRows).toBe(2)
		expect(changed.sheet('Sheet1')?.frozenCols).toBe(0)
		expect(changed.inspect().documentProperties).toMatchObject({
			core: { title: 'Before' },
			app: { company: 'Ascend' },
			custom: [{ name: 'Reviewed', value: false, type: 'bool' }],
		})

		const rollback = changed.apply(inverse.value)
		expect(rollback.errors).toEqual([])
		await changed.save(TEMP_FILE)
		const restored = await AscendWorkbook.open(TEMP_FILE)
		expect(restored.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 2 })
		expect(restored.cellStyle('Sheet1!B1')?.numberFormat).toBe('0.00')
		expect(restored.cellStyle('Sheet1!C1')?.numberFormat).toBe('0.0')
		expect(restored.sheet('Sheet1')?.dataValidations[0]).toMatchObject({
			sqref: 'F1:F3',
			formula1: '0',
			allowBlank: true,
		})
		expect(restored.sheet('Sheet1')?.conditionalFormats[0]?.rules[0]?.formulas).toEqual(['5'])
		expect(restored.sheet('Sheet1')?.autoFilter).toMatchObject({
			ref: 'A1:A3',
			columns: [{ colId: 0, kind: 'filters', values: ['Open'] }],
		})
		expect(restored.table('Revenue')).toBeUndefined()
		expect(restored.table('Sales')?.columns).toEqual(['Product', 'Qty'])
		expect(restored.table('Sales')?.columnDefs[1]?.formula).toBeUndefined()
		expect(restored.table('Sales')?.styleInfo).toEqual({ name: 'TableStyleMedium2' })
		expect(restored.inspectSheet('Sheet1')?.hyperlinks?.[0]?.target).toBe('https://example.com')
		expect(restored.workbookViews()[0]).toMatchObject({
			activeTab: 0,
			firstSheet: 0,
			tabRatio: 600,
		})
		expect(restored.getWorkbookModel().workbookProtection).toMatchObject({
			lockStructure: true,
			workbookPassword: 'ABCD',
		})
		expect(restored.sheet('Sheet1')?.protection).toMatchObject({
			password: 'ABCD',
			formatCells: false,
			autoFilter: true,
		})
		expect(restored.sheet('Sheet1')?.tabColor).toEqual({ rgb: 'FF0000' })
		expect(restored.sheet('Sheet1')?.frozenRows).toBe(1)
		expect(restored.sheet('Sheet1')?.frozenCols).toBe(1)
		expect(restored.inspect().documentProperties).toMatchObject({
			core: { title: 'Before' },
			app: { company: 'Ascend' },
			custom: [{ name: 'Reviewed', value: false, type: 'bool' }],
		})
		expect(restored.check().valid).toBe(true)
	})

	test('write lossy journal inverse ops restore recalculated workbook truth after reopen', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 2 },
					{ ref: 'D1', value: 'Item' },
					{ ref: 'E1', value: 'Qty' },
					{ ref: 'F1', value: 'Calc' },
					{ ref: 'D2', value: 'A' },
					{ ref: 'E2', value: 5 },
					{ ref: 'D3', value: 'B' },
					{ ref: 'E3', value: 6 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*3' },
			{ op: 'createTable', sheet: 'Sheet1', ref: 'D1:F3', name: 'CalcTable', hasHeaders: true },
			{ op: 'setTableColumn', table: 'CalcTable', column: 'Calc', formula: '[@Qty]*2' },
		])
		expect(wb.recalc().errors).toEqual([])
		await wb.save(TEMP_FILE)

		const result = await postJson('/write', {
			file: TEMP_FILE,
			journal: true,
			ops: [
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
				{ op: 'setTableColumn', table: 'CalcTable', column: 'Calc', formula: '[@Qty]*3' },
			],
		})

		expect(result.status).toBe(200)
		expect(result.body.ok).toBe(true)
		expect(result.body.data?.journal?.supported).toBe(true)
		expect(result.body.data?.journal?.exact).toBe(false)
		expect(result.body.data?.journal?.issues).toContainEqual(
			expect.objectContaining({
				surface: 'package-parts',
				reason: 'package-part-preservation',
			}),
		)
		const inverse = parseOperations(result.body.data?.journal?.inverseOps)
		expect(inverse.ok).toBe(true)
		if (!inverse.ok) throw new Error('Expected exact journal inverse ops to parse')

		const changed = await AscendWorkbook.open(TEMP_FILE)
		expect(changed.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 5 })
		expect(changed.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 15 })
		expect(changed.table('CalcTable')?.columnDefs[2]?.formula).toBe('[@Qty]*3')
		expect(changed.sheet('Sheet1')?.cell('F2')?.value).toEqual({ kind: 'number', value: 15 })
		expect(changed.sheet('Sheet1')?.cell('F3')?.value).toEqual({ kind: 'number', value: 18 })

		const rollback = changed.apply(inverse.value)
		expect(rollback.errors).toEqual([])
		expect(changed.recalc().errors).toEqual([])
		await changed.save(TEMP_FILE)
		const restored = await AscendWorkbook.open(TEMP_FILE)
		expect(restored.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 2 })
		expect(restored.sheet('Sheet1')?.cell('B1')?.value).toEqual({ kind: 'number', value: 6 })
		expect(restored.table('CalcTable')?.columnDefs[2]?.formula).toBe('[@Qty]*2')
		expect(restored.sheet('Sheet1')?.cell('F2')?.value).toEqual({ kind: 'number', value: 10 })
		expect(restored.sheet('Sheet1')?.cell('F3')?.value).toEqual({ kind: 'number', value: 12 })
		expect(restored.check().valid).toBe(true)
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
		expect(plan.body.data?.writePolicy?.diagnostics).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					code: 'active-content-preserved',
					featureFamily: 'preservedMacro',
					preservationMode: 'preserve-exact',
				}),
				expect.objectContaining({
					code: 'approval-required-feature',
					featureFamily: 'preservedSignature',
					preservationMode: 'invalidated-on-edit',
				}),
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

	test('compact plan warns for embedded object and vendor security sidecars', async () => {
		const input = `${TEMP_FILE}.embedding-vendor-security.xlsx`
		await Bun.write(input, embeddingVendorSecurityWorkbook())
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 7 }] }],
			})

			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.writePolicy?.diagnostics).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						code: 'active-content-preserved',
						featureFamily: 'preservedEmbedding',
						preservationMode: 'preserve-exact',
					}),
					expect.objectContaining({
						code: 'active-content-preserved',
						featureFamily: 'preservedVendorSecurity',
						preservationMode: 'preserve-exact',
					}),
				]),
			)
		} finally {
			await unlink(input).catch(() => {})
		}
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
		expect(result.body.data?.preview?.changedRanges).toEqual([{ sheet: 'Sheet1', range: 'A1:A3' }])
	})

	test('compact prepared plan and commit preserve journal v1 issue compatibility', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = `${OUTPUT_FILE}.compact-journal.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: TEMP_FILE,
				compact: true,
				ops: JOURNAL_V1_FIXTURE.scenario.ops,
			})

			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.preparedPlan?.id).toBeString()
			expect(plan.body.data?.preview?.journalSummary).toEqual(JOURNAL_V1_FIXTURE.scenario.journal)

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.apply?.journalSummary).toEqual(JOURNAL_V1_FIXTURE.scenario.journal)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('compact plan and commit expose preservation mode summaries', async () => {
		await Bun.write(TEMP_FILE, preservedCustomWorkbook())
		const output = `${OUTPUT_FILE}.compact-preservation-modes.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: TEMP_FILE,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})

			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.writePolicy?.summary?.preservationModes).toMatchObject({
				reviewRequiredParts: 1,
				lossyApprovalRequiredFeatures: 1,
			})
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual([expect.stringMatching(/^loss:preservedother:preserved:/)])

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.writePolicy?.summary?.preservationModes).toMatchObject({
				reviewRequiredParts: 1,
				lossyApprovalRequiredFeatures: 1,
			})
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(false)
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('compact plan and commit expose generated opaque preservation mode summaries', async () => {
		const input = `${TEMP_FILE}.compact-opaque-preservation-modes-input.xlsx`
		await writeOpaqueX14Workbook(input)
		const output = `${OUTPUT_FILE}.compact-opaque-preservation-modes.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 11 }] }],
			})

			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.writePolicy?.summary?.preservationModes).toMatchObject({
				generatedWithOpaquePayloads: 2,
				reviewRequiredParts: 0,
				lossyApprovalRequiredFeatures: 0,
			})

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.writePolicy?.summary?.preservationModes).toMatchObject({
				generatedWithOpaquePayloads: 2,
				reviewRequiredParts: 0,
				lossyApprovalRequiredFeatures: 0,
			})
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.opaquePayloads).toMatchObject({
				generatedWithOpaquePayloads: 2,
				x14ConditionalFormatExtensionPayloads: 1,
				x14DataValidationExtensionPayloads: 1,
				worksheetParts: ['xl/worksheets/sheet1.xml'],
				preservationMode: 'generated-with-opaque-payload',
				verification: 'reopened-output',
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact plan and commit expose inspect-only preservation mode summaries', async () => {
		const input = `${TEMP_FILE}.compact-inspect-only-preservation-modes-input.xlsx`
		await Bun.write(input, inspectOnlyWorkbook())
		const output = `${OUTPUT_FILE}.compact-inspect-only-preservation-modes.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})

			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.writePolicy?.summary?.preservationModes).toMatchObject({
				inspectOnlyParts: 1,
				lossyApprovalRequiredFeatures: 1,
			})
			expect(plan.body.data?.writePolicy?.diagnostics).toContainEqual(
				expect.objectContaining({
					code: 'approval-required-feature',
					featureFamily: 'preservedPowerQuery',
					preservationMode: 'inspect-only',
					packageParts: [
						expect.objectContaining({
							partPath: 'xl/customData/item1.data',
							preservationPolicy: 'inspect-only',
							preservationMode: 'inspect-only',
						}),
					],
				}),
			)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual([expect.stringMatching(/^loss:preservedpowerquery:preserved:/)])

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.writePolicy?.summary?.preservationModes).toMatchObject({
				inspectOnlyParts: 1,
				lossyApprovalRequiredFeatures: 1,
			})
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened legacy comment summary', async () => {
		const input = `${TEMP_FILE}.compact-comment-summary-input.xlsx`
		const output = `${OUTPUT_FILE}.compact-comment-summary.xlsx`
		const workbook = AscendWorkbook.create()
		workbook.getWorkbookModel().sheets[0]?.comments.set('B2', {
			text: 'Review this',
			author: 'Ada',
		})
		await workbook.save(input)
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.comments).toMatchObject({
				legacyCommentLocations: 1,
				threadedCommentLocations: 0,
				legacyDrawingLocations: 1,
				locations: ['Sheet1!B2'],
				threadedCommentPartPaths: [],
				verification: 'reopened-output',
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened table summary', async () => {
		const input = `${TEMP_FILE}.compact-table-summary-input.xlsx`
		const output = `${OUTPUT_FILE}.compact-table-summary.xlsx`
		const workbook = AscendWorkbook.create()
		expect(
			workbook.apply([
				{
					op: 'setCells',
					sheet: 'Sheet1',
					updates: [
						{ ref: 'A1', value: 'Qty' },
						{ ref: 'B1', value: 'Price' },
						{ ref: 'A2', value: 2 },
						{ ref: 'B2', value: 5 },
						{ ref: 'A3', value: 3 },
						{ ref: 'B3', value: 7 },
					],
				},
				{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B3', name: 'Sales', hasHeaders: true },
			]).errors,
		).toHaveLength(0)
		await workbook.save(input)
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'D1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.tables).toMatchObject({
				tableLocations: 1,
				queryTableLocations: 0,
				tableAutoFilterLocations: 1,
				tableNames: ['Sales'],
				locations: ['Sheet1!A1:B3'],
				tablePartPaths: ['xl/tables/table1.xml'],
				queryTablePartPaths: [],
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened defined name summary', async () => {
		const input = `${TEMP_FILE}.compact-defined-name-summary-input.xlsx`
		const output = `${OUTPUT_FILE}.compact-defined-name-summary.xlsx`
		const workbook = AscendWorkbook.create()
		const sheet = workbook.getWorkbookModel().sheets[0]
		workbook.getWorkbookModel().definedNames.set('GlobalRate', 'Sheet1!$A$1')
		if (sheet) {
			workbook
				.getWorkbookModel()
				.definedNames.set(
					'LocalRate',
					'Sheet1!$B$1',
					{ kind: 'sheet', sheetId: sheet.id },
					{ hidden: true },
				)
		}
		await workbook.save(input)
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'C1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.definedNames).toMatchObject({
				total: 2,
				workbookScoped: 1,
				sheetScoped: 1,
				hidden: 1,
				names: [
					{ name: 'GlobalRate', formula: 'Sheet1!$A$1', scope: 'workbook' },
					{
						name: 'LocalRate',
						formula: 'Sheet1!$B$1',
						scope: 'sheet',
						sheet: 'Sheet1',
						hidden: true,
					},
				],
				verification: 'reopened-output',
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened workbook and sheet security summary', async () => {
		const input = `${TEMP_FILE}.compact-security-summary-input.xlsx`
		const output = `${OUTPUT_FILE}.compact-security-summary.xlsx`
		const workbook = AscendWorkbook.create()
		const model = workbook.getWorkbookModel()
		model.workbookProtection = {
			lockStructure: true,
			lockWindows: true,
			workbookPassword: 'ABCD',
		}
		const sheet = model.getSheet('Sheet1')
		if (!sheet) throw new Error('Expected Sheet1')
		sheet.protection = {
			sheet: true,
			password: 'DCBA',
			autoFilter: true,
			sort: true,
		}
		sheet.protectedRanges = [{ name: 'Editable', sqref: 'C:C', password: '1234' }]
		await workbook.save(input)
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.security).toMatchObject({
				workbookProtected: true,
				workbookLocks: ['lockStructure', 'lockWindows'],
				workbookPasswordProtected: true,
				workbookRevisionPasswordProtected: false,
				protectedSheets: 1,
				protectedSheetNames: ['Sheet1'],
				sheetPasswordProtected: 1,
				sheetStrongHashProtected: 0,
				protectedRanges: 1,
				protectedRangeLocations: ['Sheet1!C:C'],
				passwordHashVerification: 'reported-not-validated',
				preservationMode: 'generated',
				verification: 'reopened-output',
				sheets: [
					expect.objectContaining({
						sheetName: 'Sheet1',
						protected: true,
						passwordProtected: true,
						allowedActions: ['sort', 'autoFilter'],
						protectedRanges: 1,
						protectedRangeLocations: ['Sheet1!C:C'],
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened external reference binding summary', async () => {
		const input = `${TEMP_FILE}.compact-external-reference-summary-input.xlsx`
		const output = `${OUTPUT_FILE}.compact-external-reference-summary.xlsx`
		await Bun.write(input, externalLinkBoundWorkbook())
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.externalReferences).toMatchObject({
				total: 1,
				boundByExternalBookRelId: 1,
				fallbackPathRelationships: 0,
				missingPathRelationships: 0,
				partPaths: ['xl/externalLinks/externalLink1.xml'],
				targets: ['../sources/source.xlsx'],
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				parts: [
					expect.objectContaining({
						partPath: 'xl/externalLinks/externalLink1.xml',
						relId: 'rIdExternal',
						externalBookRelId: 'rIdExt',
						linkRelId: 'rIdExt',
						linkBindingStatus: 'externalBookRelId',
						target: '../sources/source.xlsx',
						targetMode: 'External',
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact plan exposes workbook-qualified 3D external references as sheet-span risk', async () => {
		const input = `${TEMP_FILE}.compact-external-3d-plan-input.xlsx`
		await Bun.write(input, externalLinkBoundWorkbook())
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [
					{
						op: 'setFormula',
						sheet: 'Sheet1',
						ref: 'B2',
						formula: '=SUM([1]FY26:FY28!B2:B10)',
					},
				],
			})

			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const dependency = plan.body.data?.writePolicy?.diagnostics?.find(
				(diagnostic) => diagnostic.code === 'external-link-dependency',
			) as
				| {
						details?: {
							relatedOperations?: readonly unknown[]
							externalLinks?: readonly unknown[]
						}
				  }
				| undefined
			expect(dependency?.details?.relatedOperations).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						operationIndex: 0,
						op: 'setFormula',
						sourceKind: 'cellFormula',
						sourceRef: 'Sheet1!B2',
						formula: '=SUM([1]FY26:FY28!B2:B10)',
						workbook: '1',
						sheetSpan: { startSheet: 'FY26', endSheet: 'FY28' },
						references: ["'[1]FY26:FY28'!B2:B10"],
					}),
				]),
			)
			expect(dependency?.details?.externalLinks).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						workbook: '1',
						sheetSpans: [{ startSheet: 'FY26', endSheet: 'FY28' }],
					}),
				]),
			)
		} finally {
			await unlink(input).catch(() => {})
		}
	})

	test('compact commit exposes reopened analytics refresh summary', async () => {
		const input = `${TEMP_FILE}.compact-analytics-summary-input.xlsx`
		const output = `${OUTPUT_FILE}.compact-analytics-summary.xlsx`
		await Bun.write(input, analyticsRefreshWorkbook())
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'PivotSheet', updates: [{ ref: 'A1', value: 'ok' }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(false)
			expect(commit.body.data?.postWrite?.unresolvedPackageGraphIssueCount).toBe(0)
			expect(commit.body.data?.postWrite?.analytics).toMatchObject({
				pivotCaches: 1,
				pivotTables: 1,
				slicerCaches: 1,
				slicers: 1,
				timelineCaches: 1,
				timelines: 1,
				partPaths: expect.arrayContaining([
					'xl/pivotCache/pivotCacheDefinition1.xml',
					'xl/pivotTables/pivotTable1.xml',
					'xl/slicerCaches/slicerCache1.xml',
					'xl/timelineCaches/timelineCache1.xml',
				]),
				requiresExternalRefresh: true,
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				pivotCacheDetails: [
					expect.objectContaining({
						partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
						cacheId: 34,
						sourceSheet: 'Raw',
						sourceRef: 'A1:B3',
						outputState: 'refresh-on-open',
						requiresExternalRefresh: true,
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened active content summary', async () => {
		const input = `${TEMP_FILE}.compact-active-content-input.xlsm`
		const output = `${OUTPUT_FILE}.compact-active-content.xlsm`
		await Bun.write(input, signedMacroWorkbook())
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(false)
			expect(commit.body.data?.postWrite?.activeContent).toMatchObject({
				total: 2,
				vbaProjects: 1,
				activeXControls: 0,
				vbaSignatures: 1,
				digitalSignatures: 0,
				partPaths: ['xl/vbaProject.bin', 'xl/vbaProjectSignature.bin'],
				executionPolicy: 'blocked',
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				entries: expect.arrayContaining([
					expect.objectContaining({
						kind: 'vbaProject',
						partPath: 'xl/vbaProject.bin',
						contentType: 'application/vnd.ms-office.vbaProject',
						anchor: 'workbook',
						opaque: true,
						executionPolicy: 'blocked',
					}),
					expect.objectContaining({
						kind: 'vbaSignature',
						partPath: 'xl/vbaProjectSignature.bin',
						invalidationPolicy: 'invalidatedByPackageEdit',
						resigningPolicy: 'notSupported',
					}),
				]),
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened control active content summary', async () => {
		const input = `${TEMP_FILE}.compact-control-active-content-input.xlsx`
		const output = `${OUTPUT_FILE}.compact-control-active-content.xlsx`
		await Bun.write(input, controlWorkbook())
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.activeContent).toMatchObject({
				total: 3,
				vbaProjects: 0,
				activeXControls: 1,
				formControls: 1,
				macroSheets: 0,
				vbaSignatures: 0,
				digitalSignatures: 0,
				partPaths: [
					'xl/activeX/activeX1.xml',
					'xl/activeX/activeX1.bin',
					'xl/ctrlProps/ctrlProp1.xml',
				],
				executionPolicy: 'blocked',
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				entries: expect.arrayContaining([
					expect.objectContaining({
						kind: 'activeX',
						partPath: 'xl/activeX/activeX1.xml',
						contentType: 'application/vnd.ms-office.activeX+xml',
						anchor: 'sheet',
						sheetName: 'Data',
						sourceRelationshipId: 'rIdActiveX',
						relType: 'http://schemas.microsoft.com/office/2006/relationships/activeXControl',
					}),
					expect.objectContaining({
						kind: 'activeX',
						partPath: 'xl/activeX/activeX1.bin',
						contentType: 'application/vnd.ms-office.activeX',
						sourcePartPath: 'xl/activeX/activeX1.xml',
						sourceRelationshipId: 'rId1',
						relType: 'http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary',
					}),
					expect.objectContaining({
						kind: 'formControl',
						partPath: 'xl/ctrlProps/ctrlProp1.xml',
						contentType: 'application/vnd.ms-excel.controlproperties+xml',
						anchor: 'sheet',
						sheetName: 'Data',
						sourceRelationshipId: 'rIdCtrl',
						relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp',
					}),
				]),
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened macro sheet active content summary', async () => {
		const input = `${TEMP_FILE}.compact-macro-sheet-input.xlsm`
		const output = `${OUTPUT_FILE}.compact-macro-sheet.xlsm`
		await Bun.write(input, macroSheetWorkbook())
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.activeContent).toMatchObject({
				total: 1,
				vbaProjects: 0,
				activeXControls: 0,
				formControls: 0,
				macroSheets: 1,
				vbaSignatures: 0,
				digitalSignatures: 0,
				partPaths: ['xl/macrosheets/sheet1.xml'],
				executionPolicy: 'blocked',
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				entries: [
					expect.objectContaining({
						kind: 'macroSheet',
						partPath: 'xl/macrosheets/sheet1.xml',
						contentType: 'application/vnd.ms-excel.macrosheet+xml',
						anchor: 'sheet',
						sheetName: 'Macro1',
						sourceRelationshipId: 'rIdMacro',
						relType: 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet',
						opaque: true,
						executionPolicy: 'blocked',
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened custom UI active content summary', async () => {
		const input = `${TEMP_FILE}.compact-custom-ui-input.xlsm`
		const output = `${OUTPUT_FILE}.compact-custom-ui.xlsm`
		await Bun.write(input, customUiWorkbook())
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.activeContent).toMatchObject({
				total: 1,
				vbaProjects: 0,
				activeXControls: 0,
				formControls: 0,
				macroSheets: 0,
				vbaSignatures: 0,
				digitalSignatures: 0,
				customUi: 1,
				unknownActiveContent: 0,
				partPaths: ['customUI/customUI2.xml'],
				executionPolicy: 'blocked',
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				entries: [
					expect.objectContaining({
						kind: 'customUi',
						partPath: 'customUI/customUI2.xml',
						contentType: 'application/vnd.ms-office.customUI+xml',
						anchor: 'workbook',
						executionPolicy: 'blocked',
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened visual summary', async () => {
		const input = `${TEMP_FILE}.compact-visual-summary-input.xlsx`
		const output = `${OUTPUT_FILE}.compact-visual-summary.xlsx`
		await Bun.write(input, visualWorkbook())
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.visuals).toMatchObject({
				sheetsWithVisuals: 1,
				images: 1,
				drawingObjects: 0,
				drawingMlObjects: 0,
				vmlObjects: 0,
				chartParts: 0,
				chartSheets: 0,
				drawingPartPaths: ['xl/drawings/drawing1.xml'],
				mediaPartPaths: ['xl/media/image1.png'],
				chartPartPaths: [],
				vmlPartPaths: [],
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
				sheets: [
					expect.objectContaining({
						sheetName: 'Sheet1',
						hasDrawingMl: true,
						hasVml: false,
						imageCount: 1,
						drawingPartPaths: ['xl/drawings/drawing1.xml'],
						mediaPartPaths: ['xl/media/image1.png'],
					}),
				],
			})
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commit exposes reopened chartsheet visual summary', async () => {
		const input = CHARTSHEET_FIXTURE
		const output = `${OUTPUT_FILE}.compact-chartsheet-summary.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
			expect(commit.body.data?.postWrite?.visuals).toMatchObject({
				sheetsWithVisuals: 0,
				images: 0,
				chartParts: 1,
				chartSheets: 1,
				drawingPartPaths: [],
				mediaPartPaths: [],
				chartPartPaths: ['xl/charts/chart1.xml'],
				vmlPartPaths: [],
				preservationMode: 'preserve-exact',
				verification: 'reopened-output',
			})
		} finally {
			await unlink(output).catch(() => {})
		}
	})

	test('compact plan and commit expose invalidated signature preservation mode summaries', async () => {
		const input = `${TEMP_FILE}.compact-signature-invalidation-input.xlsx`
		await Bun.write(input, signedPackageWorkbook())
		const output = `${OUTPUT_FILE}.compact-signature-invalidation.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 11 }] }],
			})

			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.writePolicy?.summary?.preservationModes).toMatchObject({
				invalidatedOnEditParts: 2,
				lossyApprovalRequiredFeatures: 1,
			})
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual([expect.stringMatching(/^loss:preservedsignature:preserved:/)])

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: approvalIds,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.writePolicy?.summary?.preservationModes).toMatchObject({
				invalidatedOnEditParts: 2,
				lossyApprovalRequiredFeatures: 1,
			})
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact plan and commit expose discarded calc-chain preservation mode summaries', async () => {
		const input = `${TEMP_FILE}.compact-calc-chain-discard-input.xlsx`
		await Bun.write(input, calcChainWorkbook())
		const output = `${OUTPUT_FILE}.compact-calc-chain-discard.xlsx`
		try {
			const plan = await postJson('/plan', {
				file: input,
				compact: true,
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1+A1' }],
			})

			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.writePolicy?.summary?.calcChainPolicy).toBe(
				'discarded-for-formula-topology',
			)
			expect(plan.body.data?.writePolicy?.summary?.preservationModes).toMatchObject({
				discardedForRecalcParts: 1,
				lossyApprovalRequiredFeatures: 0,
			})

			const commit = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				compact: true,
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.writePolicy?.summary?.calcChainPolicy).toBe(
				'discarded-for-formula-topology',
			)
			expect(commit.body.data?.writePolicy?.summary?.preservationModes).toMatchObject({
				discardedForRecalcParts: 1,
				lossyApprovalRequiredFeatures: 0,
			})
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('compact commits expose bounded affected refs and ranges', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const output = `${OUTPUT_FILE}.compact-affected.xlsx`
		try {
			const commit = await postJson('/commit', {
				file: TEMP_FILE,
				output,
				compact: true,
				maxAffectedCells: 2,
				ops: [
					{
						op: 'setCells',
						sheet: 'Sheet1',
						updates: [
							{ ref: 'A1', value: 1 },
							{ ref: 'A2', value: 2 },
							{ ref: 'A3', value: 3 },
						],
					},
				],
			})

			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.apply?.affectedCellCount).toBe(3)
			expect(commit.body.data?.apply?.emittedAffectedCellCount).toBe(2)
			expect(commit.body.data?.apply?.affectedCellRefs).toEqual(['A1', 'A2'])
			expect(commit.body.data?.apply?.affectedRanges).toEqual([{ sheet: 'Sheet1', range: 'A1:A3' }])
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(true)
		} finally {
			await unlink(output).catch(() => {})
		}
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
			expect(plan.body.data?.preparedPlan?.file).toBe(TEMP_FILE)
			expect(plan.body.data?.preparedPlan?.inputSha256).toBe(plan.body.data?.inputSha256)
			expect(plan.body.data?.preparedPlan?.planDigest).toBe(plan.body.data?.planDigest)
			expect(plan.body.data?.preparedPlan?.operationCount).toBe(plan.body.data?.operationCount)
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

	test('prepared plan handles report failed writes and remain retryable', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(TEMP_FILE)
		const blockedOutput = join(
			tmpdir(),
			`ascend-api-prepared-missing-parent-${Date.now()}`,
			'out.xlsx',
		)
		const retryOutput = `${OUTPUT_FILE}.prepared-failed-write-retry.xlsx`
		const ops = [{ op: 'addSheet' as const, name: 'PreparedRetry' }]
		try {
			const plan = await postJson('/plan', { file: TEMP_FILE, ops })
			expect(plan.status).toBe(200)
			expect(plan.body.data?.preparedPlan?.id).toBeString()

			const blocked = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output: blockedOutput,
				approvals: [],
				compact: true,
			})
			expect(blocked.status).toBe(409)
			expect(blocked.body.ok).toBe(false)
			expect(blocked.body.error).toMatchObject({
				code: 'EXPORT_ERROR',
				retryable: true,
				retryStrategy: 'modified',
				details: {
					output: blockedOutput,
					operation: 'atomic-workbook-write',
				},
			})
			expect(await Bun.file(blockedOutput).exists()).toBe(false)

			const retried = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output: retryOutput,
				approvals: [],
				compact: true,
			})
			expect(retried.status).toBe(200)
			expect(retried.body.data?.postWrite?.valid).toBe(true)
			expect(retried.body.data?.postWrite?.auditsPassed).toBe(true)
			const reopened = await AscendWorkbook.open(retryOutput)
			expect(reopened.sheets).toContain('PreparedRetry')

			const reused = await postJson('/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output: `${retryOutput}.reuse.xlsx`,
				approvals: [],
			})
			expect(reused.status).toBe(400)
			expect(reused.body.error?.message).toBe('Prepared plan handle has already been used')
		} finally {
			await unlink(retryOutput).catch(() => {})
			await unlink(`${retryOutput}.reuse.xlsx`).catch(() => {})
		}
	})

	test('path mutation plans reuse a guarded open until a prepared handle owns it', async () => {
		const input = join(
			tmpdir(),
			`ascend-api-plan-cache-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		const output = `${input}.prepared.xlsx`
		const wb = AscendWorkbook.create()
		await wb.save(input)
		const apiFetch = createApiFetch()
		const originalOpenSourceBytes = AscendWorkbook.openSourceBytes.bind(AscendWorkbook)
		let openSourceBytesCalls = 0
		Object.defineProperty(AscendWorkbook, 'openSourceBytes', {
			configurable: true,
			value: (async (...args: Parameters<typeof AscendWorkbook.openSourceBytes>) => {
				openSourceBytesCalls += 1
				return originalOpenSourceBytes(...args)
			}) satisfies typeof AscendWorkbook.openSourceBytes,
		})
		try {
			const first = await postApiFetch(apiFetch, '/plan', {
				file: input,
				prepare: false,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 1 }],
			})
			expect(first.status).toBe(200)
			const second = await postApiFetch(apiFetch, '/plan', {
				file: input,
				prepare: false,
				compact: true,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 2 }],
			})
			expect(second.status).toBe(200)
			expect(openSourceBytesCalls).toBe(1)

			const prepared = await postApiFetch(apiFetch, '/plan', {
				file: input,
				compact: true,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 3 }],
			})
			expect(prepared.status).toBe(200)
			expect(prepared.body.data?.preparedPlan?.id).toBeString()
			expect(openSourceBytesCalls).toBe(1)

			const committed = await postApiFetch(apiFetch, '/commit', {
				planHandle: prepared.body.data?.preparedPlan?.id,
				output,
				approvals: [],
				compact: true,
			})
			expect(committed.status).toBe(200)
			expect(committed.body.data?.postWrite?.valid).toBe(true)
			expect(openSourceBytesCalls).toBe(1)

			const afterPrepared = await postApiFetch(apiFetch, '/plan', {
				file: input,
				prepare: false,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 4 }],
			})
			expect(afterPrepared.status).toBe(200)
			expect(openSourceBytesCalls).toBe(2)
		} finally {
			Object.defineProperty(AscendWorkbook, 'openSourceBytes', {
				configurable: true,
				value: originalOpenSourceBytes,
			})
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
		}
	})

	test('check reuses the guarded post-write verification result after commit', async () => {
		const input = join(
			tmpdir(),
			`ascend-api-check-cache-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
		)
		const output = `${input}.out.xlsx`
		const wb = AscendWorkbook.create()
		await wb.save(input)
		const apiFetch = createApiFetch()
		try {
			const plan = await postApiFetch(apiFetch, '/plan', {
				file: input,
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 7 }],
			})
			expect(plan.status).toBe(200)
			const commit = await postApiFetch(apiFetch, '/commit', {
				planHandle: plan.body.data?.preparedPlan?.id,
				output,
				approvals: [],
				compact: true,
			})
			expect(commit.status).toBe(200)
			expect(commit.body.data?.postWrite?.valid).toBe(true)

			const originalOpen = WorkbookDocument.open.bind(WorkbookDocument)
			let documentOpenCalls = 0
			Object.defineProperty(WorkbookDocument, 'open', {
				configurable: true,
				value: (async (...args: Parameters<typeof WorkbookDocument.open>) => {
					documentOpenCalls += 1
					return originalOpen(...args)
				}) satisfies typeof WorkbookDocument.open,
			})
			try {
				const check = await postApiFetch(apiFetch, '/check', { file: output })
				expect(check.status).toBe(200)
				expect(check.body.data?.valid).toBe(true)
				expect(documentOpenCalls).toBe(0)
			} finally {
				Object.defineProperty(WorkbookDocument, 'open', {
					configurable: true,
					value: originalOpen,
				})
			}
		} finally {
			await unlink(input).catch(() => {})
			await unlink(output).catch(() => {})
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
			expect(commit.body.data?.outputSha256).toMatch(/^[a-f0-9]{64}$/)
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
			expect(
				exactCommit.body.data?.postWrite?.packageGraphAudit?.emittedIssueCount,
			).toBeGreaterThan(0)
			expect(exactCommit.body.data?.postWrite?.packageGraphAudit?.issues).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						code: 'package_feature_classification',
						partPath: 'xl/custom/custom1.xml',
						preservationPolicy: 'unknown-review-required',
						preservationMode: 'review-required',
					}),
				]),
			)
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

	test('direct path mutation commits surface post-write audit failures as blocked output', async () => {
		await Bun.write(TEMP_FILE, preservedCustomWorkbook())
		const output = `${OUTPUT_FILE}.direct-preserved.xlsx`
		const mutations = [{ path: '/sheets/Sheet1/cells/A1/value', value: 17 }]
		const canonicalOps = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 17 }] }]
		try {
			const plan = await postJson('/plan', { file: TEMP_FILE, mutations })
			expect(plan.status).toBe(200)
			expect(plan.body.ok).toBe(true)
			expect(plan.body.data?.pathMutations?.ops).toEqual(canonicalOps)
			const approvalIds = plan.body.data?.approvals?.map((approval) => approval.id) ?? []
			expect(approvalIds).toEqual([expect.stringMatching(/^loss:preservedother:preserved:/)])

			const commit = await postJson('/commit', {
				file: TEMP_FILE,
				mutations,
				output,
				approvals: approvalIds,
			})
			expect(commit.status).toBe(200)
			expect(commit.body.ok).toBe(true)
			expect(commit.body.data?.pathMutations?.ops).toEqual(canonicalOps)
			expect(commit.body.data?.approvals?.map((approval) => approval.id)).toEqual(approvalIds)
			expect(commit.body.data?.postWrite?.valid).toBe(true)
			expect(commit.body.data?.postWrite?.auditsPassed).toBe(false)
			expect(commit.body.data?.postWrite?.outputSha256).toBe(commit.body.data?.outputSha256)
			expect(commit.body.data?.postWrite?.packageGraphAudit?.ok).toBe(false)
			expect(commit.body.data?.postWrite?.packageGraphAudit?.issues).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						code: 'package_feature_classification',
						partPath: 'xl/custom/custom1.xml',
						preservationPolicy: 'unknown-review-required',
						preservationMode: 'review-required',
					}),
				]),
			)
			expect(commit.body.data?.postWrite?.expectedPackageGraphIssueCount).toBe(0)
			expect(commit.body.data?.postWrite?.unresolvedPackageGraphIssueCount).toBeGreaterThan(0)
			expect(commit.body.data?.modelOutput?.blocked).toBe(true)
			expect(commit.body.data?.modelOutput?.counts?.postWritePackageGraphIssues).toBeGreaterThan(0)
			expect(commit.body.data?.modelOutput?.nextActions?.join('\n')).toContain(
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

function sharedFormulaWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rIdSheet"/></sheets>
  <definedNames><definedName name="BudgetTotal">Calc!$A$1:$A$2</definedName></definedNames>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:A2">SUM(B1:B2)</f><v>3</v></c>
      <c r="B1"><f t="shared" si="0">A1*2</f><v>6</v></c>
      <c r="C1"><f>SUM(Sales[[Revenue]:[Quantity]])</f><v>14</v></c>
      <c r="D1"><f>BudgetTotal*2</f><v>6</v></c>
    </row>
    <row r="2">
      <c r="B2"><f t="shared" si="0"/><v>8</v></c>
    </row>
  </sheetData>
</worksheet>`,
	})
}

function sharedOnlyFormulaWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="B1"><f t="shared" si="0">A1*2</f><v>6</v></c></row>
    <row r="2"><c r="B2"><f t="shared" si="0"/><v>8</v></c></row>
  </sheetData>
</worksheet>`,
	})
}

function dynamicArrayWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdMetadata" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata" Target="metadata.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/metadata.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
  <metadataTypes count="1">
    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>
  </metadataTypes>
  <futureMetadata name="XLDAPR" count="1">
    <bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>
  </futureMetadata>
  <cellMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </cellMetadata>
</metadata>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1"><f>_xlfn.SEQUENCE(3)</f><v>1</v></c>
      <c r="B1"><f>SUM(_xlfn.ANCHORARRAY(A1))</f><v>6</v></c>
      <c r="C1"><f>_xlfn.SINGLE(A1)</f><v>1</v></c>
    </row>
  </sheetData>
</worksheet>`,
	})
}

function staleSpillCacheWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="e"><f>_xlfn.SEQUENCE(3)</f><v>#SPILL!</v></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>blocker</t></is></c></row>
  </sheetData>
</worksheet>`,
	})
}

function controlWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdActiveX" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml"/>
  <Relationship Id="rIdCtrl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
</Relationships>`,
		'xl/activeX/activeX1.xml': `<?xml version="1.0"?><ax:ocx ax:classid="{8BD21D40-EC42-11CE-9E0D-00AA006002F3}" ax:persistence="persistStreamInit" r:id="rId1" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`,
		'xl/activeX/_rels/activeX1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
</Relationships>`,
		'xl/activeX/activeX1.bin': 'active-binary',
		'xl/ctrlProps/ctrlProp1.xml': `<?xml version="1.0"?><formControlPr macro="Module1.Run" fmlaLink="$A$1" fmlaRange="$A$2:$A$4"/>`,
	})
}

function macroSheetWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/macrosheets/sheet1.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdMacro" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rIdData"/>
    <sheet name="Macro1" sheetId="2" r:id="rIdMacro" state="hidden"/>
  </sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/macrosheets/sheet1.xml': `<?xml version="1.0"?>
<xm:macrosheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1"><f>RUN("Task")</f><v>0</v></c></row></sheetData>
</xm:macrosheet>`,
		'xl/macrosheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
	})
}

function embeddingVendorSecurityWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="xen" ContentType="application/octet-stream"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rDellEncryptedDoc" Type="http://schemas.dell.com/ddp/2016/relationships/xenFile" Target="ddp/ddpfile.xen"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <oleObjects><oleObject progId="Package" r:id="rIdOle" shapeId="1025"/></oleObjects>
</worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>`,
		'xl/embeddings/oleObject1.bin': 'embedded-payload',
		'ddp/ddpfile.xen': 'opaque-vendor-security',
	})
}

function customUiWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/customUI/customUI2.xml" ContentType="application/vnd.ms-office.customUI+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rIdCustomUi" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="/customUI/customUI2.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'customUI/customUI2.xml': `<?xml version="1.0"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon.OnLoad" loadImage="Ribbon.LoadImage">
  <ribbon><tabs><tab id="tabAscend" label="Ascend">
    <group id="grpActions" label="Actions">
      <button id="runReport" label="Run" onAction="Module1.RunReport" getEnabled="Ribbon.CanRun"/>
    </group>
  </tab></tabs></ribbon>
</customUI>`,
		'customUI/_rels/customUI2.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../xl/media/image1.png"/>
</Relationships>`,
		'xl/media/image1.png': 'image-bytes',
	})
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

function inspectOnlyWorkbook(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/customData/item1.data" ContentType="application/vnd.ms-excel.customData"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdPowerQuery" Type="http://schemas.microsoft.com/office/2014/relationships/powerQueryMashup" Target="customData/item1.data"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/customData/item1.data': encode('power-query-mashup-bytes'),
			}),
		),
	)
}

async function writeOpaqueX14Workbook(path: string): Promise<void> {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().sheets[0]
	if (!sheet) throw new Error('expected default sheet')
	sheet.x14ConditionalFormats.push({
		index: 0,
		sqref: 'A1:A5',
		type: 'dataBar',
		priority: 4,
		formulas: [],
		preservedRuleAttributes: { 'xr:uid': '{CF-UID}' },
		preservedRuleChildXml: [
			'<x14:extLst><x14:ext uri="{cf-extension}"><x14ac:metadata flag="1"/></x14:ext></x14:extLst>',
		],
	})
	sheet.x14DataValidations.push({
		index: 0,
		sqref: 'C2:C5',
		type: 'list',
		operator: 'between',
		formula1: '$A$1:$A$4',
		preservedAttributes: { 'xr:uid': '{DV-UID}' },
		preservedChildXml: ['<x14ac:metadata flag="1"><x14ac:item val="keep"/></x14ac:metadata>'],
	})
	await wb.save(path)
}

function signedPackageWorkbook(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/_xmlsignatures/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>
  <Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rIdSignatureOrigin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/>
</Relationships>`),
				'_xmlsignatures/_rels/origin.sigs.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSignature" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="sig1.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'_xmlsignatures/origin.sigs': encode(''),
				'_xmlsignatures/sig1.xml': encode(
					'<?xml version="1.0"?><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>',
				),
			}),
		),
	)
}

function calcChainWorkbook(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdCalcChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row>
  </sheetData>
</worksheet>`),
				'xl/calcChain.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <c r="B1" i="1"/>
</calcChain>`),
			}),
		),
	)
}

function externalLinkBoundWorkbook(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets>
  <externalReferences><externalReference r:id="rIdExternal"/></externalReferences>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/externalLinks/externalLink1.xml':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rIdExt"/>
</externalLink>`),
				'xl/externalLinks/_rels/externalLink1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdExt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="../sources/source.xlsx" TargetMode="External"/>
</Relationships>`),
			}),
		),
	)
}

function analyticsRefreshWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCache1.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>
  <Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdSheet2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rIdPivotCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rIdSlicerCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
  <Relationship Id="rIdTimelineCache" Type="http://schemas.microsoft.com/office/2011/relationships/timelineCache" Target="timelineCaches/timelineCache1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <pivotCaches><pivotCache cacheId="34" r:id="rIdPivotCache"/></pivotCaches>
  <sheets>
    <sheet name="PivotSheet" sheetId="1" r:id="rIdSheet1"/>
    <sheet name="Raw" sheetId="2" r:id="rIdSheet2"/>
  </sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/worksheets/sheet2.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Sales</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>West</t></is></c><c r="B2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>East</t></is></c><c r="B3"><v>20</v></c></row>
  </sheetData>
</worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
  <Relationship Id="rIdSlicer" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
  <Relationship Id="rIdTimeline" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline1.xml"/>
</Relationships>`,
		'xl/pivotTables/pivotTable1.xml': `<?xml version="1.0"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="34">
  <location ref="A3:C8" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="1"><pivotField axis="axisPage" multipleItemSelectionAllowed="1" showAll="0"><items count="2"><item x="0"/><item x="1"/></items></pivotField></pivotFields>
  <pageFields count="1"><pageField fld="0" item="0" name="Region"/></pageFields>
</pivotTableDefinition>`,
		'xl/pivotCache/pivotCacheDefinition1.xml': `<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  r:id="rIdRecords" recordCount="2" refreshOnLoad="1" enableRefresh="1">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Raw"/></cacheSource>
  <cacheFields count="2">
    <cacheField name="Region" databaseField="1"><sharedItems count="2"><s v="West"/><s v="East"/></sharedItems></cacheField>
    <cacheField name="Sales" databaseField="1"><sharedItems containsNumber="1" count="2"><n v="10"/><n v="20"/></sharedItems></cacheField>
  </cacheFields>
</pivotCacheDefinition>`,
		'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`,
		'xl/pivotCache/pivotCacheRecords1.xml': `<?xml version="1.0"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
  <r><x v="0"/><n v="10"/></r>
  <r><x v="1"/><n v="20"/></r>
</pivotCacheRecords>`,
		'xl/slicerCaches/slicerCache1.xml': `<?xml version="1.0"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_Region" sourceName="Region">
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
  <data><tabular pivotCacheId="34"><items count="2"><i x="0" s="1"/><i x="1"/></items></tabular></data>
</slicerCacheDefinition>`,
		'xl/slicerCaches/_rels/slicerCache1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSlicerUi" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
</Relationships>`,
		'xl/slicers/slicer1.xml': `<?xml version="1.0"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><slicer name="Region" cache="Slicer_Region" caption="Region"/></slicers>`,
		'xl/timelineCaches/timelineCache1.xml': `<?xml version="1.0"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="Timeline_Order_Date" sourceName="Order Date">
  <data><tabular pivotCacheId="34"/></data>
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
  <state filterId="7" filterPivotName="PivotTable1" filterType="dateRange" filterTabId="2" pivotCacheId="34" singleRangeFilterState="1">
    <selection startDate="2023-01-01T00:00:00" endDate="2023-12-31T00:00:00"/>
  </state>
</timelineCacheDefinition>`,
		'xl/timelineCaches/_rels/timelineCache1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTimelineUi" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline1.xml"/>
</Relationships>`,
		'xl/timelines/timeline1.xml': `<?xml version="1.0"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><timeline name="Order_Date" cache="Timeline_Order_Date" caption="Order Date"/></timelines>`,
	})
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

function visualWorkbook(): Uint8Array {
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
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/worksheets/_rels/sheet1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`),
				'xl/drawings/_rels/drawing1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rIdDrawing1"/>
</worksheet>`),
				'xl/drawings/drawing1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
</xdr:wsDr>`),
				'xl/media/image1.png': encode('png-bytes'),
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
