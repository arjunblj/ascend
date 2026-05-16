import { afterAll, beforeAll, describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendWorkbook } from '@ascend/sdk'
import { makeXlsx } from '../../packages/io-xlsx/test/helpers.ts'
import { createApiFetch } from './src/server.ts'

const apiFetch = createApiFetch()
let tempDir = ''
const PIVOT_FIXTURE = join(
	import.meta.dir,
	'../../fixtures/xlsx/libreoffice/PivotTable_CachedDefinitionAndDataInSync.xlsx',
)
const ENCRYPTED_FIXTURE = join(import.meta.dir, '../../fixtures/xlsx/calamine/pass_protected.xlsx')

beforeAll(() => {
	tempDir = mkdtempSync(join(tmpdir(), 'ascend-api-test-'))
})

afterAll(() => {
	if (tempDir.length > 0) rmSync(tempDir, { recursive: true, force: true })
})

function api(path: string, init?: RequestInit): Promise<Response> {
	return apiFetch(new Request(`http://ascend.test${path}`, init))
}

describe('API', () => {
	test('health check returns 200', async () => {
		const res = await api(`/health`)
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body).toEqual({
			formatVersion: 1,
			ok: true,
			data: { status: 'ok' },
		})
	})

	test('operations and capabilities endpoints expose agent schemas', async () => {
		const operations = await api(`/operations`)
		expect(operations.status).toBe(200)
		const opsBody = await operations.json()
		expect(opsBody.ok).toBe(true)
		expect(opsBody.data.operations.some((op: { op: string }) => op.op === 'setCells')).toBe(true)
		expect(opsBody.data.schemas[0].examples.length).toBeGreaterThan(0)

		const capabilities = await api(`/capabilities?feature=pivots`)
		expect(capabilities.status).toBe(200)
		const capsBody = await capabilities.json()
		expect(capsBody.ok).toBe(true)
		expect(
			capsBody.data.capabilities.some(
				(capability: { id: string }) => capability.id === 'analytics.pivots',
			),
		).toBe(true)

		const workflow = await api(`/agent-workflow`)
		expect(workflow.status).toBe(200)
		const workflowBody = await workflow.json()
		expect(workflowBody.ok).toBe(true)
		expect(workflowBody.data.workflow).toContainEqual(
			expect.objectContaining({
				step: 'plan',
				endpoint: 'POST /plan',
				proof: expect.arrayContaining(['inputSha256', 'planDigest', 'preparedPlan']),
			}),
		)
		expect(workflowBody.data.workflow).toContainEqual(
			expect.objectContaining({
				step: 'reopen-verify',
				endpoints: expect.arrayContaining(['POST /check', 'POST /lint', 'POST /diff']),
			}),
		)
		expect(workflowBody.data.examples).toMatchObject({
			installedSdkSafeEdit:
				'bun node_modules/@ascend/sdk/examples/package-install-safe-edit.ts <file.xlsx> <out.xlsx>',
			sdkSafeEdit: 'bun run example:safe-edit <file.xlsx> <out.xlsx>',
			apiSafeEdit: 'bun run example:safe-edit:http <file.xlsx> <out.xlsx>',
			mcpSafeEdit: 'bun run example:safe-edit:mcp <file.xlsx> <out.xlsx>',
		})
		expect(workflowBody.data.packageInstallExampleContext).toMatchObject({
			workdir: 'consumer-project',
			requires: expect.arrayContaining(['@ascend/sdk installed']),
			proofOutput: expect.arrayContaining([
				'proofBundle.safeToUse',
				'proofBundle.whatChanged',
				'proofBundle.whySafe',
			]),
		})
		expect(workflowBody.data.exampleContext).toMatchObject({
			workdir: 'repository-root',
			requires: expect.arrayContaining(['source checkout', 'bun install']),
			proofCommand: 'bun test examples/root-scripts.test.ts',
		})
		expect(workflowBody.data.preparedHandles).toMatchObject({
			scope: 'process-local',
			oneShot: true,
		})
	})

	test('plan and commit endpoints provide the safe write workflow', async () => {
		const tempFile = join(tempDir, 'agent-workflow.xlsx')
		const outputFile = join(tempDir, 'agent-workflow-output.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)
		const ops = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'api-plan' }] }]

		const plan = await api(`/plan`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, ops, includePackageActions: true }),
		})
		expect(plan.status).toBe(200)
		const planBody = await plan.json()
		expect(planBody.ok).toBe(true)
		expect(planBody.data.inputSha256).toMatch(/^[a-f0-9]{64}$/)
		expect(planBody.data.preview.wouldSucceed).toBe(true)
		expect(planBody.data.packageActions.kind).toBe('ascend-package-action-proof')
		expect(planBody.data.packageActions.byAction.regenerate).toBeGreaterThan(0)

		const commit = await api(`/commit`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				ops,
				output: outputFile,
				expectSha256: planBody.data.inputSha256,
				includePackageActions: true,
			}),
		})
		expect(commit.status).toBe(200)
		const commitBody = await commit.json()
		expect(commitBody.ok).toBe(true)
		expect(commitBody.data.outputSha256).toMatch(/^[a-f0-9]{64}$/)
		expect(commitBody.data.packageActions.kind).toBe('ascend-package-action-proof')
		expect(commitBody.data.packageActions.byAction.regenerate).toBeGreaterThan(0)
		expect(commitBody.data.packageActions.coverage.sourceByteDigestCount).toBeGreaterThan(0)
		expect(commitBody.data.packageActions.coverage.outputByteDigestCount).toBeGreaterThan(0)
		expect(
			commitBody.data.packageActions.coverage.matchingByteDigestCount +
				commitBody.data.packageActions.coverage.mismatchedByteDigestCount,
		).toBeGreaterThan(0)
		expect(
			commitBody.data.packageActions.actions.some(
				(action: { outputSha256?: string }) => action.outputSha256 !== undefined,
			),
		).toBe(true)
	})

	test('plan endpoint rejects malformed operation field types', async () => {
		const tempFile = join(tempDir, 'agent-validation.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)

		const plan = await api(`/plan`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				ops: [{ op: 'insertRows', sheet: 'Sheet1', at: 0, count: '2' }],
			}),
		})
		expect(plan.status).toBe(400)
		const body = await plan.json()
		expect(body.ok).toBe(false)
		expect(body.error.code).toBe('VALIDATION_ERROR')
		expect(body.error.details.issues[0]).toContain('count must be a positive integer')
	})

	test('commit endpoint requires approval for destructive operations', async () => {
		const tempFile = join(tempDir, 'agent-approval.xlsx')
		const outputFile = join(tempDir, 'agent-approval-output.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Scratch' }])
		await wb.save(tempFile)
		const ops = [{ op: 'deleteSheet', sheet: 'Scratch' }]

		const plan = await api(`/plan`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, ops }),
		})
		const planBody = await plan.json()
		const approvalId = planBody.data.approvals[0].id
		expect(planBody.data.needsApproval).toBe(true)

		const blocked = await api(`/commit`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, ops, output: outputFile }),
		})
		expect(blocked.status).toBe(400)

		const committed = await api(`/commit`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, ops, output: outputFile, approvals: [approvalId] }),
		})
		expect(committed.status).toBe(200)
		const committedBody = await committed.json()
		expect(committedBody.ok).toBe(true)
		expect(committedBody.data.approvals[0].id).toBe(approvalId)
	})

	test('inspect returns workbook info', async () => {
		const tempFile = join(tempDir, 'test.xlsx')
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as {
			wb: {
				workbookProtection: unknown
				sheets: Array<{
					comments: Map<string, { text: string }>
					conditionalFormats: Array<Record<string, unknown>>
					dataValidations: Array<Record<string, unknown>>
					imageRefs: Array<Record<string, unknown>>
				}>
			}
		}
		internal.wb.workbookProtection = { lockStructure: true }
		internal.wb.sheets[0]?.comments.set('A1', { text: 'note' })
		internal.wb.sheets[0]?.conditionalFormats.push({
			sqref: 'A1',
			rules: [{ type: 'cellIs', formulas: ['1'] }],
		})
		internal.wb.sheets[0]?.dataValidations.push({ sqref: 'B1', type: 'list', formula1: '"A,B"' })
		internal.wb.sheets[0]?.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rId1',
			targetPath: 'xl/media/image1.png',
		})
		await wb.save(tempFile)

		const res = await api(`/inspect`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(true)
		expect(body.data.sheetCount).toBe(1)
		expect(body.data.loadedSheetCount).toBe(1)
		expect(body.data.commentCount).toBeNull()
		expect(body.data.conditionalFormatCount).toBeNull()
		expect(body.data.dataValidationCount).toBeNull()
		expect(body.data.imageCount).toBeNull()
		expect(body.data.pivotTableCount).toBe(0)
		expect(body.data.pivotCacheCount).toBe(0)
		expect(body.data.slicerCount).toBe(0)
		expect(body.data.slicerCacheCount).toBe(0)
		expect(body.data.hasWorkbookProtection).toBe(true)
		expect(body.data.sheets).toHaveLength(1)
		expect(body.data.sheets[0].name).toBe('Sheet1')
		expect(body.data.sheets[0].commentCount).toBeNull()
		expect(body.data.sheets[0].conditionalFormatCount).toBeNull()
		expect(body.data.sheets[0].dataValidationCount).toBeNull()
		expect(body.data.sheets[0].imageCount).toBeNull()
		expect(body.data.cellCount).toBeNull()
		expect(body.data.load.mode).toBe('metadata-only')
	})

	test('active-content endpoint returns metadata-only active content risk', async () => {
		const tempFile = join(tempDir, 'api-active-content.xlsm')
		await Bun.write(tempFile, signedMacroWorkbook())

		const res = await api(`/active-content`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.ok).toBe(true)
		expect(body.data.activeContentCount).toBe(2)
		expect(body.data.activeContent).toContainEqual(
			expect.objectContaining({
				kind: 'vbaProject',
				partPath: 'xl/vbaProject.bin',
				sourceRelationshipId: 'rIdVba',
			}),
		)
		expect(body.data.activeContent).toContainEqual(
			expect.objectContaining({
				kind: 'vbaSignature',
				partPath: 'xl/vbaProjectSignature.bin',
				sourcePartPath: 'xl/vbaProject.bin',
				sourceRelationshipId: 'rIdVbaSignature',
			}),
		)
		expect(body.data.compatibilityFeatures).toContainEqual(
			expect.objectContaining({
				feature: 'preservedMacro',
				locations: ['xl/vbaProject.bin'],
			}),
		)
		expect(body.data.compatibilityFeatures).toContainEqual(
			expect.objectContaining({
				feature: 'preservedSignature',
				locations: ['xl/vbaProjectSignature.bin'],
			}),
		)
	})

	test('package-graph endpoint exposes package identity for agents', async () => {
		const tempFile = join(tempDir, 'api-package-graph.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)

		const res = await api(`/package-graph`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.ok).toBe(true)
		expect(body.data.parts).toContainEqual(
			expect.objectContaining({
				path: 'xl/workbook.xml',
				featureFamily: 'workbook',
				ownerScope: 'workbook',
				sourceRelationshipId: 'rId1',
			}),
		)
		expect(body.data.relationships).toContainEqual(
			expect.objectContaining({
				relationshipPartPath: '_rels/.rels',
				id: 'rId1',
				resolvedTarget: 'xl/workbook.xml',
			}),
		)
	})

	test('open-plan endpoint recommends a load mode without hydrating workbook cells', async () => {
		const tempFile = join(tempDir, 'api-open-plan.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 42 }] }])
		await wb.save(tempFile)

		const res = await api(`/open-plan`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, intent: 'read-values' }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.ok).toBe(true)
		expect(body.data.intent).toBe('read-values')
		expect(body.data.recommendedLoadOptions).toEqual({ mode: 'values' })
		expect(body.data.reviewBeforeHydration).toBe(false)
	})

	test('open-plan endpoint routes macro workbooks to metadata review', async () => {
		const tempFile = join(tempDir, 'api-open-plan-macro.xlsm')
		await Bun.write(tempFile, signedMacroWorkbook())

		const res = await api(`/open-plan`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.ok).toBe(true)
		expect(body.data.recommendedLoadOptions).toEqual({ mode: 'metadata-only' })
		expect(body.data.reviewBeforeHydration).toBe(true)
		expect(body.data.riskFeatures).toContainEqual(
			expect.objectContaining({ featureFamily: 'preservedMacro' }),
		)
	})

	test('open-plan endpoint accepts encrypted workbook passwords without echoing them', async () => {
		const res = await api(`/open-plan`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: ENCRYPTED_FIXTURE, password: '123' }),
		})
		expect(res.status).toBe(200)
		const responseText = await res.text()
		const body = JSON.parse(responseText)
		expect(body.ok).toBe(true)
		expect(body.data.recommendedLoadOptions).toEqual({ mode: 'full' })
		expect(body.data.partCount).toBeGreaterThan(0)
		expect(responseText).not.toContain('123')
	})

	test('open-plan endpoint rejects non-string encrypted workbook passwords', async () => {
		const res = await api(`/open-plan`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: ENCRYPTED_FIXTURE, password: 123 }),
		})

		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body).toMatchObject({
			ok: false,
			error: {
				code: 'VALIDATION_ERROR',
				message: 'Invalid open-plan password',
			},
		})
	})

	test('open-plan endpoint rejects invalid intents', async () => {
		const tempFile = join(tempDir, 'api-open-plan-invalid.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)

		const res = await api(`/open-plan`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, intent: 'fast' }),
		})
		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body.ok).toBe(false)
		expect(body.error.code).toBe('VALIDATION_ERROR')
		expect(body.error.details.allowedIntents).toContain('edit-plan')
	})

	test('pivots endpoint exposes output audits and materialization ops for agents', async () => {
		const res = await api(`/pivots`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: PIVOT_FIXTURE, pivotTable: 'PivotTable1' }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.ok).toBe(true)
		expect(body.data.pivotTables[0].name).toBe('PivotTable1')
		expect(body.data.pivotOutputAudits[0]).toMatchObject({
			pivotTable: 'PivotTable1',
			status: 'passed',
		})
		expect(body.data.pivotOutputMaterializePlan).toEqual({
			ops: [],
			plannedCellCount: 0,
			unsupported: [],
		})
	})

	test('inspect can return a values-loaded sheet summary', async () => {
		const tempFile = join(tempDir, 'sheet-inspect.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B2', value: 42 }] }])
		await wb.save(tempFile)

		const res = await api(`/inspect`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, sheet: 'Sheet1' }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(true)
		expect(body.data.name).toBe('Sheet1')
		expect(body.data.cellDataLoaded).toBe(true)
		expect(body.data.cellCount).toBe(1)
		expect(body.data.rowCount).toBe(2)
		expect(body.data.colCount).toBe(2)
	})

	test('visuals returns full workbook visual inventory', async () => {
		const tempFile = join(tempDir, 'visuals.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)

		const res = await api(`/visuals`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.ok).toBe(true)
		expect(body.data.load.mode).toBe('full')
		expect(body.data.sheetImageCount).toBe(0)
		expect(body.data.sheets[0]).toMatchObject({
			sheet: 'Sheet1',
			hasDrawing: false,
			hasLegacyDrawing: false,
			imageCount: 0,
		})
	})

	test('check returns structured issue metadata for agent repair', async () => {
		const tempFile = join(tempDir, 'check-issues.xlsx')
		await Bun.write(tempFile, threadedCommentMissingPersonsWorkbook())

		const res = await api(`/check`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.ok).toBe(true)
		expect(body.data.valid).toBe(false)
		const issue = body.data.issues.find(
			(entry: { rule?: string }) => entry.rule === 'threaded-comment-integrity',
		)
		expect(issue.refs).toEqual(['Sheet1!A1'])
		expect(issue.details.kind).toBe('threaded-comment-unknown-person-id')
		expect(issue.suggestedFix).toContain('persons part')
	})

	test('read returns versioned machine envelope', async () => {
		const tempFile = join(tempDir, 'read.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 42 }] }])
		await wb.save(tempFile)

		const res = await api(`/read`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, sheet: 'Sheet1', range: 'A1:A1' }),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(true)
		expect(body.data.cells).toHaveLength(1)
		expect(body.data.cells[0].ref).toBe('A1')
	})

	test('read honors pagination and display options', async () => {
		const tempFile = join(tempDir, 'read-paged.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 10 },
					{ ref: 'A3', value: 'Bob' },
					{ ref: 'B3', value: 20 },
				],
			},
		])
		await wb.save(tempFile)

		const res = await api(`/read`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				sheet: 'Sheet1',
				range: 'A1:B3',
				format: 'objects',
				headers: ['Name', 'Score'],
				rowOffset: 1,
				rowLimit: 1,
				display: true,
			}),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.ok).toBe(true)
		expect(body.data.rowOffset).toBe(1)
		expect(body.data.rowLimit).toBe(1)
		expect(body.data.rows).toEqual([{ Name: 'Alice', Score: '10' }])
	})

	test('read errors return machine failure envelope', async () => {
		const res = await api(`/read`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ range: 'A1:A1' }),
		})
		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(false)
		expect(body.error).toMatchObject({
			code: 'VALIDATION_ERROR',
			message: 'Missing or invalid read workbook reference',
			retryable: true,
			retryStrategy: 'modified',
			details: { required: ['file'] },
		})
	})

	test('preview returns a diff without mutating the workbook on disk', async () => {
		const tempFile = join(tempDir, 'preview.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*2' },
		])
		await wb.recalc()
		await wb.save(tempFile)

		const res = await api(`/preview`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				journal: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] }],
			}),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(true)
		expect(body.data.cellChanges.length).toBeGreaterThan(0)
		expect(body.data.journal.supported).toBe(true)
		expect(body.data.journal.inverseOps).toEqual([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
		])
		expect(body.data.writePlan.totalParts).toBeGreaterThan(0)

		const reopened = await AscendWorkbook.open(tempFile)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 2 })
		expect(reopened.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 4 })
	})

	test('preview returns a failure envelope when recalculation reports errors', async () => {
		const tempFile = join(tempDir, 'preview-recalc-error.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)

		const res = await api(`/preview`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }],
			}),
		})
		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body.ok).toBe(false)
		expect(body.error.message).toContain('Circular reference detected')
		expect(body.error.details.preview.errors.length).toBeGreaterThan(0)

		const reopened = await AscendWorkbook.open(tempFile)
		expect(reopened.sheet('Sheet1')?.cell('A1')).toBeUndefined()
	})

	test('write returns a failure envelope on operation errors', async () => {
		const tempFile = join(tempDir, 'write-error.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)

		const res = await api(`/write`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				ops: [{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 1 }] }],
			}),
		})
		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body.ok).toBe(false)
		expect(body.error.message).toContain('Sheet')
	})

	test('write can include mutation journal metadata in the success envelope', async () => {
		const tempFile = join(tempDir, 'write-journal.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'before' }] }])
		await wb.save(tempFile)

		const res = await api(`/write`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				journal: true,
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'after' }] }],
			}),
		})
		expect(res.status).toBe(200)
		const body = await res.json()
		expect(body.ok).toBe(true)
		expect(body.data.journal.supported).toBe(true)
		expect(body.data.journal.inverseOps).toEqual([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'before' }] },
		])

		const reopened = await AscendWorkbook.open(tempFile)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
			kind: 'string',
			value: 'after',
		})
	})

	test('write returns a failure envelope when recalculation reports errors', async () => {
		const tempFile = join(tempDir, 'write-recalc-error.xlsx')
		const wb = AscendWorkbook.create()
		await wb.save(tempFile)

		const res = await api(`/write`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				file: tempFile,
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }],
			}),
		})
		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body.ok).toBe(false)
		expect(body.error.message).toContain('Circular reference detected')
	})

	test('calc returns a failure envelope when recalculation reports errors', async () => {
		const tempFile = join(tempDir, 'calc-recalc-error.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }])
		await wb.save(tempFile)

		const res = await api(`/calc`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile }),
		})
		expect(res.status).toBe(400)
		const body = await res.json()
		expect(body.ok).toBe(false)
		expect(body.error.message).toContain('Circular reference detected')
	})

	test('unknown route returns 404', async () => {
		const res = await api(`/unknown`)
		expect(res.status).toBe(404)
		const body = await res.json()
		expect(body.formatVersion).toBe(1)
		expect(body.ok).toBe(false)
		expect(body.error).toMatchObject({
			code: 'INVALID_ARGUMENT',
			message: 'Unsupported API route: GET /unknown',
			retryable: true,
			retryStrategy: 'modified',
			details: {
				method: 'GET',
				path: '/unknown',
				supportedRoutes: expect.arrayContaining([
					{ method: 'GET', path: '/operations' },
					{ method: 'GET', path: '/agent-workflow' },
					{ method: 'POST', path: '/open-plan' },
					{ method: 'POST', path: '/plan' },
					{ method: 'POST', path: '/commit' },
				]),
				workflow: ['open-plan', 'inspect', 'plan', 'commit'],
			},
			suggestedFix: expect.stringContaining('GET /operations'),
		})
	})

	test('export returns TSV bytes and rejects unsupported formats', async () => {
		const tempFile = join(tempDir, 'export.xlsx')
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 10 },
				],
			},
		])
		await wb.save(tempFile)

		const tsvRes = await api(`/export`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, format: 'tsv' }),
		})
		expect(tsvRes.status).toBe(200)
		expect(tsvRes.headers.get('Content-Type')).toContain('text/tab-separated-values')
		expect(await tsvRes.text()).toContain('Name\tScore')

		const badRes = await api(`/export`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ file: tempFile, format: 'weird' }),
		})
		expect(badRes.status).toBe(400)
		const badBody = await badRes.json()
		expect(badBody.ok).toBe(false)
		expect(badBody.error.message).toContain('Unsupported format')
	})
})

function signedMacroWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>`,
		'xl/_rels/vbaProject.bin.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVbaSignature" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/vbaProject.bin': 'macro-bytes',
		'xl/vbaProjectSignature.bin': 'signature-bytes',
	})
}

function threadedCommentMissingPersonsWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdThreaded" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
</Relationships>`,
		'xl/threadedComments/threadedComment1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="0" id="tc1" dT="2024-01-01T00:00:00.000">
    <text>Please review</text>
  </threadedComment>
</ThreadedComments>`,
	})
}
