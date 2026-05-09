import { afterAll, describe, expect, test } from 'bun:test'
import { existsSync, unlinkSync } from 'node:fs'
import { AscendWorkbook } from '@ascend/sdk'

const CLI = new URL('./index.ts', import.meta.url).pathname
const TEST_FILE = 'test-output.xlsx'
const MULTI_SHEET_FILE = 'test-multi.xlsx'
const NAMED_RANGE_FILE = 'test-named.xlsx'
const PIVOT_CORPUS_FILE = '../../../research/excel-corpus/ms-excel-formulas-and-pivot-tables.xlsx'
const SLICER_CORPUS_FILE = '../../../research/excel-corpus/excel-dashboard-v2.xlsx'

function run(...args: string[]): Promise<{ stdout: string; stderr: string; exitCode: number }> {
	return new Promise((resolve) => {
		const proc = Bun.spawn(['bun', 'run', CLI, ...args], {
			stdout: 'pipe',
			stderr: 'pipe',
			cwd: import.meta.dir,
		})

		proc.exited.then(async (exitCode) => {
			const stdout = await new Response(proc.stdout).text()
			const stderr = await new Response(proc.stderr).text()
			resolve({ stdout: stdout.trim(), stderr: stderr.trim(), exitCode })
		})
	})
}

afterAll(() => {
	for (const f of [
		TEST_FILE,
		MULTI_SHEET_FILE,
		NAMED_RANGE_FILE,
		'exported.tsv',
		'exported.json',
		'plan-ops.json',
		'commit-ops.json',
		'commit-output.xlsx',
	]) {
		const path = `${import.meta.dir}/${f}`
		if (existsSync(path)) unlinkSync(path)
	}
})

describe('ascend cli', () => {
	test('--version prints version', async () => {
		const { stdout, exitCode } = await run('--version')
		expect(exitCode).toBe(0)
		expect(stdout).toMatch(/^\d+\.\d+\.\d+$/)
	})

	test('--help shows help text', async () => {
		const { stdout, exitCode } = await run('--help')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('ascend')
		expect(stdout).toContain('Commands:')
		expect(stdout).toContain('Global flags:')
	})

	test('unknown command exits 1', async () => {
		const { exitCode, stderr } = await run('nonexistent')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Unknown command')
	})

	test('unknown command suggests the closest match', async () => {
		const { exitCode, stderr } = await run('inspcet')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Did you mean "inspect"?')
	})

	test('unknown command --json returns a failure envelope', async () => {
		const { exitCode, stdout } = await run('inspcet', '--json')
		expect(exitCode).toBe(1)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(false)
		expect(parsed.error.code).toBe('INVALID_ARGUMENT')
		expect(parsed.error.details.suggestion).toBe('inspect')
	})

	test('create makes a workbook file', async () => {
		const { stdout, exitCode } = await run('create', TEST_FILE)
		expect(exitCode).toBe(0)
		expect(stdout).toContain('Created')
		expect(existsSync(`${import.meta.dir}/${TEST_FILE}`)).toBe(true)
	})

	test('ops --json exposes operation schemas with examples', async () => {
		const { stdout, exitCode } = await run('ops', '--op', 'setCells', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.operations[0].op).toBe('setCells')
		expect(parsed.data.schemas[0].examples[0].op).toBe('setCells')
	})

	test('capabilities --json exposes the Excel capability matrix', async () => {
		const { stdout, exitCode } = await run('capabilities', '--feature', 'pivots', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(true)
		expect(
			parsed.data.capabilities.some(
				(capability: { id: string }) => capability.id === 'analytics.pivots',
			),
		).toBe(true)
	})

	test('plan and commit implement safe agent workflow', async () => {
		const opsFile = `${import.meta.dir}/commit-ops.json`
		await Bun.write(
			opsFile,
			JSON.stringify([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'agent-safe' }] },
			]),
		)
		const plan = await run('plan', TEST_FILE, '--ops', 'commit-ops.json', '--json')
		expect(plan.exitCode).toBe(0)
		const planned = JSON.parse(plan.stdout)
		expect(planned.ok).toBe(true)
		expect(planned.data.inputSha256).toMatch(/^[a-f0-9]{64}$/)
		expect(planned.data.preview.wouldSucceed).toBe(true)

		const commit = await run(
			'commit',
			TEST_FILE,
			'--ops',
			'commit-ops.json',
			'--output',
			'commit-output.xlsx',
			'--expect-sha256',
			planned.data.inputSha256,
			'--json',
		)
		expect(commit.exitCode).toBe(0)
		const committed = JSON.parse(commit.stdout)
		expect(committed.ok).toBe(true)
		expect(committed.data.outputSha256).toMatch(/^[a-f0-9]{64}$/)
		expect(existsSync(`${import.meta.dir}/commit-output.xlsx`)).toBe(true)
	})

	test('inspect shows sheet info', async () => {
		const { stdout, exitCode } = await run('inspect', TEST_FILE)
		expect(exitCode).toBe(0)
		expect(stdout).toContain('Sheet1')
	})

	test('inspect --json outputs valid JSON', async () => {
		const { stdout, exitCode } = await run('inspect', TEST_FILE, '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.formatVersion).toBe(1)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.sheetCount).toBe(1)
		expect(parsed.data.loadedSheetCount).toBe(1)
		expect(parsed.data.commentCount).toBeNull()
		expect(parsed.data.conditionalFormatCount).toBeNull()
		expect(parsed.data.dataValidationCount).toBeNull()
		expect(parsed.data.imageCount).toBeNull()
		expect(parsed.data.pivotTableCount).toBe(0)
		expect(parsed.data.pivotCacheCount).toBe(0)
		expect(parsed.data.slicerCount).toBe(0)
		expect(parsed.data.slicerCacheCount).toBe(0)
		expect(parsed.data.load.mode).toBe('metadata-only')
		expect(parsed.data.sheets).toBeArray()
		expect(parsed.data.sheets[0].name).toBe('Sheet1')
	})

	test('unknown inspect flag suggests the closest supported flag', async () => {
		const { exitCode, stderr } = await run('inspect', TEST_FILE, '--shet', 'Sheet1')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Unknown flag for "inspect": --shet')
		expect(stderr).toContain('Did you mean "--sheet"?')
	})

	test('inspect --mode full loads full workbook detail', async () => {
		const { stdout, exitCode } = await run('inspect', TEST_FILE, '--mode', 'full', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data.load.mode).toBe('full')
		expect(parsed.data.cellCount).toBe(0)
		expect(parsed.data.commentCount).toBe(0)
	})

	test('inspect --mode values reports rich sheet metadata as unknown', async () => {
		const { stdout, exitCode } = await run('inspect', TEST_FILE, '--mode', 'values', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data.load.mode).toBe('values')
		expect(parsed.data.load.richSheetMetadataHydrated).toBe(false)
		expect(parsed.data.commentCount).toBeNull()
	})

	test('inspect --detail pivots --json returns pivot inventory', async () => {
		const { exitCode, stdout } = await run(
			'inspect',
			PIVOT_CORPUS_FILE,
			'--detail',
			'pivots',
			'--json',
		)
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data.pivotTables.length).toBeGreaterThan(0)
		expect(parsed.data.pivotCaches.length).toBeGreaterThan(0)
	})

	test('inspect --detail slicers --json returns slicer inventory', { timeout: 10000 }, async () => {
		const { exitCode, stdout } = await run(
			'inspect',
			SLICER_CORPUS_FILE,
			'--detail',
			'slicers',
			'--json',
		)
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data.slicerCaches.length).toBeGreaterThan(0)
		expect(parsed.data.slicers.length).toBeGreaterThan(0)
	})

	test('inspect --detail drawings --json returns drawing flags', async () => {
		const { exitCode, stdout } = await run(
			'inspect',
			PIVOT_CORPUS_FILE,
			'Source data',
			'--detail',
			'drawings',
			'--json',
		)
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(typeof parsed.data.drawingRefs.hasDrawing).toBe('boolean')
	})

	test('inspect --detail names --json returns defined name inventory', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A3' }])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('inspect', TEST_FILE, '--detail', 'names', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data[0].name).toBe('Budget')
		expect(parsed.data[0].references[0].text).toBe('Sheet1!A1:A3')
	})

	test('inspect --detail views --json returns workbook view inventory', async () => {
		const wb = AscendWorkbook.create()
		const internal = wb as unknown as { wb: { workbookViews: Array<Record<string, unknown>> } }
		internal.wb.workbookViews.push({ activeTab: 0, visibility: 'visible' })
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('inspect', TEST_FILE, '--detail', 'views', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data[0].activeTab).toBe(0)
	})

	test('inspect --detail external-refs --json returns workbook external reference inventory', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run(
			'inspect',
			TEST_FILE,
			'--detail',
			'external-refs',
			'--json',
		)
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(Array.isArray(parsed.data)).toBe(true)
	})

	test('read requires an explicit sheet when workbook has multiple sheets', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'main' }] },
			{ op: 'addSheet', name: 'Archive' },
			{ op: 'setCells', sheet: 'Archive', updates: [{ ref: 'A1', value: 'extra' }] },
		])
		await wb.save(`${import.meta.dir}/${MULTI_SHEET_FILE}`)

		const { exitCode, stderr } = await run('read', MULTI_SHEET_FILE, 'A1')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Multiple sheets available')
	})

	test('read accepts explicit sheet selectors', async () => {
		const { exitCode, stdout } = await run('read', TEST_FILE, 'Sheet1!A1')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('A')
	})

	test('read --json returns versioned machine envelope', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'alpha' }] }])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('read', TEST_FILE, 'Sheet1!A1', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.formatVersion).toBe(1)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.cells).toHaveLength(1)
		expect(parsed.data.cells[0].ref).toBe('A1')
	})

	test('read accepts --display and rejects invalid pagination integers', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 2 },
					{ ref: 'A2', value: 3 },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1*2' },
		])
		await wb.recalc()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const display = await run('read', TEST_FILE, 'Sheet1!A1:B1', '--display')
		expect(display.exitCode).toBe(0)

		const badOffset = await run('read', TEST_FILE, 'Sheet1!A1:B1', '--row-offset', 'nope')
		expect(badOffset.exitCode).toBe(1)
		expect(badOffset.stderr).toContain('Invalid --row-offset')

		const badLimit = await run('read', TEST_FILE, 'Sheet1!A1:B1', '--row-limit', '0')
		expect(badLimit.exitCode).toBe(1)
		expect(badLimit.stderr).toContain('Invalid --row-limit')
	})

	test('read rejects unsupported modes', async () => {
		const { exitCode, stderr } = await run('read', TEST_FILE, 'Sheet1!A1', '--mode', 'metadata')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Invalid --mode')
	})

	test('read supports named range selectors', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'alpha' }] },
			{ op: 'setDefinedName', name: 'MyRange', ref: 'Sheet1!A1:A1' },
		])
		await wb.save(`${import.meta.dir}/${NAMED_RANGE_FILE}`)

		const { exitCode, stdout } = await run('read', NAMED_RANGE_FILE, 'name:MyRange')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('MyRange: Sheet1!A1:A1')
	})

	test('read supports named range selectors with anchored refs', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'alpha' }] },
			{ op: 'setDefinedName', name: 'AnchoredRange', ref: 'Sheet1!$A$1:$A$1' },
		])
		await wb.save(`${import.meta.dir}/${NAMED_RANGE_FILE}`)

		const { exitCode, stdout } = await run('read', NAMED_RANGE_FILE, 'name:AnchoredRange')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('AnchoredRange: Sheet1!$A$1:$A$1')
		expect(stdout).toContain('alpha')
	})

	test('read name --json exposes parsed name metadata', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A3' }])
		await wb.save(`${import.meta.dir}/${NAMED_RANGE_FILE}`)

		const { exitCode, stdout } = await run('read', NAMED_RANGE_FILE, 'name:Budget', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data.kind).toBe('name')
		expect(parsed.data.normalizedFormula).toBe('Sheet1!A1:A3')
		expect(parsed.data.references[0].text).toBe('Sheet1!A1:A3')
		expect(parsed.data.resolutionKind).toBe('range')
	})

	test('read table --json exposes table metadata and paginated rows', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Name' },
					{ ref: 'B1', value: 'Score' },
					{ ref: 'A2', value: 'Alice' },
					{ ref: 'B2', value: 90 },
					{ ref: 'A3', value: 'Bob' },
					{ ref: 'B3', value: 80 },
					{ ref: 'A4', value: 'Total' },
					{ ref: 'B4', value: 170 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B3', name: 'MyTable', hasHeaders: true },
		])
		const internal = wb as unknown as {
			wb: { sheets: Array<{ tables: Array<Record<string, unknown>> }> }
		}
		const tableModel = internal.wb.sheets[0]?.tables[0] as
			| {
					sortState?: { ref: string; conditions: readonly { ref: string }[] }
					ref: { start: { row: number; col: number }; end: { row: number; col: number } }
					hasTotals: boolean
			  }
			| undefined
		if (tableModel) {
			tableModel.sortState = { ref: 'A1:B4', conditions: [{ ref: 'B2:B3' }] }
			tableModel.hasTotals = true
			tableModel.ref = { start: { row: 0, col: 0 }, end: { row: 3, col: 1 } }
		}
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run(
			'read',
			TEST_FILE,
			'table:MyTable',
			'--json',
			'--row-offset',
			'1',
			'--row-limit',
			'1',
		)
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data.kind).toBe('table')
		expect(parsed.data.hasHeaders).toBe(true)
		expect(parsed.data.hasTotals).toBe(true)
		expect(parsed.data.headerRow[0]).toEqual({ kind: 'string', value: 'Name' })
		expect(parsed.data.totalsRow[1]).toEqual({ kind: 'number', value: 170 })
		expect(parsed.data.sortState.ref).toBe('A1:B4')
		expect(parsed.data.page.rowOffset).toBe(1)
		expect(parsed.data.page.rowLimit).toBe(1)
		expect(parsed.data.page.returnedRows).toBe(1)
		expect(parsed.data.page.totalRows).toBe(2)
		expect(parsed.data.page.hasMore).toBe(false)
		expect(parsed.data.rows).toHaveLength(1)
		expect(parsed.data.rows[0].Name).toEqual({ kind: 'string', value: 'Bob' })
	})

	test('write requires an explicit sheet when workbook has multiple sheets', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'main' }] },
			{ op: 'addSheet', name: 'Archive' },
		])
		await wb.save(`${import.meta.dir}/${MULTI_SHEET_FILE}`)

		const { exitCode, stderr } = await run('write', MULTI_SHEET_FILE, 'A1', '123')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Multiple sheets available')
	})

	test('write accepts explicit sheet selectors and auto-recalculates', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*2' },
		])
		await wb.recalc()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode } = await run('write', TEST_FILE, 'Sheet1!A1', '5')
		expect(exitCode).toBe(0)

		const reopened = await AscendWorkbook.open(`${import.meta.dir}/${TEST_FILE}`)
		expect(reopened.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 10 })
	})

	test('write fails when recalculation reports errors', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const opsFile = `${import.meta.dir}/write-ops.json`
		await Bun.write(
			opsFile,
			JSON.stringify([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }]),
		)
		const { exitCode, stderr } = await run('write', TEST_FILE, '--ops', 'write-ops.json')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Circular reference detected')
		if (existsSync(opsFile)) unlinkSync(opsFile)
	})

	test('write --json returns a failure envelope on operation errors', async () => {
		const tempFile = `${import.meta.dir}/${MULTI_SHEET_FILE}`
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Sheet2' }])
		await wb.save(tempFile)

		const opsFile = `${import.meta.dir}/write-ops.json`
		await Bun.write(
			opsFile,
			JSON.stringify([{ op: 'setCells', sheet: 'Missing', updates: [{ ref: 'A1', value: 1 }] }]),
		)

		const { exitCode, stdout } = await run(
			'write',
			MULTI_SHEET_FILE,
			'--ops',
			'write-ops.json',
			'--json',
		)
		expect(exitCode).toBe(1)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(false)
		expect(parsed.error.message).toContain('Sheet')
		if (existsSync(opsFile)) unlinkSync(opsFile)
	})

	test('write --json returns a failure envelope when recalculation reports errors', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const opsFile = `${import.meta.dir}/write-ops.json`
		await Bun.write(
			opsFile,
			JSON.stringify([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }]),
		)

		const { exitCode, stdout } = await run('write', TEST_FILE, '--ops', 'write-ops.json', '--json')
		expect(exitCode).toBe(1)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(false)
		expect(parsed.error.message).toContain('Circular reference detected')
		expect(parsed.error.details.recalc.errors.length).toBeGreaterThan(0)
		if (existsSync(opsFile)) unlinkSync(opsFile)
	})

	test('preview shows changes without mutating the workbook file', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'A1*2' },
		])
		await wb.recalc()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('preview', TEST_FILE, 'Sheet1!A1', '5', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.formatVersion).toBe(1)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.cellChanges.length).toBeGreaterThan(0)
		expect(parsed.data.writePlan.totalParts).toBeGreaterThan(0)

		const reopened = await AscendWorkbook.open(`${import.meta.dir}/${TEST_FILE}`)
		expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 2 })
		expect(reopened.sheet('Sheet1')?.cell('A2')?.value).toEqual({ kind: 'number', value: 4 })
	})

	test('preview --json returns a failure envelope when recalculation reports errors', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const opsFile = `${import.meta.dir}/preview-ops.json`
		await Bun.write(
			opsFile,
			JSON.stringify([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }]),
		)

		const { exitCode, stdout } = await run(
			'preview',
			TEST_FILE,
			'--ops',
			'preview-ops.json',
			'--json',
		)
		expect(exitCode).toBe(1)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(false)
		expect(parsed.error.message).toContain('Circular reference detected')
		expect(parsed.error.details.preview.errors.length).toBeGreaterThan(0)

		const reopened = await AscendWorkbook.open(`${import.meta.dir}/${TEST_FILE}`)
		expect(reopened.sheet('Sheet1')?.cell('A1')).toBeUndefined()
		if (existsSync(opsFile)) unlinkSync(opsFile)
	})

	test('formula show returns parsed formula info', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=SUM(B1:B2)' }])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('formula', 'show', TEST_FILE, 'Sheet1!A1')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('Normalized')
		expect(stdout).toContain('SUM')
		expect(stdout).toContain('References')
		expect(stdout).toContain('B1:B2')
	})

	test('formula set fails when recalculation reports errors', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stderr } = await run('formula', 'set', TEST_FILE, 'Sheet1!A1', '=A1+1')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Circular reference detected')
	})

	test('formula set --json returns a failure envelope when recalculation reports errors', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run(
			'formula',
			'set',
			TEST_FILE,
			'Sheet1!A1',
			'=A1+1',
			'--json',
		)
		expect(exitCode).toBe(1)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(false)
		expect(parsed.error.message).toContain('Circular reference detected')
		expect(parsed.error.details.recalc.errors.length).toBeGreaterThan(0)
	})

	test('trace shows values and respects max depth', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1*2' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: '=B1+1' },
		])
		wb.recalc()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('trace', TEST_FILE, 'Sheet1!C1', '--max-depth', '1')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('Value')
		expect(stdout).toContain('[1] Sheet1!B1')
		expect(stdout).not.toContain('Sheet1!A1')
	})

	test('trace rejects invalid max depth values', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=1+1' }])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stderr } = await run('trace', TEST_FILE, 'Sheet1!A1', '--max-depth', 'bad')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Invalid --max-depth')
	})

	test('formula suggests the closest subcommand', async () => {
		const { exitCode, stderr } = await run('formula', 'sho', TEST_FILE, 'Sheet1!A1')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Unknown formula subcommand: sho')
		expect(stderr).toContain('Did you mean "show"?')
	})

	test('doctor rejects unsupported flags', async () => {
		const { exitCode, stderr } = await run('doctor', '--verbose')
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Unknown flag for "doctor": --verbose')
	})

	test('calc exits nonzero when recalculation reports errors', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stderr } = await run('calc', TEST_FILE)
		expect(exitCode).toBe(1)
		expect(stderr).toContain('Circular reference detected')
	})

	test('calc --json returns a failure envelope when recalculation reports errors', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=A1+1' }])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('calc', TEST_FILE, '--json')
		expect(exitCode).toBe(1)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(false)
		expect(parsed.error.message).toContain('Circular reference detected')
		expect(parsed.error.details.recalc.errors.length).toBeGreaterThan(0)
	})

	test('formula set and fill edit formulas from the CLI', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 2 },
					{ ref: 'A2', value: 3 },
				],
			},
		])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		expect((await run('formula', 'set', TEST_FILE, 'Sheet1!B1', '=A1*2')).exitCode).toBe(0)
		expect((await run('formula', 'fill', TEST_FILE, 'Sheet1!B1:B2', '=A1*2')).exitCode).toBe(0)

		const reopened = await AscendWorkbook.open(`${import.meta.dir}/${TEST_FILE}`)
		expect(reopened.sheet('Sheet1')?.cell('B1')?.formula).toBe('A1*2')
		expect(reopened.sheet('Sheet1')?.cell('B2')?.formula).toBe('A2*2')
		expect(reopened.sheet('Sheet1')?.cell('B2')?.value).toEqual({ kind: 'number', value: 6 })
	})

	test('doctor runs without error', async () => {
		const { exitCode, stdout } = await run('doctor')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('ascend doctor')
		expect(stdout).toContain('[+] bun')
	})

	test('check on fresh workbook passes', async () => {
		const { exitCode, stdout } = await run('check', TEST_FILE)
		expect(exitCode).toBe(0)
		expect(stdout).toContain('all checks passed')
	})

	test('check --json returns versioned machine envelope', async () => {
		const { exitCode, stdout } = await run('check', TEST_FILE, '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.formatVersion).toBe(1)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.valid).toBe(true)
	})

	test('export writes TSV output and rejects unsupported formats', async () => {
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
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const tsv = await run('export', TEST_FILE, 'exported.tsv')
		expect(tsv.exitCode).toBe(0)
		const tsvText = await Bun.file(`${import.meta.dir}/exported.tsv`).text()
		expect(tsvText).toContain('Name\tScore')
		expect(tsvText).toContain('Alice\t10')

		const bad = await run('export', TEST_FILE, 'out.weird', '--format', 'weird')
		expect(bad.exitCode).toBe(1)
		expect(bad.stderr).toContain('Invalid export format')
	})
})
