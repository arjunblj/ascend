import { afterAll, describe, expect, test } from 'bun:test'
import { existsSync, unlinkSync } from 'node:fs'
import { AscendWorkbook } from '@ascend/sdk'

const CLI = new URL('./index.ts', import.meta.url).pathname
const TEST_FILE = 'test-output.xlsx'
const MULTI_SHEET_FILE = 'test-multi.xlsx'
const NAMED_RANGE_FILE = 'test-named.xlsx'

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
	for (const f of [TEST_FILE, MULTI_SHEET_FILE, NAMED_RANGE_FILE]) {
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

	test('create makes a workbook file', async () => {
		const { stdout, exitCode } = await run('create', TEST_FILE)
		expect(exitCode).toBe(0)
		expect(stdout).toContain('Created')
		expect(existsSync(`${import.meta.dir}/${TEST_FILE}`)).toBe(true)
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

	test('formula show returns parsed formula info', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=SUM(B1:B2)' }])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('formula', 'show', TEST_FILE, 'Sheet1!A1')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('Normalized')
		expect(stdout).toContain('SUM')
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
})
