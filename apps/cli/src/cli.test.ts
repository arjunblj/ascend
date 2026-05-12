import { afterAll, describe, expect, test } from 'bun:test'
import { existsSync, unlinkSync } from 'node:fs'
import { inspect as inspectValue } from 'node:util'
import { AscendWorkbook } from '@ascend/sdk'
import { runCli } from './index.ts'

const CLI = new URL('./index.ts', import.meta.url).pathname
const TEST_FILE = 'test-output.xlsx'
const MULTI_SHEET_FILE = 'test-multi.xlsx'
const NAMED_RANGE_FILE = 'test-named.xlsx'
const TUI_TEST_FILE = 'test-tui.xlsx'
const PIVOT_CORPUS_FILE = '../../../research/excel-corpus/ms-excel-formulas-and-pivot-tables.xlsx'
const SLICER_CORPUS_FILE = '../../../research/excel-corpus/excel-dashboard-v2.xlsx'
const HAS_PIVOT_CORPUS_FILE = existsSync(`${import.meta.dir}/${PIVOT_CORPUS_FILE}`)
const HAS_SLICER_CORPUS_FILE = existsSync(`${import.meta.dir}/${SLICER_CORPUS_FILE}`)

interface CliRunResult {
	stdout: string
	stderr: string
	exitCode: number
}

function runProcess(...args: string[]): Promise<CliRunResult> {
	return new Promise((resolve) => {
		const proc = Bun.spawn([Bun.argv[0], CLI, ...args], {
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

async function run(...args: string[]): Promise<CliRunResult> {
	const cwd = process.cwd()
	const originalLog = console.log
	const originalError = console.error
	const originalStderrWrite = process.stderr.write
	const stdout: string[] = []
	const stderr: string[] = []

	console.log = (...values: unknown[]) => {
		stdout.push(formatConsole(values))
	}
	console.error = (...values: unknown[]) => {
		stderr.push(formatConsole(values))
	}
	process.stderr.write = ((
		chunk: string | Uint8Array,
		encodingOrCallback?: unknown,
		callback?: unknown,
	) => {
		stderr.push(typeof chunk === 'string' ? chunk : Buffer.from(chunk).toString())
		const done = typeof encodingOrCallback === 'function' ? encodingOrCallback : callback
		if (typeof done === 'function') done()
		return true
	}) as typeof process.stderr.write

	try {
		process.chdir(import.meta.dir)
		const exitCode = await runCli(args)
		return {
			stdout: stdout.join('\n').trim(),
			stderr: stderr.join('\n').trim(),
			exitCode,
		}
	} finally {
		process.chdir(cwd)
		console.log = originalLog
		console.error = originalError
		process.stderr.write = originalStderrWrite
	}
}

function formatConsole(values: readonly unknown[]): string {
	return values
		.map((value) => (typeof value === 'string' ? value : inspectValue(value, { colors: false })))
		.join(' ')
}

afterAll(() => {
	for (const f of [
		TEST_FILE,
		MULTI_SHEET_FILE,
		NAMED_RANGE_FILE,
		TUI_TEST_FILE,
		'exported.tsv',
		'exported.json',
		'plan-ops.json',
		'commit-ops.json',
		'commit-output.xlsx',
		'progress-ops.json',
		'progress-output.xlsx',
	]) {
		const path = `${import.meta.dir}/${f}`
		if (existsSync(path)) unlinkSync(path)
	}
})

function parseJsonl(text: string): Array<Record<string, unknown>> {
	return text
		.split('\n')
		.map((line) => line.trim())
		.filter((line) => line.length > 0)
		.map((line) => JSON.parse(line))
}

function parseTrailingTelemetry(text: string): unknown {
	const start = text.lastIndexOf('\n[')
	return JSON.parse(text.slice(start >= 0 ? start + 1 : 0))
}

describe('ascend cli', () => {
	test('--version prints version through the executable boundary', async () => {
		const { stdout, exitCode } = await runProcess('--version')
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

	test('tui renders a non-TTY first frame and telemetry JSON', async () => {
		const { stdout, exitCode } = await run('tui', '--telemetry-json')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('Ascend')
		expect(stdout).toContain('File  Home')
		const telemetry = parseTrailingTelemetry(stdout)
		expect(Array.isArray(telemetry)).toBe(true)
		expect((telemetry as unknown[]).length).toBeGreaterThan(0)
	})

	test('tui accepts renderer selection and startup calibration flags', async () => {
		const { stdout, exitCode } = await run('tui', '--renderer', 'ansi', '--calibrate')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('Terminal Calibration')
		expect(stdout).toContain('Keyboard:')
	})

	test('tui boolean flags do not consume a following workbook path', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'loaded' }] }])
		await wb.save(`${import.meta.dir}/${TUI_TEST_FILE}`)

		const { stdout, exitCode } = await run('tui', '--calibrate', TUI_TEST_FILE, '--renderer=ansi')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('Terminal Calibration')
		expect(stdout).toContain('loaded')
	})

	test('open forwards --sheet into the TUI entrypoint', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'addSheet', name: 'Data' },
			{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 'data' }] },
		])
		await wb.save(`${import.meta.dir}/${TUI_TEST_FILE}`)

		const { stdout, exitCode } = await run('open', TUI_TEST_FILE, '--sheet', 'Data')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('data')
		expect(stdout).toContain('Data')
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

	test('docs --json searches bundled agent docs', async () => {
		const { stdout, exitCode } = await run('docs', 'plan', 'commit', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(true)
		expect(
			parsed.data.results.some((result: { path: string }) => result.path.includes('llms')),
		).toBe(true)
	})

	test('docs can list and read bundled docs', async () => {
		const listed = await run('docs', '--list', '--json')
		expect(listed.exitCode).toBe(0)
		const list = JSON.parse(listed.stdout)
		expect(list.data.docs.some((doc: { path: string }) => doc.path === 'llms.txt')).toBe(true)

		const read = await run('docs', '--path', 'llms.txt', '--json')
		expect(read.exitCode).toBe(0)
		const doc = JSON.parse(read.stdout)
		expect(doc.data.text).toContain('Ascend')
	})

	test('docs --examples searches examples only', async () => {
		const { stdout, exitCode } = await run('docs', 'mcp', 'setup', '--examples', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.results.every((result: { kind: string }) => result.kind === 'example')).toBe(
			true,
		)
		expect(
			parsed.data.results.some(
				(result: { path: string }) => result.path === 'examples/mcp-setup.md',
			),
		).toBe(true)
	})

	test('agent-init prints the canonical agent workflow contract', async () => {
		const { stdout, exitCode } = await run('agent-init', '--json')
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.commands.docs).toContain('ascend docs')
		expect(parsed.data.mcpResources).toContain('ascend://llms.txt')
		expect(parsed.data.commands.plan).toContain('--progress jsonl')
		expect(parsed.data.mcpResources).toContain('ascend://operations')
		expect(parsed.data.safetyDefaults.join('\n')).toContain('--expect-sha256')
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

	test('plan, commit, and check can emit JSONL progress events', async () => {
		const opsFile = `${import.meta.dir}/progress-ops.json`
		await Bun.write(
			opsFile,
			JSON.stringify([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 7 }] }]),
		)

		const plan = await run(
			'plan',
			TEST_FILE,
			'--ops',
			'progress-ops.json',
			'--progress',
			'jsonl',
			'--json',
		)
		expect(plan.exitCode).toBe(0)
		const planEvents = parseJsonl(plan.stderr)
		expect(planEvents[0]).toMatchObject({
			type: 'progress',
			kind: 'plan',
			phase: 'hash-input',
			status: 'started',
		})
		expect(planEvents.some((event) => event.phase === 'finalize')).toBe(true)

		const planned = JSON.parse(plan.stdout)
		const commit = await run(
			'commit',
			TEST_FILE,
			'--ops',
			'progress-ops.json',
			'--output',
			'progress-output.xlsx',
			'--expect-sha256',
			planned.data.inputSha256,
			'--progress',
			'jsonl',
			'--json',
		)
		expect(commit.exitCode).toBe(0)
		const commitEvents = parseJsonl(commit.stderr)
		expect(commitEvents.some((event) => event.kind === 'commit' && event.phase === 'write')).toBe(
			true,
		)

		const check = await run('check', 'progress-output.xlsx', '--progress', 'jsonl', '--json')
		expect(check.exitCode).toBe(0)
		const checkEvents = parseJsonl(check.stderr)
		expect(checkEvents.at(-1)).toMatchObject({
			type: 'progress',
			kind: 'check',
			phase: 'check',
			status: 'ok',
		})
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
		expect(parsed.data.capabilityWarnings).toEqual([])
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

	test.skipIf(!HAS_PIVOT_CORPUS_FILE)(
		'inspect --detail pivots --json returns pivot inventory',
		{ timeout: 30_000 },
		async () => {
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
			expect(parsed.data.pivotOutputAudits.length).toBeGreaterThan(0)
			expect(parsed.data.pivotOutputMaterializePlan.ops).toBeArray()
			expect(parsed.data.pivotOutputMaterializePlan.unsupported).toBeArray()
			expect(typeof parsed.data.pivotOutputMaterializePlan.plannedCellCount).toBe('number')
			expect(parsed.data.pivotRefreshPlans.length).toBe(parsed.data.pivotCaches.length)
			expect(parsed.data.pivotRefreshPlans[0].canRefreshHeadlessly).toBe(false)
			expect(parsed.data.pivotRefreshPlans[0].cacheRecords).toMatchObject({
				partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
				declaredCount: 4115,
				parsedCount: 4115,
			})
		},
	)

	test.skipIf(!HAS_PIVOT_CORPUS_FILE)(
		'inspect --json reports registry-backed capability warnings',
		{ timeout: 30_000 },
		async () => {
			const { exitCode, stdout } = await run(
				'inspect',
				PIVOT_CORPUS_FILE,
				'--mode',
				'full',
				'--json',
			)
			expect(exitCode).toBe(0)
			const parsed = JSON.parse(stdout)
			expect(parsed.data.capabilityWarnings).toContainEqual(
				expect.objectContaining({
					capabilityId: 'analytics.pivots',
					status: 'inspectable',
				}),
			)
		},
	)

	test.skipIf(!HAS_SLICER_CORPUS_FILE)(
		'inspect --detail slicers --json returns slicer inventory',
		{ timeout: 30_000 },
		async () => {
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
			expect(Array.isArray(parsed.data.timelineCaches)).toBe(true)
			expect(Array.isArray(parsed.data.timelines)).toBe(true)
		},
	)

	test.skipIf(!HAS_SLICER_CORPUS_FILE)(
		'inspect --detail compatibility --json reports preserved package features',
		{ timeout: 30_000 },
		async () => {
			const { exitCode, stdout } = await run(
				'inspect',
				SLICER_CORPUS_FILE,
				'--detail',
				'compatibility',
				'--json',
			)
			expect(exitCode).toBe(0)
			const parsed = JSON.parse(stdout)
			const features = parsed.data.features.map((feature: { feature: string }) => feature.feature)
			expect(parsed.data.status).toBe('has-preserved')
			expect(features).toContain('preservedPivot')
			expect(features).toContain('preservedSlicer')
		},
	)

	test.skipIf(!HAS_PIVOT_CORPUS_FILE)(
		'inspect --detail drawings --json returns drawing flags',
		async () => {
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
		},
	)

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
		expect(Array.isArray(parsed.data.references)).toBe(true)
		expect(Array.isArray(parsed.data.details)).toBe(true)
		expect(Array.isArray(parsed.data.usages)).toBe(true)
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

	test(
		'check surfaces structured issue metadata for agent repair',
		async () => {
			const wb = AscendWorkbook.create()
			wb.apply([{ op: 'renameSheet', sheet: 'Sheet1', newName: 'SummaryData' }])
			wb.apply([{ op: 'setFormula', sheet: 'SummaryData', ref: 'A1', formula: '=Summary!B1' }])
			await wb.save(`${import.meta.dir}/${TEST_FILE}`)

			const json = await run('check', TEST_FILE, '--json')
			expect(json.exitCode).toBe(2)
			const parsed = JSON.parse(json.stdout)
			expect(parsed.ok).toBe(true)
			expect(parsed.data.valid).toBe(false)
			const issue = parsed.data.issues.find(
				(entry: { rule?: string }) => entry.rule === 'broken-refs',
			)
			expect(issue.ref).toBe('SummaryData!A1')
			expect(issue.refs).toEqual(['SummaryData!A1'])
			expect(issue.suggestedFix).toContain('SummaryData')

			const pretty = await run('check', TEST_FILE)
			expect(pretty.exitCode).toBe(2)
			expect(pretty.stdout).toContain('broken-refs')
			expect(pretty.stdout).toContain('Suggested Fix')
		},
		{ timeout: 15_000 },
	)

	test(
		'export writes TSV output and rejects unsupported formats',
		async () => {
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
		},
		{ timeout: 15_000 },
	)

	test(
		'lint on clean workbook passes',
		async () => {
			const { exitCode, stdout } = await run('lint', TEST_FILE)
			expect(exitCode).toBe(0)
			expect(stdout).toContain('no lint warnings')
		},
		{ timeout: 15_000 },
	)

	test(
		'lint --json returns machine envelope',
		async () => {
			const { exitCode, stdout } = await run('lint', TEST_FILE, '--json')
			expect(exitCode).toBe(0)
			const parsed = JSON.parse(stdout)
			expect(parsed.ok).toBe(true)
		},
		{ timeout: 15_000 },
	)

	test('trace shows precedents for formula cell', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1*2' },
		])
		wb.recalc()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('trace', TEST_FILE, 'Sheet1!B1')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('A1')
	})

	test('diff two workbooks shows changes', async () => {
		const wb1 = AscendWorkbook.create()
		wb1.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
		await wb1.save(`${import.meta.dir}/${TEST_FILE}`)

		const wb2 = AscendWorkbook.create()
		wb2.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 99 }] }])
		await wb2.save(`${import.meta.dir}/${MULTI_SHEET_FILE}`)

		const { exitCode, stdout } = await run('diff', TEST_FILE, MULTI_SHEET_FILE)
		expect(exitCode).toBe(0)
		expect(stdout).toContain('Sheet1')
	})

	test('calc recalculates formulas', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 10 }] },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1+5' },
		])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('calc', TEST_FILE)
		expect(exitCode).toBe(0)
		expect(stdout.toLowerCase()).toContain('recalculated')
	})

	test('list shows sheet names', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'MySheet' }])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('list', TEST_FILE)
		expect(exitCode).toBe(0)
		expect(stdout).toContain('MySheet')
	})

	test('find searches for values', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'hello' },
					{ ref: 'A2', value: 'world' },
				],
			},
		])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('find', TEST_FILE, 'hello')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('A1')
	})

	test('export --json writes JSON output', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'x' },
					{ ref: 'B1', value: 1 },
				],
			},
		])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode } = await run('export', TEST_FILE, 'exported.json', '--format', 'json')
		expect(exitCode).toBe(0)
		const jsonText = await Bun.file(`${import.meta.dir}/exported.json`).text()
		const parsed = JSON.parse(jsonText)
		expect(parsed).toBeDefined()
	})
})
