import { afterAll, describe, expect, test } from 'bun:test'
import { existsSync, readFileSync, unlinkSync } from 'node:fs'
import { inspect as inspectValue } from 'node:util'
import { AscendWorkbook } from '@ascend/sdk'
import { makeXlsx } from '../../../packages/io-xlsx/test/helpers.ts'
import { runCli } from './index.ts'

const CLI = new URL('./index.ts', import.meta.url).pathname
const TEST_FILE = 'test-output.xlsx'
const DUMP_TEST_FILE = 'test-dump.xlsx'
const TEMPLATE_TEST_FILE = 'test-template.xlsx'
const APPROVAL_TEST_FILE = 'test-approval.xlsx'
const MULTI_SHEET_FILE = 'test-multi.xlsx'
const NAMED_RANGE_FILE = 'test-named.xlsx'
const TUI_TEST_FILE = 'test-tui.xlsx'
const ACTIVE_CONTENT_FILE = 'test-active-content.xlsm'
const TRUST_REPORT_FILE = 'test-trust-report.xlsm'
const OPEN_PLAN_FILE = 'test-open-plan.xlsx'
const AGENT_VIEW_FILE = 'test-agent-view.xlsx'
const JOURNAL_V1_OPS_FILE = 'journal-v1-ops.json'
const JOURNAL_V1_OUTPUT_FILE = 'journal-v1-output.xlsx'
const JOURNAL_BUILD_FAILURE_OPS_FILE = 'journal-build-failure-ops.json'
const PIVOT_CORPUS_FILE = '../../../research/excel-corpus/ms-excel-formulas-and-pivot-tables.xlsx'
const SLICER_CORPUS_FILE = '../../../research/excel-corpus/excel-dashboard-v2.xlsx'
const HAS_PIVOT_CORPUS_FILE = existsSync(`${import.meta.dir}/${PIVOT_CORPUS_FILE}`)
const HAS_SLICER_CORPUS_FILE = existsSync(`${import.meta.dir}/${SLICER_CORPUS_FILE}`)
const JOURNAL_V1_FIXTURE = JSON.parse(
	readFileSync(`${import.meta.dir}/../../../fixtures/journal/mutation-journal-v1.json`, 'utf-8'),
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

function compactJournal(journal: {
	readonly schemaVersion: number
	readonly schemaId: string
	readonly supported: boolean
	readonly exact: boolean
	readonly inverseOps: readonly unknown[]
	readonly issues: readonly unknown[]
}) {
	return {
		schemaVersion: journal.schemaVersion,
		schemaId: journal.schemaId,
		supported: journal.supported,
		exact: journal.exact,
		inverseOpCount: journal.inverseOps.length,
		issueCount: journal.issues.length,
		issues: journal.issues,
	}
}

afterAll(() => {
	for (const f of [
		TEST_FILE,
		DUMP_TEST_FILE,
		TEMPLATE_TEST_FILE,
		MULTI_SHEET_FILE,
		NAMED_RANGE_FILE,
		TUI_TEST_FILE,
		ACTIVE_CONTENT_FILE,
		TRUST_REPORT_FILE,
		OPEN_PLAN_FILE,
		AGENT_VIEW_FILE,
		APPROVAL_TEST_FILE,
		JOURNAL_V1_OPS_FILE,
		JOURNAL_V1_OUTPUT_FILE,
		JOURNAL_BUILD_FAILURE_OPS_FILE,
		'exported.tsv',
		'exported.json',
		'plan-ops.json',
		'plan-risk-ops.json',
		'invalid-agent-ops.json',
		'commit-ops.json',
		'commit-output.xlsx',
		'commit-compact-output.xlsx',
		'commit-pretty-output.xlsx',
		'approval-ops.json',
		'approval-alias-out.xlsx',
		'approval-exact-out.xlsx',
		'lossy-ops.json',
		'lossy-alias-out.xlsm',
		'lossy-exact-out.xlsm',
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
		expect(stdout).toContain('open-plan <file>')
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

	test('agent-view --tokens returns budget metadata', async () => {
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
		await wb.save(`${import.meta.dir}/${AGENT_VIEW_FILE}`)

		const { stdout, exitCode } = await run(
			'agent-view',
			AGENT_VIEW_FILE,
			'--range',
			'A1:D20',
			'--tokens',
			'384',
			'--json',
		)

		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.budget.requestedApproxTokens).toBe(384)
		expect(parsed.data.budget.truncated).toBe(true)
		expect(parsed.data.rowCount).toBe(20)
		expect(parsed.data.colCount).toBe(4)
		expect(
			parsed.data.budget.omittedSampleRows + parsed.data.budget.omittedColumnSampleValues,
		).toBeGreaterThan(0)
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

	test('tui accepts explicit preview rows for row-limited first paint', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'visible' },
					{ ref: 'A2', value: 'capped' },
				],
			},
		])
		await wb.save(`${import.meta.dir}/${TUI_TEST_FILE}`)

		const { stdout, exitCode } = await run('tui', TUI_TEST_FILE, '--preview-rows', '1')
		expect(exitCode).toBe(0)
		expect(stdout).toContain('visible')
		expect(stdout).toContain('first 1 rows')
		expect(stdout).toContain('read-only')
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

	test('open defaults file loads to a preview-first read-only window', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'preview' }] }])
		await wb.save(`${import.meta.dir}/${TUI_TEST_FILE}`)

		const { stdout, exitCode } = await run('open', TUI_TEST_FILE)
		expect(exitCode).toBe(0)
		expect(stdout).toContain('preview')
		expect(stdout).toContain('first 500 rows')
		expect(stdout).toContain('read-only')
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
		expect(parsed.data.commands.openPlan).toContain('open-plan')
		expect(parsed.data.commands.trust).toContain('inspect <file> --agent --json')
		expect(parsed.data.workflow.join('\n')).toContain('open-plan')
		expect(parsed.data.mcpResources).toContain('ascend://llms.txt')
		expect(parsed.data.commands.plan).toContain('--progress jsonl')
		expect(parsed.data.mcpResources).toContain('ascend://operations')
		expect(parsed.data.safetyDefaults.join('\n')).toContain('--expect-sha256')
		expect(parsed.data.safetyDefaults.join('\n')).toContain('planHandle')
		expect(parsed.data.safetyDefaults.join('\n')).toContain('untrusted data')
	})

	test('plan and commit implement safe agent workflow', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)
		const opsFile = `${import.meta.dir}/commit-ops.json`
		await Bun.write(
			opsFile,
			JSON.stringify([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'agent-safe' }] },
			]),
		)
		const plan = await run(
			'plan',
			TEST_FILE,
			'--ops',
			'commit-ops.json',
			'--package-actions',
			'--json',
		)
		expect(plan.exitCode).toBe(0)
		const planned = JSON.parse(plan.stdout)
		expect(planned.ok).toBe(true)
		expect(planned.data.inputSha256).toMatch(/^[a-f0-9]{64}$/)
		expect(planned.data.preview.wouldSucceed).toBe(true)
		expect(planned.data.packageActions.kind).toBe('ascend-package-action-proof')
		expect(planned.data.packageActions.byAction.regenerate).toBeGreaterThan(0)

		const commit = await run(
			'commit',
			TEST_FILE,
			'--ops',
			'commit-ops.json',
			'--output',
			'commit-output.xlsx',
			'--expect-sha256',
			planned.data.inputSha256,
			'--package-actions',
			'--json',
		)
		expect(commit.exitCode).toBe(0)
		const committed = JSON.parse(commit.stdout)
		expect(committed.ok).toBe(true)
		expect(committed.data.outputSha256).toMatch(/^[a-f0-9]{64}$/)
		expect(committed.data.packageActions.kind).toBe('ascend-package-action-proof')
		expect(committed.data.packageActions.byAction.regenerate).toBeGreaterThan(0)
		expect(committed.data.packageActions.coverage.sourceByteDigestCount).toBeGreaterThan(0)
		expect(committed.data.packageActions.coverage.outputByteDigestCount).toBeGreaterThan(0)
		expect(
			committed.data.packageActions.coverage.matchingByteDigestCount +
				committed.data.packageActions.coverage.mismatchedByteDigestCount,
		).toBeGreaterThan(0)
		expect(
			committed.data.packageActions.actions.some(
				(action: { outputSha256?: string }) => action.outputSha256 !== undefined,
			),
		).toBe(true)
		expect(existsSync(`${import.meta.dir}/commit-output.xlsx`)).toBe(true)

		const compactCommit = await run(
			'commit',
			TEST_FILE,
			'--ops',
			'commit-ops.json',
			'--output',
			'commit-compact-output.xlsx',
			'--expect-sha256',
			planned.data.inputSha256,
			'--compact',
			'--json',
		)
		expect(compactCommit.exitCode).toBe(0)
		const compact = JSON.parse(compactCommit.stdout)
		expect(compact.ok).toBe(true)
		expect(compact.data.outputSha256).toMatch(/^[a-f0-9]{64}$/)
		expect(compact.data.apply.affectedCellCount).toBe(1)
		expect(compact.data.check.valid).toBe(true)
		expect(compact.data.postWrite.valid).toBe(true)
		expect(compact.data.postWrite.reopened).toBe(true)
		expect(compact.data.postWrite.check.valid).toBe(true)
		expect(compact.data.postWrite.packageGraphAudit.ok).toBe(true)
		expect(compact.data.trace.artifactCount).toBeNumber()
		expect(compact.data.trace.artifacts).toBeUndefined()
		expect(compact.data.apply.affectedCells).toBeUndefined()
		expect(existsSync(`${import.meta.dir}/commit-compact-output.xlsx`)).toBe(true)

		const prettyCommit = await run(
			'commit',
			TEST_FILE,
			'--ops',
			'commit-ops.json',
			'--output',
			'commit-pretty-output.xlsx',
			'--expect-sha256',
			planned.data.inputSha256,
		)
		expect(prettyCommit.exitCode).toBe(0)
		expect(prettyCommit.stdout).toContain('Post-write package graph issues: 0')
	})

	test('plan and compact commit JSON preserve journal v1 issue compatibility', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)
		await Bun.write(
			`${import.meta.dir}/${JOURNAL_V1_OPS_FILE}`,
			JSON.stringify(JOURNAL_V1_FIXTURE.scenario.ops),
		)

		const planned = await run('plan', TEST_FILE, '--ops', JOURNAL_V1_OPS_FILE, '--json')

		expect(planned.exitCode).toBe(0)
		const parsed = JSON.parse(planned.stdout)
		expect(parsed.ok).toBe(true)
		expect(compactJournal(parsed.data.preview.journal)).toEqual(JOURNAL_V1_FIXTURE.scenario.journal)

		const committed = await run(
			'commit',
			TEST_FILE,
			'--ops',
			JOURNAL_V1_OPS_FILE,
			'--output',
			JOURNAL_V1_OUTPUT_FILE,
			'--compact',
			'--json',
		)

		expect(committed.exitCode).toBe(0)
		const committedParsed = JSON.parse(committed.stdout)
		expect(committedParsed.ok).toBe(true)
		expect(committedParsed.data.apply.journalSummary).toEqual(JOURNAL_V1_FIXTURE.scenario.journal)
	})

	test('plan JSON preserves structured journal build failures for agents', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)
		await Bun.write(
			`${import.meta.dir}/${JOURNAL_BUILD_FAILURE_OPS_FILE}`,
			JSON.stringify([{ op: 'clearRange', sheet: 'Sheet1', range: 'A1:', what: 'all' }]),
		)

		const planned = await run('plan', TEST_FILE, '--ops', JOURNAL_BUILD_FAILURE_OPS_FILE, '--json')

		expect(planned.exitCode).toBe(1)
		const parsed = JSON.parse(planned.stdout)
		expect(parsed.ok).toBe(false)
		expect(parsed.error.details.plan.preview.journal).toMatchObject({
			schemaVersion: JOURNAL_V1_FIXTURE.scenario.journal.schemaVersion,
			schemaId: JOURNAL_V1_FIXTURE.scenario.journal.schemaId,
			supported: false,
			exact: false,
			inverseOps: [],
			issues: [
				{
					code: 'JOURNAL_BUILD_FAILED',
					surface: 'package-parts',
					reason: 'journal-build-failed',
				},
			],
			undoPolicy: {
				undoable: false,
				exact: false,
				reason: 'build-failed',
				riskLevel: 'high',
			},
		})
	})

	test('plan invalid ops return structured batch repair details', async () => {
		await Bun.write(
			`${import.meta.dir}/invalid-agent-ops.json`,
			JSON.stringify([
				{ op: 'insertRows', sheet: 'Sheet1', at: 0, count: '2' },
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: { nested: true } }] },
				{ op: 'missingOp', sheet: 'Sheet1' },
			]),
		)

		const result = await run('plan', TEST_FILE, '--ops', 'invalid-agent-ops.json', '--json')
		expect(result.exitCode).toBe(1)
		const parsed = JSON.parse(result.stdout)
		expect(parsed.ok).toBe(false)
		expect(parsed.error.code).toBe('VALIDATION_ERROR')
		expect(parsed.error.details.issueCount).toBe(3)
		expect(parsed.error.details.issues).toEqual(
			expect.arrayContaining([
				'ops[0].count must be a positive integer',
				'ops[1].updates[0].value must be a scalar value or null',
				'ops[2].op "missingOp" is not supported',
			]),
		)
		expect(parsed.error.details.issueDetails).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ code: 'invalid_type', opIndex: 0, path: 'ops[0].count' }),
				expect.objectContaining({ code: 'invalid_type', opIndex: 1 }),
				expect.objectContaining({ code: 'invalid_operation', opIndex: 2, path: 'ops[2].op' }),
			]),
		)
	})

	test('dump --json emits a replayable operation batch', async () => {
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
		await wb.save(`${import.meta.dir}/${DUMP_TEST_FILE}`)

		const result = await run('dump', DUMP_TEST_FILE, '--json')
		expect(result.exitCode).toBe(0)
		const parsed = JSON.parse(result.stdout)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.replayable).toBe(true)
		expect(parsed.data.ops).toEqual([
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

	test('template-merge --json emits replayable operations and unresolved placeholders', async () => {
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
		await wb.save(`${import.meta.dir}/${TEMPLATE_TEST_FILE}`)

		const result = await run(
			'template-merge',
			TEMPLATE_TEST_FILE,
			'--data',
			'{"amount":10,"tax":2}',
			'--json',
		)
		expect(result.exitCode).toBe(2)
		const parsed = JSON.parse(result.stdout)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.replayable).toBe(false)
		expect(parsed.data.unresolved).toEqual([
			{
				sheet: 'Sheet1',
				ref: 'A2',
				source: 'value',
				placeholder: '{{client}}',
				key: 'client',
			},
		])
		expect(parsed.data.ops).toEqual([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'A1', value: 10 }],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1+2' },
		])
	})

	test('commit accepts only exact approval ids emitted by plan', async () => {
		const workbook = AscendWorkbook.create()
		workbook.apply([{ op: 'addSheet', name: 'Scratch' }])
		await workbook.save(`${import.meta.dir}/${APPROVAL_TEST_FILE}`)
		await Bun.write(
			`${import.meta.dir}/approval-ops.json`,
			JSON.stringify([{ op: 'deleteSheet', sheet: 'Scratch' }]),
		)

		const plan = await run('plan', APPROVAL_TEST_FILE, '--ops', 'approval-ops.json', '--json')
		expect(plan.exitCode).toBe(0)
		const planned = JSON.parse(plan.stdout)
		const approvalId = planned.data.approvals[0].id
		expect(approvalId).toBe('op:0:deletesheet')

		const aliasCommit = await run(
			'commit',
			APPROVAL_TEST_FILE,
			'--ops',
			'approval-ops.json',
			'--output',
			'approval-alias-out.xlsx',
			'--approval',
			'deleteSheet',
			'--json',
		)
		expect(aliasCommit.exitCode).toBe(1)
		const aliasFailure = JSON.parse(aliasCommit.stdout)
		expect(aliasFailure.ok).toBe(false)
		expect(aliasFailure.error.message).toBe('Commit requires explicit approval')

		const exactCommit = await run(
			'commit',
			APPROVAL_TEST_FILE,
			'--ops',
			'approval-ops.json',
			'--output',
			'approval-exact-out.xlsx',
			'--approval',
			approvalId,
			'--json',
		)
		expect(exactCommit.exitCode).toBe(0)
		const exact = JSON.parse(exactCommit.stdout)
		expect(exact.ok).toBe(true)
		expect(exact.data.approvals[0].id).toBe(approvalId)
	})

	test('plan pretty output exposes write-policy package part details', async () => {
		await Bun.write(`${import.meta.dir}/${ACTIVE_CONTENT_FILE}`, signedMacroWorkbook())
		await Bun.write(
			`${import.meta.dir}/plan-risk-ops.json`,
			JSON.stringify([
				{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 'risk-visible' }] },
			]),
		)

		const plan = await run('plan', ACTIVE_CONTENT_FILE, '--ops', 'plan-risk-ops.json')
		expect(plan.exitCode).toBe(0)
		expect(plan.stdout).toContain('Write policy diagnostics:')
		expect(plan.stdout).toContain('Write policy active-content-preserved:')
		expect(plan.stdout).toContain('active-content-preserved package parts:')
		expect(plan.stdout).toContain('xl/vbaProject.bin (preservedMacro')
		expect(plan.stdout).toContain('Write policy approval-required-feature:')
	})

	test('lossy preserved-feature commits require exact approval ids', async () => {
		await Bun.write(`${import.meta.dir}/${ACTIVE_CONTENT_FILE}`, signedMacroWorkbook())
		await Bun.write(
			`${import.meta.dir}/lossy-ops.json`,
			JSON.stringify([{ op: 'setCells', sheet: 'Data', updates: [{ ref: 'A1', value: 7 }] }]),
		)

		const plan = await run('plan', ACTIVE_CONTENT_FILE, '--ops', 'lossy-ops.json', '--json')
		expect(plan.exitCode).toBe(0)
		const planned = JSON.parse(plan.stdout)
		const approvalIds = planned.data.approvals.map((approval: { id: string }) => approval.id)
		expect(approvalIds).toEqual(
			expect.arrayContaining([
				expect.stringMatching(/^loss:preservedmacro:preserved:/),
				expect.stringMatching(/^loss:preservedsignature:preserved:/),
			]),
		)

		const aliasCommit = await run(
			'commit',
			ACTIVE_CONTENT_FILE,
			'--ops',
			'lossy-ops.json',
			'--output',
			'lossy-alias-out.xlsm',
			'--approval',
			'preservedMacro,preservedSignature',
			'--json',
		)
		expect(aliasCommit.exitCode).toBe(1)
		const aliasFailure = JSON.parse(aliasCommit.stdout)
		expect(aliasFailure.ok).toBe(false)
		expect(aliasFailure.error.message).toBe('Commit requires explicit approval')

		const exactCommit = await run(
			'commit',
			ACTIVE_CONTENT_FILE,
			'--ops',
			'lossy-ops.json',
			'--output',
			'lossy-exact-out.xlsm',
			'--approval',
			approvalIds.join(','),
			'--json',
		)
		expect(exactCommit.exitCode).toBe(0)
		const exact = JSON.parse(exactCommit.stdout)
		expect(exact.ok).toBe(true)
		expect(exact.data.approvals.map((approval: { id: string }) => approval.id)).toEqual(approvalIds)
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

	test('inspect --detail package-graph --json returns package graph identity', async () => {
		const { stdout, exitCode } = await run(
			'inspect',
			TEST_FILE,
			'--detail',
			'package-graph',
			'--json',
		)
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.formatVersion).toBe(1)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.parts).toContainEqual(
			expect.objectContaining({
				path: 'xl/workbook.xml',
				featureFamily: 'workbook',
				ownerScope: 'workbook',
				sourceRelationshipId: 'rId1',
			}),
		)
		expect(parsed.data.relationships).toContainEqual(
			expect.objectContaining({
				relationshipPartPath: '_rels/.rels',
				id: 'rId1',
				resolvedTarget: 'xl/workbook.xml',
			}),
		)
	})

	test('inspect --detail active-content --json returns metadata-only risk inventory', async () => {
		await Bun.write(`${import.meta.dir}/${ACTIVE_CONTENT_FILE}`, signedMacroWorkbook())

		const { stdout, exitCode } = await run(
			'inspect',
			ACTIVE_CONTENT_FILE,
			'--detail',
			'active-content',
			'--json',
		)
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.formatVersion).toBe(1)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.activeContentCount).toBe(2)
		expect(parsed.data.activeContent).toContainEqual(
			expect.objectContaining({
				kind: 'vbaProject',
				partPath: 'xl/vbaProject.bin',
				sourceRelationshipId: 'rIdVba',
			}),
		)
		expect(parsed.data.activeContent).toContainEqual(
			expect.objectContaining({
				kind: 'vbaSignature',
				partPath: 'xl/vbaProjectSignature.bin',
				sourcePartPath: 'xl/vbaProject.bin',
				sourceRelationshipId: 'rIdVbaSignature',
			}),
		)
		expect(parsed.data.compatibilityFeatures).toContainEqual(
			expect.objectContaining({
				feature: 'preservedMacro',
				locations: ['xl/vbaProject.bin'],
			}),
		)
		expect(parsed.data.compatibilityFeatures).toContainEqual(
			expect.objectContaining({
				feature: 'preservedSignature',
				locations: ['xl/vbaProjectSignature.bin'],
			}),
		)
	})

	test('inspect --agent --json returns an untrusted workbook trust report', async () => {
		await Bun.write(`${import.meta.dir}/${TRUST_REPORT_FILE}`, signedMacroWorkbook())

		const { stdout, exitCode } = await run('inspect', TRUST_REPORT_FILE, '--agent', '--json')

		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.formatVersion).toBe(1)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.trust).toBe('untrusted')
		expect(parsed.data.posture).toBe('safe-parser-preserver')
		expect(parsed.data.includedInAgentContext).toMatchObject({
			activeContent: false,
			hiddenSheets: false,
		})
		expect(parsed.data.executionPolicy).toMatchObject({
			macros: 'preserve-only',
			externalLinks: 'do-not-refresh',
		})
		expect(parsed.data.findings).toContainEqual(
			expect.objectContaining({ code: 'workbook.vbaProject' }),
		)
		expect(parsed.data).not.toHaveProperty('riskScore')
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

	test('open-plan --json recommends values for read intent', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 42 }] }])
		await Bun.write(`${import.meta.dir}/${OPEN_PLAN_FILE}`, wb.toBytes())

		const { stdout, exitCode } = await run(
			'open-plan',
			OPEN_PLAN_FILE,
			'--intent',
			'read-values',
			'--json',
		)
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.formatVersion).toBe(1)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.recommendedLoadOptions).toEqual({ mode: 'values' })
		expect(parsed.data.reviewBeforeHydration).toBe(false)
	})

	test('open-plan routes macro packages to metadata-only review', async () => {
		await Bun.write(`${import.meta.dir}/${ACTIVE_CONTENT_FILE}`, signedMacroWorkbook())

		const { stdout, exitCode } = await run('open-plan', ACTIVE_CONTENT_FILE, '--json')

		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data.intent).toBe('edit-plan')
		expect(parsed.data.recommendedLoadOptions).toEqual({ mode: 'metadata-only' })
		expect(parsed.data.reviewBeforeHydration).toBe(true)
		expect(parsed.data.riskFeatures).toContainEqual(
			expect.objectContaining({ featureFamily: 'preservedMacro' }),
		)
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

	test('read --json row-limit opens a capped partial view', async () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: Array.from({ length: 5 }, (_, index) => ({
					ref: `A${index + 1}`,
					value: index + 1,
				})),
			},
		])
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run(
			'read',
			TEST_FILE,
			'Sheet1!A1:A5',
			'--json',
			'--row-limit',
			'2',
		)
		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data.cells).toHaveLength(2)
		expect(parsed.data.load.isPartial).toBe(true)
		expect(parsed.data.load.maxRows).toBe(2)
		expect(parsed.data.load.partialReasons).toContain(
			'only the first 2 row(s) are hydrated per loaded sheet',
		)
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

	test('formula assist returns formula IDE help without opening a workbook', async () => {
		const { exitCode, stdout } = await run(
			'formula',
			'assist',
			'=SUM(A1:B2',
			'--cursor',
			'5',
			'--prefix',
			'SU',
			'--completion-limit',
			'3',
			'--function-name',
			'SUM',
			'--reference',
			'C3',
			'--replace-reference-at-cursor',
			'--cycle-reference',
			'--json',
		)

		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(true)
		expect(parsed.data.diagnostics.parseOk).toBe(false)
		expect(
			parsed.data.completions.some((completion: { name: string }) => completion.name === 'SUM'),
		).toBe(true)
		expect(parsed.data.signature.name).toBe('SUM')
		expect(parsed.data.signatureHelp.signature.name).toBe('SUM')
		expect(parsed.data.insertion.formula).toContain('C3')
		expect(parsed.data.cycle.changed).toBe(true)

		const refusal = await run(
			'formula',
			'assist',
			'=Budget+Sales[Amount]',
			'--cursor',
			'10',
			'--json',
		)
		expect(refusal.exitCode).toBe(0)
		const refusalData = JSON.parse(refusal.stdout)
		expect(refusalData.data.renameTarget).toMatchObject({
			ok: false,
			reason: 'workbook-context-required',
			role: { role: 'table-name-use', text: 'Sales' },
		})
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
			await Bun.write(`${import.meta.dir}/${TEST_FILE}`, threadedCommentMissingPersonsWorkbook())

			const json = await run('check', TEST_FILE, '--json')
			expect(json.exitCode).toBe(2)
			const parsed = JSON.parse(json.stdout)
			expect(parsed.ok).toBe(true)
			expect(parsed.data.valid).toBe(false)
			const issue = parsed.data.issues.find(
				(entry: { rule?: string }) => entry.rule === 'threaded-comment-integrity',
			)
			expect(issue.refs).toEqual(['Sheet1!A1'])
			expect(issue.details.kind).toBe('threaded-comment-unknown-person-id')
			expect(issue.suggestedFix).toContain('persons part')

			const pretty = await run('check', TEST_FILE)
			expect(pretty.exitCode).toBe(2)
			expect(pretty.stdout).toContain('threaded-comment-integrity')
			expect(pretty.stdout).toContain('Kind')
			expect(pretty.stdout).toContain('threaded-comment-unknown-person-id')
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
