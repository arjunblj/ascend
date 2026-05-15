import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'
import { practicalLatencyContractsTestHooks } from './practical-latency-contracts.ts'

const runnerPath = fileURLToPath(new URL('./practical-latency-contracts.ts', import.meta.url))

describe('practical latency contracts benchmark', () => {
	test('accepts --preset as a public-tracked input-preset alias', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--preset',
				'public-tracked',
				'--contract',
				'first-view',
				'--repeat',
				'1',
				'--warmup',
				'0',
				'--dry-run',
				'--json',
			],
			{ cwd: process.cwd(), stderr: 'pipe', stdout: 'pipe' },
		)
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])
		expect(exitCode, stderr).toBe(0)
		const payload = JSON.parse(stdout) as {
			readonly args?: {
				readonly inputPreset?: string
				readonly inputFile?: string
				readonly sheet?: string
				readonly range?: string
			}
		}
		expect(payload.args).toMatchObject({
			inputPreset: 'public-tracked',
			inputFile: 'fixtures/xlsx/calamine/issue_174.xlsx',
			sheet: 'Sheet1',
			range: 'A1:K65536',
		})
	})

	test('public-tracked preset separates tracked fixtures from generated edit input', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--input-preset',
				'public-tracked',
				'--contract',
				'edit-verify',
				'--repeat',
				'1',
				'--warmup',
				'0',
				'--dry-run',
				'--json',
			],
			{ cwd: process.cwd(), stderr: 'pipe', stdout: 'pipe' },
		)
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])
		expect(exitCode, stderr).toBe(0)
		const payload = JSON.parse(stdout) as {
			readonly args?: {
				readonly inputPreset?: string
				readonly inputFile?: string
				readonly tableInputFile?: string
				readonly editInputFile?: string
				readonly generatedEditInputLabel?: string
			}
			readonly inputs?: readonly {
				readonly role?: string
				readonly kind?: string
				readonly releaseClaimable?: boolean
			}[]
			readonly results?: readonly { readonly id?: string; readonly command?: string }[]
		}
		expect(payload.args?.inputPreset).toBe('public-tracked')
		expect(payload.args?.inputFile).toBe('fixtures/xlsx/calamine/issue_174.xlsx')
		expect(payload.args?.tableInputFile).toBe('fixtures/xlsx/calamine/table-multiple.xlsx')
		expect(payload.args?.editInputFile).toContain('generated-edit-mixed-10pct-text-65536x10.xlsx')
		expect(payload.args?.generatedEditInputLabel).toContain('generated:mixed-10pct-text:65536x10')
		expect(payload.inputs?.find((input) => input.role === 'first-view')).toMatchObject({
			kind: 'tracked-file',
			releaseClaimable: true,
		})
		expect(payload.inputs?.find((input) => input.role === 'edit')).toMatchObject({
			kind: 'generated',
			releaseClaimable: true,
		})
		expect(payload.inputs?.find((input) => input.role === 'table-inspection')).toMatchObject({
			kind: 'tracked-file',
			releaseClaimable: true,
		})
		for (const result of payload.results ?? []) {
			if (result.id === 'workflow-commit' || result.id === 'post-write-breakdown') {
				expect(result.command).toContain('--input-file')
				expect(result.command).toContain('generated-edit-mixed-10pct-text-65536x10.xlsx')
			}
		}
	})

	test('ranks measured plan sub-phases ahead of aggregate prepared plan timing', () => {
		const decisions = practicalLatencyContractsTestHooks.envelopeDecisions([
			{
				contract: 'edit-verify',
				id: 'workflow-commit',
				label: 'workflow',
				status: 'ok',
				command: 'workflow',
				elapsedMs: 0,
				profileCommand: 'profile workflow',
				summary: {
					preparedTotalMedianMs: 200,
					preparedPlanMedianMs: 100,
					preparedPlanStats: { p95: 105, cv: 0.03 },
					preparedCommitMedianMs: 40,
				},
			},
			{
				contract: 'edit-verify',
				id: 'agent-phase-profile',
				label: 'phase',
				status: 'ok',
				command: 'phase',
				elapsedMs: 0,
				profileCommand: 'profile phase',
				summary: {
					sharedPlanPhaseMedianMs: {
						'hash-input': 2,
						'load-workbook': 60,
						preview: 10,
						'preservation-audit': 5,
					},
					sharedPlanPhaseStats: {
						'load-workbook': { p95: 68, cv: 0.04 },
					},
				},
			},
		] as Parameters<typeof practicalLatencyContractsTestHooks.envelopeDecisions>[0])
		const edit = decisions.find((decision) => decision.contract === 'edit-verify')
		expect(edit?.largestPhase).toBe('Shared plan load-workbook/open')
		expect(edit?.phaseMedianMs).toBe(60)
		expect(edit?.profileCommand).toBe('profile phase')
	})

	test('falls back to aggregate prepared plan timing when phase split is unavailable', () => {
		const decisions = practicalLatencyContractsTestHooks.envelopeDecisions([
			{
				contract: 'edit-verify',
				id: 'workflow-commit',
				label: 'workflow',
				status: 'ok',
				command: 'workflow',
				elapsedMs: 0,
				profileCommand: 'profile workflow',
				summary: {
					preparedTotalMedianMs: 200,
					preparedPlanMedianMs: 100,
					preparedPlanStats: { p95: 105, cv: 0.03 },
					preparedCommitMedianMs: 80,
				},
			},
		] as Parameters<typeof practicalLatencyContractsTestHooks.envelopeDecisions>[0])
		const edit = decisions.find((decision) => decision.contract === 'edit-verify')
		expect(edit?.largestPhase).toBe('Prepared plan/open')
		expect(edit?.phaseMedianMs).toBe(100)
		expect(edit?.profileCommand).toBe('profile workflow')
	})
})
