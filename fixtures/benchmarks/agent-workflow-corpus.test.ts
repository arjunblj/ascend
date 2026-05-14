import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./agent-workflow-corpus.ts', import.meta.url))

describe('agent workflow corpus benchmark', () => {
	test('runs a selected real workbook target', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--target',
				'poi-with-various-data-approved',
				'--surface',
				'both',
				'--repeat',
				'1',
				'--warmup',
				'0',
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
			readonly tool?: string
			readonly summary?: {
				readonly targetCount?: number
				readonly okCount?: number
				readonly failedCount?: number
				readonly validCount?: number
				readonly preparedValidCount?: number
				readonly mcpValidCount?: number
				readonly mcpPreparedValidCount?: number
			}
			readonly results?: Array<{
				readonly name?: string
				readonly status?: string
				readonly reproCommand?: string
				readonly profileCommand?: string
				readonly approvals?: string
				readonly summary?: {
					readonly preparedCommitVerifiedTotalMedianMs?: number
					readonly mcpPreparedCommitVerifiedTotalMedianMs?: number
					readonly preparedCommitVerifiedHydratedOpenCountMedian?: number
					readonly mcpPreparedCommitVerifiedHydratedOpenCountMedian?: number
				}
			}>
		}
		expect(payload.tool).toBe('agent-workflow-corpus')
		expect(payload.summary?.targetCount).toBe(1)
		expect(payload.summary?.okCount).toBe(1)
		expect(payload.summary?.failedCount).toBe(0)
		expect(payload.summary?.validCount).toBe(1)
		expect(payload.summary?.preparedValidCount).toBe(1)
		expect(payload.summary?.mcpValidCount).toBe(1)
		expect(payload.summary?.mcpPreparedValidCount).toBe(1)
		expect(payload.results?.[0]?.name).toBe('poi-with-various-data-approved')
		expect(payload.results?.[0]?.status).toBe('ok')
		expect(payload.results?.[0]?.approvals).toBe('all')
		expect(payload.results?.[0]?.reproCommand).toContain(
			'bun run fixtures/benchmarks/agent-workflow.ts',
		)
		expect(payload.results?.[0]?.reproCommand).toContain(
			'--input-file fixtures/xlsx/poi/WithVariousData.xlsx',
		)
		expect(payload.results?.[0]?.profileCommand).toContain(
			'bun run fixtures/benchmarks/profile-bun.ts',
		)
		expect(payload.results?.[0]?.profileCommand).toContain('--mode all-md')
		expect(payload.results?.[0]?.profileCommand).toContain(
			'bun run fixtures/benchmarks/agent-workflow.ts',
		)
		expect(payload.results?.[0]?.profileCommand).not.toContain('agent-workflow-corpus.ts')
		expect(payload.results?.[0]?.summary?.preparedCommitVerifiedTotalMedianMs).toBeNumber()
		expect(payload.results?.[0]?.summary?.mcpPreparedCommitVerifiedTotalMedianMs).toBeNumber()
		expect(
			payload.results?.[0]?.summary?.preparedCommitVerifiedHydratedOpenCountMedian,
		).toBeNumber()
		expect(
			payload.results?.[0]?.summary?.mcpPreparedCommitVerifiedHydratedOpenCountMedian,
		).toBeNumber()
	}, 20_000)
})
