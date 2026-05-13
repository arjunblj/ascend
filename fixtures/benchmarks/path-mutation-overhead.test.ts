import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./path-mutation-overhead.ts', import.meta.url))

describe('path mutation overhead benchmark', () => {
	test('measures compile, journal, plan, and commit phases', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--rows',
				'80',
				'--cols',
				'6',
				'--mutations',
				'12',
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
			readonly summary?: {
				readonly directCompileMedianMs?: number
				readonly directPreviewJournalMedianMs?: number
				readonly directApplyJournalMedianMs?: number
				readonly apiPlanMedianMs?: number
				readonly apiCommitMedianMs?: number
				readonly mutationCountMedian?: number
				readonly compiledOpsMedian?: number
				readonly compileIssuesMedian?: number
				readonly previewJournalEntriesMedian?: number
				readonly previewPreimagesMedian?: number
				readonly applyJournalEntriesMedian?: number
				readonly applyPreimagesMedian?: number
				readonly commitOk?: boolean
			}
		}
		expect(payload.summary?.directCompileMedianMs).toBeNumber()
		expect(payload.summary?.directPreviewJournalMedianMs).toBeNumber()
		expect(payload.summary?.directApplyJournalMedianMs).toBeNumber()
		expect(payload.summary?.apiPlanMedianMs).toBeNumber()
		expect(payload.summary?.apiCommitMedianMs).toBeNumber()
		expect(payload.summary?.mutationCountMedian).toBe(12)
		expect(payload.summary?.compiledOpsMedian).toBe(1)
		expect(payload.summary?.compileIssuesMedian).toBe(0)
		expect(payload.summary?.previewJournalEntriesMedian).toBe(1)
		expect(payload.summary?.previewPreimagesMedian).toBe(12)
		expect(payload.summary?.applyJournalEntriesMedian).toBe(1)
		expect(payload.summary?.applyPreimagesMedian).toBe(12)
		expect(payload.summary?.commitOk).toBe(true)
	})
})
