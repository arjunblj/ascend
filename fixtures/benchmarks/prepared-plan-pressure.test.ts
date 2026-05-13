import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./prepared-plan-pressure.ts', import.meta.url))

describe('prepared plan pressure benchmark', () => {
	test('measures bounded prepared handle retention and eviction', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--rows',
				'80',
				'--cols',
				'6',
				'--mutations',
				'8',
				'--handles',
				'4',
				'--max-handles',
				'2',
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
				readonly planTotalMedianMs?: number
				readonly planPerHandleMedianMs?: number
				readonly planPayloadBytesMedian?: number
				readonly preparedHandlesCreatedMedian?: number
				readonly maxHandlesMedian?: number
				readonly estimatedEvictedHandlesMedian?: number
				readonly rssRetainedAfterPlansMedianMb?: number
				readonly rssPerRetainedHandleMedianMb?: number
				readonly firstHandleEvicted?: boolean
				readonly latestCommitMedianMs?: number
				readonly latestCommitPayloadBytesMedian?: number
				readonly latestCommitOk?: boolean
			}
		}
		expect(payload.summary?.planTotalMedianMs).toBeNumber()
		expect(payload.summary?.planPerHandleMedianMs).toBeNumber()
		expect(payload.summary?.planPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.preparedHandlesCreatedMedian).toBe(4)
		expect(payload.summary?.maxHandlesMedian).toBe(2)
		expect(payload.summary?.estimatedEvictedHandlesMedian).toBe(2)
		expect(payload.summary?.rssRetainedAfterPlansMedianMb).toBeNumber()
		expect(payload.summary?.rssPerRetainedHandleMedianMb).toBeNumber()
		expect(payload.summary?.firstHandleEvicted).toBe(true)
		expect(payload.summary?.latestCommitMedianMs).toBeNumber()
		expect(payload.summary?.latestCommitPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.latestCommitOk).toBe(true)
	}, 20_000)
})
