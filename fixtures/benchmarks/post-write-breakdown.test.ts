import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./post-write-breakdown.ts', import.meta.url))

describe('post-write breakdown benchmark', () => {
	test('splits safe commit post-write verification into subphases', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--rows',
				'80',
				'--cols',
				'6',
				'--updates',
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
				readonly commitPostWriteMedianMs?: number
				readonly commitPostWriteReopenMedianMs?: number
				readonly commitPostWriteCheckMedianMs?: number
				readonly commitPostWriteLintMedianMs?: number
				readonly commitPostWritePreservationMedianMs?: number
				readonly commitPostWritePackageGraphMedianMs?: number
				readonly commitPostWritePackageGraphAuditMedianMs?: number
				readonly reopenOutputMedianMs?: number
				readonly checkMedianMs?: number
				readonly lintMedianMs?: number
				readonly preservationMedianMs?: number
				readonly outputPackageGraphMedianMs?: number
				readonly packageGraphAuditMedianMs?: number
				readonly outputBytesMedian?: number
				readonly checkIssuesMedian?: number
				readonly packageGraphIssuesMedian?: number
				readonly valid?: boolean
			}
		}
		expect(payload.summary?.commitPostWriteMedianMs).toBeNumber()
		expect(payload.summary?.commitPostWriteReopenMedianMs).toBeNumber()
		expect(payload.summary?.commitPostWriteCheckMedianMs).toBeNumber()
		expect(payload.summary?.commitPostWriteLintMedianMs).toBeNumber()
		expect(payload.summary?.commitPostWritePreservationMedianMs).toBeNumber()
		expect(payload.summary?.commitPostWritePackageGraphMedianMs).toBeNumber()
		expect(payload.summary?.commitPostWritePackageGraphAuditMedianMs).toBeNumber()
		expect(payload.summary?.reopenOutputMedianMs).toBeNumber()
		expect(payload.summary?.checkMedianMs).toBeNumber()
		expect(payload.summary?.lintMedianMs).toBeNumber()
		expect(payload.summary?.preservationMedianMs).toBeNumber()
		expect(payload.summary?.outputPackageGraphMedianMs).toBeNumber()
		expect(payload.summary?.packageGraphAuditMedianMs).toBeNumber()
		expect(payload.summary?.outputBytesMedian).toBeGreaterThan(0)
		expect(payload.summary?.checkIssuesMedian).toBe(0)
		expect(payload.summary?.packageGraphIssuesMedian).toBe(0)
		expect(payload.summary?.valid).toBe(true)
	})
})
