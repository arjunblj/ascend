import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./cli-open-preview.ts', import.meta.url))

describe('CLI open preview benchmark', () => {
	test('measures actual open command preview behavior', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--rows',
				'80',
				'--cols',
				'6',
				'--preview-rows',
				'10',
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
				readonly tuiFullMedianMs?: number
				readonly tuiPreviewMedianMs?: number
				readonly openDefaultMedianMs?: number
				readonly openDefaultSpeedupVsTuiFull?: number
				readonly openDefaultOutputBytesMedian?: number
			}
		}
		expect(payload.summary?.tuiFullMedianMs).toBeNumber()
		expect(payload.summary?.tuiPreviewMedianMs).toBeNumber()
		expect(payload.summary?.openDefaultMedianMs).toBeNumber()
		expect(payload.summary?.openDefaultSpeedupVsTuiFull).toBeNumber()
		expect(payload.summary?.openDefaultOutputBytesMedian).toBeGreaterThan(0)
	})
})
