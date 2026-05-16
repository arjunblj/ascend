import { describe, expect, test } from 'bun:test'

const CLI = new URL('./index.ts', import.meta.url).pathname

interface CliRunResult {
	readonly stdout: string
	readonly stderr: string
	readonly exitCode: number
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

describe('ascend cli file errors', () => {
	test('open-plan reports missing files without raw ENOENT noise', async () => {
		const missing = `missing-open-plan-${Date.now()}.xlsx`
		const { stdout, stderr, exitCode } = await runProcess('open-plan', missing)

		expect(exitCode).toBe(1)
		expect(stdout).toBe('')
		expect(stderr).toContain(`Error: File not found: ${missing}`)
		expect(stderr).toContain('Pass an existing workbook path')
		expect(stderr).not.toContain('ENOENT')
	})
})
