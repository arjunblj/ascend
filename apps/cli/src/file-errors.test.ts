import { afterAll, describe, expect, test } from 'bun:test'
import { existsSync, unlinkSync } from 'node:fs'
import { AscendWorkbook } from '@ascend/sdk'

const CLI = new URL('./index.ts', import.meta.url).pathname
const SIDE_CAR_TEST_FILE = 'file-error-sidecar.xlsx'
const SIDE_CAR_OUTPUT_FILE = 'file-error-sidecar-output.xlsx'

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

afterAll(() => {
	for (const file of [SIDE_CAR_TEST_FILE, SIDE_CAR_OUTPUT_FILE]) {
		const path = `${import.meta.dir}/${file}`
		if (existsSync(path)) unlinkSync(path)
	}
})

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

	test('plan reports a missing ops sidecar instead of blaming the workbook', async () => {
		const workbook = AscendWorkbook.create()
		await workbook.save(`${import.meta.dir}/${SIDE_CAR_TEST_FILE}`)
		const missingOps = `missing-plan-ops-${Date.now()}.json`

		const { stdout, exitCode } = await runProcess(
			'plan',
			SIDE_CAR_TEST_FILE,
			'--ops',
			missingOps,
			'--json',
		)

		expect(exitCode).toBe(1)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(false)
		expect(parsed.error).toMatchObject({
			code: 'FILE_NOT_FOUND',
			message: `File not found: ${missingOps}`,
			retryable: true,
			retryStrategy: 'modified',
			details: { file: missingOps },
		})
		expect(stdout).not.toContain('ENOENT')
	})

	test('commit with a missing ops sidecar reports that path and leaves no output proof', async () => {
		const workbook = AscendWorkbook.create()
		await workbook.save(`${import.meta.dir}/${SIDE_CAR_TEST_FILE}`)
		const missingOps = `missing-commit-ops-${Date.now()}.json`
		const outputPath = `${import.meta.dir}/${SIDE_CAR_OUTPUT_FILE}`
		if (existsSync(outputPath)) unlinkSync(outputPath)

		const { stdout, exitCode } = await runProcess(
			'commit',
			SIDE_CAR_TEST_FILE,
			'--ops',
			missingOps,
			'--output',
			SIDE_CAR_OUTPUT_FILE,
			'--json',
		)

		expect(exitCode).toBe(1)
		const parsed = JSON.parse(stdout)
		expect(parsed.ok).toBe(false)
		expect(parsed.error.details.file).toBe(missingOps)
		expect(parsed.error.message).toBe(`File not found: ${missingOps}`)
		expect(stdout).not.toContain('ENOENT')
		expect(existsSync(outputPath)).toBe(false)
	})
})
