import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./run.ts', import.meta.url))

describe('synthetic benchmark runner CLI', () => {
	test('--help prints usage without running benchmarks', () => {
		const proc = Bun.spawnSync({
			cmd: [Bun.argv[0], runnerPath, '--help'],
			stdout: 'pipe',
			stderr: 'pipe',
		})
		const stdout = new TextDecoder().decode(proc.stdout)
		const stderr = new TextDecoder().decode(proc.stderr)

		expect(proc.exitCode, stderr).toBe(0)
		expect(stdout).toContain('Ascend synthetic benchmark runner')
		expect(stdout).toContain('--scenario <name>')
		expect(stdout).toContain('Scenario sets:')
		expect(stdout).toContain('Scenarios:')
		expect(stdout).toContain('write-xlsx-100k-rows')
		expect(stdout).not.toContain('completed ')
		expect(stdout).not.toContain('Ascend benchmark summary')
	})
})
