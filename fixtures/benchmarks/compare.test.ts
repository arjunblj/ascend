import { afterEach, describe, expect, test } from 'bun:test'
import { mkdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { BenchmarkSuiteResult } from './results.ts'

const TEMP_DIR = join(
	tmpdir(),
	`ascend-compare-${Date.now()}-${Math.random().toString(16).slice(2)}`,
)

afterEach(async () => {
	await rm(TEMP_DIR, { force: true, recursive: true })
})

describe('benchmark compare CLI', () => {
	test('--explain prints repro and profile commands for regressions and new cases', async () => {
		const baseline: BenchmarkSuiteResult = {
			formatVersion: 1,
			suite: 'baseline',
			kind: 'synthetic',
			generatedAt: new Date(0).toISOString(),
			runtime: { platform: 'test', arch: 'arm64' },
			git: {},
			cases: [
				{
					name: 'regressed-case',
					category: 'read',
					dimensions: {},
					metrics: {
						sampleCount: 1,
						minMs: 100,
						medianMs: 100,
						meanMs: 100,
						p95Ms: 100,
						maxMs: 100,
					},
				},
			],
		}
		const candidate: BenchmarkSuiteResult = {
			...baseline,
			suite: 'candidate',
			cases: [
				{
					name: 'regressed-case',
					category: 'read',
					dimensions: {},
					metrics: {
						sampleCount: 1,
						minMs: 125,
						medianMs: 125,
						meanMs: 125,
						p95Ms: 125,
						maxMs: 125,
					},
					reproCommand: 'bun run bench --scenario regressed-case --json',
					profileCommand:
						'bun run fixtures/benchmarks/profile-bun.ts -- bun run bench --scenario regressed-case --json',
				},
				{
					name: 'new-case',
					category: 'write',
					dimensions: {},
					metrics: {
						sampleCount: 1,
						minMs: 10,
						medianMs: 10,
						meanMs: 10,
						p95Ms: 10,
						maxMs: 10,
					},
					reproCommand: 'bun run bench --scenario new-case --json',
					profileCommand:
						'bun run fixtures/benchmarks/profile-bun.ts -- bun run bench --scenario new-case --json',
				},
			],
		}
		const baselinePath = join(TEMP_DIR, 'baseline.json')
		const candidatePath = join(TEMP_DIR, 'candidate.json')
		await mkdir(TEMP_DIR, { recursive: true })
		await writeFile(baselinePath, JSON.stringify(baseline))
		await writeFile(candidatePath, JSON.stringify(candidate))

		const proc = Bun.spawn(
			['bun', 'run', 'fixtures/benchmarks/compare.ts', baselinePath, candidatePath, '--explain'],
			{ cwd: process.cwd(), stderr: 'pipe', stdout: 'pipe' },
		)
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])

		expect(stderr).toBe('')
		expect(exitCode).toBe(1)
		expect(stdout).toContain('REGRESSED regressed-case')
		expect(stdout).toContain('repro: bun run bench --scenario regressed-case --json')
		expect(stdout).toContain('profile: bun run fixtures/benchmarks/profile-bun.ts')
		expect(stdout).toContain('NEW new-case')
		expect(stdout).toContain('repro: bun run bench --scenario new-case --json')
	})
})
