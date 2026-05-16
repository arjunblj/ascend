import { describe, expect, test } from 'bun:test'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

describe('agent-safe-edit example', () => {
	test('runs inspect-plan-commit-reopen-verify and prints machine proof', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-agent-safe-edit-'))
		const input = join(dir, 'input.xlsx')
		const output = join(dir, 'output.xlsx')

		try {
			const proc = Bun.spawn(
				[process.execPath, 'run', 'examples/agent-safe-edit.ts', input, output],
				{
					cwd: join(import.meta.dir, '..'),
					stdout: 'pipe',
					stderr: 'pipe',
				},
			)
			const [stdout, stderr, exitCode] = await Promise.all([
				new Response(proc.stdout).text(),
				new Response(proc.stderr).text(),
				proc.exited,
			])

			expect(stderr).toBe('')
			expect(exitCode).toBe(0)

			const result = JSON.parse(stdout) as {
				ok?: boolean
				workflow?: string
				input?: {
					openPlan?: { reviewBeforeHydration?: boolean }
					read?: { cellCount?: number }
					operationCount?: number
				}
				plan?: { wouldSucceed?: boolean; changedCells?: string[] }
				commit?: {
					output?: string
					postWriteValid?: boolean
					auditsPassed?: boolean
					checkValid?: boolean
					lintClean?: boolean
				}
				verify?: {
					reopened?: boolean
					cell?: {
						ref?: string
						formula?: string | null
						value?: { kind?: string; value?: number }
					}
					checkValid?: boolean
					checkIssueCount?: number
					lintClean?: boolean
					lintWarningCount?: number
					commands?: { check?: string; diff?: string; repair?: string }
				}
			}

			expect(result.ok).toBe(true)
			expect(result.workflow).toBe(
				'open-plan-trust-inspect-read-plan-prepared-commit-verify-repair',
			)
			expect(result.input?.openPlan?.reviewBeforeHydration).toBe(false)
			expect(result.input?.read?.cellCount).toBeGreaterThan(0)
			expect(result.input?.operationCount).toBe(1)
			expect(result.plan).toMatchObject({
				wouldSucceed: true,
				changedCells: ['B2'],
			})
			expect(result.commit).toMatchObject({
				output,
				postWriteValid: true,
				auditsPassed: true,
				checkValid: true,
				lintClean: true,
			})
			expect(result.verify).toMatchObject({
				reopened: true,
				cell: {
					ref: 'Sheet1!B2',
					formula: 'SUM(A2:A4)',
					value: { kind: 'number', value: 450 },
				},
				checkValid: true,
				checkIssueCount: 0,
				lintClean: true,
				lintWarningCount: 0,
			})
			expect(result.verify?.commands?.check).toContain(output)
			expect(result.verify?.commands?.diff).toContain(input)
			expect(result.verify?.commands?.repair).toContain(output)
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})
})
