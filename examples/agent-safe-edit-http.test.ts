import { describe, expect, test } from 'bun:test'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

describe('agent-safe-edit HTTP example', () => {
	test('runs the API inspect-plan-commit-reopen-verify workflow', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-agent-safe-edit-http-'))
		const input = join(dir, 'input.xlsx')
		const output = join(dir, 'output.xlsx')

		try {
			const proc = Bun.spawn(
				[process.execPath, 'run', 'examples/agent-safe-edit-http.ts', input, output],
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
				}
				plan?: {
					inputSha256?: string
					planDigest?: string
					operationCount?: number
					changedCells?: string[]
					preparedPlanId?: string
				}
				commit?: {
					output?: string
					outputSha256?: string
					postWriteValid?: boolean
					auditsPassed?: boolean
					checkValid?: boolean
					lintClean?: boolean
				}
				verify?: {
					checkValid?: boolean
					checkIssueCount?: number
					lintClean?: boolean
					lintWarningCount?: number
					cell?: { ref?: string; formula?: string; value?: { kind?: string; value?: number } }
				}
			}

			expect(result.ok).toBe(true)
			expect(result.workflow).toBe('api-open-plan-inspect-read-plan-prepared-commit-reopen-verify')
			expect(result.input?.openPlan?.reviewBeforeHydration).toBe(false)
			expect(result.input?.read?.cellCount).toBeGreaterThan(0)
			expect(result.plan?.inputSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(result.plan?.planDigest).toMatch(/^[a-f0-9]{64}$/)
			expect(result.plan?.operationCount).toBe(1)
			expect(result.plan?.changedCells).toEqual(['B2'])
			expect(result.plan?.preparedPlanId).toBeString()
			expect(result.commit).toMatchObject({
				output,
				postWriteValid: true,
				auditsPassed: true,
				checkValid: true,
				lintClean: true,
			})
			expect(result.commit?.outputSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(result.verify).toMatchObject({
				checkValid: true,
				checkIssueCount: 0,
				lintClean: true,
				lintWarningCount: 0,
				cell: {
					ref: 'B2',
					formula: 'SUM(A2:A4)',
					value: { kind: 'number', value: 450 },
				},
			})
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})
})
