import { describe, expect, test } from 'bun:test'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

describe('package-install-safe-edit example', () => {
	test('runs the installed SDK inspect-plan-commit-reopen-verify workflow', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-package-install-safe-edit-'))
		const input = join(dir, 'input.xlsx')
		const output = join(dir, 'output.xlsx')

		try {
			const proc = Bun.spawn(
				[process.execPath, 'run', 'examples/package-install-safe-edit.ts', input, output],
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
				install?: { package?: string; example?: string }
				input?: { read?: { cellCount?: number } }
				plan?: { changedCells?: string[]; approvalCount?: number }
				commit?: { output?: string; postWriteValid?: boolean; auditsPassed?: boolean }
				verify?: {
					reopened?: boolean
					checkValid?: boolean
					lintClean?: boolean
					cell?: { formula?: string | null; value?: { kind?: string; value?: number } }
				}
			}

			expect(result.ok).toBe(true)
			expect(result.workflow).toBe(
				'installed-sdk-open-plan-trust-inspect-read-plan-commit-reopen-verify',
			)
			expect(result.install).toMatchObject({
				package: '@ascend/sdk',
				example: 'node_modules/@ascend/sdk/examples/package-install-safe-edit.ts',
			})
			expect(result.input?.read?.cellCount).toBeGreaterThan(0)
			expect(result.plan).toMatchObject({ changedCells: ['B2'], approvalCount: 0 })
			expect(result.commit).toMatchObject({
				output,
				postWriteValid: true,
				auditsPassed: true,
			})
			expect(result.verify).toMatchObject({
				reopened: true,
				checkValid: true,
				lintClean: true,
				cell: {
					formula: 'SUM(A2:A4)',
					value: { kind: 'number', value: 450 },
				},
			})
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})
})
