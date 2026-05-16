import { describe, expect, test } from 'bun:test'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const repoRoot = join(import.meta.dir, '..')
const workflows = [
	{
		script: 'example:safe-edit',
		expectedWorkflow: 'open-plan-trust-inspect-read-plan-prepared-commit-verify-repair',
		expectedProofRef: 'Sheet1!B2',
	},
	{
		script: 'example:safe-edit:http',
		expectedWorkflow: 'api-open-plan-inspect-read-plan-prepared-commit-reopen-verify',
		expectedProofRef: 'B2',
	},
	{
		script: 'example:safe-edit:mcp',
		expectedWorkflow: 'mcp-open-plan-inspect-read-plan-prepared-commit-reopen-verify',
		expectedProofRef: 'B2',
	},
]

describe('root package safe-edit scripts', () => {
	for (const workflow of workflows) {
		test(`${workflow.script} runs the agent safe edit workflow`, async () => {
			const dir = await mkdtemp(join(tmpdir(), `ascend-${workflow.script.replace(':', '-')}-`))
			const input = join(dir, 'input.xlsx')
			const output = join(dir, 'output.xlsx')

			try {
				const proc = Bun.spawn([process.execPath, 'run', workflow.script, input, output], {
					cwd: repoRoot,
					stdout: 'pipe',
					stderr: 'pipe',
				})
				const [stdout, stderr, exitCode] = await Promise.all([
					new Response(proc.stdout).text(),
					new Response(proc.stderr).text(),
					proc.exited,
				])

				expect(exitCode).toBe(0)
				expect(stderr).toContain('bun run --cwd examples')
				expect(stderr).not.toContain('error:')

				const result = JSON.parse(stdout) as {
					ok?: boolean
					workflow?: string
					plan?: { changedCells?: string[] }
					commit?: { output?: string; postWriteValid?: boolean }
					proofBundle?: { safeToUse?: boolean; whatChanged?: Array<{ ref?: string }> }
					verify?: {
						checkValid?: boolean
						lintClean?: boolean
						cell?: { formula?: string | null; value?: { kind?: string; value?: number } }
					}
				}

				expect(result.ok).toBe(true)
				expect(result.workflow).toBe(workflow.expectedWorkflow)
				expect(result.plan?.changedCells).toEqual(['B2'])
				expect(result.commit).toMatchObject({
					output,
					postWriteValid: true,
				})
				expect(result.proofBundle).toMatchObject({
					safeToUse: true,
					whatChanged: [{ ref: workflow.expectedProofRef }],
				})
				expect(result.verify).toMatchObject({
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
	}
})
