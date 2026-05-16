import { describe, expect, test } from 'bun:test'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

describe('agent-safe-edit MCP example', () => {
	test('runs the MCP inspect-plan-commit-reopen-verify workflow', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-agent-safe-edit-mcp-'))
		const input = join(dir, 'input.xlsx')
		const output = join(dir, 'output.xlsx')

		try {
			const proc = Bun.spawn(
				[process.execPath, 'run', 'examples/agent-safe-edit-mcp.ts', input, output],
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
				discovery?: { workflowSteps?: number; planTool?: string; commitTool?: string }
				input?: {
					openPlan?: { reviewBeforeHydration?: boolean }
					trust?: { trust?: string; posture?: string; findingCount?: number }
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
				postWriteProof?: {
					dataConnections?: { total?: number; verification?: string }
					formulaState?: { formulaCells?: number; verification?: string }
					security?: { workbookProtected?: boolean; verification?: string }
					visuals?: { chartParts?: number; verification?: string }
				}
				proofBundle?: {
					safeToUse?: boolean
					whatChanged?: Array<{ ref?: string }>
					whySafe?: Array<{ gate?: string; ok?: boolean }>
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
			expect(result.workflow).toBe(
				'mcp-open-plan-trust-inspect-read-plan-prepared-commit-reopen-verify',
			)
			expect(result.discovery).toMatchObject({
				planTool: 'ascend.plan',
				commitTool: 'ascend.commit',
			})
			expect(result.discovery?.workflowSteps).toBeGreaterThan(0)
			expect(result.input?.openPlan?.reviewBeforeHydration).toBe(false)
			expect(result.input?.trust).toMatchObject({
				trust: 'untrusted',
				posture: 'safe-parser-preserver',
				findingCount: 0,
			})
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
			expect(result.postWriteProof).toMatchObject({
				dataConnections: { total: 0, verification: 'reopened-output' },
				formulaState: { formulaCells: 1, verification: 'reopened-output' },
				security: { workbookProtected: false, verification: 'reopened-output' },
				visuals: { chartParts: 0, verification: 'reopened-output' },
			})
			expect(result.proofBundle).toMatchObject({
				safeToUse: true,
				whatChanged: [{ ref: 'B2' }],
			})
			expect(result.proofBundle?.whySafe?.map((gate) => [gate.gate, gate.ok])).toEqual([
				['input-guard', true],
				['approval', true],
				['write-policy', true],
				['commit', true],
				['reopen-verify', true],
				['package-graph', true],
			])
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
