import { describe, expect, test } from 'bun:test'
import { readFile } from 'node:fs/promises'
import { loadAgentDocs, searchAgentDocs } from './index.ts'

const REPO_ROOT = new URL('../../../', import.meta.url)

describe('agent documentation surface', () => {
	test('OpenAPI lists every implemented HTTP endpoint', async () => {
		const [serverSource, openapi] = await Promise.all([
			readFile(new URL('apps/api/src/server.ts', REPO_ROOT), 'utf-8'),
			readFile(new URL('docs/openapi.yaml', REPO_ROOT), 'utf-8'),
		])
		const implementedPaths = [
			...new Set(
				[...serverSource.matchAll(/path === '([^']+)'/g)].map((match) => match[1] as string),
			),
		].sort()

		for (const path of implementedPaths) {
			expect(openapi).toContain(`  ${path}:`)
		}
	})

	test('bundled docs expose the same prepared workflow vocabulary', async () => {
		const docs = await loadAgentDocs()
		const paths = docs.map((doc) => doc.path)
		expect(paths).toContain('examples/agent-safe-edit.ts')
		expect(paths).toContain('examples/agent-safe-edit-http.md')
		expect(paths).toContain('examples/agent-safe-edit-mcp.md')
		expect(paths).toContain('examples/untrusted-workbook-report.md')

		const agentApi = docs.find((doc) => doc.path === 'docs/AGENT_API.md')?.text ?? ''
		const workflow = docs.find((doc) => doc.path === 'docs/AGENT_WORKFLOW.md')?.text ?? ''
		const llms = docs.find((doc) => doc.path === 'llms.txt')?.text ?? ''
		const llmsFull = docs.find((doc) => doc.path === 'llms-full.txt')?.text ?? ''

		for (const text of [agentApi, workflow, llms, llmsFull]) {
			expect(text).toContain('planHandle')
			expect(text).toContain('process-local')
			expect(text).toContain('expectSha256')
			expect(text).toContain('formula')
			expect(text).toContain('trust')
			expect(text).toContain('ascend.trust_report')
			expect(text).toContain('untrusted')
		}
	})

	test('OpenAPI documents agent schemas and formula assist surface', async () => {
		const openapi = await readFile(new URL('docs/openapi.yaml', REPO_ROOT), 'utf-8')

		for (const required of [
			'  /trust-report:',
			'  /formula-assist:',
			'PlanRequest:',
			'CommitRequest:',
			'TrustReportResponse:',
			'TrustReportSuccessEnvelope:',
			'FormulaAssistRequest:',
			'FormulaAssistResponse:',
			'PathMutation:',
			'PreparedPlan:',
			'MachineFailure:',
			'/sheets/Revenue/cells/H2/formula',
		]) {
			expect(openapi).toContain(required)
		}
	})

	test('example search finds the runnable golden path by agent workflow terms', async () => {
		const results = await searchAgentDocs({
			query: 'golden path prepared planHandle verify repair',
			kind: 'example',
			limit: 5,
		})

		expect(results.some((result) => result.path === 'examples/agent-safe-edit.ts')).toBe(true)
	})

	test('example search finds the untrusted workbook preflight', async () => {
		const results = await searchAgentDocs({
			query: 'untrusted workbook trust_report hidden comments active content',
			kind: 'example',
			limit: 10,
		})

		expect(results.some((result) => result.path === 'examples/untrusted-workbook-report.md')).toBe(
			true,
		)
	})

	test('example search finds API and MCP safe edit transcripts', async () => {
		const results = await searchAgentDocs({
			query: 'formula assist planHandle path mutation http mcp',
			kind: 'example',
			limit: 10,
		})
		const paths = results.map((result) => result.path)

		expect(paths).toContain('examples/agent-safe-edit-http.md')
		expect(paths).toContain('examples/agent-safe-edit-mcp.md')
	})
})
