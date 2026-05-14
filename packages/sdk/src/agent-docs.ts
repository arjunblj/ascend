import { join } from 'node:path'

const REPO_ROOT = new URL('../../../', import.meta.url).pathname

export type AgentDocKind = 'docs' | 'example' | 'llms' | 'reference'

export interface AgentDocEntry {
	readonly id: string
	readonly title: string
	readonly path: string
	readonly kind: AgentDocKind
	readonly text: string
}

export interface AgentDocSearchResult {
	readonly id: string
	readonly title: string
	readonly path: string
	readonly kind: AgentDocKind
	readonly score: number
	readonly snippet: string
}

const DOC_SOURCES: readonly Omit<AgentDocEntry, 'text'>[] = [
	{ id: 'llms', title: 'Ascend llms.txt', path: 'llms.txt', kind: 'llms' },
	{ id: 'llms-full', title: 'Ascend llms-full.txt', path: 'llms-full.txt', kind: 'llms' },
	{ id: 'agent-api', title: 'Agent API Reference', path: 'docs/AGENT_API.md', kind: 'reference' },
	{ id: 'agent-workflow', title: 'Agent Workflow', path: 'docs/AGENT_WORKFLOW.md', kind: 'docs' },
	{ id: 'security', title: 'Security Policy', path: 'docs/SECURITY.md', kind: 'docs' },
	{ id: 'versioning', title: 'Versioning', path: 'docs/VERSIONING.md', kind: 'docs' },
	{ id: 'openapi', title: 'OpenAPI Schema', path: 'docs/openapi.yaml', kind: 'reference' },
	{ id: 'examples-readme', title: 'Examples README', path: 'examples/README.md', kind: 'example' },
	{
		id: 'example-agent-safe-edit',
		title: 'Agent Safe Edit Golden Path',
		path: 'examples/agent-safe-edit.ts',
		kind: 'example',
	},
	{
		id: 'example-agent-safe-edit-http',
		title: 'Agent Safe Edit HTTP Transcript',
		path: 'examples/agent-safe-edit-http.md',
		kind: 'example',
	},
	{
		id: 'example-agent-safe-edit-mcp',
		title: 'Agent Safe Edit MCP Transcript',
		path: 'examples/agent-safe-edit-mcp.md',
		kind: 'example',
	},
	{
		id: 'example-untrusted-workbook-report',
		title: 'Untrusted Workbook Report',
		path: 'examples/untrusted-workbook-report.md',
		kind: 'example',
	},
	{ id: 'mcp-setup', title: 'MCP Setup Example', path: 'examples/mcp-setup.md', kind: 'example' },
	{
		id: 'example-create',
		title: 'Create Workbook Example',
		path: 'examples/create-from-scratch.ts',
		kind: 'example',
	},
	{
		id: 'example-read-modify-save',
		title: 'Read Modify Save Example',
		path: 'examples/read-modify-save.ts',
		kind: 'example',
	},
	{
		id: 'example-batch-ops',
		title: 'Batch Operations Example',
		path: 'examples/batch-ops.ts',
		kind: 'example',
	},
	{
		id: 'example-formula-eval',
		title: 'Formula Eval Example',
		path: 'examples/formula-eval.ts',
		kind: 'example',
	},
	{
		id: 'example-csv-convert',
		title: 'CSV Convert Example',
		path: 'examples/csv-convert.ts',
		kind: 'example',
	},
] as const

let cachedDocs: readonly AgentDocEntry[] | null = null

export async function loadAgentDocs(): Promise<readonly AgentDocEntry[]> {
	if (cachedDocs) return cachedDocs
	const docs: AgentDocEntry[] = []
	for (const source of DOC_SOURCES) {
		const file = Bun.file(join(REPO_ROOT, source.path))
		if (!(await file.exists())) continue
		docs.push({ ...source, text: await file.text() })
	}
	cachedDocs = docs
	return docs
}

export async function readAgentDoc(path: string): Promise<string | undefined> {
	const normalized = normalizeDocPath(path)
	const docs = await loadAgentDocs()
	return docs.find((doc) => doc.path === normalized || doc.id === normalized)?.text
}

export async function searchAgentDocs(options: {
	readonly query: string
	readonly limit?: number
	readonly tokens?: number
	readonly kind?: AgentDocKind
}): Promise<readonly AgentDocSearchResult[]> {
	const query = options.query.trim()
	if (!query) return []
	const terms = tokenize(query)
	if (terms.length === 0) return []
	const limit = Math.min(Math.max(options.limit ?? 5, 1), 20)
	const tokenBudget = Math.min(Math.max(options.tokens ?? 1200, 100), 8000)
	const docs = (await loadAgentDocs()).filter((doc) => !options.kind || doc.kind === options.kind)
	const scored = docs
		.map((doc) => scoreDoc(doc, terms, tokenBudget))
		.filter((result): result is AgentDocSearchResult => result !== null)
		.sort((a, b) => b.score - a.score || a.path.localeCompare(b.path))
	return scored.slice(0, limit)
}

function scoreDoc(
	doc: AgentDocEntry,
	terms: readonly string[],
	tokenBudget: number,
): AgentDocSearchResult | null {
	const haystack = `${doc.title}\n${doc.path}\n${doc.text}`.toLowerCase()
	let score = 0
	for (const term of terms) {
		const titleHits = countOccurrences(doc.title.toLowerCase(), term)
		const pathHits = countOccurrences(doc.path.toLowerCase(), term)
		const bodyHits = countOccurrences(haystack, term)
		score += titleHits * 8 + pathHits * 5 + bodyHits
	}
	if (score === 0) return null
	return {
		id: doc.id,
		title: doc.title,
		path: doc.path,
		kind: doc.kind,
		score,
		snippet: snippetFor(doc.text, terms, tokenBudget),
	}
}

function snippetFor(text: string, terms: readonly string[], tokenBudget: number): string {
	const lower = text.toLowerCase()
	let index = -1
	for (const term of terms) {
		const termIndex = lower.indexOf(term)
		if (termIndex !== -1 && (index === -1 || termIndex < index)) index = termIndex
	}
	const approxChars = tokenBudget * 4
	const start = index === -1 ? 0 : Math.max(0, index - Math.floor(approxChars / 3))
	const end = Math.min(text.length, start + approxChars)
	const prefix = start > 0 ? '...' : ''
	const suffix = end < text.length ? '...' : ''
	return `${prefix}${text.slice(start, end).trim()}${suffix}`
}

function tokenize(query: string): readonly string[] {
	return [...new Set(query.toLowerCase().match(/[a-z0-9_.:-]+/g) ?? [])].filter(
		(term) => term.length > 1,
	)
}

function countOccurrences(value: string, term: string): number {
	let count = 0
	let cursor = 0
	while (true) {
		const index = value.indexOf(term, cursor)
		if (index === -1) return count
		count += 1
		cursor = index + term.length
	}
}

function normalizeDocPath(path: string): string {
	return path.replace(/^ascend:\/\//, '').replace(/^\/+/, '')
}
