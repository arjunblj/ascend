import { ascendError } from '@ascend/schema'
import {
	type AgentDocEntry,
	type AgentDocKind,
	loadAgentDocs,
	readAgentDoc,
	searchAgentDocs,
} from '@ascend/sdk'
import { cliError, jsonOut } from '../output/json.ts'
import { table } from '../output/pretty.ts'

export const usage = `Usage: ascend docs [query] [flags]

  Search or read bundled Ascend agent documentation and examples.

Flags:
  --query <text>     Search query (or pass query as positional args)
  --examples         Search examples only
  --path <path|id>   Print a specific doc by path or id
  --list             List indexed docs
  --limit <n>        Maximum search results (default: 5, max: 20)
  --tokens <n>       Approximate snippet tokens per result (default: 1200)
  --json             Output as JSON
`

export async function docsCommand(args: string[], flags: Map<string, string>): Promise<number> {
	if (flags.has('list')) return listDocs(flags)
	if (flags.get('path')) return readDoc(flags.get('path') as string, flags)

	const examplesQuery = flags.get('examples') ?? ''
	const query = (flags.get('query') ?? [examplesQuery, ...args].join(' ')).trim()
	if (!query) {
		cliError(
			ascendError('INVALID_ARGUMENT', 'Missing docs query', {
				suggestedFix: 'Run ascend docs "plan commit" --json or ascend docs --list.',
			}),
			flags,
		)
		return 1
	}

	const limit = parsePositiveIntFlag('limit', flags, 20)
	if (limit === null) return 1
	const tokens = parsePositiveIntFlag('tokens', flags, 8000)
	if (tokens === null) return 1
	const kind: AgentDocKind | undefined = flags.has('examples') ? 'example' : undefined
	const results = await searchAgentDocs({
		query,
		...(kind ? { kind } : {}),
		...(limit !== undefined ? { limit } : {}),
		...(tokens !== undefined ? { tokens } : {}),
	})

	if (flags.has('json')) {
		console.log(jsonOut({ query, results }))
		return 0
	}

	console.log(
		table(
			['Score', 'Kind', 'Path', 'Title'],
			results.map((result) => [String(result.score), result.kind, result.path, result.title]),
		),
	)
	for (const result of results) {
		console.log(`\n${result.path}\n${result.snippet}`)
	}
	return 0
}

async function listDocs(flags: Map<string, string>): Promise<number> {
	const docs = await loadAgentDocs()
	const entries = docs.map(publicDocEntry)
	if (flags.has('json')) {
		console.log(jsonOut({ docs: entries }))
		return 0
	}
	console.log(
		table(
			['ID', 'Kind', 'Path', 'Title'],
			entries.map((doc) => [doc.id, doc.kind, doc.path, doc.title]),
		),
	)
	return 0
}

async function readDoc(path: string, flags: Map<string, string>): Promise<number> {
	const text = await readAgentDoc(path)
	if (text === undefined) {
		cliError(
			ascendError('FILE_NOT_FOUND', `Documentation entry not found: ${path}`, {
				suggestedFix: 'Run ascend docs --list to see indexed docs.',
			}),
			flags,
		)
		return 1
	}
	if (flags.has('json')) {
		console.log(jsonOut({ path: normalizeDocPath(path), text }))
		return 0
	}
	console.log(text)
	return 0
}

function parsePositiveIntFlag(
	name: 'limit' | 'tokens',
	flags: Map<string, string>,
	max: number,
): number | undefined | null {
	const raw = flags.get(name)
	if (raw === undefined) return undefined
	const value = Number.parseInt(raw, 10)
	if (!Number.isSafeInteger(value) || value <= 0 || value > max) {
		cliError(
			ascendError('INVALID_ARGUMENT', `--${name} must be an integer between 1 and ${max}`, {
				suggestedFix: `Use --${name} ${name === 'limit' ? '5' : '1200'}.`,
			}),
			flags,
		)
		return null
	}
	return value
}

function publicDocEntry(doc: AgentDocEntry) {
	return {
		id: doc.id,
		title: doc.title,
		path: doc.path,
		kind: doc.kind,
	}
}

function normalizeDocPath(path: string): string {
	return path.replace(/^ascend:\/\//, '').replace(/^\/+/, '')
}
