#!/usr/bin/env bun
import { existsSync } from 'node:fs'

type ProjectName =
	| 'schema'
	| 'core'
	| 'formulas'
	| 'engine'
	| 'io-xlsx'
	| 'io-csv'
	| 'verify'
	| 'sdk'
	| 'cli'
	| 'api'
	| 'mcp'
	| 'tui'
	| 'fixtures-corpus'
	| 'fixtures-benchmarks'
	| 'fixtures-formulas'
	| 'examples'

interface Project {
	readonly name: ProjectName
	readonly dir: string
	readonly deps: readonly ProjectName[]
}

const PROJECTS: readonly Project[] = [
	{ name: 'schema', dir: 'packages/schema', deps: [] },
	{ name: 'core', dir: 'packages/core', deps: ['schema'] },
	{ name: 'formulas', dir: 'packages/formulas', deps: ['schema', 'core'] },
	{ name: 'engine', dir: 'packages/engine', deps: ['schema', 'core', 'formulas'] },
	{ name: 'io-xlsx', dir: 'packages/io-xlsx', deps: ['schema', 'core', 'formulas'] },
	{ name: 'io-csv', dir: 'packages/io-csv', deps: ['schema', 'core', 'formulas'] },
	{ name: 'verify', dir: 'packages/verify', deps: ['schema', 'core', 'formulas', 'engine'] },
	{
		name: 'sdk',
		dir: 'packages/sdk',
		deps: ['schema', 'core', 'formulas', 'engine', 'io-xlsx', 'io-csv', 'verify'],
	},
	{ name: 'tui', dir: 'apps/tui', deps: ['schema', 'core', 'sdk'] },
	{ name: 'cli', dir: 'apps/cli', deps: ['schema', 'core', 'sdk', 'tui'] },
	{ name: 'api', dir: 'apps/api', deps: ['schema', 'sdk'] },
	{ name: 'mcp', dir: 'apps/mcp', deps: ['schema', 'sdk'] },
	{ name: 'fixtures-corpus', dir: 'fixtures/corpus', deps: ['schema', 'io-xlsx', 'sdk'] },
	{
		name: 'fixtures-benchmarks',
		dir: 'fixtures/benchmarks',
		deps: ['schema', 'core', 'engine', 'io-xlsx'],
	},
	{
		name: 'fixtures-formulas',
		dir: 'fixtures/formulas',
		deps: ['schema', 'core', 'formulas', 'engine'],
	},
	{ name: 'examples', dir: 'examples', deps: ['sdk'] },
]

const PROJECT_BY_NAME = new Map(PROJECTS.map((project) => [project.name, project]))
const args = process.argv.slice(2)
const baseArg = args.find((arg) => arg.startsWith('--base='))
const headArg = args.find((arg) => arg.startsWith('--head='))
const passthrough = args.filter((arg) => !arg.startsWith('--base=') && !arg.startsWith('--head='))
const base = baseArg?.slice('--base='.length) || process.env.ASCEND_TEST_BASE || 'origin/main'
const head = headArg?.slice('--head='.length) || process.env.ASCEND_TEST_HEAD || 'HEAD'

function git(args: readonly string[]): string[] {
	const proc = Bun.spawnSync(['git', ...args], { stdout: 'pipe', stderr: 'pipe' })
	if (proc.exitCode !== 0) return []
	return new TextDecoder()
		.decode(proc.stdout)
		.split('\n')
		.map((line) => line.trim())
		.filter(Boolean)
}

function projectForPath(path: string): Project | null {
	return (
		PROJECTS.find((project) => path === project.dir || path.startsWith(`${project.dir}/`)) ?? null
	)
}

function dependentsOf(projectName: ProjectName): ProjectName[] {
	const result = new Set<ProjectName>()
	const visit = (target: ProjectName) => {
		for (const project of PROJECTS) {
			if (!project.deps.includes(target) || result.has(project.name)) continue
			result.add(project.name)
			visit(project.name)
		}
	}
	visit(projectName)
	return [...result]
}

function shouldRunFullSuite(path: string): boolean {
	return (
		path === 'package.json' ||
		path === 'bun.lock' ||
		path === 'tsconfig.json' ||
		path === 'tsconfig.base.json' ||
		path === 'AGENTS.md' ||
		path.startsWith('.github/') ||
		path.startsWith('scripts/') ||
		path.startsWith('docs/')
	)
}

const changed = new Set([
	...git(['diff', '--name-only', `${base}...${head}`]),
	...git(['diff', '--name-only']),
	...git(['diff', '--name-only', '--cached']),
])

if (changed.size === 0) {
	console.log(`No changed files detected against ${base}; running no tests.`)
	process.exit(0)
}

if ([...changed].some(shouldRunFullSuite)) {
	console.log('Repository-level change detected; running the full test suite.')
	const proc = Bun.spawn(['bun', 'run', 'test:ci', ...passthrough], {
		stdout: 'inherit',
		stderr: 'inherit',
	})
	process.exit(await proc.exited)
}

const affected = new Set<ProjectName>()
for (const path of changed) {
	const project = projectForPath(path)
	if (!project) continue
	affected.add(project.name)
	for (const dependent of dependentsOf(project.name)) affected.add(dependent)
}

const filters = [...affected]
	.map((name) => PROJECT_BY_NAME.get(name)?.dir)
	.filter((dir): dir is string => dir !== undefined && existsSync(dir))
	.sort()

if (filters.length === 0) {
	console.log(`No test-owning project matched ${changed.size} changed file(s); running no tests.`)
	process.exit(0)
}

console.log(`Changed files: ${changed.size}`)
console.log(`Affected test filters: ${filters.join(', ')}`)

const proc = Bun.spawn(
	[
		'bun',
		'test',
		'--recursive',
		'--parallel=8',
		'--timeout',
		'30000',
		'--only-failures',
		'--pass-with-no-tests',
		...passthrough,
		...filters,
	],
	{ stdout: 'inherit', stderr: 'inherit' },
)
process.exit(await proc.exited)
