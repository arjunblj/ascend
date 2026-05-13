#!/usr/bin/env bun
import { existsSync, readFileSync } from 'node:fs'
import { join, normalize, relative } from 'node:path'

interface PackageJson {
	readonly name: string
	readonly dependencies?: Record<string, string>
}

interface TsConfig {
	readonly references?: readonly { readonly path: string }[]
}

interface Project {
	readonly name: string
	readonly dir: string
	readonly deps: readonly string[]
}

const packageProjectNames = new Map<string, string>([
	['@ascend/schema', 'schema'],
	['@ascend/core', 'core'],
	['@ascend/formulas', 'formulas'],
	['@ascend/engine', 'engine'],
	['@ascend/io-xlsx', 'io-xlsx'],
	['@ascend/io-csv', 'io-csv'],
	['@ascend/verify', 'verify'],
	['@ascend/sdk', 'sdk'],
	['@ascend/tui', 'tui'],
	['@ascend/cli', 'cli'],
	['@ascend/api', 'api'],
	['@ascend/mcp', 'mcp'],
	['@ascend/benchmarks', 'fixtures-benchmarks'],
	['@ascend/corpus-tests', 'fixtures-corpus'],
	['@ascend/formula-conformance', 'fixtures-formulas'],
	['@ascend/examples', 'examples'],
])

const fixtureOnlyProjects = new Map<string, Project>([
	[
		'fixtures-xlsx',
		{
			name: 'fixtures-xlsx',
			dir: 'fixtures/xlsx',
			deps: [
				'schema',
				'core',
				'engine',
				'io-xlsx',
				'sdk',
				'fixtures-benchmarks',
				'fixtures-corpus',
			],
		},
	],
])

const errors: string[] = []

function readJson<T>(path: string): T {
	return JSON.parse(readFileSync(path, 'utf8')) as T
}

function normalizePath(path: string): string {
	return normalize(path).replaceAll('\\', '/')
}

function sorted(values: Iterable<string>): string[] {
	return [...values].sort((a, b) => a.localeCompare(b))
}

function formatList(values: readonly string[]): string {
	return values.length === 0 ? '(none)' : values.join(', ')
}

function sameSet(actual: readonly string[], expected: readonly string[]): boolean {
	if (actual.length !== expected.length) return false
	return actual.every((value, index) => value === expected[index])
}

function projectForPackageJson(packageJsonPath: string): Project {
	const dir = normalizePath(packageJsonPath.replace(/\/package\.json$/, ''))
	const packageJson = readJson<PackageJson>(packageJsonPath)
	const name = packageProjectNames.get(packageJson.name)
	if (!name) {
		errors.push(`${packageJsonPath}: unknown workspace package name ${packageJson.name}`)
		return { name: packageJson.name, dir, deps: [] }
	}
	return {
		name,
		dir,
		deps: workspaceDeps(packageJson),
	}
}

function workspaceDeps(packageJson: PackageJson): string[] {
	const deps = packageJson.dependencies ?? {}
	return sorted(
		Object.keys(deps)
			.map((depName) => packageProjectNames.get(depName))
			.filter((depName): depName is string => depName !== undefined),
	)
}

function packageJsonPaths(): string[] {
	const paths: string[] = []
	for (const root of ['packages', 'apps', 'fixtures']) {
		const proc = Bun.spawnSync(
			['find', root, '-mindepth', '2', '-maxdepth', '2', '-name', 'package.json'],
			{
				stdout: 'pipe',
				stderr: 'pipe',
			},
		)
		if (proc.exitCode !== 0) {
			errors.push(`failed to scan ${root} package.json files`)
			continue
		}
		paths.push(
			...new TextDecoder()
				.decode(proc.stdout)
				.split('\n')
				.map((line) => line.trim())
				.filter(Boolean),
		)
	}
	if (existsSync('examples/package.json')) paths.push('examples/package.json')
	return sorted(paths)
}

function parseAffectedProjects(): Project[] {
	const source = readFileSync('scripts/affected-tests.ts', 'utf8')
	const projects: Project[] = []
	const pattern = /\{\s*name:\s*'([^']+)'\s*,\s*dir:\s*'([^']+)'\s*,\s*deps:\s*\[([^\]]*)\]/g
	for (const match of source.matchAll(pattern)) {
		const [, name, dir, depsSource] = match
		if (!name || !dir || depsSource === undefined) continue
		const deps = sorted(
			[...depsSource.matchAll(/'([^']+)'/g)].map((depMatch) => depMatch[1] as string),
		)
		projects.push({ name, dir, deps })
	}
	return projects
}

function tsconfigReferences(project: Project, dirByName: Map<string, string>): string[] {
	const tsconfigPath = join(project.dir, 'tsconfig.json')
	if (!existsSync(tsconfigPath)) return []
	const tsconfig = readJson<TsConfig>(tsconfigPath)
	const refs = tsconfig.references ?? []
	return sorted(
		refs
			.map((ref) => normalizePath(relative('.', join(project.dir, ref.path))))
			.map((refDir) => {
				for (const [name, dir] of dirByName) {
					if (normalizePath(dir) === refDir) return name
				}
				return undefined
			})
			.filter((name): name is string => name !== undefined),
	)
}

const packageProjects = packageJsonPaths().map(projectForPackageJson)
const expectedProjects = [...packageProjects, ...fixtureOnlyProjects.values()].sort((a, b) =>
	a.name.localeCompare(b.name),
)
const expectedByName = new Map(expectedProjects.map((project) => [project.name, project]))
const dirByName = new Map(expectedProjects.map((project) => [project.name, project.dir]))
const affectedProjects = parseAffectedProjects()
const affectedByName = new Map(affectedProjects.map((project) => [project.name, project]))

for (const project of expectedProjects) {
	if (!existsSync(project.dir))
		errors.push(`${project.name}: expected directory is missing: ${project.dir}`)
}

for (const affected of affectedProjects) {
	if (!expectedByName.has(affected.name)) {
		errors.push(`scripts/affected-tests.ts: unknown affected project ${affected.name}`)
	}
	for (const dep of affected.deps) {
		if (!expectedByName.has(dep)) {
			errors.push(`scripts/affected-tests.ts: ${affected.name} has unknown dependency ${dep}`)
		}
	}
}

for (const project of expectedProjects) {
	const affected = affectedByName.get(project.name)
	if (!affected) {
		errors.push(`scripts/affected-tests.ts: missing project ${project.name}`)
		continue
	}
	const actualDeps = sorted(affected.deps)
	const expectedDeps = sorted(project.deps)
	if (!sameSet(actualDeps, expectedDeps)) {
		errors.push(
			`scripts/affected-tests.ts: ${project.name} deps differ; expected ${formatList(expectedDeps)}, got ${formatList(actualDeps)}`,
		)
	}
}

for (const project of packageProjects) {
	if (!existsSync(join(project.dir, 'tsconfig.json'))) continue
	const expectedRefs = sorted(
		project.deps.filter((dep) => existsSync(join(dirByName.get(dep) ?? '', 'tsconfig.json'))),
	)
	const actualRefs = tsconfigReferences(project, dirByName)
	if (!sameSet(actualRefs, expectedRefs)) {
		errors.push(
			`${project.dir}/tsconfig.json: references differ; expected ${formatList(expectedRefs)}, got ${formatList(actualRefs)}`,
		)
	}
}

if (errors.length > 0) {
	console.error('CI graph check failed:')
	for (const error of errors) console.error(`- ${error}`)
	process.exit(1)
}

console.log(`CI graph check passed (${expectedProjects.length} projects).`)
