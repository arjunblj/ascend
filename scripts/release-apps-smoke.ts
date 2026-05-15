import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = fileURLToPath(new URL('..', import.meta.url))
const releaseRoot = '/private/tmp/ascend-apps-local-release'
const artifactRoot = join(releaseRoot, 'artifacts')
const artifactPackagesRoot = join(artifactRoot, 'packages')
const appRoot = join(releaseRoot, 'consumer')

const packageNames = ['schema', 'core', 'formulas', 'engine', 'io-xlsx', 'io-csv', 'verify', 'sdk']
const appNames = ['cli', 'api', 'mcp']
const internalPackageNames = new Map(packageNames.map((name) => [`@ascend/${name}`, name]))
const appTarballs = new Map(
	appNames.map((name) => [name, join(artifactRoot, `ascend-${name}-0.0.0.tgz`)]),
)

type PackageJson = {
	name: string
	version: string
	private?: boolean
	main?: string
	dependencies?: Record<string, string>
	bundledDependencies?: string[]
	[key: string]: unknown
}

async function main(): Promise<void> {
	await run(['bun', 'run', 'build:js'], repoRoot)
	await prepareArtifacts()
	await prepareConsumerApp()
	await run(['bun', 'install'], appRoot)
	await run([join(appRoot, 'node_modules', '.bin', 'ascend'), '--version'], appRoot)
	await run(['bun', 'run', 'app-smoke.ts'], appRoot)
	console.log(`App release smoke passed: ${appRoot}`)
	for (const [name, tarball] of appTarballs) console.log(`${name} artifact: ${tarball}`)
}

async function prepareArtifacts(): Promise<void> {
	await rm(releaseRoot, { recursive: true, force: true })
	await mkdir(artifactPackagesRoot, { recursive: true })

	for (const name of packageNames) {
		const sourceDist = join(repoRoot, 'packages', name, 'dist')
		const target = join(artifactPackagesRoot, name)
		await cp(sourceDist, target, { recursive: true })
		await rewriteArtifactManifest(target)
	}

	for (const name of appNames) {
		const sourceDist = join(repoRoot, 'apps', name, 'dist')
		const target = join(artifactPackagesRoot, name)
		await cp(sourceDist, target, { recursive: true })
		await rewriteArtifactManifest(target)
		await bundleInternalPackages(target)
		await run(
			[
				'bun',
				'pm',
				'pack',
				'--filename',
				appTarballs.get(name) as string,
				'--ignore-scripts',
				'--quiet',
			],
			target,
		)
	}
}

async function rewriteArtifactManifest(packageRoot: string): Promise<void> {
	const manifestPath = join(packageRoot, 'package.json')
	const manifest = JSON.parse(await readFile(manifestPath, 'utf8')) as PackageJson
	const dependencies = manifest.dependencies
		? Object.fromEntries(
				Object.entries(manifest.dependencies).filter(([dep]) => !internalPackageNames.has(dep)),
			)
		: undefined
	const next: PackageJson = {
		...manifest,
		private: false,
		dependencies,
	}
	await writeFile(join(packageRoot, 'package.json'), `${JSON.stringify(next, null, '\t')}\n`)
}

async function bundleInternalPackages(packageRoot: string): Promise<void> {
	const publicDeps: Record<string, string> = {}
	for (const name of packageNames) {
		const manifest = JSON.parse(
			await readFile(join(artifactPackagesRoot, name, 'package.json'), 'utf8'),
		) as PackageJson
		for (const [dep, version] of Object.entries(manifest.dependencies ?? {})) {
			if (!internalPackageNames.has(dep)) publicDeps[dep] = version
		}
	}

	const nodeModulesRoot = join(packageRoot, 'node_modules', '@ascend')
	await mkdir(nodeModulesRoot, { recursive: true })
	for (const name of packageNames) {
		await cp(join(artifactPackagesRoot, name), join(nodeModulesRoot, name), { recursive: true })
	}

	const manifestPath = join(packageRoot, 'package.json')
	const manifest = JSON.parse(await readFile(manifestPath, 'utf8')) as PackageJson
	const ownPublicDeps = Object.fromEntries(
		Object.entries(manifest.dependencies ?? {}).filter(([dep]) => !internalPackageNames.has(dep)),
	)
	const next: PackageJson = {
		...manifest,
		dependencies: { ...publicDeps, ...ownPublicDeps },
		bundledDependencies: packageNames.map((name) => `@ascend/${name}`),
	}
	await writeFile(manifestPath, `${JSON.stringify(next, null, '\t')}\n`)
}

async function prepareConsumerApp(): Promise<void> {
	await mkdir(appRoot, { recursive: true })
	await writeFile(
		join(appRoot, 'package.json'),
		`${JSON.stringify(
			{
				name: 'ascend-apps-release-consumer',
				private: true,
				type: 'module',
				dependencies: Object.fromEntries(
					appNames.map((name) => [`@ascend/${name}`, `file:${appTarballs.get(name)}`]),
				),
			},
			null,
			'\t',
		)}\n`,
	)
	await writeFile(
		join(appRoot, 'app-smoke.ts'),
		`import { createApiFetch, createServer as createApiServer } from '@ascend/api'
import { createServer as createMcpServer } from '@ascend/mcp'

if (typeof createApiFetch !== 'function') throw new Error('missing createApiFetch export')
if (typeof createApiServer !== 'function') throw new Error('missing createApiServer export')
if (typeof createMcpServer !== 'function') throw new Error('missing createMcpServer export')

const apiFetch = createApiFetch()
const capabilitiesResponse = await apiFetch(new Request('http://ascend.local/capabilities'))
const capabilities = await capabilitiesResponse.json()
if (!capabilities.ok) throw new Error('installed API capabilities request failed')
if (!Array.isArray(capabilities.data?.capabilities) || capabilities.data.capabilities.length === 0) {
	throw new Error('installed API capabilities request returned no capabilities')
}

const mcpServer = createMcpServer()
const tools = mcpServer._registeredTools
const resources = mcpServer._registeredResources
const mcpCapabilities = await tools['ascend.capabilities']?.handler({})
if (!mcpCapabilities?.structuredContent?.ok) {
	throw new Error('installed MCP capabilities tool failed')
}
const mcpResource = await resources['ascend://capabilities']?.readCallback(
	new URL('ascend://capabilities'),
)
if (!mcpResource?.contents?.[0]?.text?.includes('"capabilities"')) {
	throw new Error('installed MCP capabilities resource failed')
}

console.log(JSON.stringify({
	apiFetchExport: typeof createApiFetch,
	apiServerExport: typeof createApiServer,
	apiCapabilities: capabilities.data.capabilities.length,
	mcpServerExport: typeof createMcpServer,
	mcpTools: Object.keys(tools).length,
	mcpCapabilities:
		mcpCapabilities.structuredContent.data.capabilities.length,
}))
`,
	)
}

async function run(cmd: string[], cwd: string): Promise<void> {
	console.log(`$ ${cmd.join(' ')}`)
	const proc = Bun.spawn(cmd, {
		cwd,
		stdout: 'inherit',
		stderr: 'inherit',
	})
	const exitCode = await proc.exited
	if (exitCode !== 0) {
		throw new Error(`${cmd.join(' ')} failed with exit code ${exitCode}`)
	}
}

main().catch((error) => {
	console.error(error instanceof Error ? error.message : error)
	process.exit(1)
})
