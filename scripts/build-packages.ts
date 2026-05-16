import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const packageNames = ['schema', 'core', 'formulas', 'engine', 'io-xlsx', 'io-csv', 'verify', 'sdk']
const appNames = ['cli', 'api', 'mcp']
const appBins: Record<string, Record<string, string>> = {
	cli: { ascend: './index.js' },
	api: { 'ascend-api': './index.js' },
	mcp: { 'ascend-mcp': './index.js' },
}
const sdkAgentDocAssets = [
	'llms.txt',
	'llms-full.txt',
	'docs/AGENT_API.md',
	'docs/AGENT_WORKFLOW.md',
	'docs/SECURITY.md',
	'docs/VERSIONING.md',
	'docs/openapi.yaml',
	'examples/README.md',
	'examples/package.json',
	'examples/agent-safe-edit.ts',
	'examples/package-install-safe-edit.ts',
	'examples/agent-safe-edit-http.ts',
	'examples/agent-safe-edit-http.md',
	'examples/agent-safe-edit-mcp.ts',
	'examples/agent-safe-edit-mcp.md',
	'examples/untrusted-workbook-report.md',
	'examples/mcp-setup.md',
	'examples/create-from-scratch.ts',
	'examples/read-modify-save.ts',
	'examples/batch-ops.ts',
	'examples/formula-eval.ts',
	'examples/csv-convert.ts',
] as const
const repoRoot = fileURLToPath(new URL('..', import.meta.url))
const packageVersions = new Map<string, string>()

type PackageJson = {
	name: string
	version: string
	private?: boolean
	type?: string
	main?: string
	bin?: Record<string, string>
	dependencies?: Record<string, string>
	devDependencies?: Record<string, string>
}

for (const name of packageNames) {
	const packageRoot = join(repoRoot, 'packages', name)
	const packageJson = JSON.parse(
		await readFile(join(packageRoot, 'package.json'), 'utf8'),
	) as PackageJson
	packageVersions.set(packageJson.name, packageJson.version)
}

function rewriteInternalDeps(
	deps: Record<string, string> | undefined,
): Record<string, string> | undefined {
	if (!deps) return undefined
	const next: Record<string, string> = {}
	for (const [dep, version] of Object.entries(deps)) {
		next[dep] = dep.startsWith('@ascend/') ? (packageVersions.get(dep) ?? version) : version
	}
	return next
}

async function writePublishManifest(
	packageRoot: string,
	distDir: string,
	bin?: Record<string, string>,
): Promise<void> {
	const source = JSON.parse(
		await readFile(join(packageRoot, 'package.json'), 'utf8'),
	) as PackageJson & Record<string, unknown>

	const publishManifest = {
		name: source.name,
		version: source.version,
		private: false,
		type: source.type ?? 'module',
		main: './index.js',
		exports: {
			'.': {
				import: './index.js',
				default: './index.js',
				types: './index.d.ts',
			},
		},
		types: './index.d.ts',
		...(rewriteInternalDeps(source.dependencies)
			? { dependencies: rewriteInternalDeps(source.dependencies) }
			: {}),
		...(bin ? { bin } : {}),
	}

	await writeFile(join(distDir, 'package.json'), `${JSON.stringify(publishManifest, null, '\t')}\n`)
}

async function buildPackage(name: string): Promise<void> {
	const packageRoot = join(repoRoot, 'packages', name)
	const distDir = join(packageRoot, 'dist')
	await mkdir(distDir, { recursive: true })
	await rm(join(distDir, 'index.js'), { force: true }).catch(() => {})
	await rm(join(distDir, 'index.js.map'), { force: true }).catch(() => {})

	const result = await Bun.build({
		entrypoints: [join(packageRoot, 'src', 'index.ts')],
		outdir: distDir,
		format: 'esm',
		target: 'node',
		sourcemap: 'external',
		packages: 'external',
	})

	if (!result.success) {
		for (const log of result.logs) console.error(log)
		throw new Error(`build failed for ${name}`)
	}

	await writePublishManifest(packageRoot, distDir)
	if (name === 'sdk') await copySdkAgentDocs(distDir)
}

async function copySdkAgentDocs(distDir: string): Promise<void> {
	for (const path of sdkAgentDocAssets) {
		const target = join(distDir, path)
		await mkdir(dirname(target), { recursive: true })
		await cp(join(repoRoot, path), target)
	}
}

async function buildApp(name: string): Promise<void> {
	const appRoot = join(repoRoot, 'apps', name)
	const distDir = join(appRoot, 'dist')
	await mkdir(distDir, { recursive: true })
	await rm(join(distDir, 'index.js'), { force: true }).catch(() => {})
	await rm(join(distDir, 'index.js.map'), { force: true }).catch(() => {})

	const result = await Bun.build({
		entrypoints: [join(appRoot, 'src', 'index.ts')],
		outdir: distDir,
		format: 'esm',
		target: 'node',
		sourcemap: 'external',
		packages: 'external',
	})

	if (!result.success) {
		for (const log of result.logs) console.error(log)
		throw new Error(`build failed for ${name}`)
	}

	await writePublishManifest(appRoot, distDir, appBins[name])
}

for (const name of packageNames) {
	console.log(`building @ascend/${name}`)
	await buildPackage(name)
}

for (const name of appNames) {
	console.log(`building @ascend/${name}`)
	await buildApp(name)
}
