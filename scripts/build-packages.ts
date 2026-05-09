import { mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const packageNames = ['schema', 'core', 'formulas', 'engine', 'io-xlsx', 'io-csv', 'verify', 'sdk']
const repoRoot = fileURLToPath(new URL('..', import.meta.url))
const packageVersions = new Map<string, string>()

type PackageJson = {
	name: string
	version: string
	private?: boolean
	type?: string
	main?: string
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

async function writePublishManifest(packageRoot: string, distDir: string): Promise<void> {
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
}

for (const name of packageNames) {
	console.log(`building @ascend/${name}`)
	await buildPackage(name)
}
