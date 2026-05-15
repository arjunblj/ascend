import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = fileURLToPath(new URL('..', import.meta.url))
const releaseRoot = '/private/tmp/ascend-rc-gate'
const artifactRoot = join(releaseRoot, 'artifacts')
const artifactPackagesRoot = join(artifactRoot, 'packages')
const consumerRoot = join(releaseRoot, 'consumer')

const libraryNames = ['schema', 'core', 'formulas', 'engine', 'io-xlsx', 'io-csv', 'verify', 'sdk']
const packageNames = ['sdk', 'cli', 'api', 'mcp']
const appNames = ['cli', 'api', 'mcp']
const internalPackageNames = new Map(libraryNames.map((name) => [`@ascend/${name}`, name]))
const packageTarballs = new Map(
	packageNames.map((name) => [name, join(artifactRoot, `ascend-${name}-0.0.0.tgz`)]),
)

type PackageJson = {
	name: string
	version: string
	private?: boolean
	dependencies?: Record<string, string>
	bundledDependencies?: string[]
	bin?: Record<string, string>
	exports?: unknown
	[key: string]: unknown
}

class GateError extends Error {
	constructor(
		message: string,
		readonly blocker: {
			lane: string
			files: string[]
			command: string
			output: string
		},
	) {
		super(message)
		this.name = 'GateError'
	}
}

async function main(): Promise<void> {
	try {
		await mkdir(releaseRoot, { recursive: true })
		await runStep('build release JS artifacts', ['bun', 'run', 'build:js'], repoRoot, {
			lane: 'release packageability',
			files: ['scripts/build-packages.ts', 'packages/*/package.json', 'apps/*/package.json'],
		})
		await prepareArtifacts()
		await prepareConsumerApp()
		await runStep(
			'install RC consumer app',
			['bun', 'install', '--backend=copyfile', '--cache-dir', join(releaseRoot, 'cache')],
			consumerRoot,
			{
				lane: 'release packageability',
				files: ['scripts/release-rc-gate.ts', 'package.json'],
			},
		)
		await assertInstalledManifestContracts()
		await runStep(
			'run packaged SDK/CLI/API/MCP workbook proof',
			['bun', 'run', 'rc-gate.ts'],
			consumerRoot,
			{
				lane: 'SDK/CLI/API/MCP contract',
				files: [
					'packages/sdk/src/index.ts',
					'apps/cli/src/index.ts',
					'apps/api/src/index.ts',
					'apps/mcp/src/index.ts',
				],
			},
		)

		console.log(
			JSON.stringify(
				{
					ok: true,
					gate: 'release:rc:gate',
					consumerRoot,
					artifacts: Object.fromEntries(packageTarballs),
					proof: join(consumerRoot, 'rc-gate-proof.json'),
				},
				null,
				'\t',
			),
		)
		console.log(`RC gate passed: ${consumerRoot}`)
	} catch (error) {
		emitFailure(error)
		process.exit(1)
	}
}

async function prepareArtifacts(): Promise<void> {
	await rm(releaseRoot, { recursive: true, force: true })
	await mkdir(artifactPackagesRoot, { recursive: true })

	for (const name of libraryNames) {
		const sourceDist = join(repoRoot, 'packages', name, 'dist')
		const target = join(artifactPackagesRoot, name)
		await cp(sourceDist, target, { recursive: true })
		await rewriteArtifactManifest(target, { removeInternalDeps: false })
	}

	await bundleInternalPackages(join(artifactPackagesRoot, 'sdk'), {
		excludePackage: 'sdk',
		includeOwnPublicDeps: false,
	})
	await packArtifact('sdk', join(artifactPackagesRoot, 'sdk'))

	for (const name of appNames) {
		const sourceDist = join(repoRoot, 'apps', name, 'dist')
		const target = join(artifactPackagesRoot, name)
		await cp(sourceDist, target, { recursive: true })
		await rewriteArtifactManifest(target, { removeInternalDeps: true })
		await bundleInternalPackages(target, { includeOwnPublicDeps: true })
		await packArtifact(name, target)
	}
}

async function rewriteArtifactManifest(
	packageRoot: string,
	options: { removeInternalDeps: boolean },
): Promise<void> {
	const manifestPath = join(packageRoot, 'package.json')
	const manifest = JSON.parse(await readFile(manifestPath, 'utf8')) as PackageJson
	const dependencies = manifest.dependencies
		? Object.fromEntries(
				Object.entries(manifest.dependencies)
					.map(([dep, version]) => [
						dep,
						internalPackageNames.has(dep) ? `file:../${internalPackageNames.get(dep)}` : version,
					])
					.filter(([dep]) => !(options.removeInternalDeps && internalPackageNames.has(dep))),
			)
		: undefined
	const next: PackageJson = {
		...manifest,
		private: false,
		...(dependencies ? { dependencies } : {}),
	}
	await writeFile(manifestPath, `${JSON.stringify(next, null, '\t')}\n`)
}

async function bundleInternalPackages(
	packageRoot: string,
	options: { excludePackage?: string; includeOwnPublicDeps: boolean },
): Promise<void> {
	const publicDeps: Record<string, string> = {}
	for (const name of libraryNames) {
		const manifest = JSON.parse(
			await readFile(join(artifactPackagesRoot, name, 'package.json'), 'utf8'),
		) as PackageJson
		for (const [dep, version] of Object.entries(manifest.dependencies ?? {})) {
			if (!internalPackageNames.has(dep)) publicDeps[dep] = version
		}
	}

	const nodeModulesRoot = join(packageRoot, 'node_modules', '@ascend')
	await mkdir(nodeModulesRoot, { recursive: true })
	for (const name of libraryNames) {
		if (name === options.excludePackage) continue
		await cp(join(artifactPackagesRoot, name), join(nodeModulesRoot, name), { recursive: true })
	}

	const manifestPath = join(packageRoot, 'package.json')
	const manifest = JSON.parse(await readFile(manifestPath, 'utf8')) as PackageJson
	const ownPublicDeps = options.includeOwnPublicDeps
		? Object.fromEntries(
				Object.entries(manifest.dependencies ?? {}).filter(
					([dep]) => !internalPackageNames.has(dep),
				),
			)
		: {}
	const bundled = libraryNames
		.filter((name) => name !== options.excludePackage)
		.map((name) => `@ascend/${name}`)
	const next: PackageJson = {
		...manifest,
		dependencies: { ...publicDeps, ...ownPublicDeps },
		bundledDependencies: bundled,
	}
	await writeFile(manifestPath, `${JSON.stringify(next, null, '\t')}\n`)
}

async function packArtifact(name: string, packageRoot: string): Promise<void> {
	await runStep(
		`pack @ascend/${name}`,
		[
			'bun',
			'pm',
			'pack',
			'--filename',
			packageTarballs.get(name) as string,
			'--ignore-scripts',
			'--quiet',
		],
		packageRoot,
		{
			lane: 'release packageability',
			files: ['scripts/release-rc-gate.ts', 'scripts/build-packages.ts'],
		},
	)
}

async function prepareConsumerApp(): Promise<void> {
	await mkdir(consumerRoot, { recursive: true })
	await writeFile(
		join(consumerRoot, 'package.json'),
		`${JSON.stringify(
			{
				name: 'ascend-rc-gate-consumer',
				private: true,
				type: 'module',
				dependencies: Object.fromEntries(
					packageNames.map((name) => [`@ascend/${name}`, `file:${packageTarballs.get(name)}`]),
				),
			},
			null,
			'\t',
		)}\n`,
	)
	await writeFile(join(consumerRoot, 'rc-gate.ts'), consumerProofSource())
}

async function assertInstalledManifestContracts(): Promise<void> {
	const consumerManifest = JSON.parse(
		await readFile(join(consumerRoot, 'package.json'), 'utf8'),
	) as PackageJson
	if ('overrides' in consumerManifest || 'resolutions' in consumerManifest) {
		throw new GateError('consumer app uses package manager overrides', {
			lane: 'release packageability',
			files: ['scripts/release-rc-gate.ts'],
			command: 'bun run release:rc:gate',
			output: JSON.stringify(consumerManifest, null, 2),
		})
	}
	for (const dep of Object.keys(consumerManifest.dependencies ?? {})) {
		if (!['@ascend/sdk', '@ascend/cli', '@ascend/api', '@ascend/mcp'].includes(dep)) {
			throw new GateError(`consumer app has unexpected dependency ${dep}`, {
				lane: 'release packageability',
				files: ['scripts/release-rc-gate.ts'],
				command: 'bun run release:rc:gate',
				output: JSON.stringify(consumerManifest.dependencies, null, 2),
			})
		}
	}

	for (const name of packageNames) {
		const manifest = JSON.parse(
			await readFile(join(consumerRoot, 'node_modules', '@ascend', name, 'package.json'), 'utf8'),
		) as PackageJson
		for (const [dep, version] of Object.entries(manifest.dependencies ?? {})) {
			if (version.startsWith('workspace:') || version.startsWith('file:')) {
				throw new GateError(
					`installed @ascend/${name} has repo-local dependency ${dep}: ${version}`,
					{
						lane: 'release packageability',
						files: [
							`${name === 'sdk' ? 'packages' : 'apps'}/${name}/package.json`,
							'scripts/build-packages.ts',
						],
						command: 'bun run release:rc:gate',
						output: JSON.stringify(manifest, null, 2),
					},
				)
			}
		}
		if (!manifest.exports) {
			throw new GateError(`installed @ascend/${name} is missing root exports`, {
				lane: 'release packageability',
				files: ['scripts/build-packages.ts'],
				command: 'bun run release:rc:gate',
				output: JSON.stringify(manifest, null, 2),
			})
		}
	}
}

function consumerProofSource(): string {
	return `import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { AscendWorkbook } from '@ascend/sdk'
import { createApiFetch, createServer as createApiServer } from '@ascend/api'
import { createServer as createMcpServer } from '@ascend/mcp'

if (typeof AscendWorkbook?.open !== 'function') throw new Error('missing SDK AscendWorkbook.open export')
if (typeof createApiFetch !== 'function') throw new Error('missing API createApiFetch export')
if (typeof createApiServer !== 'function') throw new Error('missing API createServer export')
if (typeof createMcpServer !== 'function') throw new Error('missing MCP createServer export')

const cwd = process.cwd()
const ascendBin = join(cwd, 'node_modules', '.bin', 'ascend')
const input = join(cwd, 'rc-input.xlsx')
const output = join(cwd, 'rc-output.xlsx')
const ops = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 125 }] }]
const opsPath = join(cwd, 'rc-ops.json')
await writeFile(opsPath, JSON.stringify(ops))

async function runCli(args) {
	const proc = Bun.spawn([ascendBin, ...args], { cwd, stdout: 'pipe', stderr: 'pipe' })
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error('ascend ' + args.join(' ') + ' failed with exit code ' + exitCode + '\\n' + stdout + stderr)
	}
	return stdout
}

async function runCliJson(args) {
	const parsed = JSON.parse(await runCli([...args, '--json']))
	if (!parsed.ok) throw new Error('ascend ' + args.join(' ') + ' failed: ' + JSON.stringify(parsed.error))
	return parsed.data
}

async function apiJson(apiFetch, path, body) {
	const response = await apiFetch(new Request('http://ascend.local' + path, {
		method: 'POST',
		headers: { 'Content-Type': 'application/json' },
		body: JSON.stringify(body),
	}))
	const payload = await response.json()
	if (!payload.ok) throw new Error('API ' + path + ' failed: ' + JSON.stringify(payload.error))
	return payload.data
}

async function mcpTool(tools, name, args) {
	const handler = tools[name]?.handler
	if (typeof handler !== 'function') throw new Error('missing MCP tool ' + name)
	const result = await handler(args)
	const envelope = result.structuredContent
	if (!envelope?.ok) throw new Error('MCP ' + name + ' failed: ' + JSON.stringify(envelope?.error))
	return envelope.data
}

function numericCell(readResult, ref) {
	const cell = readResult.cells?.find((entry) => entry.ref === ref)
	const value = cell?.value
	if (typeof value === 'number') return value
	if (value?.kind === 'number') return value.value
	throw new Error('missing numeric cell ' + ref + ': ' + JSON.stringify(cell))
}

const created = AscendWorkbook.create()
created.set('Sheet1!A1', 'Revenue')
created.set('Sheet1!B1', 100)
created.setFormula('Sheet1!C1', '=B1*2')
await created.save(input)

const sdkInspect = (await AscendWorkbook.open(input)).inspect()
const cliPlan = await runCliJson(['plan', input, '--ops', opsPath])
if (cliPlan.preview?.wouldSucceed !== true) {
	throw new Error('CLI plan failed: ' + JSON.stringify(cliPlan.preview?.errors))
}

const apiFetch = createApiFetch()
const apiCommit = await apiJson(apiFetch, '/commit', {
	file: input,
	ops,
	output,
	expectSha256: cliPlan.inputSha256,
	compact: true,
})

const reopened = await AscendWorkbook.open(output)
const sdkCheck = reopened.check()
const b1 = reopened.get('Sheet1!B1')
const c1 = reopened.get('Sheet1!C1')
if (!sdkCheck.valid) throw new Error('SDK reopen check failed: ' + JSON.stringify(sdkCheck.issues))
if (b1.kind !== 'number' || b1.value !== 125) throw new Error('unexpected SDK B1: ' + JSON.stringify(b1))
if (c1.kind !== 'number' || c1.value !== 250) throw new Error('unexpected SDK C1: ' + JSON.stringify(c1))

const cliCheck = await runCliJson(['check', output])
const cliRead = await runCliJson(['read', output, 'Sheet1!B1:C1'])
if (cliCheck.valid !== true) throw new Error('CLI check failed: ' + JSON.stringify(cliCheck))
if (numericCell(cliRead, 'B1') !== 125 || numericCell(cliRead, 'C1') !== 250) {
	throw new Error('CLI read values failed: ' + JSON.stringify(cliRead))
}

const mcpServer = createMcpServer()
const tools = mcpServer._registeredTools
const mcpInspect = await mcpTool(tools, 'ascend.inspect', { file: output })
const mcpCheck = await mcpTool(tools, 'ascend.check', { file: output })
const mcpRead = await mcpTool(tools, 'ascend.read', {
	file: output,
	sheet: 'Sheet1',
	range: 'B1:C1',
	format: 'cells',
})
if (mcpCheck.valid !== true) throw new Error('MCP check failed: ' + JSON.stringify(mcpCheck))
if (numericCell(mcpRead, 'B1') !== 125 || numericCell(mcpRead, 'C1') !== 250) {
	throw new Error('MCP read values failed: ' + JSON.stringify(mcpRead))
}

const apiCapabilities = await (await apiFetch(new Request('http://ascend.local/capabilities'))).json()
if (!apiCapabilities.ok || !Array.isArray(apiCapabilities.data?.capabilities)) {
	throw new Error('API capabilities contract failed: ' + JSON.stringify(apiCapabilities))
}
const cliDocs = await runCliJson(['docs', 'plan commit'])
if (!Array.isArray(cliDocs.results) || cliDocs.results.length === 0) {
	throw new Error('CLI docs contract failed: ' + JSON.stringify(cliDocs))
}
const mcpDocs = await mcpTool(tools, 'ascend.search_docs', { query: 'plan commit' })
if (!Array.isArray(mcpDocs.results) || mcpDocs.results.length === 0) {
	throw new Error('MCP docs contract failed: ' + JSON.stringify(mcpDocs))
}

const proof = {
	ok: true,
	workflow: ['inspect', 'plan', 'commit', 'reopen', 'verify'],
	surfaces: {
		sdk: {
			inspectSheets: sdkInspect.sheets.map((sheet) => sheet.name),
			reopenedValid: sdkCheck.valid,
			values: { b1: b1.value, c1: c1.value },
		},
		cli: {
			planWouldSucceed: cliPlan.preview.wouldSucceed,
			reopenedValid: cliCheck.valid,
			values: { b1: numericCell(cliRead, 'B1'), c1: numericCell(cliRead, 'C1') },
			docHits: cliDocs.results.length,
		},
		api: {
			createApiFetchExport: typeof createApiFetch,
			createServerExport: typeof createApiServer,
			outputSha256: apiCommit.outputSha256,
			capabilities: apiCapabilities.data.capabilities.length,
		},
		mcp: {
			createServerExport: typeof createMcpServer,
			inspectSheets: mcpInspect.sheets?.map((sheet) => sheet.name),
			reopenedValid: mcpCheck.valid,
			values: { b1: numericCell(mcpRead, 'B1'), c1: numericCell(mcpRead, 'C1') },
			docHits: mcpDocs.results.length,
			tools: Object.keys(tools).length,
		},
	},
	files: { input, output },
}

await writeFile(join(cwd, 'rc-gate-proof.json'), JSON.stringify(proof, null, '\\t') + '\\n')
console.log(JSON.stringify(proof))
`
}

async function runStep(
	label: string,
	cmd: string[],
	cwd: string,
	blocker: { lane: string; files: string[] },
): Promise<string> {
	console.log(`$ ${cmd.join(' ')}`)
	const proc = Bun.spawn(cmd, {
		cwd,
		stdout: 'pipe',
		stderr: 'pipe',
		env: { ...process.env, TMPDIR: releaseRoot },
	})
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (stdout) process.stdout.write(stdout)
	if (stderr) process.stderr.write(stderr)
	if (exitCode !== 0) {
		throw new GateError(`${label} failed`, {
			...blocker,
			command: `cd ${cwd} && ${cmd.join(' ')}`,
			output: `${stdout}${stderr}`.trim(),
		})
	}
	return stdout
}

function emitFailure(error: unknown): void {
	if (error instanceof GateError) {
		console.error(
			JSON.stringify(
				{
					ok: false,
					gate: 'release:rc:gate',
					blockers: [
						{
							ownerLane: error.blocker.lane,
							files: error.blocker.files,
							command: error.blocker.command,
							failureOutput: error.blocker.output,
							summary: error.message,
						},
					],
				},
				null,
				'\t',
			),
		)
		return
	}
	console.error(
		JSON.stringify(
			{
				ok: false,
				gate: 'release:rc:gate',
				blockers: [
					{
						ownerLane: 'release packageability',
						files: ['scripts/release-rc-gate.ts'],
						command: 'bun run release:rc:gate',
						failureOutput: error instanceof Error ? error.stack || error.message : String(error),
						summary: 'RC gate failed before a classified step completed',
					},
				],
			},
			null,
			'\t',
		),
	)
}

await main()
