import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = fileURLToPath(new URL('..', import.meta.url))
const releaseRoot = '/private/tmp/ascend-sdk-local-release'
const artifactRoot = join(releaseRoot, 'artifacts')
const artifactPackagesRoot = join(artifactRoot, 'packages')
const sdkTarball = join(artifactRoot, 'ascend-sdk-0.0.0.tgz')
const appRoot = join(releaseRoot, 'consumer')

const packageNames = ['schema', 'core', 'formulas', 'engine', 'io-xlsx', 'io-csv', 'verify', 'sdk']
const internalPackageNames = new Map(packageNames.map((name) => [`@ascend/${name}`, name]))

type PackageJson = {
	name: string
	version: string
	private?: boolean
	main?: string
	dependencies?: Record<string, string>
	[key: string]: unknown
}

async function main(): Promise<void> {
	await run(['bun', 'run', 'build:js'], repoRoot)
	await prepareArtifacts()
	await prepareConsumerApp()
	await run(['bun', 'install'], appRoot)
	await run(['bun', 'run', 'sdk-smoke.ts'], appRoot)
	console.log(`SDK release smoke passed: ${appRoot}`)
	console.log(`SDK artifact: ${sdkTarball}`)
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
	await bundleInternalPackagesForSdk()
	await run(
		['bun', 'pm', 'pack', '--filename', sdkTarball, '--ignore-scripts', '--quiet'],
		join(artifactPackagesRoot, 'sdk'),
	)
}

async function rewriteArtifactManifest(packageRoot: string): Promise<void> {
	const manifestPath = join(packageRoot, 'package.json')
	const manifest = JSON.parse(await readFile(manifestPath, 'utf8')) as PackageJson
	const dependencies = manifest.dependencies
		? Object.fromEntries(
				Object.entries(manifest.dependencies).map(([dep, version]) => [
					dep,
					internalPackageNames.has(dep) ? `file:../${internalPackageNames.get(dep)}` : version,
				]),
			)
		: undefined
	const next: PackageJson = {
		...manifest,
		private: false,
		...(dependencies ? { dependencies } : {}),
	}
	await writeFile(manifestPath, `${JSON.stringify(next, null, '\t')}\n`)
}

async function bundleInternalPackagesForSdk(): Promise<void> {
	const sdkRoot = join(artifactPackagesRoot, 'sdk')
	const publicDeps: Record<string, string> = {}
	for (const name of packageNames) {
		const manifest = JSON.parse(
			await readFile(join(artifactPackagesRoot, name, 'package.json'), 'utf8'),
		) as PackageJson
		for (const [dep, version] of Object.entries(manifest.dependencies ?? {})) {
			if (!internalPackageNames.has(dep)) publicDeps[dep] = version
		}
	}

	const sdkNodeModules = join(sdkRoot, 'node_modules', '@ascend')
	await mkdir(sdkNodeModules, { recursive: true })
	for (const name of packageNames) {
		if (name === 'sdk') continue
		await cp(join(artifactPackagesRoot, name), join(sdkNodeModules, name), { recursive: true })
	}

	const manifestPath = join(sdkRoot, 'package.json')
	const manifest = JSON.parse(await readFile(manifestPath, 'utf8')) as PackageJson
	const bundledInternalDeps = packageNames
		.filter((name) => name !== 'sdk')
		.map((name) => `@ascend/${name}`)
	const next: PackageJson = {
		...manifest,
		dependencies: publicDeps,
		bundledDependencies: bundledInternalDeps,
	}
	await writeFile(manifestPath, `${JSON.stringify(next, null, '\t')}\n`)
}

async function prepareConsumerApp(): Promise<void> {
	await mkdir(appRoot, { recursive: true })
	await writeFile(
		join(appRoot, 'package.json'),
		`${JSON.stringify(
			{
				name: 'ascend-sdk-release-consumer',
				private: true,
				type: 'module',
				dependencies: {
					'@ascend/sdk': `file:${sdkTarball}`,
				},
			},
			null,
			'\t',
		)}\n`,
	)
	await writeFile(
		join(appRoot, 'sdk-smoke.ts'),
		`import { AscendWorkbook, commitAgentPlan, createAgentPlan, createPreparedAgentPlan, inspectWorkbookOpenPlan, readAgentDoc, searchAgentDocs } from '@ascend/sdk'

const input = '${join(appRoot, 'input.xlsx')}'
const output = '${join(appRoot, 'output.xlsx')}'
const preparedOutput = '${join(appRoot, 'prepared-output.xlsx')}'
const installedExampleInput = '${join(appRoot, 'installed-example-input.xlsx')}'
const installedExampleOutput = '${join(appRoot, 'installed-example-output.xlsx')}'
const installedExampleBin = '${join(appRoot, 'node_modules', '.bin', 'ascend-sdk-safe-edit')}'
const installedSdkManifest = '${join(appRoot, 'node_modules', '@ascend', 'sdk', 'package.json')}'

const created = AscendWorkbook.create()
created.set('Sheet1!A1', 'Revenue')
created.set('Sheet1!B1', 100)
created.setFormula('Sheet1!C1', '=B1*2')
await created.save(input)

const sourceBytes = await Bun.file(input).bytes()
const openPlan = inspectWorkbookOpenPlan(new Uint8Array(sourceBytes), { intent: 'edit-plan' })
if (openPlan.recommendedLoadOptions.mode !== 'full') {
	throw new Error(\`unexpected open plan: \${JSON.stringify(openPlan)}\`)
}

const opened = await AscendWorkbook.open(input, openPlan.recommendedLoadOptions)
const info = opened.inspect()
const ops = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 125 }] }]
const plan = await createAgentPlan(input, ops)
if (!plan.preview.wouldSucceed) throw new Error(\`plan failed: \${JSON.stringify(plan.preview.errors)}\`)

const commit = await commitAgentPlan(input, ops, { output, expectSha256: plan.inputSha256 })
const reopened = await AscendWorkbook.open(output)
const check = reopened.check()
const b1 = reopened.get('Sheet1!B1')
const c1 = reopened.get('Sheet1!C1')
const preparedOps = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 150 }] }]
const prepared = await createPreparedAgentPlan(input, preparedOps)
if (prepared.plan.planDigest !== prepared.planDigest) {
	throw new Error(\`prepared plan digest mismatch: \${prepared.plan.planDigest} !== \${prepared.planDigest}\`)
}
const preparedCommit = await prepared.commit({
	output: preparedOutput,
	expectSha256: prepared.inputSha256,
})
const preparedReopened = await AscendWorkbook.open(preparedOutput)
const preparedCheck = preparedReopened.check()
const preparedB1 = preparedReopened.get('Sheet1!B1')
const preparedC1 = preparedReopened.get('Sheet1!C1')
const llms = await readAgentDoc('llms.txt')
const examplesReadme = await readAgentDoc('examples/README.md')
const rootExample = await readAgentDoc('examples/agent-safe-edit.ts')
const packageInstallExample = await readAgentDoc('examples/package-install-safe-edit.ts')
const apiExample = await readAgentDoc('examples/agent-safe-edit-http.ts')
const mcpExample = await readAgentDoc('examples/agent-safe-edit-mcp.ts')
const sdkManifest = await Bun.file(installedSdkManifest).json()
const docHits = await searchAgentDocs({ query: 'plan commit' })
const rootExampleHits = await searchAgentDocs({ query: 'example:safe-edit root workflow' })
const packageInstallExampleHits = await searchAgentDocs({ query: 'installed SDK package safe edit workflow' })
const apiExampleHits = await searchAgentDocs({ query: 'runnable HTTP safe edit workflow' })
const mcpExampleHits = await searchAgentDocs({ query: 'runnable MCP safe edit workflow' })

if (!check.valid) throw new Error(\`check failed: \${JSON.stringify(check.issues)}\`)
if (b1.kind !== 'number' || b1.value !== 125) throw new Error(\`unexpected B1: \${JSON.stringify(b1)}\`)
if (c1.kind !== 'number' || c1.value !== 250) throw new Error(\`unexpected C1: \${JSON.stringify(c1)}\`)
if (!preparedCheck.valid) throw new Error(\`prepared check failed: \${JSON.stringify(preparedCheck.issues)}\`)
if (!preparedCommit.postWrite.reopened || !preparedCommit.postWrite.auditsPassed) {
	throw new Error(\`prepared commit did not reopen cleanly: \${JSON.stringify(preparedCommit.postWrite)}\`)
}
if (preparedB1.kind !== 'number' || preparedB1.value !== 150) {
	throw new Error(\`unexpected prepared B1: \${JSON.stringify(preparedB1)}\`)
}
if (preparedC1.kind !== 'number' || preparedC1.value !== 300) {
	throw new Error(\`unexpected prepared C1: \${JSON.stringify(preparedC1)}\`)
}
if (!llms?.includes('Ascend')) throw new Error('installed SDK could not read bundled llms.txt')
if (!examplesReadme?.includes('bun run example:safe-edit')) {
	throw new Error('installed SDK missing root safe-edit workflow examples')
}
if (
	!rootExample?.includes('createAgentWorkflowProofSummary') ||
	!rootExample?.includes('proofBundle')
) {
	throw new Error('installed SDK root safe-edit example missing proof bundle contract')
}
if (!packageInstallExample?.includes('ascend-sdk-safe-edit')) {
	throw new Error('installed SDK missing package-install safe-edit example')
}
if (sdkManifest.bin?.['ascend-sdk-safe-edit'] !== './examples/package-install-safe-edit.ts') {
	throw new Error('installed SDK manifest missing safe-edit bin: ' + JSON.stringify(sdkManifest.bin))
}
if (!apiExample?.includes("from '@ascend/api'")) {
	throw new Error('installed SDK missing runnable HTTP API safe-edit example')
}
if (!mcpExample?.includes("from '@ascend/mcp'")) {
	throw new Error('installed SDK missing runnable MCP safe-edit example')
}
if (docHits.length === 0) throw new Error('installed SDK docs search returned no hits')
if (!rootExampleHits.some((hit) => hit.path === 'examples/README.md')) {
	throw new Error('installed SDK docs search did not find root safe-edit workflow examples')
}
if (!packageInstallExampleHits.some((hit) => hit.path === 'examples/package-install-safe-edit.ts')) {
	throw new Error('installed SDK docs search did not find package-install safe-edit example')
}
if (!apiExampleHits.some((hit) => hit.path === 'examples/agent-safe-edit-http.ts')) {
	throw new Error('installed SDK docs search did not find runnable HTTP safe-edit example')
}
if (!mcpExampleHits.some((hit) => hit.path === 'examples/agent-safe-edit-mcp.ts')) {
	throw new Error('installed SDK docs search did not find runnable MCP safe-edit example')
}
const installedExampleProc = Bun.spawn(
	[installedExampleBin, installedExampleInput, installedExampleOutput],
	{
		cwd: process.cwd(),
		stdout: 'pipe',
		stderr: 'pipe',
	},
)
const [installedExampleStdout, installedExampleStderr, installedExampleExitCode] = await Promise.all([
	new Response(installedExampleProc.stdout).text(),
	new Response(installedExampleProc.stderr).text(),
	installedExampleProc.exited,
])
if (installedExampleExitCode !== 0) {
	throw new Error(
		\`installed SDK safe-edit bin failed with exit \${installedExampleExitCode}: \${installedExampleStdout}\${installedExampleStderr}\`,
	)
}
const installedExampleProof = JSON.parse(installedExampleStdout)
if (
	installedExampleProof.workflow !== 'installed-sdk-open-plan-trust-inspect-read-plan-commit-reopen-verify' ||
	installedExampleProof.verify?.cell?.value?.value !== 450 ||
	installedExampleProof.proofBundle?.safeToUse !== true
) {
	throw new Error(\`installed SDK safe-edit proof failed: \${installedExampleStdout}\`)
}
const installedProofGates = installedExampleProof.proofBundle?.whySafe?.map(
	(gate) => [gate.gate, gate.ok],
)
if (
	JSON.stringify(installedProofGates) !==
	JSON.stringify([
		['open-plan', true],
		['trust', true],
		['plan-linked', true],
		['plan', true],
		['write-policy', true],
		['commit', true],
		['reopen-verify', true],
		['package-graph', true],
	])
) {
	throw new Error(\`installed SDK safe-edit proof gates failed: \${installedExampleStdout}\`)
}
if (installedExampleProof.proofBundle?.whatChanged?.[0]?.ref !== 'Sheet1!B2') {
	throw new Error(\`installed SDK safe-edit proof did not explain the changed cell: \${installedExampleStdout}\`)
}

console.log(JSON.stringify({
	openPlanMode: openPlan.recommendedLoadOptions.mode,
	sheets: info.sheets.map((sheet) => sheet.name),
	planWouldSucceed: plan.preview.wouldSucceed,
	commitOutputSha256: commit.outputSha256,
	reopenedValid: check.valid,
	preparedPlanDigest: prepared.planDigest,
	preparedOutputSha256: preparedCommit.outputSha256,
	preparedReopenedValid: preparedCheck.valid,
	docHits: docHits.length,
	rootExampleHits: rootExampleHits.length,
	packageInstallExampleHits: packageInstallExampleHits.length,
	installedExampleCommand: 'node_modules/.bin/ascend-sdk-safe-edit',
	installedExampleWorkflow: installedExampleProof.workflow,
	installedExampleCell: installedExampleProof.verify?.cell,
	installedExampleProofBundle: {
		safeToUse: installedExampleProof.proofBundle?.safeToUse,
		whatChanged: installedExampleProof.proofBundle?.whatChanged,
		whySafe: installedExampleProof.proofBundle?.whySafe?.map((gate) => ({
			gate: gate.gate,
			ok: gate.ok,
		})),
	},
	apiExampleHits: apiExampleHits.length,
	mcpExampleHits: mcpExampleHits.length,
	b1,
	c1,
	preparedB1,
	preparedC1,
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
