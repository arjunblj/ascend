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
		`import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { createApiFetch, createServer as createApiServer } from '@ascend/api'
import { createServer as createMcpServer } from '@ascend/mcp'

if (typeof createApiFetch !== 'function') throw new Error('missing createApiFetch export')
if (typeof createApiServer !== 'function') throw new Error('missing createApiServer export')
if (typeof createMcpServer !== 'function') throw new Error('missing createMcpServer export')

const cwd = process.cwd()
const ascendBin = join(cwd, 'node_modules', '.bin', 'ascend')
const setupOps = [
	{
		op: 'setCells',
		sheet: 'Sheet1',
		updates: [
			{ ref: 'A1', value: 'Revenue' },
			{ ref: 'B1', value: 100 },
		],
	},
	{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: '=B1*2' },
]
const planOps = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 125 }] }]
const setupOpsPath = join(cwd, 'setup-ops.json')
const planOpsPath = join(cwd, 'plan-ops.json')
await writeFile(setupOpsPath, JSON.stringify(setupOps))
await writeFile(planOpsPath, JSON.stringify(planOps))

async function runCli(args) {
	const proc = Bun.spawn([ascendBin, ...args], {
		cwd,
		stdout: 'pipe',
		stderr: 'pipe',
	})
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
	const stdout = await runCli([...args, '--json'])
	const parsed = JSON.parse(stdout)
	if (!parsed.ok) throw new Error('ascend ' + args.join(' ') + ' failed: ' + JSON.stringify(parsed.error))
	return parsed.data
}

async function apiJson(apiFetch, path, body, method = 'POST') {
	const response = await apiFetch(
		new Request('http://ascend.local' + path, {
			method,
			headers: { 'Content-Type': 'application/json' },
			...(body === undefined ? {} : { body: JSON.stringify(body) }),
		}),
	)
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

function numberAt(readResult, ref) {
	const cell = readResult.cells?.find((entry) => entry.ref === ref)
	const value = cell?.value
	if (typeof value === 'number') return value
	if (value?.kind === 'number') return value.value
	throw new Error('missing numeric cell ' + ref + ': ' + JSON.stringify(cell))
}

function assertWorkbookResult(label, readResult, checkResult) {
	if (checkResult.valid !== true) throw new Error(label + ' check failed: ' + JSON.stringify(checkResult))
	const b1 = numberAt(readResult, 'B1')
	const c1 = numberAt(readResult, 'C1')
	if (b1 !== 125 || c1 !== 250) {
		throw new Error(label + ' reopened values failed: ' + JSON.stringify({ b1, c1 }))
	}
	return { b1, c1 }
}

const cliInput = join(cwd, 'cli-input.xlsx')
const cliOutput = join(cwd, 'cli-output.xlsx')
await runCliJson(['create', cliInput])
await runCliJson(['write', cliInput, '--ops', setupOpsPath])
const cliOpenPlan = await runCliJson(['open-plan', cliInput])
if (cliOpenPlan.recommendedLoadOptions?.mode !== 'full') {
	throw new Error('CLI open-plan returned unexpected load options: ' + JSON.stringify(cliOpenPlan))
}
const cliInspect = await runCliJson(['inspect', cliInput])
const cliPlan = await runCliJson(['plan', cliInput, '--ops', planOpsPath])
if (cliPlan.preview?.wouldSucceed !== true) {
	throw new Error('CLI plan failed: ' + JSON.stringify(cliPlan.preview?.errors))
}
const cliCommit = await runCliJson([
	'commit',
	cliInput,
	'--ops',
	planOpsPath,
	'--output',
	cliOutput,
	'--expect-sha256',
	cliPlan.inputSha256,
	'--compact',
	'--proof',
])
if (cliCommit.proofBundle?.safeToUse !== true) {
	throw new Error('CLI commit proof bundle did not pass: ' + JSON.stringify(cliCommit))
}
if (!cliCommit.proofBundle?.whatChanged?.some((cell) => cell.ref === 'B1')) {
	throw new Error('CLI commit proof bundle did not explain changed B1: ' + JSON.stringify(cliCommit))
}
const cliCheck = await runCliJson(['check', cliOutput])
const cliRead = await runCliJson(['read', cliOutput, 'Sheet1!B1:C1'])
const cliValues = assertWorkbookResult('CLI', cliRead, cliCheck)
const cliDocs = await runCliJson(['docs', 'plan commit'])
if (!Array.isArray(cliDocs.results) || cliDocs.results.length === 0) {
	throw new Error('CLI installed docs search returned no hits')
}
const cliAgentInit = await runCliJson(['agent-init'])
if (cliAgentInit.apiEndpoints?.workflow !== 'GET /agent-workflow') {
	throw new Error('CLI agent-init missing API workflow endpoint: ' + JSON.stringify(cliAgentInit))
}
if (cliAgentInit.mcpTools?.workflow !== 'ascend.agent_workflow') {
	throw new Error('CLI agent-init missing MCP workflow tool: ' + JSON.stringify(cliAgentInit))
}
if (
	cliAgentInit.examples?.installedCliSafeEdit !==
	'ascend example-safe-edit <file.xlsx> <out.xlsx>'
) {
	throw new Error('CLI agent-init missing installed CLI safe-edit example: ' + JSON.stringify(cliAgentInit))
}
if (
	cliAgentInit.examples?.installedSdkSafeEdit !==
	'node_modules/.bin/ascend-sdk-safe-edit <file.xlsx> <out.xlsx>'
) {
	throw new Error('CLI agent-init missing installed SDK safe-edit example: ' + JSON.stringify(cliAgentInit))
}
if (!cliAgentInit.packageInstallExampleContext?.proofOutput?.includes('proofBundle.safeToUse')) {
	throw new Error('CLI agent-init missing installed SDK proof output context: ' + JSON.stringify(cliAgentInit))
}
if (!cliAgentInit.packageInstallExampleContext?.proofOutput?.includes('postWrite.dataConnections')) {
	throw new Error(
		'CLI agent-init missing installed connection proof output context: ' + JSON.stringify(cliAgentInit),
	)
}
if (!cliAgentInit.packageInstallExampleContext?.proofOutput?.includes('postWrite.formulaState')) {
	throw new Error(
		'CLI agent-init missing installed formula proof output context: ' + JSON.stringify(cliAgentInit),
	)
}
if (!cliAgentInit.packageInstallExampleContext?.proofOutput?.includes('postWrite.security')) {
	throw new Error(
		'CLI agent-init missing installed security proof output context: ' +
			JSON.stringify(cliAgentInit),
	)
}
if (!cliAgentInit.packageInstallExampleContext?.proofOutput?.includes('postWrite.visuals')) {
	throw new Error(
		'CLI agent-init missing installed visual proof output context: ' + JSON.stringify(cliAgentInit),
	)
}
if (
	!cliAgentInit.packageInstallExampleContext?.requires?.includes(
		'@ascend/cli installed for ascend example-safe-edit',
	)
) {
	throw new Error('CLI agent-init missing installed CLI safe-edit requirement: ' + JSON.stringify(cliAgentInit))
}
if (cliAgentInit.examples?.sdkSafeEdit !== 'bun run example:safe-edit <file.xlsx> <out.xlsx>') {
	throw new Error('CLI agent-init missing SDK safe-edit example: ' + JSON.stringify(cliAgentInit))
}
if (cliAgentInit.examples?.apiSafeEdit !== 'bun run example:safe-edit:http <file.xlsx> <out.xlsx>') {
	throw new Error('CLI agent-init missing API safe-edit example: ' + JSON.stringify(cliAgentInit))
}
if (cliAgentInit.examples?.mcpSafeEdit !== 'bun run example:safe-edit:mcp <file.xlsx> <out.xlsx>') {
	throw new Error('CLI agent-init missing MCP safe-edit example: ' + JSON.stringify(cliAgentInit))
}
if (cliAgentInit.exampleContext?.proofCommand !== 'bun test examples/root-scripts.test.ts') {
	throw new Error('CLI agent-init missing runnable example proof context: ' + JSON.stringify(cliAgentInit))
}
const cliSafeEdit = await runCliJson([
	'example-safe-edit',
	join(cwd, 'cli-safe-edit-input.xlsx'),
	join(cwd, 'cli-safe-edit-output.xlsx'),
])
if (cliSafeEdit.workflow !== 'installed-cli-open-plan-trust-inspect-read-plan-commit-reopen-verify') {
	throw new Error('installed CLI safe-edit returned unexpected workflow: ' + JSON.stringify(cliSafeEdit))
}
if (cliSafeEdit.proofBundle?.safeToUse !== true) {
	throw new Error('installed CLI safe-edit proof did not pass: ' + JSON.stringify(cliSafeEdit))
}
if (!cliSafeEdit.proofBundle?.whatChanged?.some((cell) => cell.ref === 'Sheet1!B2')) {
	throw new Error('installed CLI safe-edit proof did not explain the changed cell: ' + JSON.stringify(cliSafeEdit))
}
const cliSafeEditProofGates = cliSafeEdit.proofBundle?.whySafe?.map((gate) => [
	gate.gate,
	gate.ok,
])
if (
	JSON.stringify(cliSafeEditProofGates) !==
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
	throw new Error('installed CLI safe-edit proof gates were incomplete: ' + JSON.stringify(cliSafeEdit))
}
if (
	cliSafeEdit.postWriteProof?.formulaState?.formulaCells !== 1 ||
	cliSafeEdit.postWriteProof?.security?.workbookProtected !== false ||
	cliSafeEdit.postWriteProof?.dataConnections?.verification !== 'reopened-output' ||
	cliSafeEdit.postWriteProof?.visuals?.verification !== 'reopened-output'
) {
	throw new Error(
		'installed CLI safe-edit did not expose advertised post-write proof slices: ' +
			JSON.stringify(cliSafeEdit),
	)
}

const apiFetch = createApiFetch()
const apiInput = join(cwd, 'api-input.xlsx')
const apiOutput = join(cwd, 'api-output.xlsx')
await runCliJson(['create', apiInput])
const apiOpenPlan = await apiJson(apiFetch, '/open-plan', { file: apiInput })
if (apiOpenPlan.recommendedLoadOptions?.mode !== 'full') {
	throw new Error('API open-plan returned unexpected load options: ' + JSON.stringify(apiOpenPlan))
}
await apiJson(apiFetch, '/write', { file: apiInput, ops: setupOps })
const apiTrust = await apiJson(apiFetch, '/trust-report', { file: apiInput, maxFindings: 20 })
if (apiTrust.trust === 'unsafe') {
	throw new Error('API trust preflight unexpectedly unsafe: ' + JSON.stringify(apiTrust))
}
const apiInspect = await apiJson(apiFetch, '/inspect', { file: apiInput })
const apiPlan = await apiJson(apiFetch, '/plan', { file: apiInput, ops: planOps })
if (apiPlan.preview?.wouldSucceed !== true) {
	throw new Error('API plan failed: ' + JSON.stringify(apiPlan.preview?.errors))
}
const apiCommit = await apiJson(apiFetch, '/commit', {
	file: apiInput,
	ops: planOps,
	output: apiOutput,
	expectSha256: apiPlan.inputSha256,
	compact: true,
	includeProofBundle: true,
})
if (apiCommit.proofBundle?.safeToUse !== true) {
	throw new Error('API commit proof bundle did not pass: ' + JSON.stringify(apiCommit))
}
const apiCheck = await apiJson(apiFetch, '/check', { file: apiOutput })
const apiRead = await apiJson(apiFetch, '/read', {
	file: apiOutput,
	sheet: 'Sheet1',
	range: 'B1:C1',
})
const apiValues = assertWorkbookResult('API', apiRead, apiCheck)

const capabilitiesResponse = await apiFetch(new Request('http://ascend.local/capabilities'))
const capabilities = await capabilitiesResponse.json()
if (!capabilities.ok) throw new Error('installed API capabilities request failed')
if (!Array.isArray(capabilities.data?.capabilities) || capabilities.data.capabilities.length === 0) {
	throw new Error('installed API capabilities request returned no capabilities')
}
const apiWorkflow = await apiJson(apiFetch, '/agent-workflow', undefined, 'GET')
if (apiWorkflow.endpoints?.plan !== 'POST /plan' || apiWorkflow.endpoints?.commit !== 'POST /commit') {
	throw new Error('installed API agent workflow contract missing plan/commit endpoints')
}
if (!apiWorkflow.workflow?.some((step) => step.step === 'reopen-verify')) {
	throw new Error('installed API agent workflow contract missing reopen-verify step')
}
if (!apiWorkflow.workflow?.some((step) => step.step === 'trust-preflight')) {
	throw new Error('installed API agent workflow contract missing trust-preflight step')
}
if (apiWorkflow.examples?.apiSafeEdit !== 'bun run example:safe-edit:http <file.xlsx> <out.xlsx>') {
	throw new Error('installed API agent workflow contract missing API safe-edit example')
}
if (
	apiWorkflow.examples?.installedCliSafeEdit !==
	'ascend example-safe-edit <file.xlsx> <out.xlsx>'
) {
	throw new Error('installed API agent workflow contract missing installed CLI safe-edit example')
}
if (
	apiWorkflow.examples?.installedSdkSafeEdit !==
	'node_modules/.bin/ascend-sdk-safe-edit <file.xlsx> <out.xlsx>'
) {
	throw new Error('installed API agent workflow contract missing installed SDK safe-edit example')
}
if (!apiWorkflow.packageInstallExampleContext?.proofOutput?.includes('proofBundle.whatChanged')) {
	throw new Error('installed API agent workflow contract missing installed SDK proof output context')
}
if (!apiWorkflow.packageInstallExampleContext?.proofOutput?.includes('postWrite.dataConnections')) {
	throw new Error('installed API agent workflow contract missing connection proof output context')
}
if (!apiWorkflow.packageInstallExampleContext?.proofOutput?.includes('postWrite.formulaState')) {
	throw new Error('installed API agent workflow contract missing formula proof output context')
}
if (!apiWorkflow.packageInstallExampleContext?.proofOutput?.includes('postWrite.security')) {
	throw new Error('installed API agent workflow contract missing security proof output context')
}
if (!apiWorkflow.packageInstallExampleContext?.proofOutput?.includes('postWrite.visuals')) {
	throw new Error('installed API agent workflow contract missing visual proof output context')
}
if (apiWorkflow.exampleContext?.workdir !== 'repository-root') {
	throw new Error('installed API agent workflow contract missing example workdir context')
}

const mcpInput = join(cwd, 'mcp-input.xlsx')
const mcpOutput = join(cwd, 'mcp-output.xlsx')
await runCliJson(['create', mcpInput])
const mcpServer = createMcpServer()
const tools = mcpServer._registeredTools
const resources = mcpServer._registeredResources
await mcpTool(tools, 'ascend.write', { file: mcpInput, ops: setupOps })
const mcpOpenPlan = await mcpTool(tools, 'ascend.open_plan', { file: mcpInput })
if (mcpOpenPlan.recommendedLoadOptions?.mode !== 'full') {
	throw new Error('MCP open_plan returned unexpected load options: ' + JSON.stringify(mcpOpenPlan))
}
const mcpTrust = await mcpTool(tools, 'ascend.trust_report', {
	file: mcpInput,
	maxFindings: 20,
})
if (mcpTrust.trust === 'unsafe') {
	throw new Error('MCP trust preflight unexpectedly unsafe: ' + JSON.stringify(mcpTrust))
}
const mcpInspect = await mcpTool(tools, 'ascend.inspect', { file: mcpInput })
const mcpPlan = await mcpTool(tools, 'ascend.plan', { file: mcpInput, ops: planOps })
if (mcpPlan.preview?.wouldSucceed !== true) {
	throw new Error('MCP plan failed: ' + JSON.stringify(mcpPlan.preview?.errors))
}
const mcpCommit = await mcpTool(tools, 'ascend.commit', {
	file: mcpInput,
	ops: planOps,
	output: mcpOutput,
	expectSha256: mcpPlan.inputSha256,
	compact: true,
	includeProofBundle: true,
})
if (mcpCommit.proofBundle?.safeToUse !== true) {
	throw new Error('MCP commit proof bundle did not pass: ' + JSON.stringify(mcpCommit))
}
const mcpCheck = await mcpTool(tools, 'ascend.check', { file: mcpOutput })
const mcpRead = await mcpTool(tools, 'ascend.read', {
	file: mcpOutput,
	sheet: 'Sheet1',
	range: 'B1:C1',
	format: 'cells',
})
const mcpValues = assertWorkbookResult('MCP', mcpRead, mcpCheck)
const mcpDocs = await mcpTool(tools, 'ascend.search_docs', { query: 'plan commit' })
if (!Array.isArray(mcpDocs.results) || mcpDocs.results.length === 0) {
	throw new Error('MCP installed docs search returned no hits')
}
const mcpCapabilities = await tools['ascend.capabilities']?.handler({})
if (!mcpCapabilities?.structuredContent?.ok) {
	throw new Error('installed MCP capabilities tool failed')
}
const mcpWorkflow = await mcpTool(tools, 'ascend.agent_workflow', {})
if (mcpWorkflow.tools?.plan !== 'ascend.plan' || mcpWorkflow.tools?.commit !== 'ascend.commit') {
	throw new Error('installed MCP agent workflow contract missing plan/commit tools')
}
if (!mcpWorkflow.workflow?.some((step) => step.step === 'reopen-verify')) {
	throw new Error('installed MCP agent workflow contract missing reopen-verify step')
}
if (!mcpWorkflow.workflow?.some((step) => step.step === 'trust-preflight')) {
	throw new Error('installed MCP agent workflow contract missing trust-preflight step')
}
if (mcpWorkflow.examples?.mcpSafeEdit !== 'bun run example:safe-edit:mcp <file.xlsx> <out.xlsx>') {
	throw new Error('installed MCP agent workflow contract missing MCP safe-edit example')
}
if (
	mcpWorkflow.examples?.installedCliSafeEdit !==
	'ascend example-safe-edit <file.xlsx> <out.xlsx>'
) {
	throw new Error('installed MCP agent workflow contract missing installed CLI safe-edit example')
}
if (
	mcpWorkflow.examples?.installedSdkSafeEdit !==
	'node_modules/.bin/ascend-sdk-safe-edit <file.xlsx> <out.xlsx>'
) {
	throw new Error('installed MCP agent workflow contract missing installed SDK safe-edit example')
}
if (!mcpWorkflow.packageInstallExampleContext?.proofOutput?.includes('proofBundle.whySafe')) {
	throw new Error('installed MCP agent workflow contract missing installed SDK proof output context')
}
if (!mcpWorkflow.packageInstallExampleContext?.proofOutput?.includes('postWrite.dataConnections')) {
	throw new Error('installed MCP agent workflow contract missing connection proof output context')
}
if (!mcpWorkflow.packageInstallExampleContext?.proofOutput?.includes('postWrite.formulaState')) {
	throw new Error('installed MCP agent workflow contract missing formula proof output context')
}
if (!mcpWorkflow.packageInstallExampleContext?.proofOutput?.includes('postWrite.security')) {
	throw new Error('installed MCP agent workflow contract missing security proof output context')
}
if (!mcpWorkflow.packageInstallExampleContext?.proofOutput?.includes('postWrite.visuals')) {
	throw new Error('installed MCP agent workflow contract missing visual proof output context')
}
if (mcpWorkflow.exampleContext?.proofCommand !== 'bun test examples/root-scripts.test.ts') {
	throw new Error('installed MCP agent workflow contract missing example proof context')
}
const mcpResource = await resources['ascend://capabilities']?.readCallback(
	new URL('ascend://capabilities'),
)
if (!mcpResource?.contents?.[0]?.text?.includes('"capabilities"')) {
	throw new Error('installed MCP capabilities resource failed')
}
const mcpWorkflowResource = await resources['ascend://agent-workflow']?.readCallback(
	new URL('ascend://agent-workflow'),
)
if (!mcpWorkflowResource?.contents?.[0]?.text?.includes('ascend.plan')) {
	throw new Error('installed MCP agent workflow resource failed')
}
if (!mcpWorkflowResource.contents[0].text.includes('root-scripts.test.ts')) {
	throw new Error('installed MCP agent workflow resource missing example proof context')
}
if (!mcpWorkflowResource.contents[0].text.includes('proofBundle.safeToUse')) {
	throw new Error('installed MCP agent workflow resource missing installed package proof output context')
}
if (!mcpWorkflowResource.contents[0].text.includes('postWrite.dataConnections')) {
	throw new Error('installed MCP agent workflow resource missing connection proof output context')
}
if (!mcpWorkflowResource.contents[0].text.includes('postWrite.formulaState')) {
	throw new Error('installed MCP agent workflow resource missing formula proof output context')
}
if (!mcpWorkflowResource.contents[0].text.includes('postWrite.security')) {
	throw new Error('installed MCP agent workflow resource missing security proof output context')
}
if (!mcpWorkflowResource.contents[0].text.includes('postWrite.visuals')) {
	throw new Error('installed MCP agent workflow resource missing visual proof output context')
}

console.log(JSON.stringify({
	cli: {
		openPlanMode: cliOpenPlan.recommendedLoadOptions.mode,
		sheets: cliInspect.sheets?.map((sheet) => sheet.name),
		planWouldSucceed: cliPlan.preview.wouldSucceed,
		outputSha256: cliCommit.outputSha256,
		values: cliValues,
		docHits: cliDocs.results.length,
		apiWorkflowEndpoint: cliAgentInit.apiEndpoints.workflow,
		mcpWorkflowTool: cliAgentInit.mcpTools.workflow,
		examples: cliAgentInit.examples,
		packageInstallExampleContext: cliAgentInit.packageInstallExampleContext,
		exampleContext: cliAgentInit.exampleContext,
		safeEdit: {
			workflow: cliSafeEdit.workflow,
			safeToUse: cliSafeEdit.proofBundle.safeToUse,
			changedCells: cliSafeEdit.proofBundle.whatChanged.map((cell) => cell.ref),
			postWriteProof: cliSafeEdit.postWriteProof,
		},
	},
	api: {
		createApiFetchExport: typeof createApiFetch,
		createServerExport: typeof createApiServer,
		openPlanMode: apiOpenPlan.recommendedLoadOptions.mode,
		trust: apiTrust.trust,
		sheets: apiInspect.sheets?.map((sheet) => sheet.name),
		planWouldSucceed: apiPlan.preview.wouldSucceed,
		outputSha256: apiCommit.outputSha256,
		values: apiValues,
		workflowSteps: apiWorkflow.workflow.length,
		examples: apiWorkflow.examples,
		packageInstallExampleContext: apiWorkflow.packageInstallExampleContext,
		exampleContext: apiWorkflow.exampleContext,
	},
	apiCapabilities: capabilities.data.capabilities.length,
	mcp: {
		createServerExport: typeof createMcpServer,
		openPlanMode: mcpOpenPlan.recommendedLoadOptions.mode,
		trust: mcpTrust.trust,
		sheets: mcpInspect.sheets?.map((sheet) => sheet.name),
		planWouldSucceed: mcpPlan.preview.wouldSucceed,
		outputSha256: mcpCommit.outputSha256,
		values: mcpValues,
		docHits: mcpDocs.results.length,
		tools: Object.keys(tools).length,
		capabilities: mcpCapabilities.structuredContent.data.capabilities.length,
		workflowSteps: mcpWorkflow.workflow.length,
		examples: mcpWorkflow.examples,
		packageInstallExampleContext: mcpWorkflow.packageInstallExampleContext,
		exampleContext: mcpWorkflow.exampleContext,
	},
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
