import { execFileSync } from 'node:child_process'
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs'
import { basename, join } from 'node:path'
import { unzipSync } from 'fflate'

export interface PackageActionFixtureScanOptions {
	readonly root?: string
}

export interface PackageActionFixtureFeatureCounts {
	readonly docPropsCore: number
	readonly docPropsCustom: number
	readonly calcChain: number
	readonly customXml: number
	readonly macro: number
	readonly chartOrDrawing: number
	readonly signaturePackage: number
	readonly syntheticUnknownPathFamily: number
}

export interface PackageActionFixtureCandidate {
	readonly fixture: string
	readonly features: readonly PackageActionFixtureFeature[]
}

export type PackageActionFixtureFeature =
	| 'docPropsCore'
	| 'docPropsCustom'
	| 'calcChain'
	| 'customXml'
	| 'macro'
	| 'chartOrDrawing'
	| 'signaturePackage'
	| 'syntheticUnknownPathFamily'

export interface PackageActionFixtureScanResult {
	readonly generatedAt: string
	readonly root: string
	readonly corpus: 'tracked-git-fixtures' | 'filesystem-fixtures'
	readonly skippedDirectories: readonly string[]
	readonly scanned: number
	readonly rejected: number
	readonly rejectedFixtures: readonly string[]
	readonly featureCounts: PackageActionFixtureFeatureCounts
	readonly replacementStatus:
		| 'remaining-generated-edge-cases'
		| 'all-edge-cases-have-public-candidates'
	readonly candidates: readonly PackageActionFixtureCandidate[]
	readonly boundary: string
}

const FEATURES: readonly PackageActionFixtureFeature[] = [
	'docPropsCore',
	'docPropsCustom',
	'calcChain',
	'customXml',
	'macro',
	'chartOrDrawing',
	'signaturePackage',
	'syntheticUnknownPathFamily',
]

const SKIPPED_DIRECTORIES = new Set(['external', 'stress'])

export function runPackageActionFixtureScan(
	options: PackageActionFixtureScanOptions = {},
): PackageActionFixtureScanResult {
	const root = options.root ?? 'fixtures/xlsx'
	if (!existsSync(root)) throw new Error(`Missing fixture scan root ${root}`)
	const listed = listFixtureFiles(root)
	const files = listed.files
	const rejectedFixtures: string[] = []
	const candidates: PackageActionFixtureCandidate[] = []
	const featureCounts = emptyCounts()
	for (const fixture of files) {
		try {
			const features = classifyFixtureParts(Object.keys(unzipSync(readFileSync(fixture))))
			for (const feature of features) featureCounts[feature] += 1
			if (features.length > 0) candidates.push({ fixture, features })
		} catch {
			rejectedFixtures.push(fixture)
		}
	}
	return {
		generatedAt: new Date().toISOString(),
		root,
		corpus: listed.corpus,
		skippedDirectories: Array.from(SKIPPED_DIRECTORIES).sort(),
		scanned: files.length,
		rejected: rejectedFixtures.length,
		rejectedFixtures,
		featureCounts,
		replacementStatus:
			featureCounts.signaturePackage === 0 || featureCounts.syntheticUnknownPathFamily === 0
				? 'remaining-generated-edge-cases'
				: 'all-edge-cases-have-public-candidates',
		candidates,
		boundary:
			'This scans the tracked public XLSX/XLSM fixture corpus when git metadata is available and skips ignored external/stress fixture folders for filesystem fallback. It does not prove that no license-clear public signature or unknown-part workbooks exist elsewhere, and it does not authorize hiding generated fixture provenance.',
	}
}

export function packageActionFixtureScanMarkdown(result: PackageActionFixtureScanResult): string {
	return [
		'# Package Action Fixture Replacement Scan',
		'',
		`Generated: ${result.generatedAt}`,
		`Root: \`${result.root}\``,
		`Corpus: ${result.corpus}`,
		`Skipped directories: ${result.skippedDirectories.join(', ') || 'none'}`,
		`Scanned fixtures: ${result.scanned}`,
		`Rejected during scan: ${result.rejected}`,
		`Replacement status: ${result.replacementStatus}`,
		'',
		result.boundary,
		'',
		'| Feature | Fixtures |',
		'| --- | ---: |',
		...FEATURES.map((feature) => `| ${feature} | ${result.featureCounts[feature]} |`),
		'',
		'| Candidate fixture | Features |',
		'| --- | --- |',
		...result.candidates
			.slice(0, 40)
			.map((candidate) => `| \`${candidate.fixture}\` | ${candidate.features.join(', ')} |`),
		...(result.candidates.length > 40
			? [`| ... | ${result.candidates.length - 40} more candidates omitted |`]
			: []),
		'',
		'Rejected fixtures:',
		...(result.rejectedFixtures.length > 0
			? result.rejectedFixtures.map((fixture) => `- \`${fixture}\``)
			: ['- none']),
	].join('\n')
}

function classifyFixtureParts(partPaths: readonly string[]): PackageActionFixtureFeature[] {
	const parts = new Set(partPaths)
	const features: PackageActionFixtureFeature[] = []
	if (parts.has('docProps/core.xml')) features.push('docPropsCore')
	if (parts.has('docProps/custom.xml')) features.push('docPropsCustom')
	if (parts.has('xl/calcChain.xml')) features.push('calcChain')
	if (partPaths.some((path) => path.startsWith('customXml/'))) features.push('customXml')
	if (partPaths.some((path) => path.endsWith('/vbaProject.bin'))) features.push('macro')
	if (partPaths.some((path) => /^xl\/(charts|drawings)\//.test(path))) {
		features.push('chartOrDrawing')
	}
	if (
		partPaths.some(
			(path) => path.startsWith('_xmlsignatures/') || path.endsWith('/_xmlsignatures/origin.sigs'),
		)
	) {
		features.push('signaturePackage')
	}
	if (partPaths.some((path) => /^xl\/custom\/.+/.test(path) || /^custom\/.+/.test(path))) {
		features.push('syntheticUnknownPathFamily')
	}
	return features
}

function listFixtureFiles(root: string): {
	readonly corpus: PackageActionFixtureScanResult['corpus']
	readonly files: readonly string[]
} {
	try {
		const output = execFileSync('git', ['ls-files', '--', root], {
			encoding: 'utf8',
			stdio: ['ignore', 'pipe', 'ignore'],
		})
		const files = output
			.split('\n')
			.filter((path) => /\.(xlsx|xlsm)$/i.test(path))
			.sort((left, right) => left.localeCompare(right))
		if (files.length > 0) return { corpus: 'tracked-git-fixtures', files }
	} catch {
		// Fall through to filesystem scan when git metadata is unavailable.
	}
	return { corpus: 'filesystem-fixtures', files: walk(root) }
}

function emptyCounts(): Record<PackageActionFixtureFeature, number> {
	return {
		docPropsCore: 0,
		docPropsCustom: 0,
		calcChain: 0,
		customXml: 0,
		macro: 0,
		chartOrDrawing: 0,
		signaturePackage: 0,
		syntheticUnknownPathFamily: 0,
	}
}

function walk(root: string): string[] {
	const out: string[] = []
	for (const name of readdirSync(root)) {
		const path = join(root, name)
		const stat = statSync(path)
		if (stat.isDirectory()) {
			if (SKIPPED_DIRECTORIES.has(name)) continue
			out.push(...walk(path))
		} else if (/\.(xlsx|xlsm)$/i.test(name)) {
			out.push(path)
		}
	}
	return out.sort((left, right) => left.localeCompare(right))
}

if (import.meta.main) {
	const json = process.argv.includes('--json')
	const root = readFlag('--root')
	const result = runPackageActionFixtureScan({ root })
	console.log(json ? JSON.stringify(result, null, 2) : packageActionFixtureScanMarkdown(result))
	if (!json) {
		console.error(`Scanned ${result.scanned} fixtures for package-action replacements.`)
		console.error(`Run with --json for machine-readable output from ${basename(import.meta.path)}.`)
	}
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	if (index === -1) return undefined
	return process.argv[index + 1]
}
