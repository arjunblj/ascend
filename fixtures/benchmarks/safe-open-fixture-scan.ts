import { execFileSync } from 'node:child_process'
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs'
import { basename, join } from 'node:path'
import { inspectWorkbookOpenPlan } from '../../packages/sdk/src/index.ts'

export interface SafeOpenFixtureScanOptions {
	readonly root?: string
}

export interface SafeOpenFixtureScanMatch {
	readonly fixture: string
	readonly riskFamilies: readonly string[]
	readonly recommendedMode: string
	readonly reviewBeforeHydration: boolean
	readonly partCount: number
}

export interface SafeOpenFixtureScanResult {
	readonly generatedAt: string
	readonly root: string
	readonly corpus: 'tracked-git-fixtures' | 'filesystem-fixtures'
	readonly skippedDirectories: readonly string[]
	readonly scanned: number
	readonly rejected: number
	readonly rejectedFixtures: readonly string[]
	readonly signatureOrUnknownMatches: readonly SafeOpenFixtureScanMatch[]
	readonly replacementStatus: 'no-public-binary-replacement-found' | 'candidate-found'
	readonly boundary: string
}

const REPLACEMENT_RISK_FAMILIES = new Set(['preservedSignature', 'preservedOther'])
const SKIPPED_DIRECTORIES = new Set(['external', 'stress'])

export function runSafeOpenFixtureScan(
	options: SafeOpenFixtureScanOptions = {},
): SafeOpenFixtureScanResult {
	const root = options.root ?? 'fixtures/xlsx'
	if (!existsSync(root)) throw new Error(`Missing fixture scan root ${root}`)
	const listed = listFixtureFiles(root)
	const matches: SafeOpenFixtureScanMatch[] = []
	const rejectedFixtures: string[] = []
	const files = listed.files
	for (const fixture of files) {
		try {
			const plan = inspectWorkbookOpenPlan(readFileSync(fixture), { intent: 'edit-plan' })
			const riskFamilies = Array.from(
				new Set(plan.riskFeatures.map((feature) => feature.featureFamily)),
			).sort()
			if (riskFamilies.some((family) => REPLACEMENT_RISK_FAMILIES.has(family))) {
				matches.push({
					fixture,
					riskFamilies,
					recommendedMode: plan.recommendedMode,
					reviewBeforeHydration: plan.reviewBeforeHydration,
					partCount: plan.partCount,
				})
			}
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
		signatureOrUnknownMatches: matches,
		replacementStatus:
			matches.length === 0 ? 'no-public-binary-replacement-found' : 'candidate-found',
		boundary:
			'This scans the tracked public XLSX/XLSM fixture corpus when git metadata is available and skips ignored external/stress fixture folders for filesystem fallback. It does not prove that no public signed or unknown-part workbooks exist elsewhere.',
	}
}

export function safeOpenFixtureScanMarkdown(result: SafeOpenFixtureScanResult): string {
	return [
		'# Safe Open Fixture Replacement Scan',
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
		'| Fixture | Risk families | Mode | Review before hydration | Parts |',
		'| --- | --- | --- | --- | ---: |',
		...(result.signatureOrUnknownMatches.length > 0
			? result.signatureOrUnknownMatches.map(
					(match) =>
						`| \`${match.fixture}\` | ${match.riskFamilies.join(', ')} | ${match.recommendedMode} | ${match.reviewBeforeHydration} | ${match.partCount} |`,
				)
			: ['| none | none | n/a | n/a | n/a |']),
		'',
		'Rejected fixtures:',
		...(result.rejectedFixtures.length > 0
			? result.rejectedFixtures.map((fixture) => `- \`${fixture}\``)
			: ['- none']),
	].join('\n')
}

function listFixtureFiles(root: string): {
	readonly corpus: SafeOpenFixtureScanResult['corpus']
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
	const result = runSafeOpenFixtureScan({ root })
	console.log(json ? JSON.stringify(result, null, 2) : safeOpenFixtureScanMarkdown(result))
	if (!json) {
		console.error(`Scanned ${result.scanned} fixtures for safe-open replacements.`)
		console.error(`Run with --json for machine-readable output from ${basename(import.meta.path)}.`)
	}
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	if (index === -1) return undefined
	return process.argv[index + 1]
}
