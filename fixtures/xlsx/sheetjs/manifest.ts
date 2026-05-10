import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CorpusManifestEntry } from '../../corpus/manifest.ts'
import { inspectOoxmlPackageFeatures } from '../../corpus/ooxml-feature-probe.ts'
import { SHEETJS_FIXTURE_FILES } from '../poi/manifest.ts'

const SHEETJS_SOURCE = 'SheetJS test files'
const SHEETJS_BASE_URL = 'https://oss.sheetjs.com/test_files'

export async function loadManifest(): Promise<CorpusManifestEntry[]> {
	const root = dirname(fileURLToPath(import.meta.url))
	const poiRoot = join(root, '../poi')
	const entries: CorpusManifestEntry[] = []
	for (const file of SHEETJS_FIXTURE_FILES) {
		const entry = await buildEntry(poiRoot, file)
		if (entry) entries.push(entry)
	}
	return entries
}

async function buildEntry(root: string, file: string): Promise<CorpusManifestEntry | null> {
	const bytes = await readFixture(root, file)
	if (!bytes) return null
	const probe = inspectOoxmlPackageFeatures(bytes)
	const features = { ...probe.features, macros: false }
	return {
		file: `../poi/${file}`,
		size_bytes: bytes.byteLength,
		features,
		counts: {
			worksheets: probe.counts.worksheets,
			charts: probe.counts.charts,
			tables: probe.counts.tables,
			drawings: probe.counts.drawings,
			pivot_tables: probe.counts.pivot_tables,
			pivot_caches: probe.counts.pivot_caches,
			comments: probe.counts.comments,
		},
		source: SHEETJS_SOURCE,
		sourceUrl: `${SHEETJS_BASE_URL}/${file}`,
		license: 'Apache-2.0',
		sha256: createHash('sha256').update(bytes).digest('hex'),
		redistributionAllowed: true,
		citation: 'SheetJS test_files XLSX fixture subset, Apache-2.0.',
		vendorable: true,
		benchmarkTier: deriveTier(features),
		assertionClass: deriveAssertionClass(features),
		riskClass: deriveRisk(features),
		featureTags: deriveTags(file, features),
	}
}

async function readFixture(root: string, file: string): Promise<Uint8Array | null> {
	try {
		return new Uint8Array(await readFile(join(root, file)))
	} catch (error) {
		if (typeof error === 'object' && error !== null && 'code' in error && error.code === 'ENOENT') {
			return null
		}
		throw error
	}
}

function deriveTier(
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['benchmarkTier'] {
	return features.tables ||
		features.conditional_formatting ||
		features.data_validations ||
		features.calc_chain
		? 'core'
		: 'smoke'
}

function deriveAssertionClass(
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['assertionClass'] {
	if (features.conditional_formatting || features.data_validations || features.tables) {
		return 'semantic-plus-package'
	}
	return 'exact-bytes'
}

function deriveRisk(features: CorpusManifestEntry['features']): CorpusManifestEntry['riskClass'] {
	return features.conditional_formatting ||
		features.data_validations ||
		features.calc_chain ||
		features.tables
		? 'medium'
		: 'low'
}

function deriveTags(file: string, features: CorpusManifestEntry['features']): string[] {
	const tags = new Set<string>(['sheetjs', 'small'])
	if (features.tables) tags.add('table')
	if (features.conditional_formatting) tags.add('conditional-formatting')
	if (features.data_validations) tags.add('data-validation')
	if (features.merged_cells) tags.add('merged-cells')
	if (features.defined_names) tags.add('defined-names')
	if (features.calc_chain || /formula/i.test(file)) tags.add('formula-fidelity')
	if (/formula/i.test(file)) tags.add('formula')
	return [...tags].sort((a, b) => a.localeCompare(b))
}
