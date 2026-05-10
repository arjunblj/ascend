import { createHash } from 'node:crypto'
import { readdir, readFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CorpusManifestEntry } from '../../corpus/manifest.ts'
import { inspectOoxmlPackageFeatures } from '../../corpus/ooxml-feature-probe.ts'

const APACHE_SOURCE = 'Apache POI spreadsheet test-data'
const APACHE_BASE_URL =
	'https://raw.githubusercontent.com/apache/poi/refs/heads/trunk/test-data/spreadsheet'
const SHEETJS_SOURCE = 'SheetJS test files'
const SHEETJS_BASE_URL = 'https://oss.sheetjs.com/test_files'
const SHEETJS_FILES = new Set([
	'AutoFilter.xlsx',
	'formula_stress_test.xlsx',
	'merge_cells.xlsx',
	'named_ranges_2011.xlsx',
])

export async function loadManifest(): Promise<CorpusManifestEntry[]> {
	const root = dirname(fileURLToPath(import.meta.url))
	const files = (await readdir(root))
		.filter((file) => file.endsWith('.xlsx'))
		.sort((a, b) => a.localeCompare(b))
	const entries: CorpusManifestEntry[] = []
	for (const file of files) {
		entries.push(await buildEntry(root, file))
	}
	return entries
}

async function buildEntry(root: string, file: string): Promise<CorpusManifestEntry> {
	const bytes = new Uint8Array(await readFile(join(root, file)))
	const probe = inspectOoxmlPackageFeatures(bytes)
	const counts = {
		worksheets: probe.counts.worksheets,
		charts: probe.counts.charts,
		tables: probe.counts.tables,
		drawings: probe.counts.drawings,
		pivot_tables: probe.counts.pivot_tables,
		pivot_caches: probe.counts.pivot_caches,
		comments: probe.counts.comments,
	}
	const features = { ...probe.features, macros: false }
	const source = SHEETJS_FILES.has(file) ? SHEETJS_SOURCE : APACHE_SOURCE
	const sourceUrl = `${SHEETJS_FILES.has(file) ? SHEETJS_BASE_URL : APACHE_BASE_URL}/${file}`
	return {
		file: basename(file),
		size_bytes: bytes.byteLength,
		features,
		counts,
		source,
		sourceUrl,
		license: 'Apache-2.0',
		sha256: createHash('sha256').update(bytes).digest('hex'),
		redistributionAllowed: true,
		citation:
			source === SHEETJS_SOURCE
				? 'SheetJS test_files XLSX fixture subset, Apache-2.0.'
				: 'Apache POI test-data/spreadsheet XLSX fixture subset, Apache-2.0.',
		vendorable: true,
		benchmarkTier: deriveTier(features),
		assertionClass: deriveAssertionClass(features),
		riskClass: deriveRisk(features),
		featureTags: deriveTags(file, source, features),
	}
}

function deriveTier(
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['benchmarkTier'] {
	return features.charts ||
		features.tables ||
		features.drawings ||
		features.conditional_formatting ||
		features.data_validations ||
		features.calc_chain
		? 'core'
		: 'smoke'
}

function deriveAssertionClass(
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['assertionClass'] {
	if (features.charts || features.drawings) return 'preservation-only'
	if (features.conditional_formatting || features.data_validations || features.tables) {
		return 'semantic-plus-package'
	}
	return 'exact-bytes'
}

function deriveRisk(features: CorpusManifestEntry['features']): CorpusManifestEntry['riskClass'] {
	return features.charts ||
		features.drawings ||
		features.conditional_formatting ||
		features.data_validations ||
		features.calc_chain ||
		features.tables
		? 'medium'
		: 'low'
}

function deriveTags(
	file: string,
	source: string,
	features: CorpusManifestEntry['features'],
): string[] {
	const tags = new Set<string>([source === SHEETJS_SOURCE ? 'sheetjs' : 'apache-poi', 'small'])
	if (features.charts) tags.add('chart')
	if (features.tables) tags.add('table')
	if (features.drawings) tags.add('drawing')
	if (features.comments) tags.add('comment')
	if (features.conditional_formatting) tags.add('conditional-formatting')
	if (features.data_validations) tags.add('data-validation')
	if (features.merged_cells) tags.add('merged-cells')
	if (features.defined_names) tags.add('defined-names')
	if (features.calc_chain || /formula/i.test(file)) tags.add('formula-fidelity')
	if (/formula/i.test(file)) tags.add('formula')
	if (/style|format|theme|colour|color/i.test(file)) tags.add('style')
	if (/protect/i.test(file)) tags.add('protection')
	return [...tags].sort((a, b) => a.localeCompare(b))
}
