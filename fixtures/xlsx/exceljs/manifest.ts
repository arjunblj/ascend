import { createHash } from 'node:crypto'
import { readdir, readFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CorpusManifestEntry } from '../../corpus/manifest.ts'
import { inspectOoxmlPackageFeatures } from '../../corpus/ooxml-feature-probe.ts'

const EXCELJS_SOURCE = 'ExcelJS integration test data'
const EXCELJS_BASE_URL =
	'https://raw.githubusercontent.com/exceljs/exceljs/master/spec/integration/data'

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
		formulas: probe.counts.formulas,
		charts: probe.counts.charts,
		tables: probe.counts.tables,
		drawings: probe.counts.drawings,
		pivot_tables: probe.counts.pivot_tables,
		pivot_caches: probe.counts.pivot_caches,
		comments: probe.counts.comments,
	}
	const features = { ...probe.features, macros: false }
	return {
		file: basename(file),
		size_bytes: bytes.byteLength,
		features,
		counts,
		source: EXCELJS_SOURCE,
		sourceUrl: `${EXCELJS_BASE_URL}/${file}`,
		license: 'MIT',
		sha256: createHash('sha256').update(bytes).digest('hex'),
		redistributionAllowed: true,
		citation: 'ExcelJS spec/integration/data XLSX fixture subset, MIT.',
		vendorable: true,
		benchmarkTier: deriveTier(bytes.byteLength, features),
		assertionClass: deriveAssertionClass(features),
		riskClass: deriveRisk(features),
		featureTags: deriveTags(file, features),
	}
}

function deriveTier(
	size: number,
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['benchmarkTier'] {
	if (size >= 250_000) return 'extended'
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
	if (features.charts || features.drawings || features.images_or_media) return 'preservation-only'
	if (features.conditional_formatting || features.data_validations || features.tables) {
		return 'semantic-plus-package'
	}
	return 'exact-bytes'
}

function deriveRisk(features: CorpusManifestEntry['features']): CorpusManifestEntry['riskClass'] {
	return features.charts ||
		features.drawings ||
		features.images_or_media ||
		features.conditional_formatting ||
		features.data_validations ||
		features.calc_chain ||
		features.tables
		? 'medium'
		: 'low'
}

function deriveTags(file: string, features: CorpusManifestEntry['features']): string[] {
	const tags = new Set<string>(['exceljs', 'small'])
	if (features.charts) tags.add('chart')
	if (features.tables) tags.add('table')
	if (features.drawings) tags.add('drawing')
	if (features.images_or_media) tags.add('media')
	if (features.comments) tags.add('comment')
	if (features.conditional_formatting) tags.add('conditional-formatting')
	if (features.data_validations) tags.add('data-validation')
	if (features.merged_cells) tags.add('merged-cells')
	if (features.defined_names) tags.add('defined-names')
	if (features.calc_chain || /formula|fibonacci/i.test(file)) tags.add('formula-fidelity')
	if (/formula|fibonacci/i.test(file)) tags.add('formula')
	if (/style|format|theme|colour|color|row-styles/i.test(file)) tags.add('style')
	if (/1904|date/i.test(file)) tags.add('date')
	if (/issue|pr-/i.test(file)) tags.add('issue-regression')
	return [...tags].sort((a, b) => a.localeCompare(b))
}
