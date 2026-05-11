import { createHash } from 'node:crypto'
import { readdir, readFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { readXlsx } from '../../../packages/io-xlsx/src/reader/index.ts'
import type { CorpusManifestEntry } from '../../corpus/manifest.ts'
import { inspectOoxmlPackageFeatures } from '../../corpus/ooxml-feature-probe.ts'

const APACHE_SOURCE = 'Apache POI spreadsheet test-data'
const APACHE_BASE_URL =
	'https://raw.githubusercontent.com/apache/poi/refs/heads/trunk/test-data/spreadsheet'
export const SHEETJS_FIXTURE_FILES = [
	'AutoFilter.xlsx',
	'formula_stress_test.xlsx',
	'merge_cells.xlsx',
	'named_ranges_2011.xlsx',
] as const

const SHEETJS_FILES = new Set<string>(SHEETJS_FIXTURE_FILES)

export async function loadManifest(): Promise<CorpusManifestEntry[]> {
	const root = dirname(fileURLToPath(import.meta.url))
	const files = (await readdir(root))
		.filter((file) => file.endsWith('.xlsx'))
		.filter((file) => !SHEETJS_FILES.has(file))
		.sort((a, b) => a.localeCompare(b))
	const entries: CorpusManifestEntry[] = []
	for (const file of files) {
		const entry = await buildPoiEntry(root, file)
		if (entry) entries.push(entry)
	}
	return entries
}

export async function buildPoiEntry(
	root: string,
	file: string,
): Promise<CorpusManifestEntry | null> {
	const bytes = new Uint8Array(await readFile(join(root, file)))
	const password = fixturePassword(file)
	const packageBytes = password ? decryptFixture(bytes, password, file) : bytes
	if (!isZipPackage(packageBytes)) return null
	const probe = inspectOoxmlPackageFeatures(packageBytes)
	const counts = {
		worksheets: probe.counts.worksheets,
		formulas: probe.counts.formulas,
		charts: probe.counts.charts,
		tables: probe.counts.tables,
		drawings: probe.counts.drawings,
		pivot_tables: probe.counts.pivot_tables,
		pivot_caches: probe.counts.pivot_caches,
		comments: probe.counts.comments,
		workbook_protection: probe.counts.workbook_protection,
		sheet_protection: probe.counts.sheet_protection,
	}
	const features = { ...probe.features, macros: false, encrypted: isCompoundFile(bytes) }
	return {
		file: basename(file),
		size_bytes: bytes.byteLength,
		features,
		counts,
		source: APACHE_SOURCE,
		sourceUrl: `${APACHE_BASE_URL}/${file}`,
		license: 'Apache-2.0',
		sha256: createHash('sha256').update(bytes).digest('hex'),
		...(password !== undefined ? { password } : {}),
		redistributionAllowed: true,
		citation: 'Apache POI test-data/spreadsheet XLSX fixture subset, Apache-2.0.',
		vendorable: true,
		benchmarkTier: deriveTier(features),
		assertionClass: deriveAssertionClass(features),
		riskClass: deriveRisk(features),
		featureTags: deriveTags(file, 'apache-poi', features),
		...(password ? { notes: 'Encrypted OOXML package; POI fixture password is tika.' } : {}),
	}
}

function isZipPackage(bytes: Uint8Array): boolean {
	return bytes[0] === 0x50 && bytes[1] === 0x4b
}

function isCompoundFile(bytes: Uint8Array): boolean {
	const magic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
	return magic.every((byte, index) => bytes[index] === byte)
}

function fixturePassword(file: string): string | undefined {
	return file === 'protected_passtika.xlsx' ? 'tika' : undefined
}

function decryptFixture(bytes: Uint8Array, password: string, file: string): Uint8Array {
	const read = readXlsx(bytes, { password })
	if (!read.ok) throw new Error(`${file}: ${read.error.message}`)
	return read.value.workbook.sourceArchiveBytes ?? bytes
}

function deriveTier(
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['benchmarkTier'] {
	return features.charts ||
		features.tables ||
		features.drawings ||
		features.conditional_formatting ||
		features.data_validations ||
		features.protection ||
		features.calc_chain
		? 'core'
		: 'smoke'
}

function deriveAssertionClass(
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['assertionClass'] {
	if (features.charts || features.drawings) return 'preservation-only'
	if (
		features.conditional_formatting ||
		features.data_validations ||
		features.tables ||
		features.protection
	) {
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
		features.tables ||
		features.protection
		? 'medium'
		: 'low'
}

function deriveTags(
	file: string,
	sourceTag: string,
	features: CorpusManifestEntry['features'],
): string[] {
	const tags = new Set<string>([sourceTag, 'small'])
	if (features.encrypted) tags.add('encrypted')
	if (features.charts) tags.add('chart')
	if (features.tables) tags.add('table')
	if (features.drawings) tags.add('drawing')
	if (features.comments) tags.add('comment')
	if (features.conditional_formatting) tags.add('conditional-formatting')
	if (features.data_validations) tags.add('data-validation')
	if (features.merged_cells) tags.add('merged-cells')
	if (features.defined_names) tags.add('defined-names')
	if (features.calc_chain || /formula/i.test(file)) tags.add('formula-fidelity')
	if (features.protection) tags.add('protection')
	if (/formula/i.test(file)) tags.add('formula')
	if (/style|format|theme|colour|color/i.test(file)) tags.add('style')
	return [...tags].sort((a, b) => a.localeCompare(b))
}
