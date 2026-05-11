import { createHash } from 'node:crypto'
import { readdir, readFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { readXlsx } from '../../../packages/io-xlsx/src/reader/index.ts'
import type { CorpusManifestEntry } from '../../corpus/manifest.ts'
import { inspectOoxmlPackageFeatures } from '../../corpus/ooxml-feature-probe.ts'

const CALAMINE_SOURCE = 'Calamine test workbooks'
const CALAMINE_BASE_URL = 'https://raw.githubusercontent.com/tafia/calamine/master/tests'

export async function loadManifest(): Promise<CorpusManifestEntry[]> {
	const root = dirname(fileURLToPath(import.meta.url))
	const files = (await readdir(root))
		.filter((file) => file.endsWith('.xlsx') || file.endsWith('.xlsm'))
		.sort((a, b) => a.localeCompare(b))
	const entries: CorpusManifestEntry[] = []
	for (const file of files) {
		entries.push(await buildEntry(root, file))
	}
	return entries
}

async function buildEntry(root: string, file: string): Promise<CorpusManifestEntry> {
	const bytes = new Uint8Array(await readFile(join(root, file)))
	const password = fixturePassword(file)
	const packageBytes = password ? decryptFixture(bytes, password, file) : bytes
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
	}
	const features = {
		...probe.features,
		encrypted: isCompoundFile(bytes),
	}
	return {
		file: basename(file),
		size_bytes: bytes.byteLength,
		features,
		counts,
		source: CALAMINE_SOURCE,
		sourceUrl: `${CALAMINE_BASE_URL}/${file}`,
		license: 'MIT',
		sha256: createHash('sha256').update(bytes).digest('hex'),
		...(password !== undefined ? { password } : {}),
		redistributionAllowed: true,
		citation: 'Calamine tests XLSX/XLSM fixture subset, MIT.',
		vendorable: true,
		benchmarkTier: deriveTier(bytes.byteLength, features),
		assertionClass: deriveAssertionClass(features),
		riskClass: deriveRisk(features),
		featureTags: deriveTags(file, counts.formulas, features),
		...(password ? { notes: 'Encrypted OOXML package; Calamine fixture password is 123.' } : {}),
	}
}

function isCompoundFile(bytes: Uint8Array): boolean {
	const magic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
	return magic.every((byte, index) => bytes[index] === byte)
}

function fixturePassword(file: string): string | undefined {
	return file === 'pass_protected.xlsx' ? '123' : undefined
}

function decryptFixture(bytes: Uint8Array, password: string, file: string): Uint8Array {
	const read = readXlsx(bytes, { password })
	if (!read.ok) throw new Error(`${file}: ${read.error.message}`)
	return read.value.workbook.sourceArchiveBytes ?? bytes
}

function deriveTier(
	size: number,
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['benchmarkTier'] {
	if (size >= 250_000) return 'extended'
	return features.macros ||
		features.charts ||
		features.tables ||
		features.drawings ||
		features.pivot_tables ||
		features.conditional_formatting ||
		features.data_validations ||
		features.calc_chain
		? 'core'
		: 'smoke'
}

function deriveAssertionClass(
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['assertionClass'] {
	if (
		features.macros ||
		features.charts ||
		features.drawings ||
		features.images_or_media ||
		features.pivot_tables
	) {
		return 'preservation-only'
	}
	if (features.conditional_formatting || features.data_validations || features.tables) {
		return 'semantic-plus-package'
	}
	return 'exact-bytes'
}

function deriveRisk(features: CorpusManifestEntry['features']): CorpusManifestEntry['riskClass'] {
	return features.macros ||
		features.charts ||
		features.drawings ||
		features.images_or_media ||
		features.pivot_tables ||
		features.conditional_formatting ||
		features.data_validations ||
		features.calc_chain ||
		features.tables ||
		features.external_links
		? 'medium'
		: 'low'
}

function deriveTags(
	file: string,
	formulaCount: number,
	features: CorpusManifestEntry['features'],
): string[] {
	const tags = new Set<string>(['calamine', 'small'])
	if (features.encrypted) tags.add('encrypted')
	if (features.macros) tags.add('macro')
	if (features.charts) tags.add('chart')
	if (features.tables) tags.add('table')
	if (features.drawings) tags.add('drawing')
	if (features.images_or_media) tags.add('media')
	if (features.comments) tags.add('comment')
	if (features.conditional_formatting) tags.add('conditional-formatting')
	if (features.data_validations) tags.add('data-validation')
	if (features.merged_cells) tags.add('merged-cells')
	if (features.hyperlinks) tags.add('hyperlink')
	if (features.defined_names) tags.add('defined-names')
	if (features.pivot_tables) tags.add('pivot-table')
	if (features.calc_chain || formulaCount > 0 || /formula/i.test(file)) {
		tags.add('formula-fidelity')
	}
	if (formulaCount > 0 || /formula/i.test(file)) tags.add('formula')
	if (/errors/i.test(file)) tags.add('formula-error')
	if (/date|1904|iso/i.test(file)) tags.add('date')
	if (/inline|shared|string|sst|richtext|rph|encoded|x000D/i.test(file)) tags.add('string')
	if (/merge/i.test(file)) tags.add('merged-cells')
	if (/table|inventory/i.test(file)) tags.add('table')
	if (/issue/i.test(file)) tags.add('issue-regression')
	if (/protect/i.test(file)) tags.add('protection')
	return [...tags].sort((a, b) => a.localeCompare(b))
}
