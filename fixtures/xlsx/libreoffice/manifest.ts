import { createHash } from 'node:crypto'
import { readdir, readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CorpusManifestEntry } from '../../corpus/manifest.ts'
import { inspectOoxmlPackageFeatures } from '../../corpus/ooxml-feature-probe.ts'

interface LibreOfficeFixture {
	readonly file: string
	readonly notes?: string
}

const LIBREOFFICE_SOURCE = 'LibreOffice Calc QA XLSX regression data'
const LIBREOFFICE_BASE_URL = 'https://raw.githubusercontent.com/LibreOffice/core/master'
const LIBREOFFICE_LICENSE = 'MPL-2.0 OR LGPL-3.0-or-later OR GPL-3.0-or-later'

const FIXTURE_NOTES = new Map<string, string>([
	[
		'MissingPathExternal.xlsx',
		'External-link regression workbook; link targets are preserved but not fetched.',
	],
	[
		'pivot-table/tdf126858-1.xlsx',
		'Calculated pivot field regression: the calculated data field is the only visible data dimension.',
	],
	[
		'pivot-table/test_diff_aggregation.xlsx',
		'Calculated pivot field regression: calculated fields use SUM aggregation even beside visible COUNT data fields.',
	],
	[
		'user_defined_function.xlsx',
		'Contains an external/user-defined formula entry; useful for parser and preservation coverage.',
	],
])

export async function loadManifest(): Promise<CorpusManifestEntry[]> {
	const root = dirname(fileURLToPath(import.meta.url))
	const entries: CorpusManifestEntry[] = []
	for (const file of await listWorkbookFiles(root)) {
		entries.push(await buildEntry(root, { file, notes: FIXTURE_NOTES.get(file) }))
	}
	return entries
}

async function buildEntry(root: string, fixture: LibreOfficeFixture): Promise<CorpusManifestEntry> {
	const bytes = new Uint8Array(await readFile(join(root, fixture.file)))
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
		active_content: probe.counts.active_content,
		sparklines: probe.counts.sparklines,
	}
	const features = {
		...probe.features,
		strict_ooxml: /strict/i.test(fixture.file),
	}
	return {
		file: fixture.file,
		size_bytes: bytes.byteLength,
		features,
		counts,
		source: LIBREOFFICE_SOURCE,
		sourceUrl: `${LIBREOFFICE_BASE_URL}/sc/qa/unit/data/xlsx/${fixture.file}`,
		license: LIBREOFFICE_LICENSE,
		sha256: createHash('sha256').update(bytes).digest('hex'),
		redistributionAllowed: true,
		citation: 'LibreOffice core sc/qa/unit/data/xlsx Calc QA fixture subset.',
		vendorable: true,
		benchmarkTier: deriveTier(bytes.byteLength, features),
		assertionClass: deriveAssertionClass(features),
		riskClass: deriveRisk(features),
		featureTags: deriveTags(fixture.file, counts.formulas, features),
		...(fixture.notes ? { notes: fixture.notes } : {}),
	}
}

async function listWorkbookFiles(root: string, dir = ''): Promise<string[]> {
	const entries = await readdir(join(root, dir), { withFileTypes: true })
	const files: string[] = []
	for (const entry of entries) {
		const path = dir ? `${dir}/${entry.name}` : entry.name
		if (entry.isDirectory()) {
			files.push(...(await listWorkbookFiles(root, path)))
		} else if (entry.isFile() && (entry.name.endsWith('.xlsx') || entry.name.endsWith('.xlsm'))) {
			files.push(path)
		}
	}
	return files.sort((a, b) => a.localeCompare(b))
}

function deriveTier(
	size: number,
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['benchmarkTier'] {
	if (size >= 100_000) return 'extended'
	return features.external_links ||
		features.pivot_tables ||
		features.tables ||
		features.drawings ||
		features.data_validations ||
		features.active_content ||
		features.sparklines ||
		features.strict_ooxml
		? 'core'
		: 'smoke'
}

function deriveAssertionClass(
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['assertionClass'] {
	if (
		features.pivot_tables ||
		features.drawings ||
		features.active_content ||
		features.external_links
	) {
		return 'preservation-only'
	}
	if (features.tables || features.data_validations || features.defined_names)
		return 'semantic-plus-package'
	return 'exact-bytes'
}

function deriveRisk(features: CorpusManifestEntry['features']): CorpusManifestEntry['riskClass'] {
	return features.external_links ||
		features.pivot_tables ||
		features.drawings ||
		features.active_content ||
		features.strict_ooxml
		? 'medium'
		: 'low'
}

function deriveTags(
	file: string,
	formulaCount: number,
	features: CorpusManifestEntry['features'],
): string[] {
	const tags = new Set<string>(['libreoffice', 'small'])
	if (features.charts) tags.add('chart')
	if (features.pivot_tables) tags.add('pivot-table')
	if (features.tables) tags.add('table')
	if (features.drawings) tags.add('drawing')
	if (features.images_or_media) tags.add('media')
	if (features.comments) tags.add('comment')
	if (features.conditional_formatting) tags.add('conditional-formatting')
	if (features.data_validations) tags.add('data-validation')
	if (features.merged_cells) tags.add('merged-cells')
	if (features.hyperlinks) tags.add('hyperlink')
	if (features.defined_names) tags.add('defined-names')
	if (features.external_links) tags.add('external-link')
	if (features.active_content) tags.add('active-content')
	if (features.sparklines) tags.add('sparkline')
	if (features.strict_ooxml) tags.add('strict-ooxml')
	if (/theme|style|color|colour|writingmode/i.test(file)) tags.add('style')
	if (/date|1904/i.test(file)) tags.add('date')
	if (/protect/i.test(file)) tags.add('protection')
	if (/autofilter/i.test(file)) tags.add('autofilter')
	if (/tdf|issue|bug|129969/i.test(file)) tags.add('issue-regression')
	if (formulaCount > 0) tags.add('formula-fidelity')
	if (formulaCount > 0) tags.add('formula')
	return [...tags].sort((a, b) => a.localeCompare(b))
}
