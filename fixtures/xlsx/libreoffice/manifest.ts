import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CorpusManifestEntry } from '../../corpus/manifest.ts'
import { inspectOoxmlPackageFeatures } from '../../corpus/ooxml-feature-probe.ts'

interface LibreOfficeFixture {
	readonly file: string
	readonly sourcePath: string
	readonly notes?: string
}

const LIBREOFFICE_SOURCE = 'LibreOffice Calc QA XLSX regression data'
const LIBREOFFICE_BASE_URL = 'https://raw.githubusercontent.com/LibreOffice/core/master'
const LIBREOFFICE_LICENSE = 'MPL-2.0 OR LGPL-3.0-or-later OR GPL-3.0-or-later'

const FIXTURES: readonly LibreOfficeFixture[] = [
	{ file: '129969-min.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/129969-min.xlsx' },
	{ file: 'CalcThemeTest.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/CalcThemeTest.xlsx' },
	{
		file: 'MissingPathExternal.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/MissingPathExternal.xlsx',
		notes: 'External-link regression workbook; link targets are preserved but not fetched.',
	},
	{
		file: 'PivotTable_CachedDefinitionAndDataInSync.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/PivotTable_CachedDefinitionAndDataInSync.xlsx',
	},
	{
		file: 'PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithCacheData.xlsx',
		sourcePath:
			'sc/qa/unit/data/xlsx/PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithCacheData.xlsx',
	},
	{
		file: 'PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithoutCacheData.xlsx',
		sourcePath:
			'sc/qa/unit/data/xlsx/PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithoutCacheData.xlsx',
	},
	{
		file: 'ProtecteSheet1234Pass.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/ProtecteSheet1234Pass.xlsx',
	},
	{ file: 'Sparklines.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/Sparklines.xlsx' },
	{ file: 'TableEmptyHeaders.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/TableEmptyHeaders.xlsx' },
	{ file: 'TableStyleTest.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/TableStyleTest.xlsx' },
	{
		file: 'Test_ThemeColor_Text_Background_Border.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/Test_ThemeColor_Text_Background_Border.xlsx',
	},
	{ file: 'activex_checkbox.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/activex_checkbox.xlsx' },
	{ file: 'autofilter-colors.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/autofilter-colors.xlsx' },
	{ file: 'autofilter.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/autofilter.xlsx' },
	{ file: 'colorscale.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/colorscale.xlsx' },
	{ file: 'complex_icon_set.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/complex_icon_set.xlsx' },
	{ file: 'condFormat_cellis.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/condFormat_cellis.xlsx' },
	{ file: 'databar.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/databar.xlsx' },
	{
		file: 'functions-excel-2010.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/functions-excel-2010.xlsx',
	},
	{
		file: 'matrix-multiplication.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/matrix-multiplication.xlsx',
	},
	{
		file: 'pivot_table_first_header_row.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/pivot_table_first_header_row.xlsx',
	},
	{
		file: 'pivottable_date_field_filter.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/pivottable_date_field_filter.xlsx',
	},
	{
		file: 'pivot-table/tdf126858-1.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/pivot-table/tdf126858-1.xlsx',
		notes:
			'Calculated pivot field regression: the calculated data field is the only visible data dimension.',
	},
	{
		file: 'pivot-table/test_diff_aggregation.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/pivot-table/test_diff_aggregation.xlsx',
		notes:
			'Calculated pivot field regression: calculated fields use SUM aggregation even beside visible COUNT data fields.',
	},
	{ file: 'sortconditionref2.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/sortconditionref2.xlsx' },
	{
		file: 'tdf143068_top10filter.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/tdf143068_top10filter.xlsx',
	},
	{
		file: 'tdf165180_date1904.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/tdf165180_date1904.xlsx',
	},
	{
		file: 'tdf167689_tableType.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/tdf167689_tableType.xlsx',
	},
	{ file: 'tdf170201.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/tdf170201.xlsx' },
	{
		file: 'textLengthDataValidity.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/textLengthDataValidity.xlsx',
	},
	{ file: 'textbox-hyperlink.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/textbox-hyperlink.xlsx' },
	{ file: 'totalsRowFunction.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/totalsRowFunction.xlsx' },
	{ file: 'totalsRowShown.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/totalsRowShown.xlsx' },
	{
		file: 'universal-content-strict.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/universal-content-strict.xlsx',
	},
	{ file: 'universal-content.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/universal-content.xlsx' },
	{
		file: 'user_defined_function.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/user_defined_function.xlsx',
		notes:
			'Contains an external/user-defined formula entry; useful for parser and preservation coverage.',
	},
	{
		file: 'value-in-column-2000.xlsx',
		sourcePath: 'sc/qa/unit/data/xlsx/value-in-column-2000.xlsx',
	},
	{ file: 'writingMode.xlsx', sourcePath: 'sc/qa/unit/data/xlsx/writingMode.xlsx' },
]

export async function loadManifest(): Promise<CorpusManifestEntry[]> {
	const root = dirname(fileURLToPath(import.meta.url))
	const entries: CorpusManifestEntry[] = []
	for (const fixture of FIXTURES) entries.push(await buildEntry(root, fixture))
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
		sourceUrl: `${LIBREOFFICE_BASE_URL}/${fixture.sourcePath}`,
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
