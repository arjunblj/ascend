import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CorpusManifestEntry } from '../../corpus/manifest.ts'
import { inspectOoxmlPackageFeatures } from '../../corpus/ooxml-feature-probe.ts'

interface ClosedXmlFixture {
	readonly file: string
	readonly sourcePath: string
}

const CLOSEDXML_SOURCE = 'ClosedXML test resources'
const CLOSEDXML_BASE_URL = 'https://raw.githubusercontent.com/ClosedXML/ClosedXML/develop'
const CACHED_FORMULA_FIXTURES = new Set([
	'Other_ExternalLinks_WorkbookWithExternalLink.xlsx',
	'Other_Formulas_BooleanFormulaValues.xlsx',
	'Other_Formulas_DataTableFormula-Excel-Input.xlsx',
	'Misc_FormulasWithEvaluation.xlsx',
])
const FIXTURES: readonly ClosedXmlFixture[] = [
	{
		file: 'AutoFilter_CustomAutoFilter.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/AutoFilter/CustomAutoFilter.xlsx',
	},
	{
		file: 'Comments_AddingComments.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Comments/AddingComments.xlsx',
	},
	{
		file: 'ConditionalFormatting_CFDataBars.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/ConditionalFormatting/CFDataBars.xlsx',
	},
	{
		file: 'ConditionalFormatting_CFIconSet.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/ConditionalFormatting/CFIconSet.xlsx',
	},
	{
		file: 'ImageHandling_ImageAnchors.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/ImageHandling/ImageAnchors.xlsx',
	},
	{
		file: 'Misc_DataValidation.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Misc/DataValidation.xlsx',
	},
	{
		file: 'Misc_Formulas.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Misc/Formulas.xlsx',
	},
	{
		file: 'Misc_FormulasWithEvaluation.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Misc/FormulasWithEvaluation.xlsx',
	},
	{
		file: 'Misc_Hyperlinks.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Misc/Hyperlinks.xlsx',
	},
	{
		file: 'Misc_MergeCells.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Misc/MergeCells.xlsx',
	},
	{
		file: 'Misc_ShiftingFormulas.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Misc/ShiftingFormulas.xlsx',
	},
	{
		file: 'Misc_SheetProtection.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Misc/SheetProtection.xlsx',
	},
	{
		file: 'PivotTables_PivotTables.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/PivotTables/PivotTables.xlsx',
	},
	{
		file: 'Ranges_DefinedNames.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Ranges/DefinedNames.xlsx',
	},
	{
		file: 'Ranges_SortExample.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Ranges/SortExample.xlsx',
	},
	{
		file: 'Sparklines_SampleSparklines.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Sparklines/SampleSparklines.xlsx',
	},
	{
		file: 'Styles_StyleNumberFormat.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Styles/StyleNumberFormat.xlsx',
	},
	{
		file: 'Styles_UsingRichText.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Styles/UsingRichText.xlsx',
	},
	{
		file: 'Tables_UsingTables.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Tables/UsingTables.xlsx',
	},
	{
		file: 'Tables_ResizingTables.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Examples/Tables/ResizingTables.xlsx',
	},
	{
		file: 'Other_Charts_PreserveCharts_inputfile.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Other/Charts/PreserveCharts/inputfile.xlsx',
	},
	{
		file: 'Other_ExternalLinks_WorkbookWithExternalLink.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Other/ExternalLinks/WorkbookWithExternalLink.xlsx',
	},
	{
		file: 'Other_Formulas_ArrayFormula.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Other/Formulas/ArrayFormula.xlsx',
	},
	{
		file: 'Other_Formulas_BooleanFormulaValues.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Other/Formulas/BooleanFormulaValues.xlsx',
	},
	{
		file: 'Other_Formulas_DataTableFormula-Excel-Input.xlsx',
		sourcePath: 'ClosedXML.Tests/Resource/Other/Formulas/DataTableFormula-Excel-Input.xlsx',
	},
	{
		file: 'Other_PivotTableReferenceFiles_ChartsheetAndPivotTable.xlsx',
		sourcePath:
			'ClosedXML.Tests/Resource/Other/PivotTableReferenceFiles/ChartsheetAndPivotTable.xlsx',
	},
]

export async function loadManifest(): Promise<CorpusManifestEntry[]> {
	const root = dirname(fileURLToPath(import.meta.url))
	const entries: CorpusManifestEntry[] = []
	for (const fixture of FIXTURES) {
		entries.push(await buildEntry(root, fixture))
	}
	return entries
}

async function buildEntry(root: string, fixture: ClosedXmlFixture): Promise<CorpusManifestEntry> {
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
	}
	const features = { ...probe.features, macros: false }
	return {
		file: basename(fixture.file),
		size_bytes: bytes.byteLength,
		features,
		counts,
		source: CLOSEDXML_SOURCE,
		sourceUrl: `${CLOSEDXML_BASE_URL}/${fixture.sourcePath}`,
		license: 'MIT',
		sha256: createHash('sha256').update(bytes).digest('hex'),
		redistributionAllowed: true,
		citation: 'ClosedXML ClosedXML.Tests/Resource XLSX fixture subset, MIT.',
		vendorable: true,
		benchmarkTier: deriveTier(bytes.byteLength, features),
		assertionClass: deriveAssertionClass(features),
		riskClass: deriveRisk(features),
		featureTags: deriveTags(fixture.file, counts.formulas, features),
	}
}

function deriveTier(
	size: number,
	features: CorpusManifestEntry['features'],
): CorpusManifestEntry['benchmarkTier'] {
	if (size >= 100_000) return 'extended'
	return features.charts ||
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
	if (features.charts || features.drawings || features.images_or_media || features.pivot_tables) {
		return 'preservation-only'
	}
	if (features.conditional_formatting || features.data_validations || features.tables) {
		return 'semantic-plus-package'
	}
	return 'exact-bytes'
}

function deriveRisk(features: CorpusManifestEntry['features']): CorpusManifestEntry['riskClass'] {
	return features.charts ||
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
	_formulaCount: number,
	features: CorpusManifestEntry['features'],
): string[] {
	const tags = new Set<string>(['closedxml', 'small'])
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
	if (CACHED_FORMULA_FIXTURES.has(file)) tags.add('formula-fidelity')
	if (/formula/i.test(file)) tags.add('formula')
	if (/style|format|richtext/i.test(file)) tags.add('style')
	if (/protect/i.test(file)) tags.add('protection')
	if (/sparkline/i.test(file)) tags.add('sparkline')
	if (/sort/i.test(file)) tags.add('sort')
	return [...tags].sort((a, b) => a.localeCompare(b))
}
