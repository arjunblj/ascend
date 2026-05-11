/**
 * Head-to-head real workbook read and no-op roundtrip benchmarks.
 *
 * Run:
 *   bun run fixtures/benchmarks/competitive-real-workbook.ts
 *   bun run fixtures/benchmarks/competitive-real-workbook.ts --json --repeat 5 path/to/file.xlsx
 *   bun run fixtures/benchmarks/competitive-real-workbook.ts --full-corpus --category read --repeat 5
 *   bun run fixtures/benchmarks/competitive-real-workbook.ts --corpus-tier core --tag pivot --category read --json
 *   bun run fixtures/benchmarks/competitive-real-workbook.ts --runner-manifest runners.json --json workbook.xlsx
 */
import { createHash } from 'node:crypto'
import { access, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, extname, resolve } from 'node:path'
import {
	indexToColumn,
	parseA1Safe,
	type RangeRef,
	type Workbook,
} from '../../packages/core/src/index.ts'
import { readXlsx } from '../../packages/io-xlsx/src/reader/index.ts'
import {
	parseRelationships,
	REL_COMMENTS,
	REL_SHARED_STRINGS,
	REL_WORKSHEET,
	resolvePath,
} from '../../packages/io-xlsx/src/reader/relationships.ts'
import {
	emptySharedStrings,
	parseSharedStrings,
	type SharedStringResolver,
} from '../../packages/io-xlsx/src/reader/shared-strings.ts'
import { parseWorkbookXml } from '../../packages/io-xlsx/src/reader/workbook.ts'
import { decodeXmlText } from '../../packages/io-xlsx/src/reader/xml-utils.ts'
import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'
import type { CellValue } from '../../packages/schema/src/index.ts'
import { topLeftScalar } from '../../packages/schema/src/index.ts'
import { Ascend } from '../../packages/sdk/src/index.ts'
import {
	type CorpusAssertionClass,
	type CorpusBenchmarkTier,
	type CorpusRiskClass,
	type CorpusSelection,
	loadCorpusManifestEntries,
	type NormalizedCorpusManifestEntry,
	normalizeManifest,
	selectManifestEntries,
} from '../corpus/manifest.ts'
import { workbookSheetEntriesForSummary } from './ascend-workbook-shape.ts'
import {
	type BenchmarkCaseResult,
	createBenchmarkSuite,
	formatBytes,
	formatRate,
	summarizeSamples,
} from './results.ts'

export { loadCorpusManifestEntries } from '../corpus/manifest.ts'

export const QUICK_TARGETS = [
	'fixtures/xlsx/stress/multi-sheet-10.xlsx',
	'fixtures/xlsx/xlsxwriter/styles_formulas.xlsx',
	'fixtures/xlsx/calamine/shared_formula_reversed.xlsx',
	'fixtures/xlsx/poi/StructuredReferences.xlsx',
	'fixtures/xlsx/poi/shared_formulas.xlsx',
	'fixtures/xlsx/poi/WithChart.xlsx',
	'fixtures/xlsx/libreoffice/universal-content-strict.xlsx',
	'fixtures/xlsx/libreoffice/PivotTable_CachedDefinitionAndDataInSync.xlsx',
	'research/excel-corpus/conditional-formatting.xlsx',
]

export const FULL_CORPUS_TARGETS = [
	...QUICK_TARGETS,
	'fixtures/xlsx/poi/ConditionalFormattingSamples.xlsx',
	'fixtures/xlsx/poi/DataValidationEvaluations.xlsx',
	'fixtures/xlsx/poi/FormulaEvalTestData_Copy.xlsx',
	'fixtures/xlsx/poi/NewStyleConditionalFormattings.xlsx',
	'fixtures/xlsx/poi/NumberFormatTests.xlsx',
	'fixtures/xlsx/poi/SimpleStrict.xlsx',
	'fixtures/xlsx/poi/SimpleWithComments.xlsx',
	'fixtures/xlsx/poi/Tables.xlsx',
	'fixtures/xlsx/poi/Themes.xlsx',
	'fixtures/xlsx/poi/AutoFilter.xlsx',
	'fixtures/xlsx/poi/formula_stress_test.xlsx',
	'fixtures/xlsx/poi/merge_cells.xlsx',
	'fixtures/xlsx/poi/named_ranges_2011.xlsx',
	'fixtures/xlsx/libreoffice/MissingPathExternal.xlsx',
	'fixtures/xlsx/libreoffice/TableStyleTest.xlsx',
	'fixtures/xlsx/libreoffice/activex_checkbox.xlsx',
	'fixtures/xlsx/libreoffice/textLengthDataValidity.xlsx',
	'fixtures/xlsx/libreoffice/totalsRowFunction.xlsx',
	'fixtures/xlsx/calamine/pivots.xlsx',
	'fixtures/xlsx/calamine/picture.xlsx',
	'fixtures/xlsx/calamine/richtext-namespaced.xlsx',
	'fixtures/xlsx/calamine/table-multiple.xlsx',
	'fixtures/xlsx/closedxml/Comments_AddingComments.xlsx',
	'fixtures/xlsx/closedxml/ConditionalFormatting_CFDataBars.xlsx',
	'fixtures/xlsx/closedxml/Misc_FormulasWithEvaluation.xlsx',
	'fixtures/xlsx/closedxml/Other_ExternalLinks_WorkbookWithExternalLink.xlsx',
	'fixtures/xlsx/closedxml/Other_PivotTableReferenceFiles_ChartsheetAndPivotTable.xlsx',
	'fixtures/xlsx/exceljs/bogus-defined-name.xlsx',
	'fixtures/xlsx/exceljs/chart-sheet.xlsx',
	'fixtures/xlsx/exceljs/fibonacci.xlsx',
	'fixtures/xlsx/exceljs/formulas.xlsx',
	'fixtures/xlsx/exceljs/test-issue-1669.xlsx',
	'fixtures/xlsx/exceljs/test-issue-1842.xlsx',
	'research/excel-corpus/ms-excel-formulas-and-pivot-tables.xlsx',
	'research/excel-corpus/bevreport-demo.xlsm',
	'research/excel-corpus/excel-dashboard-v2.xlsx',
	'research/excel-corpus/large-macro-example.xlsm',
	'fixtures/xlsx/external/sec-mmf-statistics-2022-02.xlsx',
	'fixtures/xlsx/external/uk-gov-spend-nice-2026-02.xlsx',
	'fixtures/xlsx/external/us-census-construction-2025-10.xlsx',
	'fixtures/xlsx/stress/many-strings.xlsx',
	'fixtures/xlsx/stress/many-styles.xlsx',
	'fixtures/xlsx/stress/formula-dense.xlsx',
	'fixtures/xlsx/stress/dense-100k.xlsx',
	'fixtures/xlsx/stress/merged-complex.xlsx',
	'fixtures/xlsx/xlsxwriter/strings_links.xlsx',
	'fixtures/xlsx/xlsxwriter/layout_breaks.xlsx',
	'fixtures/xlsx/xlsxwriter/multisheet_names.xlsx',
	'fixtures/xlsx/filter/poi/AutoFilter.xlsx',
	'fixtures/xlsx/filter/poi/ConditionalFormattingSamples.xlsx',
	'fixtures/xlsx/libreoffice/129969-min.xlsx',
	'fixtures/xlsx/libreoffice/CalcThemeTest.xlsx',
	'fixtures/xlsx/libreoffice/PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithCacheData.xlsx',
	'fixtures/xlsx/libreoffice/PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithoutCacheData.xlsx',
	'fixtures/xlsx/libreoffice/ProtecteSheet1234Pass.xlsx',
	'fixtures/xlsx/libreoffice/Sparklines.xlsx',
	'fixtures/xlsx/libreoffice/TableEmptyHeaders.xlsx',
	'fixtures/xlsx/libreoffice/Test_ThemeColor_Text_Background_Border.xlsx',
	'fixtures/xlsx/libreoffice/autofilter-colors.xlsx',
	'fixtures/xlsx/libreoffice/autofilter.xlsx',
	'fixtures/xlsx/libreoffice/colorscale.xlsx',
	'fixtures/xlsx/libreoffice/complex_icon_set.xlsx',
	'fixtures/xlsx/libreoffice/condFormat_cellis.xlsx',
	'fixtures/xlsx/libreoffice/databar.xlsx',
	'fixtures/xlsx/libreoffice/functions-excel-2010.xlsx',
	'fixtures/xlsx/libreoffice/matrix-multiplication.xlsx',
	'fixtures/xlsx/libreoffice/pivot-table/tdf126858-1.xlsx',
	'fixtures/xlsx/libreoffice/pivot-table/test_diff_aggregation.xlsx',
	'fixtures/xlsx/libreoffice/pivot_table_first_header_row.xlsx',
	'fixtures/xlsx/libreoffice/pivottable_date_field_filter.xlsx',
	'fixtures/xlsx/libreoffice/sortconditionref2.xlsx',
	'fixtures/xlsx/libreoffice/tdf143068_top10filter.xlsx',
	'fixtures/xlsx/libreoffice/tdf165180_date1904.xlsx',
	'fixtures/xlsx/libreoffice/tdf167689_tableType.xlsx',
	'fixtures/xlsx/libreoffice/tdf170201.xlsx',
	'fixtures/xlsx/libreoffice/textbox-hyperlink.xlsx',
	'fixtures/xlsx/libreoffice/totalsRowShown.xlsx',
	'fixtures/xlsx/libreoffice/universal-content.xlsx',
	'fixtures/xlsx/libreoffice/user_defined_function.xlsx',
	'fixtures/xlsx/libreoffice/value-in-column-2000.xlsx',
	'fixtures/xlsx/libreoffice/writingMode.xlsx',
	'fixtures/xlsx/calamine/any_sheets.xlsx',
	'fixtures/xlsx/calamine/column_row_ranges.xlsx',
	'fixtures/xlsx/calamine/date.xlsx',
	'fixtures/xlsx/calamine/date_1904.xlsx',
	'fixtures/xlsx/calamine/date_iso.xlsx',
	'fixtures/xlsx/calamine/empty_s_attribute.xlsx',
	'fixtures/xlsx/calamine/empty_shared_string.xlsx',
	'fixtures/xlsx/calamine/empty_shared_string_value.xlsx',
	'fixtures/xlsx/calamine/empty_sheet.xlsx',
	'fixtures/xlsx/calamine/encoded_entities.xlsx',
	'fixtures/xlsx/calamine/errors.xlsx',
	'fixtures/xlsx/calamine/formula.issue.xlsx',
	'fixtures/xlsx/calamine/has_x000D_.xlsx',
	'fixtures/xlsx/calamine/has_x000D_inline.xlsx',
	'fixtures/xlsx/calamine/header-row.xlsx',
	'fixtures/xlsx/calamine/inlineStr_with_value.xlsx',
	'fixtures/xlsx/calamine/inventory-table.xlsx',
	'fixtures/xlsx/calamine/issue127.xlsx',
	'fixtures/xlsx/calamine/issue221.xlsm',
	'fixtures/xlsx/calamine/issue281.xlsm',
	'fixtures/xlsx/calamine/issue3.xlsm',
	'fixtures/xlsx/calamine/issue438.xlsx',
	'fixtures/xlsx/calamine/issue446.xlsx',
	'fixtures/xlsx/calamine/issue9.xlsx',
	'fixtures/xlsx/calamine/issue_174.xlsx',
	'fixtures/xlsx/calamine/issue_261.xlsx',
	'fixtures/xlsx/calamine/issue_261_fixed_by_excel.xlsx',
	'fixtures/xlsx/calamine/issue_391.xlsx',
	'fixtures/xlsx/calamine/issue_419.xlsx',
	'fixtures/xlsx/calamine/issue_530.xlsx',
	'fixtures/xlsx/calamine/issue_553.xlsx',
	'fixtures/xlsx/calamine/issue_565_multi_axis_shared.xlsx',
	'fixtures/xlsx/calamine/issue_567_absolute_shared.xlsx',
	'fixtures/xlsx/calamine/issues.xlsx',
	'fixtures/xlsx/calamine/merge_cells.xlsx',
	'fixtures/xlsx/calamine/merged_range.xlsx',
	'fixtures/xlsx/calamine/no-header.xlsx',
	'fixtures/xlsx/calamine/non_monotonic_si.xlsx',
	'fixtures/xlsx/calamine/rph.xlsx',
	'fixtures/xlsx/calamine/string-ref.xlsx',
	'fixtures/xlsx/calamine/table_with_absolute_paths.xlsx',
	'fixtures/xlsx/calamine/table_with_insertrow_attribute.xlsx',
	'fixtures/xlsx/calamine/temperature-in-middle.xlsx',
	'fixtures/xlsx/calamine/temperature-table.xlsx',
	'fixtures/xlsx/calamine/temperature.xlsx',
	'fixtures/xlsx/calamine/vba.xlsm',
	'fixtures/xlsx/closedxml/AutoFilter_CustomAutoFilter.xlsx',
	'fixtures/xlsx/closedxml/ConditionalFormatting_CFIconSet.xlsx',
	'fixtures/xlsx/closedxml/ImageHandling_ImageAnchors.xlsx',
	'fixtures/xlsx/closedxml/Misc_DataValidation.xlsx',
	'fixtures/xlsx/closedxml/Misc_Formulas.xlsx',
	'fixtures/xlsx/closedxml/Misc_Hyperlinks.xlsx',
	'fixtures/xlsx/closedxml/Misc_MergeCells.xlsx',
	'fixtures/xlsx/closedxml/Misc_SheetProtection.xlsx',
	'fixtures/xlsx/closedxml/Misc_ShiftingFormulas.xlsx',
	'fixtures/xlsx/closedxml/Other_Charts_PreserveCharts_inputfile.xlsx',
	'fixtures/xlsx/closedxml/Other_Formulas_ArrayFormula.xlsx',
	'fixtures/xlsx/closedxml/Other_Formulas_BooleanFormulaValues.xlsx',
	'fixtures/xlsx/closedxml/Other_Formulas_DataTableFormula-Excel-Input.xlsx',
	'fixtures/xlsx/closedxml/PivotTables_PivotTables.xlsx',
	'fixtures/xlsx/closedxml/Ranges_DefinedNames.xlsx',
	'fixtures/xlsx/closedxml/Ranges_SortExample.xlsx',
	'fixtures/xlsx/closedxml/Sparklines_SampleSparklines.xlsx',
	'fixtures/xlsx/closedxml/Styles_StyleNumberFormat.xlsx',
	'fixtures/xlsx/closedxml/Styles_UsingRichText.xlsx',
	'fixtures/xlsx/closedxml/Tables_ResizingTables.xlsx',
	'fixtures/xlsx/closedxml/Tables_UsingTables.xlsx',
	'fixtures/xlsx/exceljs/1904.xlsx',
	'fixtures/xlsx/exceljs/dateIssue.xlsx',
	'fixtures/xlsx/exceljs/many-columns.xlsx',
	'fixtures/xlsx/exceljs/shared_string_with_escape.xlsx',
	'fixtures/xlsx/exceljs/test-issue-1364.xlsx',
	'fixtures/xlsx/exceljs/test-issue-1575.xlsx',
	'fixtures/xlsx/exceljs/test-issue-163.xlsx',
	'fixtures/xlsx/exceljs/test-issue-176.xlsx',
	'fixtures/xlsx/exceljs/test-issue-623.xlsx',
	'fixtures/xlsx/exceljs/test-pr-1204.xlsx',
	'fixtures/xlsx/exceljs/test-pr-1220.xlsx',
	'fixtures/xlsx/exceljs/test-pr-567.xlsx',
	'fixtures/xlsx/exceljs/test-pr-728.xlsx',
	'fixtures/xlsx/exceljs/test-row-styles.xlsx',
]

const DEFAULT_CORPUS_MANIFEST = 'research/excel-corpus/manifest.json'
const DEFAULT_CORPUS_ROOT = 'research/excel-corpus'

interface WorkbookTarget {
	readonly path: string
	readonly name: string
	readonly extension: string
	readonly bytes: Uint8Array
	readonly packageBytes: Uint8Array
	readonly sizeBytes: number
	readonly sha256: string
	readonly packageSha256: string
	readonly expectedInfo: WorkbookShapeSummary
	readonly ascendInfo?: {
		readonly sheetCount: number
		readonly cellCount: number | null
		readonly physicalCellCount: number | null
		readonly formulaCount: number
		readonly usedRanges: readonly string[]
		readonly physicalUsedRanges: readonly string[]
		readonly compatibility: string
	}
	readonly corpus?: CorpusTargetMetadata
}

interface CorpusTargetMetadata {
	readonly file: string
	readonly benchmarkTier: CorpusBenchmarkTier
	readonly assertionClass: CorpusAssertionClass
	readonly riskClass: CorpusRiskClass
	readonly featureTags: readonly string[]
	readonly vendorable: boolean
	readonly knownUnsupported: readonly string[]
	readonly password?: string
}

type CompetitiveCategory = 'read' | 'roundtrip' | 'edit-roundtrip'
type ExternalRunnerCategory = CompetitiveCategory | 'write'

export interface WorkbookShapeSummary {
	readonly sheetNames: readonly string[]
	readonly sheetCount: number
	readonly cellCount: number
	readonly physicalCellCount: number | null
	readonly formulaCount: number
	readonly usedRanges: readonly string[]
	readonly physicalUsedRanges: readonly string[]
	readonly sheetNamesHash: string
	readonly usedRangesHash: string
	readonly physicalUsedRangesHash: string
	readonly semanticCellRefsHash: string
	readonly semanticCellValuesHash: string
	readonly formulaTextHash: string
	readonly orderedSemanticCellRefsHash?: string
	readonly orderedSemanticCellValuesHash?: string
	readonly orderedFormulaTextHash?: string
	readonly packageFingerprint?: WorkbookPackageFingerprint
	readonly featureSummary?: WorkbookFeatureSummary
}

export interface WorkbookPackageFingerprint {
	readonly partCount: number
	readonly partNamesHash: string
	readonly contentTypeCount: number
	readonly contentTypesHash: string
	readonly relationshipCount: number
	readonly relationshipGraphHash: string
	readonly preservedPartCount: number
	readonly preservedPartNamesHash: string
	readonly preservedPartContentHash: string
}

export interface WorkbookFeatureSummary {
	readonly tablePartCount: number
	readonly chartPartCount: number
	readonly chartExPartCount: number
	readonly drawingPartCount: number
	readonly vmlDrawingPartCount: number
	readonly pivotTablePartCount: number
	readonly pivotCachePartCount: number
	readonly slicerPartCount: number
	readonly commentPartCount: number
	readonly packageCommentEntryCount: number
	readonly threadedCommentPartCount: number
	readonly mediaPartCount: number
	readonly externalLinkPartCount: number
	readonly connectionPartCount: number
	readonly customXmlPartCount: number
	readonly worksheetCommentCount: number
	readonly worksheetHyperlinkCount: number
	readonly worksheetDataValidationCount: number
	readonly worksheetConditionalFormattingCount: number
	readonly definedNameCount: number
	readonly featurePartNamesHash: string
	readonly featureInventoryHash: string
}

interface AscendWorkbookLike {
	readonly report: { readonly status: string }
	getWorkbookModel(): Workbook
}

interface CaseEvaluation {
	readonly status: string
	readonly assertions: Record<string, string | number | boolean | null>
}

function isRankingEligible(status: string): boolean {
	return (
		status === 'pass' || status === 'exact-package-match' || status === 'semantic-roundtrip-pass'
	)
}

function numericAssertion(value: string | number | boolean | null | undefined): number {
	return typeof value === 'number' && Number.isFinite(value) ? value : 0
}

function hasExpectedReadFeatures(summary: WorkbookFeatureSummary | undefined): boolean {
	return (
		summary !== undefined &&
		(summary.worksheetCommentCount > 0 ||
			summary.worksheetHyperlinkCount > 0 ||
			summary.worksheetDataValidationCount > 0 ||
			summary.worksheetConditionalFormattingCount > 0 ||
			summary.definedNameCount > 0)
	)
}

export function workbookReadFeatureAssertions(
	workbook: Workbook,
): Record<string, string | number | boolean | null> {
	let readCommentCount = 0
	let readHyperlinkCount = 0
	let readDataValidationCount = 0
	let readConditionalFormatCount = 0
	for (const sheet of workbook.sheets) {
		readCommentCount += sheet.comments.size
		readHyperlinkCount += sheet.hyperlinks.size
		readDataValidationCount += sheet.dataValidations.length
		readConditionalFormatCount +=
			sheet.conditionalFormats.length + sheet.x14ConditionalFormats.length
	}
	return {
		readCommentCount,
		readHyperlinkCount,
		readDataValidationCount,
		readConditionalFormatCount,
		readDefinedNameCount: workbook.definedNames.size,
	}
}

export function workbookShapeAssertions(
	shape: WorkbookShapeSummary,
): Record<string, string | number | boolean | null> {
	return {
		sheetCount: shape.sheetCount,
		sheetNamesHash: shape.sheetNamesHash,
		cellCount: shape.cellCount,
		physicalCellCount: shape.physicalCellCount,
		formulaCount: shape.formulaCount,
		usedRangeCount: shape.usedRanges.length,
		firstUsedRange: shape.usedRanges[0] ?? null,
		firstPhysicalUsedRange: shape.physicalUsedRanges[0] ?? null,
		usedRangesHash: shape.usedRangesHash,
		physicalUsedRangesHash: shape.physicalUsedRangesHash,
		semanticCellRefsHash: shape.semanticCellRefsHash,
		semanticCellValuesHash: shape.semanticCellValuesHash,
		formulaTextHash: shape.formulaTextHash,
		...featureAssertions('', shape.featureSummary),
	}
}

export function roundtripAssertions(
	bytes: Uint8Array,
	target: WorkbookTarget,
): Record<string, string | number | boolean | null> {
	const shape = extractExpectedWorkbookShape(bytes)
	const packageFingerprint = shape.packageFingerprint ?? extractWorkbookPackageFingerprint(bytes)
	return {
		bytes: bytes.byteLength,
		byteIdentical: sha256(bytes) === target.packageSha256,
		roundtripSheetCount: shape.sheetCount,
		roundtripSheetNamesHash: shape.sheetNamesHash,
		roundtripCellCount: shape.cellCount,
		roundtripPhysicalCellCount: shape.physicalCellCount,
		roundtripFormulaCount: shape.formulaCount,
		roundtripUsedRangeCount: shape.usedRanges.length,
		roundtripFirstUsedRange: shape.usedRanges[0] ?? null,
		roundtripFirstPhysicalUsedRange: shape.physicalUsedRanges[0] ?? null,
		roundtripUsedRangesHash: shape.usedRangesHash,
		roundtripPhysicalUsedRangesHash: shape.physicalUsedRangesHash,
		roundtripSemanticCellRefsHash: shape.semanticCellRefsHash,
		roundtripSemanticCellValuesHash: shape.semanticCellValuesHash,
		roundtripFormulaTextHash: shape.formulaTextHash,
		roundtripPackagePartCount: packageFingerprint.partCount,
		roundtripPackagePartNamesHash: packageFingerprint.partNamesHash,
		roundtripPackageContentTypeCount: packageFingerprint.contentTypeCount,
		roundtripPackageContentTypesHash: packageFingerprint.contentTypesHash,
		roundtripPackageRelationshipCount: packageFingerprint.relationshipCount,
		roundtripPackageRelationshipGraphHash: packageFingerprint.relationshipGraphHash,
		roundtripPreservedPartCount: packageFingerprint.preservedPartCount,
		roundtripPreservedPartNamesHash: packageFingerprint.preservedPartNamesHash,
		roundtripPreservedPartContentHash: packageFingerprint.preservedPartContentHash,
		...featureAssertions('roundtrip', shape.featureSummary),
	}
}

interface EditTarget {
	readonly sheetName: string
	readonly ref: string
	readonly mode: 'replace-cell' | 'add-cell'
	readonly valueType: 'number' | 'string' | 'boolean'
	readonly oldValue: string | number | boolean | null
	readonly newValue: string | number | boolean
	readonly expectedCellDelta: number
	readonly expectedPhysicalCellDelta: number
	readonly expectedFormulaDelta: number
}

function selectEditTarget(target: WorkbookTarget): EditTarget {
	const selected = findFirstEditableScalarCell(target.packageBytes)
	if (selected) {
		return {
			sheetName: selected.sheetName,
			ref: selected.ref,
			mode: 'replace-cell',
			valueType: selected.valueType,
			oldValue: selected.value,
			newValue: editReplacementValue(selected.value),
			expectedCellDelta: 0,
			expectedPhysicalCellDelta: 0,
			expectedFormulaDelta: 0,
		}
	}
	const addTarget = selectAddCellEditTarget(target.packageBytes)
	if (!addTarget) {
		throw new Error(`${target.name} has no worksheet for edit-roundtrip`)
	}
	return addTarget
}

function editRoundtripAssertions(
	bytes: Uint8Array,
	target: WorkbookTarget,
	edit: EditTarget,
): Record<string, string | number | boolean | null> {
	const assertions = roundtripAssertions(bytes, target)
	const observed = readCellScalar(bytes, edit.sheetName, edit.ref)
	const roundtripCellCount = numericAssertion(assertions.roundtripCellCount)
	const roundtripPhysicalCellCount = numericAssertion(assertions.roundtripPhysicalCellCount)
	const expectedPhysicalCellCount = numericAssertion(target.expectedInfo.physicalCellCount)
	const roundtripFormulaCount = numericAssertion(assertions.roundtripFormulaCount)
	return {
		...assertions,
		editSheetName: edit.sheetName,
		editRef: edit.ref,
		editMode: edit.mode,
		editValueType: edit.valueType,
		editOldValue: edit.oldValue,
		editExpectedValue: edit.newValue,
		editObservedValue: observed,
		editExpectedCellDelta: edit.expectedCellDelta,
		editObservedCellDelta: roundtripCellCount - target.expectedInfo.cellCount,
		editExpectedPhysicalCellDelta: edit.expectedPhysicalCellDelta,
		editObservedPhysicalCellDelta: roundtripPhysicalCellCount - expectedPhysicalCellCount,
		editExpectedFormulaDelta: edit.expectedFormulaDelta,
		editObservedFormulaDelta: roundtripFormulaCount - target.expectedInfo.formulaCount,
		editCellValueMatches: observed === edit.newValue,
		editFormulaTextUnchanged:
			assertions.roundtripFormulaTextHash === target.expectedInfo.formulaTextHash,
		editSemanticRefsChanged:
			assertions.roundtripSemanticCellRefsHash !== target.expectedInfo.semanticCellRefsHash,
		editPreservedPartNamesUnchanged:
			assertions.roundtripPreservedPartNamesHash ===
			target.expectedInfo.packageFingerprint?.preservedPartNamesHash,
		editPreservedPartContentUnchanged:
			assertions.roundtripPreservedPartContentHash ===
			target.expectedInfo.packageFingerprint?.preservedPartContentHash,
		editFeatureInventoryUnchanged:
			assertions.roundtripFeatureInventoryHash ===
			target.expectedInfo.featureSummary?.featureInventoryHash,
	}
}

export function evaluateAssertions(
	category: CompetitiveCategory,
	expected: WorkbookShapeSummary,
	assertions: Record<string, string | number | boolean | null> | undefined,
	operationProfile = 'default',
): CaseEvaluation {
	const observed = assertions ?? {}
	if (category === 'read') {
		const expectedFeature = expected.featureSummary
		const sheetCountMatches = observed.sheetCount === expected.sheetCount
		const sheetNamesHashMatches = observed.sheetNamesHash === expected.sheetNamesHash
		const cellCountMatches = observed.cellCount === expected.cellCount
		const formulaPreservationRequired = operationProfile !== 'read-values'
		const formulaCountMatches =
			!formulaPreservationRequired || observed.formulaCount === expected.formulaCount
		const firstUsedRangeMatches =
			expected.usedRanges.length === 0 || observed.firstUsedRange === expected.usedRanges[0]
		const usedRangesHashMatches = observed.usedRangesHash === expected.usedRangesHash
		const orderedSemanticCellRefsHashMatches =
			expected.orderedSemanticCellRefsHash !== undefined &&
			observed.orderedSemanticCellRefsHash !== undefined
				? observed.orderedSemanticCellRefsHash === expected.orderedSemanticCellRefsHash
				: null
		const orderedSemanticCellValuesHashMatches =
			expected.orderedSemanticCellValuesHash !== undefined &&
			observed.orderedSemanticCellValuesHash !== undefined
				? observed.orderedSemanticCellValuesHash === expected.orderedSemanticCellValuesHash
				: null
		const orderedFormulaTextHashMatches =
			expected.orderedFormulaTextHash !== undefined && observed.orderedFormulaTextHash !== undefined
				? observed.orderedFormulaTextHash === expected.orderedFormulaTextHash
				: null
		const semanticCellRefsHashMatches =
			orderedSemanticCellRefsHashMatches !== null
				? orderedSemanticCellRefsHashMatches
				: observed.semanticCellRefsHash === expected.semanticCellRefsHash
		const semanticCellValuesHashMatches =
			orderedSemanticCellValuesHashMatches !== null
				? orderedSemanticCellValuesHashMatches
				: observed.semanticCellValuesHash === expected.semanticCellValuesHash
		const formulaTextHashRequired =
			formulaPreservationRequired && observed.compatibility !== 'has-unsupported'
		const formulaTextHashMatches =
			!formulaTextHashRequired ||
			(orderedFormulaTextHashMatches !== null
				? orderedFormulaTextHashMatches
				: observed.formulaTextHash === expected.formulaTextHash)
		const readFeatureVerificationRequired =
			formulaPreservationRequired && hasExpectedReadFeatures(expectedFeature)
		const hasReadFeatureAssertions =
			observed.readCommentCount !== undefined &&
			observed.readHyperlinkCount !== undefined &&
			observed.readDataValidationCount !== undefined &&
			observed.readConditionalFormatCount !== undefined &&
			observed.readDefinedNameCount !== undefined
		const readCommentCountMatches =
			!readFeatureVerificationRequired ||
			(hasReadFeatureAssertions &&
				observed.readCommentCount === expectedFeature?.worksheetCommentCount)
		const readHyperlinkCountMatches =
			!readFeatureVerificationRequired ||
			(hasReadFeatureAssertions &&
				observed.readHyperlinkCount === expectedFeature?.worksheetHyperlinkCount)
		const readDataValidationCountMatches =
			!readFeatureVerificationRequired ||
			(hasReadFeatureAssertions &&
				observed.readDataValidationCount === expectedFeature?.worksheetDataValidationCount)
		const readConditionalFormatCountMatches =
			!readFeatureVerificationRequired ||
			(hasReadFeatureAssertions &&
				observed.readConditionalFormatCount ===
					expectedFeature?.worksheetConditionalFormattingCount)
		const readDefinedNameCountMatches =
			!readFeatureVerificationRequired ||
			(hasReadFeatureAssertions &&
				observed.readDefinedNameCount === expectedFeature?.definedNameCount)
		const readFeatureCountsMatch =
			readCommentCountMatches &&
			readHyperlinkCountMatches &&
			readDataValidationCountMatches &&
			readConditionalFormatCountMatches &&
			readDefinedNameCountMatches
		const matches =
			sheetCountMatches &&
			sheetNamesHashMatches &&
			cellCountMatches &&
			formulaCountMatches &&
			firstUsedRangeMatches &&
			usedRangesHashMatches &&
			semanticCellRefsHashMatches &&
			semanticCellValuesHashMatches &&
			formulaTextHashMatches &&
			readFeatureCountsMatch
		return {
			status: matches
				? 'pass'
				: readFeatureVerificationRequired && !hasReadFeatureAssertions
					? 'feature-read-unverified'
					: readFeatureVerificationRequired && !readFeatureCountsMatch
						? 'feature-read-mismatch'
						: 'semantic-mismatch',
			assertions: {
				...observed,
				expectedSheetCount: expected.sheetCount,
				expectedSheetNamesHash: expected.sheetNamesHash,
				expectedCellCount: expected.cellCount,
				expectedPhysicalCellCount: expected.physicalCellCount,
				expectedFormulaCount: expected.formulaCount,
				expectedFirstUsedRange: expected.usedRanges[0] ?? null,
				expectedFirstPhysicalUsedRange: expected.physicalUsedRanges[0] ?? null,
				expectedUsedRangesHash: expected.usedRangesHash,
				expectedSemanticCellRefsHash: expected.semanticCellRefsHash,
				expectedSemanticCellValuesHash: expected.semanticCellValuesHash,
				expectedFormulaTextHash: expected.formulaTextHash,
				expectedOrderedSemanticCellRefsHash: expected.orderedSemanticCellRefsHash ?? null,
				expectedOrderedSemanticCellValuesHash: expected.orderedSemanticCellValuesHash ?? null,
				expectedOrderedFormulaTextHash: expected.orderedFormulaTextHash ?? null,
				expectedReadCommentCount: expectedFeature?.worksheetCommentCount ?? null,
				expectedReadHyperlinkCount: expectedFeature?.worksheetHyperlinkCount ?? null,
				expectedReadDataValidationCount: expectedFeature?.worksheetDataValidationCount ?? null,
				expectedReadConditionalFormatCount:
					expectedFeature?.worksheetConditionalFormattingCount ?? null,
				expectedReadDefinedNameCount: expectedFeature?.definedNameCount ?? null,
				formulaPreservationRequired,
				formulaTextHashRequired,
				readFeatureVerificationRequired,
				hasReadFeatureAssertions,
				sheetCountMatches,
				sheetNamesHashMatches,
				cellCountMatches,
				formulaCountMatches,
				firstUsedRangeMatches,
				usedRangesHashMatches,
				semanticCellRefsHashMatches,
				semanticCellValuesHashMatches,
				formulaTextHashMatches,
				orderedSemanticCellRefsHashMatches,
				orderedSemanticCellValuesHashMatches,
				orderedFormulaTextHashMatches,
				readCommentCountMatches,
				readHyperlinkCountMatches,
				readDataValidationCountMatches,
				readConditionalFormatCountMatches,
				readDefinedNameCountMatches,
			},
		}
	}
	const isEditRoundtrip = category === 'edit-roundtrip'
	const isAddCellEdit = isEditRoundtrip && observed.editMode === 'add-cell'
	const expectedCellDelta = isEditRoundtrip ? numericAssertion(observed.editExpectedCellDelta) : 0
	const expectedPhysicalCellDelta = isEditRoundtrip
		? numericAssertion(observed.editExpectedPhysicalCellDelta)
		: 0
	const expectedFormulaDelta = isEditRoundtrip
		? numericAssertion(observed.editExpectedFormulaDelta)
		: 0
	const byteIdentical = !isEditRoundtrip && observed.byteIdentical === true
	const roundtripSheetCountMatches = observed.roundtripSheetCount === expected.sheetCount
	const roundtripSheetNamesHashMatches =
		observed.roundtripSheetNamesHash === expected.sheetNamesHash
	const roundtripCellCountMatches =
		observed.roundtripCellCount === expected.cellCount + expectedCellDelta
	const roundtripPhysicalCellCountMatches =
		!isAddCellEdit ||
		expected.physicalCellCount === null ||
		observed.roundtripPhysicalCellCount === expected.physicalCellCount + expectedPhysicalCellDelta
	const roundtripFormulaCountMatches =
		observed.roundtripFormulaCount === expected.formulaCount + expectedFormulaDelta
	const roundtripFirstUsedRangeMatches =
		isAddCellEdit ||
		expected.usedRanges.length === 0 ||
		observed.roundtripFirstUsedRange === expected.usedRanges[0]
	const roundtripUsedRangesHashMatches =
		isAddCellEdit || observed.roundtripUsedRangesHash === expected.usedRangesHash
	const roundtripSemanticCellRefsHashMatches = isAddCellEdit
		? observed.editSemanticRefsChanged === true
		: observed.roundtripSemanticCellRefsHash === expected.semanticCellRefsHash
	const roundtripSemanticCellValuesHashMatches =
		isEditRoundtrip || observed.roundtripSemanticCellValuesHash === expected.semanticCellValuesHash
	const roundtripFormulaTextHashMatches =
		observed.roundtripFormulaTextHash === expected.formulaTextHash
	const expectedPackage = expected.packageFingerprint
	const expectedFeature = expected.featureSummary
	const packageFingerprintRequired = !byteIdentical && expectedPackage !== undefined
	const featureFingerprintRequired = !byteIdentical && expectedFeature !== undefined
	const hasRoundtripPackageFingerprint =
		observed.roundtripPackagePartCount !== undefined &&
		observed.roundtripPackagePartNamesHash !== undefined &&
		observed.roundtripPackageContentTypeCount !== undefined &&
		observed.roundtripPackageContentTypesHash !== undefined &&
		observed.roundtripPackageRelationshipCount !== undefined &&
		observed.roundtripPackageRelationshipGraphHash !== undefined &&
		observed.roundtripPreservedPartCount !== undefined &&
		observed.roundtripPreservedPartNamesHash !== undefined &&
		observed.roundtripPreservedPartContentHash !== undefined
	const hasRoundtripFeatureFingerprint =
		observed.roundtripTablePartCount !== undefined &&
		observed.roundtripChartPartCount !== undefined &&
		observed.roundtripChartExPartCount !== undefined &&
		observed.roundtripDrawingPartCount !== undefined &&
		observed.roundtripVmlDrawingPartCount !== undefined &&
		observed.roundtripPivotTablePartCount !== undefined &&
		observed.roundtripPivotCachePartCount !== undefined &&
		observed.roundtripSlicerPartCount !== undefined &&
		observed.roundtripCommentPartCount !== undefined &&
		observed.roundtripPackageCommentEntryCount !== undefined &&
		observed.roundtripThreadedCommentPartCount !== undefined &&
		observed.roundtripMediaPartCount !== undefined &&
		observed.roundtripExternalLinkPartCount !== undefined &&
		observed.roundtripConnectionPartCount !== undefined &&
		observed.roundtripCustomXmlPartCount !== undefined &&
		observed.roundtripWorksheetCommentCount !== undefined &&
		observed.roundtripWorksheetHyperlinkCount !== undefined &&
		observed.roundtripWorksheetDataValidationCount !== undefined &&
		observed.roundtripWorksheetConditionalFormattingCount !== undefined &&
		observed.roundtripDefinedNameCount !== undefined &&
		observed.roundtripFeaturePartNamesHash !== undefined &&
		observed.roundtripFeatureInventoryHash !== undefined
	const roundtripPackagePartCountMatches =
		!packageFingerprintRequired || observed.roundtripPackagePartCount === expectedPackage?.partCount
	const roundtripPackagePartNamesHashMatches =
		!packageFingerprintRequired ||
		observed.roundtripPackagePartNamesHash === expectedPackage?.partNamesHash
	const roundtripPackageContentTypeCountMatches =
		!packageFingerprintRequired ||
		observed.roundtripPackageContentTypeCount === expectedPackage?.contentTypeCount
	const roundtripPackageContentTypesHashMatches =
		!packageFingerprintRequired ||
		observed.roundtripPackageContentTypesHash === expectedPackage?.contentTypesHash
	const roundtripPackageRelationshipCountMatches =
		!packageFingerprintRequired ||
		observed.roundtripPackageRelationshipCount === expectedPackage?.relationshipCount
	const roundtripPackageRelationshipGraphHashMatches =
		!packageFingerprintRequired ||
		observed.roundtripPackageRelationshipGraphHash === expectedPackage?.relationshipGraphHash
	const roundtripPreservedPartCountMatches =
		!packageFingerprintRequired ||
		observed.roundtripPreservedPartCount === expectedPackage?.preservedPartCount
	const roundtripPreservedPartNamesHashMatches =
		!packageFingerprintRequired ||
		observed.roundtripPreservedPartNamesHash === expectedPackage?.preservedPartNamesHash
	const roundtripPreservedPartContentHashMatches =
		!packageFingerprintRequired ||
		observed.roundtripPreservedPartContentHash === expectedPackage?.preservedPartContentHash
	const packageTopologyMatches =
		isEditRoundtrip ||
		(roundtripPackagePartCountMatches &&
			roundtripPackagePartNamesHashMatches &&
			roundtripPackageContentTypeCountMatches &&
			roundtripPackageContentTypesHashMatches &&
			roundtripPackageRelationshipCountMatches &&
			roundtripPackageRelationshipGraphHashMatches)
	const packageRoundtripMatches =
		!packageFingerprintRequired ||
		(hasRoundtripPackageFingerprint &&
			packageTopologyMatches &&
			roundtripPreservedPartCountMatches &&
			roundtripPreservedPartNamesHashMatches &&
			roundtripPreservedPartContentHashMatches)
	const featureRoundtripMatches =
		!featureFingerprintRequired ||
		(hasRoundtripFeatureFingerprint &&
			observed.roundtripTablePartCount === expectedFeature?.tablePartCount &&
			observed.roundtripChartPartCount === expectedFeature?.chartPartCount &&
			observed.roundtripChartExPartCount === expectedFeature?.chartExPartCount &&
			observed.roundtripDrawingPartCount === expectedFeature?.drawingPartCount &&
			observed.roundtripVmlDrawingPartCount === expectedFeature?.vmlDrawingPartCount &&
			observed.roundtripPivotTablePartCount === expectedFeature?.pivotTablePartCount &&
			observed.roundtripPivotCachePartCount === expectedFeature?.pivotCachePartCount &&
			observed.roundtripSlicerPartCount === expectedFeature?.slicerPartCount &&
			observed.roundtripCommentPartCount === expectedFeature?.commentPartCount &&
			observed.roundtripPackageCommentEntryCount === expectedFeature?.packageCommentEntryCount &&
			observed.roundtripThreadedCommentPartCount === expectedFeature?.threadedCommentPartCount &&
			observed.roundtripMediaPartCount === expectedFeature?.mediaPartCount &&
			observed.roundtripExternalLinkPartCount === expectedFeature?.externalLinkPartCount &&
			observed.roundtripConnectionPartCount === expectedFeature?.connectionPartCount &&
			observed.roundtripCustomXmlPartCount === expectedFeature?.customXmlPartCount &&
			observed.roundtripWorksheetCommentCount === expectedFeature?.worksheetCommentCount &&
			observed.roundtripWorksheetHyperlinkCount === expectedFeature?.worksheetHyperlinkCount &&
			observed.roundtripWorksheetDataValidationCount ===
				expectedFeature?.worksheetDataValidationCount &&
			observed.roundtripWorksheetConditionalFormattingCount ===
				expectedFeature?.worksheetConditionalFormattingCount &&
			observed.roundtripDefinedNameCount === expectedFeature?.definedNameCount &&
			observed.roundtripFeaturePartNamesHash === expectedFeature?.featurePartNamesHash &&
			observed.roundtripFeatureInventoryHash === expectedFeature?.featureInventoryHash)
	const hasRoundtripShape =
		observed.roundtripSheetCount !== undefined &&
		observed.roundtripCellCount !== undefined &&
		observed.roundtripFormulaCount !== undefined
	const semanticRoundtripMatches =
		hasRoundtripShape &&
		roundtripSheetCountMatches &&
		roundtripSheetNamesHashMatches &&
		roundtripCellCountMatches &&
		roundtripPhysicalCellCountMatches &&
		roundtripFormulaCountMatches &&
		roundtripFirstUsedRangeMatches &&
		roundtripUsedRangesHashMatches &&
		roundtripSemanticCellRefsHashMatches &&
		roundtripSemanticCellValuesHashMatches &&
		roundtripFormulaTextHashMatches
	const editCellValueMatches = !isEditRoundtrip || observed.editCellValueMatches === true
	const editStructuralDeltaMatches =
		!isEditRoundtrip ||
		(observed.editObservedCellDelta === expectedCellDelta &&
			(!isAddCellEdit || observed.editObservedPhysicalCellDelta === expectedPhysicalCellDelta) &&
			observed.editObservedFormulaDelta === expectedFormulaDelta)
	const semanticAndPackageRoundtripMatches =
		semanticRoundtripMatches &&
		packageRoundtripMatches &&
		featureRoundtripMatches &&
		editCellValueMatches &&
		editStructuralDeltaMatches
	return {
		status: byteIdentical
			? 'exact-package-match'
			: semanticAndPackageRoundtripMatches
				? 'semantic-roundtrip-pass'
				: hasRoundtripShape
					? semanticRoundtripMatches &&
						packageFingerprintRequired &&
						!hasRoundtripPackageFingerprint
						? 'package-roundtrip-unverified'
						: semanticRoundtripMatches &&
								featureFingerprintRequired &&
								!hasRoundtripFeatureFingerprint
							? 'feature-roundtrip-unverified'
							: semanticRoundtripMatches && !packageRoundtripMatches
								? 'package-roundtrip-mismatch'
								: semanticRoundtripMatches && !featureRoundtripMatches
									? 'feature-roundtrip-mismatch'
									: 'semantic-roundtrip-mismatch'
					: 'rewritten-package-unverified',
		assertions: {
			...observed,
			exactPackageMatch: byteIdentical,
			expectedSheetCount: expected.sheetCount,
			expectedSheetNamesHash: expected.sheetNamesHash,
			expectedCellCount: expected.cellCount,
			expectedPhysicalCellCount: expected.physicalCellCount,
			expectedFormulaCount: expected.formulaCount,
			expectedFirstUsedRange: expected.usedRanges[0] ?? null,
			expectedFirstPhysicalUsedRange: expected.physicalUsedRanges[0] ?? null,
			expectedUsedRangesHash: expected.usedRangesHash,
			expectedSemanticCellRefsHash: expected.semanticCellRefsHash,
			expectedSemanticCellValuesHash: expected.semanticCellValuesHash,
			expectedFormulaTextHash: expected.formulaTextHash,
			expectedPackagePartCount: expectedPackage?.partCount ?? null,
			expectedPackagePartNamesHash: expectedPackage?.partNamesHash ?? null,
			expectedPackageContentTypeCount: expectedPackage?.contentTypeCount ?? null,
			expectedPackageContentTypesHash: expectedPackage?.contentTypesHash ?? null,
			expectedPackageRelationshipCount: expectedPackage?.relationshipCount ?? null,
			expectedPackageRelationshipGraphHash: expectedPackage?.relationshipGraphHash ?? null,
			expectedPreservedPartCount: expectedPackage?.preservedPartCount ?? null,
			expectedPreservedPartNamesHash: expectedPackage?.preservedPartNamesHash ?? null,
			expectedPreservedPartContentHash: expectedPackage?.preservedPartContentHash ?? null,
			expectedTablePartCount: expectedFeature?.tablePartCount ?? null,
			expectedChartPartCount: expectedFeature?.chartPartCount ?? null,
			expectedChartExPartCount: expectedFeature?.chartExPartCount ?? null,
			expectedDrawingPartCount: expectedFeature?.drawingPartCount ?? null,
			expectedVmlDrawingPartCount: expectedFeature?.vmlDrawingPartCount ?? null,
			expectedPivotTablePartCount: expectedFeature?.pivotTablePartCount ?? null,
			expectedPivotCachePartCount: expectedFeature?.pivotCachePartCount ?? null,
			expectedSlicerPartCount: expectedFeature?.slicerPartCount ?? null,
			expectedCommentPartCount: expectedFeature?.commentPartCount ?? null,
			expectedPackageCommentEntryCount: expectedFeature?.packageCommentEntryCount ?? null,
			expectedThreadedCommentPartCount: expectedFeature?.threadedCommentPartCount ?? null,
			expectedMediaPartCount: expectedFeature?.mediaPartCount ?? null,
			expectedExternalLinkPartCount: expectedFeature?.externalLinkPartCount ?? null,
			expectedConnectionPartCount: expectedFeature?.connectionPartCount ?? null,
			expectedCustomXmlPartCount: expectedFeature?.customXmlPartCount ?? null,
			expectedWorksheetCommentCount: expectedFeature?.worksheetCommentCount ?? null,
			expectedWorksheetHyperlinkCount: expectedFeature?.worksheetHyperlinkCount ?? null,
			expectedWorksheetDataValidationCount: expectedFeature?.worksheetDataValidationCount ?? null,
			expectedWorksheetConditionalFormattingCount:
				expectedFeature?.worksheetConditionalFormattingCount ?? null,
			expectedDefinedNameCount: expectedFeature?.definedNameCount ?? null,
			expectedFeaturePartNamesHash: expectedFeature?.featurePartNamesHash ?? null,
			expectedFeatureInventoryHash: expectedFeature?.featureInventoryHash ?? null,
			roundtripSheetCountMatches,
			roundtripSheetNamesHashMatches,
			roundtripCellCountMatches,
			roundtripPhysicalCellCountMatches,
			roundtripFormulaCountMatches,
			roundtripFirstUsedRangeMatches,
			roundtripUsedRangesHashMatches,
			roundtripSemanticCellRefsHashMatches,
			roundtripSemanticCellValuesHashMatches,
			roundtripFormulaTextHashMatches,
			editCellValueMatches,
			editStructuralDeltaMatches,
			semanticRoundtripMatches,
			packageFingerprintRequired,
			hasRoundtripPackageFingerprint,
			roundtripPackagePartCountMatches,
			roundtripPackagePartNamesHashMatches,
			roundtripPackageContentTypeCountMatches,
			roundtripPackageContentTypesHashMatches,
			roundtripPackageRelationshipCountMatches,
			roundtripPackageRelationshipGraphHashMatches,
			packageTopologyMatches,
			roundtripPreservedPartCountMatches,
			roundtripPreservedPartNamesHashMatches,
			roundtripPreservedPartContentHashMatches,
			packageRoundtripMatches,
			featureFingerprintRequired,
			hasRoundtripFeatureFingerprint,
			roundtripTablePartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripTablePartCount === expectedFeature?.tablePartCount,
			roundtripChartPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripChartPartCount === expectedFeature?.chartPartCount,
			roundtripChartExPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripChartExPartCount === expectedFeature?.chartExPartCount,
			roundtripDrawingPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripDrawingPartCount === expectedFeature?.drawingPartCount,
			roundtripVmlDrawingPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripVmlDrawingPartCount === expectedFeature?.vmlDrawingPartCount,
			roundtripPivotTablePartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripPivotTablePartCount === expectedFeature?.pivotTablePartCount,
			roundtripPivotCachePartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripPivotCachePartCount === expectedFeature?.pivotCachePartCount,
			roundtripSlicerPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripSlicerPartCount === expectedFeature?.slicerPartCount,
			roundtripCommentPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripCommentPartCount === expectedFeature?.commentPartCount,
			roundtripPackageCommentEntryCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripPackageCommentEntryCount === expectedFeature?.packageCommentEntryCount,
			roundtripThreadedCommentPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripThreadedCommentPartCount === expectedFeature?.threadedCommentPartCount,
			roundtripMediaPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripMediaPartCount === expectedFeature?.mediaPartCount,
			roundtripExternalLinkPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripExternalLinkPartCount === expectedFeature?.externalLinkPartCount,
			roundtripConnectionPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripConnectionPartCount === expectedFeature?.connectionPartCount,
			roundtripCustomXmlPartCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripCustomXmlPartCount === expectedFeature?.customXmlPartCount,
			roundtripWorksheetCommentCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripWorksheetCommentCount === expectedFeature?.worksheetCommentCount,
			roundtripWorksheetHyperlinkCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripWorksheetHyperlinkCount === expectedFeature?.worksheetHyperlinkCount,
			roundtripWorksheetDataValidationCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripWorksheetDataValidationCount ===
					expectedFeature?.worksheetDataValidationCount,
			roundtripWorksheetConditionalFormattingCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripWorksheetConditionalFormattingCount ===
					expectedFeature?.worksheetConditionalFormattingCount,
			roundtripDefinedNameCountMatches:
				!featureFingerprintRequired ||
				observed.roundtripDefinedNameCount === expectedFeature?.definedNameCount,
			roundtripFeaturePartNamesHashMatches:
				!featureFingerprintRequired ||
				observed.roundtripFeaturePartNamesHash === expectedFeature?.featurePartNamesHash,
			roundtripFeatureInventoryHashMatches:
				!featureFingerprintRequired ||
				observed.roundtripFeatureInventoryHash === expectedFeature?.featureInventoryHash,
			featureRoundtripMatches,
		},
	}
}

interface CompetitiveCase {
	readonly name: string
	readonly library: string
	readonly category: CompetitiveCategory
	readonly executionScope?: 'in-process' | 'external-process'
	readonly runnerProvenance?: RunnerProvenance
	readonly timingModel?: string
	readonly validationModel?: string
	readonly memoryModel?: string
	readonly capabilities?: {
		readonly xlsmRoundtrip?: boolean
		readonly internalTiming?: boolean
		readonly valueOnlyRead?: boolean
		readonly metadataOnlyRead?: boolean
		readonly writeFormulas?: boolean
		readonly writeTables?: boolean
		readonly writeRichMetadata?: boolean
	}
	run(
		target: WorkbookTarget,
	): Promise<{ assertions?: Record<string, string | number | boolean | null> }>
	runMeasured?(target: WorkbookTarget): Promise<{
		durationMs: number
		assertions?: Record<string, string | number | boolean | null>
	}>
	runBatched?(
		target: WorkbookTarget,
		repeat: number,
		warmup: number,
	): Promise<{
		assertions?: Record<string, string | number | boolean | null>
		assertionsBySample?: readonly Record<string, string | number | boolean | null>[]
		samples?: readonly MetricSample[]
	}>
}

interface RunnerProvenance {
	readonly adapterVersion?: string
	readonly libraryVersion?: string
	readonly runtime?: string
}

export interface ExternalRunnerSpec {
	readonly name: string
	readonly command: readonly string[]
	readonly categories?: readonly ExternalRunnerCategory[]
	readonly workloads?: readonly string[]
	readonly adapterVersion?: string
	readonly libraryVersion?: string
	readonly runtime?: string
	readonly timingModel?: string
	readonly validationModel?: string
	readonly memoryModel?: string
	readonly installHint?: string
	readonly licenseGate?: {
		readonly env: string
		readonly value?: string
		readonly reason?: string
	}
	readonly capabilities?: {
		readonly xlsmRoundtrip?: boolean
		readonly internalTiming?: boolean
		readonly valueOnlyRead?: boolean
		readonly metadataOnlyRead?: boolean
		readonly writeFormulas?: boolean
		readonly writeTables?: boolean
		readonly writeRichMetadata?: boolean
		readonly finalValidation?: boolean
	}
}

type DimensionValue = string | number | boolean

export function benchmarkProvenanceDimensions(
	assertions: Record<string, string | number | boolean | null> | undefined,
	provenance?: RunnerProvenance,
): Record<string, DimensionValue> {
	const dimensions: Record<string, DimensionValue> = {}
	if (provenance?.adapterVersion) dimensions.runnerAdapterVersion = provenance.adapterVersion
	if (provenance?.runtime) dimensions.runnerRuntime = provenance.runtime
	if (provenance?.libraryVersion && provenance.libraryVersion !== 'reported-by-runner') {
		dimensions.runnerManifestLibraryVersion = provenance.libraryVersion
	}
	if (!assertions) return dimensions
	for (const [key, value] of Object.entries(assertions)) {
		if (value === null) continue
		if (key === 'runnerEngine' || /^[A-Za-z][A-Za-z0-9]*Version$/.test(key)) {
			dimensions[key] = value
		}
	}
	return dimensions
}

interface MetricSample {
	readonly durationMs: number
	readonly throughputPerSec?: number
	readonly rssDeltaBytes?: number
	readonly retainedRssDeltaBytes?: number
	readonly rssAfterBytes?: number
	readonly rssAfterGcBytes?: number
	readonly peakRssBytes?: number
	readonly heapDeltaBytes?: number
	readonly heapUsedBytes?: number
	readonly heapTotalBytes?: number
	readonly heapAfterGcBytes?: number
}

type LibraryAllowlist = ReadonlySet<string> | undefined

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	return index >= 0 ? process.argv[index + 1] : undefined
}

function readFlagValues(name: string): string[] {
	const values: string[] = []
	for (let i = 2; i < process.argv.length; i++) {
		if (process.argv[i] !== name) continue
		const value = process.argv[i + 1]
		if (value !== undefined && !value.startsWith('--')) values.push(value)
	}
	return values
}

function hasFlag(name: string): boolean {
	return process.argv.includes(name)
}

function competitorFilter(): 'all' | 'in-process' | 'external' {
	const raw = readFlag('--competitor')
	if (raw === undefined || raw === 'all') return 'all'
	if (raw === 'in-process' || raw === 'external') return raw
	throw new Error('--competitor must be one of: all, in-process, external')
}

export function parseLibraryAllowlist(raw: string | undefined): LibraryAllowlist {
	if (raw === undefined || raw.trim() === '') return undefined
	const libraries = raw
		.split(',')
		.map((library) => library.trim())
		.filter(Boolean)
	if (libraries.length === 0) return undefined
	return new Set(libraries)
}

export function libraryAllowed(library: string, allowlist: LibraryAllowlist): boolean {
	return allowlist === undefined || allowlist.has(library)
}

function readPositiveIntFlag(name: string, fallback: number): number {
	const raw = readFlag(name)
	if (raw === undefined) return fallback
	const value = Number.parseInt(raw, 10)
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function readNonNegativeIntFlag(name: string, fallback: number): number {
	const raw = readFlag(name)
	if (raw === undefined) return fallback
	const value = Number.parseInt(raw, 10)
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

function positionalArgs(): string[] {
	const flagsWithValues = new Set([
		'--repeat',
		'--warmup',
		'--runner-manifest',
		'--category',
		'--corpus-manifest',
		'--corpus-root',
		'--corpus-file',
		'--corpus-tier',
		'--tag',
		'--risk',
		'--assertion-class',
		'--competitor',
		'--libraries',
		'--expected-shape-sidecar',
	])
	const args: string[] = []
	for (let i = 2; i < process.argv.length; i++) {
		const arg = process.argv[i]
		if (!arg) continue
		if (arg.startsWith('--')) {
			if (flagsWithValues.has(arg)) i++
			continue
		}
		args.push(arg)
	}
	return args
}

function hasCorpusTargetMode(): boolean {
	return (
		hasFlag('--corpus-manifest') ||
		hasFlag('--corpus-root') ||
		hasFlag('--corpus-file') ||
		hasFlag('--corpus-tier') ||
		hasFlag('--tag') ||
		hasFlag('--risk') ||
		hasFlag('--assertion-class') ||
		hasFlag('--vendorable-only')
	)
}

function targetMode(): 'explicit' | 'quick' | 'full-corpus' | 'corpus' {
	if (positionalArgs().length > 0) return 'explicit'
	if (hasCorpusTargetMode()) return 'corpus'
	return hasFlag('--full-corpus') ? 'full-corpus' : 'quick'
}

function readCorpusSelection(): CorpusSelection {
	const file = readFlag('--corpus-file')
	const tags = readFlagValues('--tag')
	const tiers = readFlagValues('--corpus-tier') as CorpusBenchmarkTier[]
	const risks = readFlagValues('--risk') as CorpusRiskClass[]
	const assertionClasses = readFlagValues('--assertion-class') as CorpusAssertionClass[]
	return {
		...(file ? { file } : {}),
		...(tags.length > 0 ? { tags } : {}),
		...(tiers.length > 0 ? { tiers } : {}),
		...(risks.length > 0 ? { risks } : {}),
		...(assertionClasses.length > 0 ? { assertionClasses } : {}),
		...(hasFlag('--vendorable-only') ? { vendorableOnly: true } : {}),
	}
}

function readCategoryFilter(): CompetitiveCategory | undefined {
	const raw = readFlag('--category')
	if (raw === undefined) return undefined
	if (raw === 'read' || raw === 'roundtrip' || raw === 'edit-roundtrip') return raw
	throw new Error('--category must be "read", "roundtrip", or "edit-roundtrip"')
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		/* best effort */
	}
}

function getRssBytes(): number | undefined {
	try {
		return process.memoryUsage.rss()
	} catch {
		return undefined
	}
}

function observedPeakRssBytes(values: readonly (number | undefined)[]): number | undefined {
	const defined = values.filter((value): value is number => value !== undefined)
	return defined.length > 0 ? Math.max(...defined) : undefined
}

function sha256(bytes: Uint8Array): string {
	return createHash('sha256').update(bytes).digest('hex')
}

function formatRange(sheet: string, range: RangeRef): string {
	return `${sheet}!${indexToColumn(range.start.col)}${range.start.row + 1}:${indexToColumn(range.end.col)}${range.end.row + 1}`
}

function formatCellRef(sheet: string, row: number, col: number): string {
	return `${sheet}!${indexToColumn(col)}${row + 1}`
}

function formatLocalCellRef(row: number, col: number): string {
	return `${indexToColumn(col)}${row + 1}`
}

function hashLines(lines: readonly string[]): string {
	const hash = createHash('sha256')
	for (const line of [...lines].sort()) {
		hash.update(`${line.length}:`)
		hash.update(line)
		hash.update('\n')
	}
	return hash.digest('hex')
}

function prefixedKey(prefix: string, key: string): string {
	if (!prefix) return key
	return `${prefix}${key[0]?.toUpperCase() ?? ''}${key.slice(1)}`
}

function featureAssertions(
	prefix: string,
	summary: WorkbookFeatureSummary | undefined,
): Record<string, string | number> {
	if (!summary) return {}
	return {
		[prefixedKey(prefix, 'tablePartCount')]: summary.tablePartCount,
		[prefixedKey(prefix, 'chartPartCount')]: summary.chartPartCount,
		[prefixedKey(prefix, 'chartExPartCount')]: summary.chartExPartCount,
		[prefixedKey(prefix, 'drawingPartCount')]: summary.drawingPartCount,
		[prefixedKey(prefix, 'vmlDrawingPartCount')]: summary.vmlDrawingPartCount,
		[prefixedKey(prefix, 'pivotTablePartCount')]: summary.pivotTablePartCount,
		[prefixedKey(prefix, 'pivotCachePartCount')]: summary.pivotCachePartCount,
		[prefixedKey(prefix, 'slicerPartCount')]: summary.slicerPartCount,
		[prefixedKey(prefix, 'commentPartCount')]: summary.commentPartCount,
		[prefixedKey(prefix, 'packageCommentEntryCount')]: summary.packageCommentEntryCount,
		[prefixedKey(prefix, 'threadedCommentPartCount')]: summary.threadedCommentPartCount,
		[prefixedKey(prefix, 'mediaPartCount')]: summary.mediaPartCount,
		[prefixedKey(prefix, 'externalLinkPartCount')]: summary.externalLinkPartCount,
		[prefixedKey(prefix, 'connectionPartCount')]: summary.connectionPartCount,
		[prefixedKey(prefix, 'customXmlPartCount')]: summary.customXmlPartCount,
		[prefixedKey(prefix, 'worksheetCommentCount')]: summary.worksheetCommentCount,
		[prefixedKey(prefix, 'worksheetHyperlinkCount')]: summary.worksheetHyperlinkCount,
		[prefixedKey(prefix, 'worksheetDataValidationCount')]: summary.worksheetDataValidationCount,
		[prefixedKey(prefix, 'worksheetConditionalFormattingCount')]:
			summary.worksheetConditionalFormattingCount,
		[prefixedKey(prefix, 'definedNameCount')]: summary.definedNameCount,
		[prefixedKey(prefix, 'featurePartNamesHash')]: summary.featurePartNamesHash,
		[prefixedKey(prefix, 'featureInventoryHash')]: summary.featureInventoryHash,
	}
}

export function extractWorkbookPackageFingerprint(bytes: Uint8Array): WorkbookPackageFingerprint {
	const archive = extractZip(bytes)
	const partPaths = [...archive.entries()]
		.map((entry) => entry.path)
		.filter((path) => !path.endsWith('/') && !isIgnorablePackageEntry(path))
		.sort()
	const contentTypeLines = contentTypeFingerprintLines(archive.readText('[Content_Types].xml'))
	const relationshipLines: string[] = []
	for (const relsPath of partPaths.filter(isRelationshipPart)) {
		const relsXml = archive.readText(relsPath)
		if (!relsXml) continue
		const sourcePart = sourcePartForRelationships(relsPath)
		const targetModes = relationshipTargetModes(relsXml)
		for (const relationship of parseRelationships(relsXml)) {
			const targetMode = targetModes.get(relationship.id) ?? 'Internal'
			const resolvedTarget =
				targetMode === 'External'
					? relationship.target
					: resolvePath(sourcePart, relationship.target)
			relationshipLines.push(
				[
					sourcePart || '/',
					relationship.type,
					targetMode,
					relationship.target,
					resolvedTarget,
				].join('\t'),
			)
		}
	}
	const preservedPartPaths = partPaths.filter(isPreservedNonCellPart)
	const preservedPartContentLines = preservedPartPaths.map((path) => {
		const partBytes = archive.readBytes(path)
		return `${path}\t${partBytes ? sha256(partBytes) : 'missing'}`
	})
	return {
		partCount: partPaths.length,
		partNamesHash: hashLines(partPaths),
		contentTypeCount: contentTypeLines.length,
		contentTypesHash: hashLines(contentTypeLines),
		relationshipCount: relationshipLines.length,
		relationshipGraphHash: hashLines(relationshipLines),
		preservedPartCount: preservedPartPaths.length,
		preservedPartNamesHash: hashLines(preservedPartPaths),
		preservedPartContentHash: hashLines(preservedPartContentLines),
	}
}

export function extractWorkbookFeatureSummary(bytes: Uint8Array): WorkbookFeatureSummary {
	const archive = extractZip(bytes)
	const partPaths = [...archive.entries()]
		.map((entry) => entry.path)
		.filter((path) => !path.endsWith('/') && !isIgnorablePackageEntry(path))
		.sort()
	const workbookXml = archive.readText('xl/workbook.xml') ?? ''
	const workbookRelsXml = archive.readText('xl/_rels/workbook.xml.rels') ?? ''
	const workbookInfo = workbookXml ? parseWorkbookXml(workbookXml) : { sheets: [] }
	const workbookRels = workbookRelsXml ? parseRelationships(workbookRelsXml) : []
	const worksheetPaths = workbookInfo.sheets
		.map((sheet, index) =>
			worksheetPathForWorkbookSheet(archive, 'xl/workbook.xml', workbookRels, sheet, index),
		)
		.filter((path): path is string => path !== undefined)
	const featurePartLines: string[] = []
	for (const path of partPaths) {
		const kind = classifyFeaturePart(path)
		if (kind) featurePartLines.push(`${kind}\t${path}`)
	}
	const worksheetFeatureLines: string[] = []
	for (const path of worksheetPaths) {
		const xml = archive.readText(path)
		if (!xml) continue
		for (const entry of worksheetFeatureEntries(path, xml)) worksheetFeatureLines.push(entry)
	}
	const definedNameLines = definedNameEntries(workbookXml)
	const sharedStringFeatureLines = sharedStringFeatureEntries(
		archive.readText('xl/sharedStrings.xml') ?? '',
	)
	const packageCommentFeatureLines = packageCommentFeatureEntries(archive, partPaths)
	const worksheetCommentFeatureLines = worksheetCommentFeatureEntries(archive, worksheetPaths)
	const calcChainFeatureLines = calcChainFeatureEntries(archive.readText('xl/calcChain.xml') ?? '')
	const inventoryLines = [
		...featurePartLines,
		...worksheetFeatureLines,
		...definedNameLines,
		...sharedStringFeatureLines,
		...packageCommentFeatureLines,
		...worksheetCommentFeatureLines,
		...calcChainFeatureLines,
	]
	return {
		tablePartCount: countPartPaths(partPaths, (path) => /^xl\/tables\/.+\.xml$/.test(path)),
		chartPartCount: countPartPaths(partPaths, (path) => /^xl\/charts\/.+\.xml$/.test(path)),
		chartExPartCount: countPartPaths(partPaths, (path) => /^xl\/chartEx\/.+\.xml$/.test(path)),
		drawingPartCount: countPartPaths(partPaths, (path) => /^xl\/drawings\/.+\.xml$/.test(path)),
		vmlDrawingPartCount: countPartPaths(partPaths, (path) => /^xl\/drawings\/.+\.vml$/.test(path)),
		pivotTablePartCount: countPartPaths(partPaths, (path) =>
			/^xl\/pivotTables\/.+\.xml$/.test(path),
		),
		pivotCachePartCount: countPartPaths(partPaths, (path) =>
			/^xl\/pivotCache\/.+\.xml$/.test(path),
		),
		slicerPartCount: countPartPaths(partPaths, (path) =>
			/^xl\/(?:slicers|slicerCaches)\/.+\.xml$/.test(path),
		),
		commentPartCount: countPartPaths(partPaths, (path) =>
			/^xl\/(?:comments\d*|comments\/.+)\.xml$/.test(path),
		),
		packageCommentEntryCount: packageCommentFeatureLines.length,
		threadedCommentPartCount: countPartPaths(partPaths, (path) =>
			/^xl\/threadedComments\/.+\.xml$/.test(path),
		),
		mediaPartCount: countPartPaths(partPaths, (path) => /^xl\/media\/.+/.test(path)),
		externalLinkPartCount: countPartPaths(partPaths, (path) =>
			/^xl\/externalLinks\/.+\.xml$/.test(path),
		),
		connectionPartCount: countPartPaths(partPaths, (path) => path === 'xl/connections.xml'),
		customXmlPartCount: countPartPaths(partPaths, (path) => /^customXml\/.+/.test(path)),
		worksheetCommentCount: worksheetCommentFeatureLines.length,
		worksheetHyperlinkCount: worksheetFeatureLines.filter((line) =>
			line.startsWith('worksheet-hyperlink\t'),
		).length,
		worksheetDataValidationCount: worksheetFeatureLines.filter((line) =>
			line.startsWith('worksheet-data-validation\t'),
		).length,
		worksheetConditionalFormattingCount: worksheetFeatureLines.filter((line) =>
			line.startsWith('worksheet-conditional-formatting\t'),
		).length,
		definedNameCount: definedNameLines.length,
		featurePartNamesHash: hashLines(featurePartLines),
		featureInventoryHash: hashLines(inventoryLines),
	}
}

function worksheetPathForWorkbookSheet(
	archive: ReturnType<typeof extractZip>,
	workbookPath: string,
	workbookRels: readonly { readonly id: string; readonly type: string; readonly target: string }[],
	sheet: { readonly rId: string },
	index: number,
): string | undefined {
	const rel = workbookRels.find((entry) => entry.id === sheet.rId)
	if (rel && rel.type !== REL_WORKSHEET) return undefined
	const relPath = rel ? resolvePath(workbookPath, rel.target) : undefined
	if (relPath && archive.has(relPath)) return relPath
	return firstExistingPart(archive, [
		`xl/worksheets/sheet${index + 1}.xml`,
		`xl/worksheets/Sheet${index + 1}.xml`,
	])
}

function sharedStringsPathForWorkbook(
	archive: ReturnType<typeof extractZip>,
	workbookPath: string,
	workbookRels: readonly { readonly type: string; readonly target: string }[],
): string | undefined {
	const rel = workbookRels.find((entry) => entry.type === REL_SHARED_STRINGS)
	const relPath = rel ? resolvePath(workbookPath, rel.target) : undefined
	if (relPath && archive.has(relPath)) return relPath
	return archive.has('xl/sharedStrings.xml') ? 'xl/sharedStrings.xml' : undefined
}

function firstExistingPart(
	archive: ReturnType<typeof extractZip>,
	paths: readonly string[],
): string | undefined {
	return paths.find((path) => archive.has(path))
}

function countPartPaths(
	partPaths: readonly string[],
	predicate: (path: string) => boolean,
): number {
	let count = 0
	for (const path of partPaths) {
		if (predicate(path)) count++
	}
	return count
}

function classifyFeaturePart(path: string): string | null {
	if (/^xl\/tables\/.+\.xml$/.test(path)) return 'table-part'
	if (/^xl\/charts\/.+\.xml$/.test(path)) return 'chart-part'
	if (/^xl\/chartEx\/.+\.xml$/.test(path)) return 'chart-ex-part'
	if (/^xl\/drawings\/.+\.xml$/.test(path)) return 'drawing-part'
	if (/^xl\/drawings\/.+\.vml$/.test(path)) return 'vml-drawing-part'
	if (/^xl\/pivotTables\/.+\.xml$/.test(path)) return 'pivot-table-part'
	if (/^xl\/pivotCache\/.+\.xml$/.test(path)) return 'pivot-cache-part'
	if (/^xl\/(?:slicers|slicerCaches)\/.+\.xml$/.test(path)) return 'slicer-part'
	if (/^xl\/(?:comments\d*|comments\/.+)\.xml$/.test(path)) return 'comment-part'
	if (/^xl\/threadedComments\/.+\.xml$/.test(path)) return 'threaded-comment-part'
	if (/^xl\/media\/.+/.test(path)) return 'media-part'
	if (/^xl\/externalLinks\/.+\.xml$/.test(path)) return 'external-link-part'
	if (path === 'xl/connections.xml') return 'connection-part'
	if (path === 'xl/calcChain.xml') return 'calc-chain-part'
	if (/^customXml\/.+/.test(path)) return 'custom-xml-part'
	return null
}

const XML_NAME_PREFIX = String.raw`(?:[A-Za-z_][\w.-]*:)?`

function openTagRegex(localName: string): RegExp {
	return new RegExp(String.raw`<\s*${XML_NAME_PREFIX}${localName}\b([^>]*)>`, 'g')
}

function closeTagRegex(localName: string): RegExp {
	return new RegExp(String.raw`<\/\s*${XML_NAME_PREFIX}${localName}\s*>`, 'g')
}

function elementBodyRegex(localName: string): RegExp {
	return new RegExp(
		String.raw`<\s*${XML_NAME_PREFIX}${localName}\b[^>]*>([\s\S]*?)<\/\s*${XML_NAME_PREFIX}${localName}\s*>`,
		'g',
	)
}

function hasOpenTag(xml: string, localName: string): boolean {
	return openTagRegex(localName).test(xml)
}

function isSelfClosingTag(openTag: string): boolean {
	return /\/\s*>$/.test(openTag)
}

function findCloseTagIndex(xml: string, localName: string, start: number): number {
	const pattern = closeTagRegex(localName)
	pattern.lastIndex = start
	return pattern.exec(xml)?.index ?? -1
}

function worksheetFeatureEntries(path: string, xml: string): readonly string[] {
	const lines: string[] = []
	for (const match of xml.matchAll(openTagRegex('mergeCell'))) {
		lines.push(`worksheet-merge-cell\t${path}\t${featureRef(match[1] ?? '')}`)
	}
	for (const match of xml.matchAll(openTagRegex('autoFilter'))) {
		lines.push(`worksheet-auto-filter\t${path}\t${featureRef(match[1] ?? '')}`)
	}
	for (const match of xml.matchAll(openTagRegex('sortState'))) {
		lines.push(`worksheet-sort-state\t${path}\t${canonicalAttributes(match[1] ?? '')}`)
	}
	for (const match of xml.matchAll(openTagRegex('sheetProtection'))) {
		lines.push(`worksheet-sheet-protection\t${path}\t${canonicalAttributes(match[1] ?? '')}`)
	}
	for (const match of xml.matchAll(openTagRegex('sheetView'))) {
		lines.push(`worksheet-sheet-view\t${path}\t${canonicalAttributes(match[1] ?? '')}`)
	}
	for (const match of xml.matchAll(openTagRegex('hyperlink'))) {
		lines.push(`worksheet-hyperlink\t${path}\t${featureRef(match[1] ?? '')}`)
	}
	for (const match of xml.matchAll(openTagRegex('dataValidation'))) {
		lines.push(`worksheet-data-validation\t${path}\t${featureSqref(match[1] ?? '')}`)
	}
	for (const match of xml.matchAll(openTagRegex('conditionalFormatting'))) {
		lines.push(`worksheet-conditional-formatting\t${path}\t${featureSqref(match[1] ?? '')}`)
	}
	return lines
}

function sharedStringFeatureEntries(xml: string): readonly string[] {
	const lines: string[] = []
	let index = 0
	for (const match of xml.matchAll(elementBodyRegex('si'))) {
		const inner = match[1] ?? ''
		if (hasOpenTag(inner, 'r'))
			lines.push(`shared-string-rich-text\t${index}\t${sha256Text(inner)}`)
		index += 1
	}
	return lines
}

function calcChainFeatureEntries(xml: string): readonly string[] {
	const lines: string[] = []
	for (const match of xml.matchAll(openTagRegex('c'))) {
		lines.push(`calc-chain-cell\t${canonicalAttributes(match[1] ?? '')}`)
	}
	return lines
}

function packageCommentFeatureEntries(
	archive: ReturnType<typeof extractZip>,
	partPaths: readonly string[],
): readonly string[] {
	const lines: string[] = []
	for (const path of partPaths.filter((entry) =>
		/^xl\/(?:comments\d*|comments\/.+)\.xml$/.test(entry),
	)) {
		const xml = archive.readText(path)
		if (!xml) continue
		for (const match of xml.matchAll(openTagRegex('comment'))) {
			const attrs = parseXmlAttributes(match[1] ?? '')
			const ref = attrs.get('ref')
			if (ref) lines.push(`package-comment\t${path}\t${ref}`)
		}
	}
	return lines
}

function worksheetCommentFeatureEntries(
	archive: ReturnType<typeof extractZip>,
	worksheetPaths: readonly string[],
): readonly string[] {
	const lines: string[] = []
	for (const sheetPath of worksheetPaths) {
		const relsXml = archive.readText(relationshipsPartPath(sheetPath))
		if (!relsXml) continue
		for (const relationship of parseRelationships(relsXml)) {
			if (relationship.type !== REL_COMMENTS || relationship.targetMode === 'External') continue
			const commentsPath = resolvePath(sheetPath, relationship.target)
			const commentsXml = archive.readText(commentsPath)
			if (!commentsXml) continue
			for (const match of commentsXml.matchAll(openTagRegex('comment'))) {
				const attrs = parseXmlAttributes(match[1] ?? '')
				const ref = attrs.get('ref')
				if (ref) lines.push(`worksheet-comment\t${sheetPath}\t${commentsPath}\t${ref}`)
			}
		}
	}
	return lines
}

function relationshipsPartPath(sourcePart: string): string {
	const slash = sourcePart.lastIndexOf('/')
	const directory = slash === -1 ? '' : sourcePart.slice(0, slash + 1)
	const name = slash === -1 ? sourcePart : sourcePart.slice(slash + 1)
	return `${directory}_rels/${name}.rels`
}

function definedNameEntries(workbookXml: string): readonly string[] {
	const lines: string[] = []
	const definedNameRe = new RegExp(
		String.raw`<\s*${XML_NAME_PREFIX}definedName\b([^>]*)>([\s\S]*?)<\/\s*${XML_NAME_PREFIX}definedName\s*>`,
		'g',
	)
	for (const match of workbookXml.matchAll(definedNameRe)) {
		const attrs = parseXmlAttributes(match[1] ?? '')
		const name = attrs.get('name') ?? ''
		const localSheetId = attrs.get('localSheetId') ?? ''
		lines.push(`defined-name\t${name}\t${localSheetId}\t${decodeXmlText(match[2] ?? '')}`)
	}
	return lines
}

function featureRef(attrsText: string): string {
	return parseXmlAttributes(attrsText).get('ref') ?? ''
}

function featureSqref(attrsText: string): string {
	return parseXmlAttributes(attrsText).get('sqref') ?? ''
}

function canonicalAttributes(attrsText: string): string {
	return [...parseXmlAttributes(attrsText).entries()]
		.sort(([left], [right]) => (left < right ? -1 : left > right ? 1 : 0))
		.map(([key, value]) => `${key}=${value}`)
		.join(';')
}

function contentTypeFingerprintLines(xml: string | undefined): readonly string[] {
	if (!xml) return []
	const lines: string[] = []
	for (const match of xml.matchAll(/<(Default|Override)\b([^>]*)\/?>/g)) {
		const kind = match[1]
		const attrs = parseXmlAttributes(match[2] ?? '')
		if (kind === 'Default') {
			const extension = attrs.get('Extension')
			const contentType = attrs.get('ContentType')
			if (extension && contentType) lines.push(`Default\t${extension}\t${contentType}`)
		} else {
			const partName = attrs.get('PartName')
			const contentType = attrs.get('ContentType')
			if (partName && contentType) {
				lines.push(`Override\t${partName.replace(/^\//, '')}\t${contentType}`)
			}
		}
	}
	return lines
}

function relationshipTargetModes(xml: string): Map<string, string> {
	const modes = new Map<string, string>()
	for (const match of xml.matchAll(/<Relationship\b([^>]*)\/>/g)) {
		const attrs = parseXmlAttributes(match[1] ?? '')
		const id = attrs.get('Id')
		if (id) modes.set(id, attrs.get('TargetMode') ?? 'Internal')
	}
	return modes
}

function parseXmlAttributes(attrs: string): Map<string, string> {
	const values = new Map<string, string>()
	for (const match of attrs.matchAll(/([A-Za-z_][\w:.-]*)="([^"]*)"/g)) {
		const key = match[1]
		const value = match[2]
		if (key && value !== undefined) values.set(key, decodeXmlText(value))
	}
	return values
}

function sha256Text(text: string): string {
	return createHash('sha256').update(text).digest('hex')
}

function isRelationshipPart(path: string): boolean {
	return path === '_rels/.rels' || path.endsWith('.rels')
}

function isIgnorablePackageEntry(path: string): boolean {
	return path === '.DS_Store' || path.endsWith('/.DS_Store') || path.startsWith('__MACOSX/')
}

function sourcePartForRelationships(relsPath: string): string {
	if (relsPath === '_rels/.rels') return ''
	const marker = '/_rels/'
	const index = relsPath.lastIndexOf(marker)
	if (index < 0 || !relsPath.endsWith('.rels')) return ''
	return `${relsPath.slice(0, index)}/${relsPath.slice(index + marker.length, -'.rels'.length)}`
}

function isPreservedNonCellPart(path: string): boolean {
	if (
		path === '[Content_Types].xml' ||
		path === '_rels/.rels' ||
		path.endsWith('.rels') ||
		path === 'xl/workbook.xml' ||
		path === 'xl/sharedStrings.xml' ||
		path === 'xl/calcChain.xml' ||
		/^xl\/worksheets\/sheet\d+\.xml$/i.test(path) ||
		/^docProps\/(?:app|core)\.xml$/.test(path)
	) {
		return false
	}
	return true
}

function canonicalNumber(value: number): string {
	return Object.is(value, -0) ? '0' : String(value)
}

function serializeCellValue(value: CellValue): string {
	const scalar = topLeftScalar(value)
	switch (scalar.kind) {
		case 'empty':
			return 'empty'
		case 'number':
			return `n:${canonicalNumber(scalar.value)}`
		case 'date':
			return `n:${canonicalNumber(scalar.serial)}`
		case 'string':
			return `s:${scalar.value}`
		case 'richText':
			return `s:${scalar.runs.map((run) => run.text).join('')}`
		case 'boolean':
			return `b:${scalar.value ? 'true' : 'false'}`
		case 'error':
			return `e:${scalar.value}`
	}
}

function serializeUnknownScalarValue(value: unknown): string {
	if (value === undefined || value === null || value === '') return 'empty'
	if (typeof value === 'number') return `n:${canonicalNumber(value)}`
	if (typeof value === 'boolean') return `b:${value ? 'true' : 'false'}`
	if (typeof value === 'string') return `s:${value}`
	if (value instanceof Date) return `s:${value.toISOString()}`
	if (Array.isArray(value)) return serializeUnknownScalarValue(value[0])
	if (typeof value !== 'object') return `s:${String(value)}`
	const record = value as Record<string, unknown>
	if (typeof record.text === 'string') return `s:${record.text}`
	if (Array.isArray(record.richText)) {
		return `s:${record.richText
			.map((run) =>
				typeof run === 'object' && run !== null && 'text' in run
					? String((run as { text?: unknown }).text ?? '')
					: '',
			)
			.join('')}`
	}
	if ('result' in record) return serializeUnknownScalarValue(record.result)
	if ('hyperlink' in record && typeof record.text === 'string') return `s:${record.text}`
	if ('error' in record && typeof record.error === 'string') return `e:${record.error}`
	return `s:${String(value)}`
}

export function summarizeAscendWorkbook(workbook: AscendWorkbookLike): WorkbookShapeSummary & {
	readonly compatibility: string
} {
	const model = workbook.getWorkbookModel()
	const sheetEntries = workbookSheetEntriesForSummary(model)
	let cellCount = 0
	let formulaCount = 0
	const sheetNames = sheetEntries.map((entry) => entry.name)
	const usedRanges: string[] = []
	const semanticCellRefs: string[] = []
	const semanticCellValues: string[] = []
	const formulaTexts: string[] = []
	for (const entry of sheetEntries) {
		const sheet = entry.sheet
		if (!sheet) {
			usedRanges.push(`${entry.name}!empty`)
			continue
		}
		cellCount += sheet.cells.cellCount()
		formulaCount += sheet.cells.formulaCellCount()
		const usedRange = sheet.cells.usedRange()
		usedRanges.push(usedRange ? formatRange(sheet.name, usedRange) : `${sheet.name}!empty`)
		for (const [row, col, cell] of sheet.cells.iterate()) {
			const ref = formatCellRef(sheet.name, row, col)
			semanticCellRefs.push(ref)
			semanticCellValues.push(`${ref}\t${serializeCellValue(cell.value)}`)
			if (cell.formula) {
				formulaTexts.push(`${ref}=${storedFormulaText(sheet, row, col, cell.formula)}`)
			}
		}
	}
	return {
		sheetNames,
		sheetCount: sheetEntries.length,
		cellCount,
		physicalCellCount: null,
		formulaCount,
		usedRanges,
		physicalUsedRanges: [],
		sheetNamesHash: hashLines(sheetNames.map((name, index) => `${index}:${name}`)),
		usedRangesHash: hashLines(usedRanges),
		physicalUsedRangesHash: hashLines([]),
		semanticCellRefsHash: hashLines(semanticCellRefs),
		semanticCellValuesHash: hashLines(semanticCellValues),
		formulaTextHash: hashLines(formulaTexts),
		compatibility: workbook.report.status,
	}
}

function storedFormulaText(
	sheet: { readonly storedFormulaText?: ReadonlyMap<string, string> },
	row: number,
	col: number,
	formula: string,
): string {
	return sheet.storedFormulaText?.get(`${row}:${col}`) ?? formula
}

function extractExpectedWorkbookShape(bytes: Uint8Array): WorkbookShapeSummary {
	const archive = extractZip(bytes)
	const workbookXml = archive.readText('xl/workbook.xml')
	const workbookRelsXml = archive.readText('xl/_rels/workbook.xml.rels')
	if (!workbookXml || !workbookRelsXml) {
		throw new Error('Workbook is missing xl/workbook.xml or xl/_rels/workbook.xml.rels')
	}
	const workbookInfo = parseWorkbookXml(workbookXml)
	const workbookRels = parseRelationships(workbookRelsXml)
	const worksheetRels = new Map(
		workbookInfo.sheets
			.map((sheet, index) => [
				sheet.rId,
				worksheetPathForWorkbookSheet(archive, 'xl/workbook.xml', workbookRels, sheet, index),
			])
			.filter((entry): entry is [string, string] => entry[1] !== undefined),
	)
	const sharedStringsPath = sharedStringsPathForWorkbook(archive, 'xl/workbook.xml', workbookRels)
	const sharedStringsXml = sharedStringsPath ? archive.readText(sharedStringsPath) : undefined
	const sharedStrings = sharedStringsXml
		? parseSharedStrings(sharedStringsXml, { lazy: true })
		: emptySharedStrings()
	let cellCount = 0
	let physicalCellCount = 0
	let formulaCount = 0
	const sheetNames = workbookInfo.sheets.map((sheet) => sheet.name)
	const usedRanges: string[] = []
	const physicalUsedRanges: string[] = []
	const semanticCellRefs: string[] = []
	const semanticCellValues: string[] = []
	const formulaTexts: string[] = []
	for (const sheet of workbookInfo.sheets) {
		const partPath = worksheetRels.get(sheet.rId)
		const xml = partPath ? archive.readText(partPath) : undefined
		const sheetShape = xml ? scanSheetShapeXml(xml, sharedStrings) : null
		cellCount += sheetShape?.cellCount ?? 0
		physicalCellCount += sheetShape?.physicalCellCount ?? 0
		formulaCount += sheetShape?.formulaCount ?? 0
		usedRanges.push(
			sheetShape?.usedRange ? `${sheet.name}!${sheetShape.usedRange}` : `${sheet.name}!empty`,
		)
		physicalUsedRanges.push(
			sheetShape?.physicalUsedRange
				? `${sheet.name}!${sheetShape.physicalUsedRange}`
				: `${sheet.name}!empty`,
		)
		if (sheetShape) {
			for (const ref of sheetShape.semanticCellRefs) semanticCellRefs.push(`${sheet.name}!${ref}`)
			for (const value of sheetShape.semanticCellValues) {
				semanticCellValues.push(`${sheet.name}!${value}`)
			}
			for (const entry of sheetShape.formulaTexts) formulaTexts.push(`${sheet.name}!${entry}`)
		}
	}
	return {
		sheetNames,
		sheetCount: workbookInfo.sheets.length,
		cellCount,
		physicalCellCount,
		formulaCount,
		usedRanges,
		physicalUsedRanges,
		sheetNamesHash: hashLines(sheetNames.map((name, index) => `${index}:${name}`)),
		usedRangesHash: hashLines(usedRanges),
		physicalUsedRangesHash: hashLines(physicalUsedRanges),
		semanticCellRefsHash: hashLines(semanticCellRefs),
		semanticCellValuesHash: hashLines(semanticCellValues),
		formulaTextHash: hashLines(formulaTexts),
		packageFingerprint: extractWorkbookPackageFingerprint(bytes),
		featureSummary: extractWorkbookFeatureSummary(bytes),
	}
}

function findFirstEditableScalarCell(bytes: Uint8Array): {
	sheetName: string
	ref: string
	valueType: EditTarget['valueType']
	value: string | number | boolean
} | null {
	const archive = extractZip(bytes)
	const workbookXml = archive.readText('xl/workbook.xml')
	const workbookRelsXml = archive.readText('xl/_rels/workbook.xml.rels')
	if (!workbookXml || !workbookRelsXml) return null
	const workbookInfo = parseWorkbookXml(workbookXml)
	const workbookRels = parseRelationships(workbookRelsXml)
	const sharedStringsPath = sharedStringsPathForWorkbook(archive, 'xl/workbook.xml', workbookRels)
	const sharedStringsXml = sharedStringsPath ? archive.readText(sharedStringsPath) : undefined
	const sharedStrings = sharedStringsXml
		? parseSharedStrings(sharedStringsXml, { lazy: true })
		: emptySharedStrings()
	const worksheetRels = new Map(
		workbookInfo.sheets
			.map((sheet, index) => [
				sheet.rId,
				worksheetPathForWorkbookSheet(archive, 'xl/workbook.xml', workbookRels, sheet, index),
			])
			.filter((entry): entry is [string, string] => entry[1] !== undefined),
	)
	let firstBoolean: ReturnType<typeof editableCellCandidate> = null
	let firstString: ReturnType<typeof editableCellCandidate> = null
	for (const sheet of workbookInfo.sheets) {
		const partPath = worksheetRels.get(sheet.rId)
		const xml = partPath ? archive.readText(partPath) : undefined
		if (!xml) continue
		const blockedRanges = arrayFormulaRanges(xml)
		for (const cell of scanSheetCells(xml)) {
			const { attrs, body } = cell
			if (hasOpenTag(body, 'f')) continue
			const ref = cell.ref
			const parsedRef = parseA1Safe(ref)
			if (!parsedRef) continue
			if (blockedRanges.some((range) => containsCell(range, parsedRef.row, parsedRef.col))) {
				continue
			}
			const candidate = editableCellCandidate(sheet.name, ref, attrs, body, sharedStrings)
			if (!candidate) continue
			if (candidate.valueType === 'number') return candidate
			if (candidate.valueType === 'boolean') firstBoolean ??= candidate
			if (candidate.valueType === 'string') firstString ??= candidate
		}
	}
	return firstBoolean ?? firstString
}

function selectAddCellEditTarget(bytes: Uint8Array): EditTarget | null {
	const archive = extractZip(bytes)
	const workbookXml = archive.readText('xl/workbook.xml')
	const workbookRelsXml = archive.readText('xl/_rels/workbook.xml.rels')
	if (!workbookXml || !workbookRelsXml) return null
	const workbookInfo = parseWorkbookXml(workbookXml)
	const workbookRels = parseRelationships(workbookRelsXml)
	const worksheetRels = new Map(
		workbookInfo.sheets
			.map((sheet, index) => [
				sheet.rId,
				worksheetPathForWorkbookSheet(archive, 'xl/workbook.xml', workbookRels, sheet, index),
			])
			.filter((entry): entry is [string, string] => entry[1] !== undefined),
	)
	for (const sheet of workbookInfo.sheets) {
		const partPath = worksheetRels.get(sheet.rId)
		const xml = partPath ? archive.readText(partPath) : undefined
		const ref = firstEmptyEditRef(xml ?? '')
		if (!ref) continue
		return {
			sheetName: sheet.name,
			ref,
			mode: 'add-cell',
			valueType: 'number',
			oldValue: null,
			newValue: 424242,
			expectedCellDelta: 1,
			expectedPhysicalCellDelta: 1,
			expectedFormulaDelta: 0,
		}
	}
	return null
}

function firstEmptyEditRef(sheetXml: string): string | null {
	const blockedCells = new Set<string>()
	const blockedRanges: CellRange[] = [...arrayFormulaRanges(sheetXml)]
	for (const cell of scanSheetCells(sheetXml)) {
		const parsed = parseA1Safe(normalizeA1Token(cell.ref))
		if (parsed) blockedCells.add(cellKey(parsed.row, parsed.col))
	}
	for (const match of sheetXml.matchAll(openTagRegex('mergeCell'))) {
		const range = parseCellRange(featureRef(match[1] ?? ''))
		if (range) blockedRanges.push(range)
	}
	for (let row = 0; row < 1000; row++) {
		for (let col = 0; col < 256; col++) {
			if (blockedCells.has(cellKey(row, col))) continue
			if (blockedRanges.some((range) => containsCell(range, row, col))) continue
			return formatLocalCellRef(row, col)
		}
	}
	return null
}

function arrayFormulaRanges(sheetXml: string): readonly CellRange[] {
	const ranges: CellRange[] = []
	for (const cell of scanSheetCells(sheetXml)) {
		for (const match of cell.body.matchAll(openTagRegex('f'))) {
			const attrs = parseXmlAttributes(match[1] ?? '')
			if (attrs.get('t') !== 'array') continue
			const range = parseCellRange(attrs.get('ref') ?? cell.ref)
			if (range) ranges.push(range)
		}
	}
	return ranges
}

interface CellRange {
	readonly minRow: number
	readonly minCol: number
	readonly maxRow: number
	readonly maxCol: number
}

function parseCellRange(ref: string): CellRange | null {
	const [startRef, endRef = startRef] = ref.split(':')
	const start = parseA1Safe(normalizeA1Token(startRef))
	const end = parseA1Safe(normalizeA1Token(endRef))
	if (!start || !end) return null
	return {
		minRow: Math.min(start.row, end.row),
		minCol: Math.min(start.col, end.col),
		maxRow: Math.max(start.row, end.row),
		maxCol: Math.max(start.col, end.col),
	}
}

function normalizeA1Token(ref: string | undefined): string | undefined {
	return ref?.replace(/\$/g, '')
}

function cellKey(row: number, col: number): string {
	return `${row}:${col}`
}

function containsCell(range: CellRange, row: number, col: number): boolean {
	return row >= range.minRow && row <= range.maxRow && col >= range.minCol && col <= range.maxCol
}

function editableCellCandidate(
	sheetName: string,
	ref: string,
	attrs: string,
	body: string,
	sharedStrings: SharedStringResolver,
): {
	sheetName: string
	ref: string
	valueType: EditTarget['valueType']
	value: string | number | boolean
} | null {
	const value = readCellScalarValue(attrs, body, sharedStrings)
	if (value === null || value === '') return null
	return { sheetName, ref, valueType: typeof value, value }
}

function editReplacementValue(value: string | number | boolean): string | number | boolean {
	if (typeof value === 'number') return value === 424242 ? 424243 : 424242
	if (typeof value === 'boolean') return !value
	return value === 'Ascend edit probe' ? 'Ascend edit probe 2' : 'Ascend edit probe'
}

function readCellScalar(
	bytes: Uint8Array,
	sheetName: string,
	ref: string,
): string | number | boolean | null {
	const archive = extractZip(bytes)
	const workbookXml = archive.readText('xl/workbook.xml')
	const workbookRelsXml = archive.readText('xl/_rels/workbook.xml.rels')
	if (!workbookXml || !workbookRelsXml) return null
	const workbookRels = parseRelationships(workbookRelsXml)
	const sharedStringsPath = sharedStringsPathForWorkbook(archive, 'xl/workbook.xml', workbookRels)
	const sharedStringsXml = sharedStringsPath ? archive.readText(sharedStringsPath) : undefined
	const sharedStrings = sharedStringsXml
		? parseSharedStrings(sharedStringsXml, { lazy: true })
		: emptySharedStrings()
	const workbookInfo = parseWorkbookXml(workbookXml)
	const sheet = workbookInfo.sheets.find((entry) => entry.name === sheetName)
	if (!sheet) return null
	const sheetIndex = workbookInfo.sheets.indexOf(sheet)
	const sheetPath = worksheetPathForWorkbookSheet(
		archive,
		'xl/workbook.xml',
		workbookRels,
		sheet,
		sheetIndex,
	)
	if (!sheetPath) return null
	const xml = archive.readText(sheetPath)
	if (!xml) return null
	const cell = findCellXml(xml, ref)
	if (!cell) return null
	return readCellScalarValue(cell.attrs, cell.body, sharedStrings)
}

function readCellScalarValue(
	attrs: string,
	body: string,
	sharedStrings: SharedStringResolver,
): string | number | boolean | null {
	const type = /(?:^|\s)t="([^"]+)"/.exec(attrs)?.[1]
	const raw = extractCellValueText(body)
	if (type === 's') {
		const index = raw !== null ? Number.parseInt(raw, 10) : -1
		if (index < 0) return ''
		const text = sharedStrings.getString?.(index)
		if (text !== undefined) return text
		const value = sharedStrings.get(index)
		const scalar = value ? topLeftScalar(value) : null
		return scalar?.kind === 'string' || scalar?.kind === 'richText'
			? serializeCellValue(value).slice(2)
			: ''
	}
	if (type === 'inlineStr') return extractInlineStringText(body)
	if (type === 'str') return raw ?? ''
	if (type === 'b') return raw === '1' || raw === 'true'
	if (type === 'e') return null
	if (raw === null || raw === '') return null
	const value = Number(raw)
	return Number.isFinite(value) ? value : null
}

function findCellXml(xml: string, ref: string): { attrs: string; body: string } | null {
	for (const cell of scanSheetCells(xml)) {
		if (cell.ref === ref) return { attrs: cell.attrs, body: cell.body }
	}
	return null
}

function scanSheetCells(
	xml: string,
): Array<{ readonly ref: string; readonly attrs: string; readonly body: string }> {
	const cells: Array<{ ref: string; attrs: string; body: string }> = []
	let inferredRowIndex = 0
	for (const rowMatch of xml.matchAll(openTagRegex('row'))) {
		const rowAttrs = rowMatch[1] ?? ''
		const rowStart = rowMatch.index + rowMatch[0].length
		const rowEnd = isSelfClosingTag(rowMatch[0])
			? rowStart
			: findCloseTagIndex(xml, 'row', rowStart)
		if (rowEnd < rowStart) continue
		const rowIndexText = /(?:^|\s)r="([^"]+)"/.exec(rowAttrs)?.[1]
		const rowIndex = rowIndexText ? Number.parseInt(rowIndexText, 10) - 1 : inferredRowIndex
		if (!Number.isInteger(rowIndex) || rowIndex < 0) continue
		inferredRowIndex = rowIndex + 1
		const rowBody = xml.slice(rowStart, rowEnd)
		let nextCol = 0
		for (const cellMatch of rowBody.matchAll(openTagRegex('c'))) {
			const attrs = cellMatch[1] ?? ''
			const bodyStart = cellMatch.index + cellMatch[0].length
			const bodyEnd = isSelfClosingTag(cellMatch[0])
				? bodyStart
				: findCloseTagIndex(rowBody, 'c', bodyStart)
			if (bodyEnd < bodyStart) continue
			const explicitRef = /(?:^|\s)r="([^"]+)"/.exec(attrs)?.[1]
			const parsed = parseA1Safe(normalizeA1Token(explicitRef))
			const col = parsed?.col ?? nextCol
			const row = parsed?.row ?? rowIndex
			nextCol = col + 1
			cells.push({
				ref: formatLocalCellRef(row, col),
				attrs,
				body: rowBody.slice(bodyStart, bodyEnd),
			})
		}
	}
	return cells
}

function scanSheetShapeXml(
	xml: string,
	sharedStrings: SharedStringResolver,
): {
	readonly cellCount: number
	readonly physicalCellCount: number
	readonly formulaCount: number
	readonly usedRange: string | null
	readonly physicalUsedRange: string | null
	readonly semanticCellRefs: readonly string[]
	readonly semanticCellValues: readonly string[]
	readonly formulaTexts: readonly string[]
} {
	let cellCount = 0
	let physicalCellCount = 0
	let formulaCount = 0
	let minRow = Number.POSITIVE_INFINITY
	let minCol = Number.POSITIVE_INFINITY
	let maxRow = -1
	let maxCol = -1
	let physicalMinRow = Number.POSITIVE_INFINITY
	let physicalMinCol = Number.POSITIVE_INFINITY
	let physicalMaxRow = -1
	let physicalMaxCol = -1
	const semanticCellRefs: string[] = []
	const semanticCellValues: string[] = []
	const formulaTexts: string[] = []
	for (const cell of scanSheetCells(xml)) {
		const { attrs, body, ref } = cell
		const parsed = parseA1Safe(ref)
		physicalCellCount++
		if (parsed) {
			physicalMinRow = Math.min(physicalMinRow, parsed.row)
			physicalMinCol = Math.min(physicalMinCol, parsed.col)
			physicalMaxRow = Math.max(physicalMaxRow, parsed.row)
			physicalMaxCol = Math.max(physicalMaxCol, parsed.col)
		}
		const hasCellFormula = hasOpenTag(body, 'f')
		const semantic = hasOpenTag(body, 'v') || hasCellFormula || hasOpenTag(body, 'is')
		if (semantic) {
			cellCount++
			if (hasCellFormula) formulaCount++
			if (parsed) {
				minRow = Math.min(minRow, parsed.row)
				minCol = Math.min(minCol, parsed.col)
				maxRow = Math.max(maxRow, parsed.row)
				maxCol = Math.max(maxCol, parsed.col)
				const normalizedRef = formatLocalCellRef(parsed.row, parsed.col)
				semanticCellRefs.push(normalizedRef)
				semanticCellValues.push(
					`${normalizedRef}\t${serializeExpectedCellValue(attrs, body, sharedStrings)}`,
				)
				const formula = extractCellFormulaText(body)
				if (formula !== null) formulaTexts.push(`${normalizedRef}=${formula}`)
			}
		}
	}
	const usedRange =
		maxRow >= 0
			? `${indexToColumn(minCol)}${minRow + 1}:${indexToColumn(maxCol)}${maxRow + 1}`
			: null
	const physicalUsedRange =
		physicalMaxRow >= 0
			? `${indexToColumn(physicalMinCol)}${physicalMinRow + 1}:${indexToColumn(physicalMaxCol)}${physicalMaxRow + 1}`
			: null
	return {
		cellCount,
		physicalCellCount,
		formulaCount,
		usedRange,
		physicalUsedRange,
		semanticCellRefs,
		semanticCellValues,
		formulaTexts,
	}
}

function serializeExpectedCellValue(
	attrs: string,
	body: string,
	sharedStrings: SharedStringResolver,
): string {
	const type = /(?:^|\s)t="([^"]+)"/.exec(attrs)?.[1]
	const rawValue = extractCellValueText(body)
	if (type === 's') {
		const index = rawValue !== null ? Number.parseInt(rawValue, 10) : -1
		if (index < 0) return 's:'
		const text = sharedStrings.getString?.(index)
		if (text !== undefined) return `s:${text}`
		const value = sharedStrings.get(index)
		return value ? serializeCellValue(value) : 's:'
	}
	if (type === 'inlineStr') return `s:${extractInlineStringText(body)}`
	if (type === 'str') return `s:${rawValue ?? ''}`
	if (type === 'b') return `b:${rawValue === '1' || rawValue === 'true' ? 'true' : 'false'}`
	if (type === 'e') return `e:${rawValue ?? '#VALUE!'}`
	if (rawValue === null || rawValue === '') return 'empty'
	const parsed = Number(rawValue)
	return Number.isNaN(parsed) ? `s:${rawValue}` : `n:${canonicalNumber(parsed)}`
}

function extractCellValueText(body: string): string | null {
	const match = elementBodyRegex('v').exec(body)
	if (!match) return null
	return decodeXmlText(match[1] ?? '')
}

function extractInlineStringText(body: string): string {
	let text = ''
	for (const match of body.matchAll(elementBodyRegex('t'))) {
		text += decodeXmlText(match[1] ?? '')
	}
	return text
}

function extractCellFormulaText(body: string): string | null {
	const match = elementBodyRegex('f').exec(body)
	if (!match) return null
	const formula = decodeXmlText(match[1] ?? '')
	return formula.length > 0 ? formula : null
}

async function firstExisting(paths: readonly string[]): Promise<string[]> {
	const existing: string[] = []
	for (const candidate of paths) {
		const fullPath = resolve(candidate)
		try {
			await access(fullPath)
			existing.push(fullPath)
		} catch {
			/* try the next candidate */
		}
	}
	if (existing.length === 0) {
		throw new Error(`No default real-workbook benchmark targets found: ${paths.join(', ')}`)
	}
	return existing
}

export function selectCorpusTargets(
	entries: readonly NormalizedCorpusManifestEntry[],
	selection: CorpusSelection,
	corpusRoot: string,
): Array<{ path: string; corpus: CorpusTargetMetadata }> {
	return selectManifestEntries(entries, selection).map((entry) => ({
		path: resolve(corpusRoot, entry.file),
		corpus: corpusTargetMetadata(entry),
	}))
}

function corpusTargetMetadata(entry: NormalizedCorpusManifestEntry): CorpusTargetMetadata {
	return {
		file: entry.file,
		benchmarkTier: entry.benchmarkTier,
		assertionClass: entry.assertionClass,
		riskClass: entry.riskClass,
		featureTags: entry.featureTags,
		vendorable: entry.vendorable,
		knownUnsupported: entry.knownUnsupported,
		...(entry.password !== undefined ? { password: entry.password } : {}),
	}
}

async function loadCorpusTargetSpecs(): Promise<
	Array<{ path: string; corpus: CorpusTargetMetadata }>
> {
	const manifestPath = resolve(readFlag('--corpus-manifest') ?? DEFAULT_CORPUS_MANIFEST)
	const corpusRoot = resolve(readFlag('--corpus-root') ?? DEFAULT_CORPUS_ROOT)
	const manifest = normalizeManifest(await loadCorpusManifestEntries(manifestPath))
	const selected = selectCorpusTargets(manifest, readCorpusSelection(), corpusRoot)
	if (selected.length === 0) {
		throw new Error('No corpus entries matched the requested real-workbook benchmark filters')
	}
	return selected
}

async function loadExpectedShapeSidecar(): Promise<
	(WorkbookShapeSummary & { readonly xlsxSha256?: string }) | undefined
> {
	const sidecarPath = readFlag('--expected-shape-sidecar')
	if (!sidecarPath) return undefined
	const parsed = JSON.parse(await readFile(resolve(sidecarPath), 'utf-8')) as Partial<
		WorkbookShapeSummary & { readonly xlsxSha256?: string }
	>
	for (const key of [
		'sheetNames',
		'sheetCount',
		'cellCount',
		'physicalCellCount',
		'formulaCount',
		'usedRanges',
		'physicalUsedRanges',
		'sheetNamesHash',
		'usedRangesHash',
		'physicalUsedRangesHash',
		'semanticCellRefsHash',
		'semanticCellValuesHash',
		'formulaTextHash',
	] as const) {
		if (parsed[key] === undefined) {
			throw new Error(`Expected shape sidecar is missing "${key}"`)
		}
	}
	return parsed as WorkbookShapeSummary & { readonly xlsxSha256?: string }
}

async function loadTargets(options: {
	readonly needAscendObservation: boolean
}): Promise<WorkbookTarget[]> {
	const explicit = positionalArgs()
	const targetSpecs =
		explicit.length > 0
			? explicit.map((path) => ({ path: resolve(path) }))
			: hasCorpusTargetMode()
				? await loadCorpusTargetSpecs()
				: (await firstExisting(hasFlag('--full-corpus') ? FULL_CORPUS_TARGETS : QUICK_TARGETS)).map(
						(path) => ({ path }),
					)
	const targets: WorkbookTarget[] = []
	const expectedShapeSidecar = await loadExpectedShapeSidecar()
	for (const spec of targetSpecs) {
		const path = spec.path
		const bytes = new Uint8Array(await readFile(path))
		const packageBytes = decryptTargetPackageBytes(bytes, spec.corpus)
		const targetSha256 = sha256(bytes)
		const packageSha256 = sha256(packageBytes)
		if (
			expectedShapeSidecar?.xlsxSha256 !== undefined &&
			expectedShapeSidecar.xlsxSha256 !== targetSha256
		) {
			throw new Error(
				`Expected shape sidecar SHA-256 ${expectedShapeSidecar.xlsxSha256} does not match ${path} SHA-256 ${targetSha256}`,
			)
		}
		const expectedInfo = expectedShapeSidecar ?? extractExpectedWorkbookShape(packageBytes)
		const ascendSummary = options.needAscendObservation
			? summarizeAscendWorkbook(
					await Ascend.open(bytes, targetOpenOptions(spec.corpus, { mode: 'formula' })),
				)
			: undefined
		targets.push({
			path,
			name: basename(path),
			extension: extname(path).toLowerCase(),
			bytes,
			packageBytes,
			sizeBytes: bytes.byteLength,
			sha256: targetSha256,
			packageSha256,
			expectedInfo,
			...(ascendSummary
				? {
						ascendInfo: {
							sheetCount: ascendSummary.sheetCount,
							cellCount: ascendSummary.cellCount,
							physicalCellCount: null,
							formulaCount: ascendSummary.formulaCount,
							usedRanges: ascendSummary.usedRanges,
							physicalUsedRanges: [],
							compatibility: ascendSummary.compatibility,
						},
					}
				: {}),
			...('corpus' in spec ? { corpus: spec.corpus } : {}),
		})
	}
	return targets
}

function decryptTargetPackageBytes(
	bytes: Uint8Array,
	corpus: CorpusTargetMetadata | undefined,
): Uint8Array {
	if (corpus?.password === undefined) return bytes
	const result = readXlsx(bytes, { password: corpus.password })
	if (!result.ok) throw new Error(`${corpus.file}: ${result.error.message}`)
	return result.value.workbook.sourceArchiveBytes ?? bytes
}

function targetOpenOptions<T extends Record<string, unknown>>(
	corpus: CorpusTargetMetadata | undefined,
	options: T,
): T & { readonly password?: string } {
	return corpus?.password === undefined ? options : { ...options, password: corpus.password }
}

async function loadExternalRunnerSpecs(): Promise<ExternalRunnerSpec[]> {
	const manifestPaths = readFlagValues('--runner-manifest')
	if (manifestPaths.length === 0) return []
	const parsedManifests: unknown[] = []
	for (const manifestPath of manifestPaths) {
		const raw = await readFile(resolve(manifestPath), 'utf-8')
		parsedManifests.push(JSON.parse(raw) as unknown)
	}
	return normalizeExternalRunnerManifestSet(parsedManifests).filter(
		(spec) =>
			externalRunnerLicenseGateSatisfied(spec) &&
			(spec.categories === undefined ||
				spec.categories.some((category) => isRunnableExternalRunnerCategory(category))),
	)
}

function isRunnableExternalRunnerCategory(
	category: ExternalRunnerCategory,
): category is CompetitiveCategory {
	return category === 'read' || category === 'roundtrip' || category === 'edit-roundtrip'
}

export function externalRunnerLicenseGateSatisfied(spec: ExternalRunnerSpec): boolean {
	const gate = spec.licenseGate
	if (!gate) return true
	const expected = gate.value ?? '1'
	return process.env[gate.env] === expected
}

export function resolveExternalRunnerCommand(
	command: readonly string[],
	env: Pick<NodeJS.ProcessEnv, 'ASCEND_BENCH_PYTHON'> = process.env,
): string[] {
	const python = env.ASCEND_BENCH_PYTHON?.trim()
	if (python && command[0] === 'python3') {
		return [python, ...command.slice(1)]
	}
	return [...command]
}

export function normalizeExternalRunnerSpecs(parsed: unknown): ExternalRunnerSpec[] {
	if (!Array.isArray(parsed)) throw new Error('--runner-manifest must be a JSON array')
	const names = new Set<string>()
	return parsed.map((entry, index) => {
		if (typeof entry !== 'object' || entry === null) {
			throw new Error(`External runner at index ${index} must be an object`)
		}
		const spec = entry as Partial<ExternalRunnerSpec>
		if (typeof spec.name !== 'string' || spec.name.length === 0) {
			throw new Error(`External runner at index ${index} is missing "name"`)
		}
		if (names.has(spec.name)) {
			throw new Error(`External runner "${spec.name}" is declared more than once`)
		}
		names.add(spec.name)
		if (
			!Array.isArray(spec.command) ||
			spec.command.length === 0 ||
			spec.command.some((part) => typeof part !== 'string' || part.length === 0)
		) {
			throw new Error(`External runner "${spec.name}" must provide command as a string array`)
		}
		if (
			spec.categories &&
			(!Array.isArray(spec.categories) ||
				spec.categories.some(
					(category) =>
						category !== 'read' &&
						category !== 'roundtrip' &&
						category !== 'edit-roundtrip' &&
						category !== 'write',
				))
		) {
			throw new Error(`External runner "${spec.name}" has invalid categories`)
		}
		const capabilities = normalizeRunnerCapabilities(spec.name, spec.capabilities)
		const licenseGate = normalizeRunnerLicenseGate(spec.name, spec.licenseGate)
		const workloads = normalizeRunnerWorkloads(spec.name, spec.workloads)
		return {
			name: spec.name,
			command: spec.command,
			...(spec.categories ? { categories: spec.categories } : {}),
			...(workloads ? { workloads } : {}),
			...(typeof spec.adapterVersion === 'string' ? { adapterVersion: spec.adapterVersion } : {}),
			...(typeof spec.libraryVersion === 'string' ? { libraryVersion: spec.libraryVersion } : {}),
			...(typeof spec.runtime === 'string' ? { runtime: spec.runtime } : {}),
			...(typeof spec.timingModel === 'string' ? { timingModel: spec.timingModel } : {}),
			...(typeof spec.validationModel === 'string'
				? { validationModel: spec.validationModel }
				: {}),
			...(typeof spec.memoryModel === 'string' ? { memoryModel: spec.memoryModel } : {}),
			...(typeof spec.installHint === 'string' ? { installHint: spec.installHint } : {}),
			...(licenseGate ? { licenseGate } : {}),
			...(capabilities ? { capabilities } : {}),
		}
	})
}

export function normalizeExternalRunnerManifestSet(
	parsedManifests: readonly unknown[],
): ExternalRunnerSpec[] {
	const entries: unknown[] = []
	for (const parsed of parsedManifests) {
		if (!Array.isArray(parsed)) throw new Error('--runner-manifest must be a JSON array')
		entries.push(...parsed)
	}
	return normalizeExternalRunnerSpecs(entries)
}

function normalizeRunnerWorkloads(
	name: string,
	workloads: ExternalRunnerSpec['workloads'] | undefined,
): readonly string[] | undefined {
	if (workloads === undefined) return undefined
	if (
		!Array.isArray(workloads) ||
		workloads.length === 0 ||
		workloads.some((workload) => typeof workload !== 'string' || workload.length === 0)
	) {
		throw new Error(`External runner "${name}" workloads must be a non-empty string array`)
	}
	return workloads
}

function normalizeRunnerLicenseGate(
	name: string,
	value: ExternalRunnerSpec['licenseGate'] | undefined,
): ExternalRunnerSpec['licenseGate'] | undefined {
	if (value === undefined) return undefined
	if (typeof value !== 'object' || value === null || Array.isArray(value)) {
		throw new Error(`External runner "${name}" licenseGate must be an object`)
	}
	if (typeof value.env !== 'string' || value.env.length === 0) {
		throw new Error(`External runner "${name}" licenseGate.env must be a non-empty string`)
	}
	if (value.value !== undefined && typeof value.value !== 'string') {
		throw new Error(`External runner "${name}" licenseGate.value must be a string`)
	}
	if (value.reason !== undefined && typeof value.reason !== 'string') {
		throw new Error(`External runner "${name}" licenseGate.reason must be a string`)
	}
	return {
		env: value.env,
		...(value.value !== undefined ? { value: value.value } : {}),
		...(value.reason !== undefined ? { reason: value.reason } : {}),
	}
}

function normalizeRunnerCapabilities(
	name: string,
	value: ExternalRunnerSpec['capabilities'] | undefined,
): ExternalRunnerSpec['capabilities'] | undefined {
	if (value === undefined) return undefined
	if (typeof value !== 'object' || value === null || Array.isArray(value)) {
		throw new Error(`External runner "${name}" capabilities must be an object`)
	}
	const result: NonNullable<ExternalRunnerSpec['capabilities']> = {}
	for (const key of [
		'xlsmRoundtrip',
		'internalTiming',
		'valueOnlyRead',
		'metadataOnlyRead',
		'writeFormulas',
		'writeTables',
		'writeRichMetadata',
		'finalValidation',
	] as const) {
		const flag = value[key]
		if (flag === undefined) continue
		if (typeof flag !== 'boolean') {
			throw new Error(`External runner "${name}" capability "${key}" must be boolean`)
		}
		result[key] = flag
	}
	return Object.keys(result).length > 0 ? result : undefined
}

function summarizeSheetJsWorkbook(
	sheetJs: typeof import('xlsx'),
	workbook: import('xlsx').WorkBook,
): WorkbookShapeSummary {
	let cellCount = 0
	let physicalCellCount = 0
	let formulaCount = 0
	const sheetNames = [...workbook.SheetNames]
	const usedRanges: string[] = []
	const physicalUsedRanges: string[] = []
	const semanticCellRefs: string[] = []
	const semanticCellValues: string[] = []
	const formulaTexts: string[] = []
	for (const sheetName of workbook.SheetNames) {
		const worksheet = workbook.Sheets[sheetName]
		const ref = typeof worksheet?.['!ref'] === 'string' ? worksheet['!ref'] : null
		if (ref) {
			const range = sheetJs.utils.decode_range(ref)
			physicalUsedRanges.push(`${sheetName}!${sheetJs.utils.encode_range(range)}`)
		} else {
			physicalUsedRanges.push(`${sheetName}!empty`)
		}
		let minRow = Number.POSITIVE_INFINITY
		let minCol = Number.POSITIVE_INFINITY
		let maxRow = -1
		let maxCol = -1
		if (!worksheet) {
			usedRanges.push(`${sheetName}!empty`)
			continue
		}
		for (const [key, cell] of Object.entries(worksheet)) {
			if (key.startsWith('!')) continue
			physicalCellCount++
			const formula = formulaTextOf(cell)
			if (formula !== null) formulaCount++
			if (!isSheetJsSemanticCell(cell)) continue
			cellCount++
			const parsed = parseA1Safe(key)
			if (!parsed) continue
			const ref = formatCellRef(sheetName, parsed.row, parsed.col)
			semanticCellRefs.push(ref)
			semanticCellValues.push(`${ref}\t${serializeSheetJsCellValue(cell)}`)
			if (formula !== null) formulaTexts.push(`${ref}=${formula}`)
			minRow = Math.min(minRow, parsed.row)
			minCol = Math.min(minCol, parsed.col)
			maxRow = Math.max(maxRow, parsed.row)
			maxCol = Math.max(maxCol, parsed.col)
		}
		usedRanges.push(
			maxRow >= 0
				? `${sheetName}!${indexToColumn(minCol)}${minRow + 1}:${indexToColumn(maxCol)}${maxRow + 1}`
				: `${sheetName}!empty`,
		)
	}
	return {
		sheetNames,
		sheetCount: workbook.SheetNames.length,
		cellCount,
		physicalCellCount,
		formulaCount,
		usedRanges,
		physicalUsedRanges,
		sheetNamesHash: hashLines(sheetNames.map((name, index) => `${index}:${name}`)),
		usedRangesHash: hashLines(usedRanges),
		physicalUsedRangesHash: hashLines(physicalUsedRanges),
		semanticCellRefsHash: hashLines(semanticCellRefs),
		semanticCellValuesHash: hashLines(semanticCellValues),
		formulaTextHash: hashLines(formulaTexts),
	}
}

export function sheetJsReadFeatureAssertions(
	workbook: import('xlsx').WorkBook,
): Record<string, string | number | boolean | null> {
	let readCommentCount = 0
	let readHyperlinkCount = 0
	for (const sheetName of workbook.SheetNames) {
		const worksheet = workbook.Sheets[sheetName]
		if (!worksheet) continue
		for (const [key, cell] of Object.entries(worksheet)) {
			if (key.startsWith('!') || typeof cell !== 'object' || cell === null) continue
			const record = cell as Record<string, unknown>
			if (Array.isArray(record.c)) readCommentCount += record.c.length
			else if (record.c !== undefined) readCommentCount++
			if (record.l !== undefined) readHyperlinkCount++
		}
	}
	return {
		readCommentCount,
		readHyperlinkCount,
		readDataValidationCount: 0,
		readConditionalFormatCount: 0,
		readDefinedNameCount: workbook.Workbook?.Names?.length ?? 0,
	}
}

interface ExcelJsCellLike {
	readonly value: unknown
	readonly address?: string
	readonly note?: unknown
	readonly hyperlink?: unknown
}

interface ExcelJsRowLike {
	eachCell(
		options: { readonly includeEmpty: boolean },
		callback: (cell: ExcelJsCellLike) => void,
	): void
}

interface ExcelJsWorksheetLike {
	readonly name: string
	readonly actualRowCount: number
	readonly actualColumnCount: number
	readonly dimensions: { readonly range?: string }
	readonly dataValidations?: { readonly model?: Record<string, unknown> }
	readonly conditionalFormattings?: readonly unknown[]
	eachRow(
		options: { readonly includeEmpty: boolean },
		callback: (row: ExcelJsRowLike) => void,
	): void
}

interface ExcelJsWorkbookLike {
	readonly worksheets: readonly ExcelJsWorksheetLike[]
	readonly definedNames?: { readonly model?: readonly unknown[] }
}

function summarizeExcelJsWorkbook(workbook: ExcelJsWorkbookLike): WorkbookShapeSummary {
	let cellCount = 0
	let physicalCellCount = 0
	let formulaCount = 0
	const sheetNames = workbook.worksheets.map((worksheet) => worksheet.name)
	const usedRanges: string[] = []
	const physicalUsedRanges: string[] = []
	const semanticCellRefs: string[] = []
	const semanticCellValues: string[] = []
	const formulaTexts: string[] = []
	for (const worksheet of workbook.worksheets) {
		physicalUsedRanges.push(
			worksheet.dimensions.range
				? `${worksheet.name}!${worksheet.dimensions.range}`
				: `${worksheet.name}!empty`,
		)
		let minRow = Number.POSITIVE_INFINITY
		let minCol = Number.POSITIVE_INFINITY
		let maxRow = -1
		let maxCol = -1
		worksheet.eachRow({ includeEmpty: false }, (row) => {
			row.eachCell({ includeEmpty: false }, (cell) => {
				physicalCellCount++
				const formula = formulaTextOf(cell.value)
				if (formula !== null) formulaCount++
				if (!isExcelJsSemanticCell(cell.value)) return
				cellCount++
				const parsed = parseA1Safe(cell.address)
				if (!parsed) return
				const ref = formatCellRef(worksheet.name, parsed.row, parsed.col)
				semanticCellRefs.push(ref)
				semanticCellValues.push(`${ref}\t${serializeUnknownScalarValue(cell.value)}`)
				if (formula !== null) formulaTexts.push(`${ref}=${formula}`)
				minRow = Math.min(minRow, parsed.row)
				minCol = Math.min(minCol, parsed.col)
				maxRow = Math.max(maxRow, parsed.row)
				maxCol = Math.max(maxCol, parsed.col)
			})
		})
		usedRanges.push(
			maxRow >= 0
				? `${worksheet.name}!${indexToColumn(minCol)}${minRow + 1}:${indexToColumn(maxCol)}${maxRow + 1}`
				: `${worksheet.name}!empty`,
		)
	}
	return {
		sheetNames,
		sheetCount: workbook.worksheets.length,
		cellCount,
		physicalCellCount,
		formulaCount,
		usedRanges,
		physicalUsedRanges,
		sheetNamesHash: hashLines(sheetNames.map((name, index) => `${index}:${name}`)),
		usedRangesHash: hashLines(usedRanges),
		physicalUsedRangesHash: hashLines(physicalUsedRanges),
		semanticCellRefsHash: hashLines(semanticCellRefs),
		semanticCellValuesHash: hashLines(semanticCellValues),
		formulaTextHash: hashLines(formulaTexts),
	}
}

export function excelJsReadFeatureAssertions(
	workbook: ExcelJsWorkbookLike,
): Record<string, string | number | boolean | null> {
	let readCommentCount = 0
	let readHyperlinkCount = 0
	let readDataValidationCount = 0
	let readConditionalFormatCount = 0
	for (const worksheet of workbook.worksheets) {
		readDataValidationCount += Object.keys(worksheet.dataValidations?.model ?? {}).length
		readConditionalFormatCount += worksheet.conditionalFormattings?.length ?? 0
		worksheet.eachRow({ includeEmpty: false }, (row) => {
			row.eachCell({ includeEmpty: false }, (cell) => {
				if (cell.note !== undefined) readCommentCount++
				const value = cell.value
				const valueHasHyperlink =
					typeof value === 'object' && value !== null && 'hyperlink' in value
				if (cell.hyperlink !== undefined || valueHasHyperlink) readHyperlinkCount++
			})
		})
	}
	return {
		readCommentCount,
		readHyperlinkCount,
		readDataValidationCount,
		readConditionalFormatCount,
		readDefinedNameCount: workbook.definedNames?.model?.length ?? 0,
	}
}

function serializeSheetJsCellValue(value: unknown): string {
	if (typeof value !== 'object' || value === null) return 'empty'
	const cell = value as Record<string, unknown>
	if (!('v' in cell)) return 'empty'
	if (cell.t === 'b') return `b:${cell.v ? 'true' : 'false'}`
	if (cell.t === 'e') return `e:${String(cell.v ?? '#VALUE!')}`
	if (cell.t === 'n') return `n:${canonicalNumber(Number(cell.v))}`
	return serializeUnknownScalarValue(cell.v)
}

function isSheetJsSemanticCell(value: unknown): boolean {
	return typeof value === 'object' && value !== null && ('v' in value || 'f' in value)
}

function isExcelJsSemanticCell(value: unknown): boolean {
	return value !== undefined && value !== null && value !== ''
}

function sheetJsEditCell(value: string | number | boolean): {
	t: 'b' | 'n' | 's'
	v: string | number | boolean
} {
	if (typeof value === 'number') return { t: 'n', v: value }
	if (typeof value === 'boolean') return { t: 'b', v: value }
	return { t: 's', v: value }
}

function formulaTextOf(value: unknown): string | null {
	if (typeof value !== 'object' || value === null) return null
	const record = value as Record<string, unknown>
	if (typeof record.f === 'string') return record.f
	if (typeof record.formula === 'string') return stripFormulaPrefix(record.formula)
	if (typeof record.sharedFormula === 'string') return stripFormulaPrefix(record.sharedFormula)
	return null
}

function stripFormulaPrefix(value: string): string {
	return value.startsWith('=') ? value.slice(1) : value
}

async function loadCases(): Promise<{
	cases: CompetitiveCase[]
	skipped: Array<{ library: string; reason: string }>
	externalRunnerSpecs: readonly ExternalRunnerSpec[]
}> {
	const cases: CompetitiveCase[] = [
		{
			name: 'ascend:read-formula',
			library: 'ascend',
			category: 'read',
			async runMeasured(target) {
				const start = performance.now()
				const workbook = await Ascend.open(
					target.bytes,
					targetOpenOptions(target.corpus, { mode: 'formula', richMetadata: true }),
				)
				const durationMs = performance.now() - start
				const summary = summarizeAscendWorkbook(workbook)
				return {
					durationMs,
					assertions: {
						...workbookShapeAssertions(summary),
						...workbookReadFeatureAssertions(workbook.getWorkbookModel()),
						compatibility: summary.compatibility,
					},
				}
			},
			async run(target) {
				const workbook = await Ascend.open(
					target.bytes,
					targetOpenOptions(target.corpus, { mode: 'formula', richMetadata: true }),
				)
				const summary = summarizeAscendWorkbook(workbook)
				return {
					assertions: {
						...workbookShapeAssertions(summary),
						...workbookReadFeatureAssertions(workbook.getWorkbookModel()),
						compatibility: summary.compatibility,
					},
				}
			},
		},
		{
			name: 'ascend:read-values',
			library: 'ascend',
			category: 'read',
			async runMeasured(target) {
				const start = performance.now()
				const workbook = await Ascend.open(
					target.bytes,
					targetOpenOptions(target.corpus, { mode: 'values' }),
				)
				const durationMs = performance.now() - start
				const summary = summarizeAscendWorkbook(workbook)
				return {
					durationMs,
					assertions: {
						...workbookShapeAssertions(summary),
						compatibility: summary.compatibility,
					},
				}
			},
			async run(target) {
				const workbook = await Ascend.open(
					target.bytes,
					targetOpenOptions(target.corpus, { mode: 'values' }),
				)
				const summary = summarizeAscendWorkbook(workbook)
				return {
					assertions: {
						...workbookShapeAssertions(summary),
						compatibility: summary.compatibility,
					},
				}
			},
		},
		{
			name: 'ascend:no-op-roundtrip',
			library: 'ascend',
			category: 'roundtrip',
			capabilities: { xlsmRoundtrip: true },
			async runMeasured(target) {
				const start = performance.now()
				const workbook = await Ascend.open(target.bytes, targetOpenOptions(target.corpus, {}))
				const bytes = workbook.toBytes()
				const durationMs = performance.now() - start
				return {
					durationMs,
					assertions: roundtripAssertions(bytes, target),
				}
			},
			async run(target) {
				const workbook = await Ascend.open(target.bytes, targetOpenOptions(target.corpus, {}))
				const bytes = workbook.toBytes()
				return {
					assertions: roundtripAssertions(bytes, target),
				}
			},
		},
		{
			name: 'ascend:edit-roundtrip',
			library: 'ascend',
			category: 'edit-roundtrip',
			capabilities: { xlsmRoundtrip: true },
			async runMeasured(target) {
				const edit = selectEditTarget(target)
				const start = performance.now()
				const workbook = await Ascend.open(target.bytes, targetOpenOptions(target.corpus, {}))
				workbook.apply([
					{
						op: 'setCells',
						sheet: edit.sheetName,
						updates: [{ ref: edit.ref, value: edit.newValue }],
					},
				])
				const bytes = workbook.toBytes()
				const durationMs = performance.now() - start
				return {
					durationMs,
					assertions: editRoundtripAssertions(bytes, target, edit),
				}
			},
			async run(target) {
				const edit = selectEditTarget(target)
				const workbook = await Ascend.open(target.bytes, targetOpenOptions(target.corpus, {}))
				workbook.apply([
					{
						op: 'setCells',
						sheet: edit.sheetName,
						updates: [{ ref: edit.ref, value: edit.newValue }],
					},
				])
				const bytes = workbook.toBytes()
				return {
					assertions: editRoundtripAssertions(bytes, target, edit),
				}
			},
		},
	]
	const skipped: Array<{ library: string; reason: string }> = []

	let sheetJs: typeof import('xlsx') | undefined
	try {
		sheetJs = await import('xlsx')
	} catch (error) {
		skipped.push({
			library: 'sheetjs',
			reason: error instanceof Error ? error.message : 'module not available',
		})
	}
	if (sheetJs) {
		cases.push(
			{
				name: 'sheetjs:read',
				library: 'sheetjs',
				category: 'read',
				async runMeasured(target) {
					const start = performance.now()
					const workbook = sheetJs.read(target.bytes, {
						type: 'buffer',
						bookVBA: true,
						cellStyles: true,
					})
					const durationMs = performance.now() - start
					return {
						durationMs,
						assertions: {
							...workbookShapeAssertions(summarizeSheetJsWorkbook(sheetJs, workbook)),
							...sheetJsReadFeatureAssertions(workbook),
						},
					}
				},
				async run(target) {
					const workbook = sheetJs.read(target.bytes, {
						type: 'buffer',
						bookVBA: true,
						cellStyles: true,
					})
					return {
						assertions: {
							...workbookShapeAssertions(summarizeSheetJsWorkbook(sheetJs, workbook)),
							...sheetJsReadFeatureAssertions(workbook),
						},
					}
				},
			},
			{
				name: 'sheetjs:no-op-roundtrip',
				library: 'sheetjs',
				category: 'roundtrip',
				capabilities: { xlsmRoundtrip: true },
				async runMeasured(target) {
					const start = performance.now()
					const workbook = sheetJs.read(target.bytes, {
						type: 'buffer',
						bookVBA: true,
						cellStyles: true,
					})
					const bookType = target.extension === '.xlsm' ? 'xlsm' : 'xlsx'
					const bytes = sheetJs.write(workbook, { type: 'buffer', bookType }) as Uint8Array
					const durationMs = performance.now() - start
					return {
						durationMs,
						assertions: roundtripAssertions(bytes, target),
					}
				},
				async run(target) {
					const workbook = sheetJs.read(target.bytes, {
						type: 'buffer',
						bookVBA: true,
						cellStyles: true,
					})
					const bookType = target.extension === '.xlsm' ? 'xlsm' : 'xlsx'
					const bytes = sheetJs.write(workbook, { type: 'buffer', bookType }) as Uint8Array
					return {
						assertions: roundtripAssertions(bytes, target),
					}
				},
			},
			{
				name: 'sheetjs:edit-roundtrip',
				library: 'sheetjs',
				category: 'edit-roundtrip',
				capabilities: { xlsmRoundtrip: true },
				async runMeasured(target) {
					const edit = selectEditTarget(target)
					const start = performance.now()
					const workbook = sheetJs.read(target.bytes, {
						type: 'buffer',
						bookVBA: true,
						cellStyles: true,
					})
					const worksheet = workbook.Sheets[edit.sheetName]
					if (!worksheet) throw new Error(`SheetJS missing sheet ${edit.sheetName}`)
					worksheet[edit.ref] = sheetJsEditCell(edit.newValue)
					const bookType = target.extension === '.xlsm' ? 'xlsm' : 'xlsx'
					const bytes = sheetJs.write(workbook, { type: 'buffer', bookType }) as Uint8Array
					const durationMs = performance.now() - start
					return {
						durationMs,
						assertions: editRoundtripAssertions(bytes, target, edit),
					}
				},
				async run(target) {
					const edit = selectEditTarget(target)
					const workbook = sheetJs.read(target.bytes, {
						type: 'buffer',
						bookVBA: true,
						cellStyles: true,
					})
					const worksheet = workbook.Sheets[edit.sheetName]
					if (!worksheet) throw new Error(`SheetJS missing sheet ${edit.sheetName}`)
					worksheet[edit.ref] = sheetJsEditCell(edit.newValue)
					const bookType = target.extension === '.xlsm' ? 'xlsm' : 'xlsx'
					const bytes = sheetJs.write(workbook, { type: 'buffer', bookType }) as Uint8Array
					return {
						assertions: editRoundtripAssertions(bytes, target, edit),
					}
				},
			},
		)
	}

	let ExcelJS: typeof import('exceljs') | undefined
	try {
		ExcelJS = await import('exceljs')
	} catch (error) {
		skipped.push({
			library: 'exceljs',
			reason: error instanceof Error ? error.message : 'module not available',
		})
	}
	if (ExcelJS) {
		cases.push(
			{
				name: 'exceljs:read',
				library: 'exceljs',
				category: 'read',
				async runMeasured(target) {
					const workbook = new ExcelJS.Workbook()
					const start = performance.now()
					await workbook.xlsx.load(Buffer.from(target.bytes))
					const durationMs = performance.now() - start
					return {
						durationMs,
						assertions: {
							...workbookShapeAssertions(summarizeExcelJsWorkbook(workbook)),
							...excelJsReadFeatureAssertions(workbook),
						},
					}
				},
				async run(target) {
					const workbook = new ExcelJS.Workbook()
					await workbook.xlsx.load(Buffer.from(target.bytes))
					return {
						assertions: {
							...workbookShapeAssertions(summarizeExcelJsWorkbook(workbook)),
							...excelJsReadFeatureAssertions(workbook),
						},
					}
				},
			},
			{
				name: 'exceljs:no-op-roundtrip',
				library: 'exceljs',
				category: 'roundtrip',
				capabilities: { xlsmRoundtrip: false },
				async runMeasured(target) {
					const workbook = new ExcelJS.Workbook()
					const start = performance.now()
					await workbook.xlsx.load(Buffer.from(target.bytes))
					const bytes = (await workbook.xlsx.writeBuffer()) as ArrayBuffer
					const durationMs = performance.now() - start
					return {
						durationMs,
						assertions: roundtripAssertions(new Uint8Array(bytes), target),
					}
				},
				async run(target) {
					const workbook = new ExcelJS.Workbook()
					await workbook.xlsx.load(Buffer.from(target.bytes))
					const bytes = (await workbook.xlsx.writeBuffer()) as ArrayBuffer
					return {
						assertions: roundtripAssertions(new Uint8Array(bytes), target),
					}
				},
			},
			{
				name: 'exceljs:edit-roundtrip',
				library: 'exceljs',
				category: 'edit-roundtrip',
				capabilities: { xlsmRoundtrip: false },
				async runMeasured(target) {
					const edit = selectEditTarget(target)
					const workbook = new ExcelJS.Workbook()
					const start = performance.now()
					await workbook.xlsx.load(Buffer.from(target.bytes))
					const worksheet = workbook.getWorksheet(edit.sheetName)
					if (!worksheet) throw new Error(`ExcelJS missing sheet ${edit.sheetName}`)
					worksheet.getCell(edit.ref).value = edit.newValue
					const bytes = (await workbook.xlsx.writeBuffer()) as ArrayBuffer
					const durationMs = performance.now() - start
					return {
						durationMs,
						assertions: editRoundtripAssertions(new Uint8Array(bytes), target, edit),
					}
				},
				async run(target) {
					const edit = selectEditTarget(target)
					const workbook = new ExcelJS.Workbook()
					await workbook.xlsx.load(Buffer.from(target.bytes))
					const worksheet = workbook.getWorksheet(edit.sheetName)
					if (!worksheet) throw new Error(`ExcelJS missing sheet ${edit.sheetName}`)
					worksheet.getCell(edit.ref).value = edit.newValue
					const bytes = (await workbook.xlsx.writeBuffer()) as ArrayBuffer
					return {
						assertions: editRoundtripAssertions(new Uint8Array(bytes), target, edit),
					}
				},
			},
		)
	}

	const externalRunnerSpecs = await loadExternalRunnerSpecs()
	for (const spec of externalRunnerSpecs) {
		const categories = (spec.categories ?? ['read', 'roundtrip']).filter(
			(category): category is CompetitiveCategory => isRunnableExternalRunnerCategory(category),
		)
		for (const category of categories) {
			cases.push({
				name: `${spec.name}:${category}`,
				library: spec.name,
				category,
				executionScope: 'external-process',
				runnerProvenance: {
					...(spec.adapterVersion ? { adapterVersion: spec.adapterVersion } : {}),
					...(spec.libraryVersion ? { libraryVersion: spec.libraryVersion } : {}),
					...(spec.runtime ? { runtime: spec.runtime } : {}),
				},
				...(spec.timingModel ? { timingModel: spec.timingModel } : {}),
				...(spec.validationModel ? { validationModel: spec.validationModel } : {}),
				...(spec.memoryModel ? { memoryModel: spec.memoryModel } : {}),
				...(spec.capabilities ? { capabilities: spec.capabilities } : {}),
				async run(target) {
					return {
						assertions: await runExternalRunner(spec, category, target),
					}
				},
				...(spec.capabilities?.internalTiming
					? {
							async runBatched(target, repeat, warmup) {
								return runExternalRunnerBatched(spec, category, target, repeat, warmup)
							},
						}
					: {}),
			})
		}
	}

	return { cases, skipped, externalRunnerSpecs }
}

async function runExternalRunner(
	spec: ExternalRunnerSpec,
	category: CompetitiveCategory,
	target: WorkbookTarget,
): Promise<Record<string, string | number | boolean | null>> {
	const proc = Bun.spawn(externalRunnerCommand(spec, category, target, { json: true }), {
		stdout: 'pipe',
		stderr: 'pipe',
		cwd: process.cwd(),
	})
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(stderr.trim() || `External runner "${spec.name}" exited with code ${exitCode}`)
	}
	const parsed = JSON.parse(stdout) as unknown
	return normalizeAssertions(parsed)
}

async function runExternalRunnerBatched(
	spec: ExternalRunnerSpec,
	category: CompetitiveCategory,
	target: WorkbookTarget,
	repeat: number,
	warmup: number,
): Promise<{
	assertions?: Record<string, string | number | boolean | null>
	samples?: readonly MetricSample[]
}> {
	const proc = Bun.spawn(
		externalRunnerCommand(spec, category, target, { repeat, warmup, json: true }),
		{
			stdout: 'pipe',
			stderr: 'pipe',
			cwd: process.cwd(),
		},
	)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(stderr.trim() || `External runner "${spec.name}" exited with code ${exitCode}`)
	}
	const parsed = JSON.parse(stdout) as unknown
	return {
		assertions: normalizeAssertions(parsed),
		assertionsBySample: normalizeExternalSampleAssertions(parsed, repeat, spec.name),
		samples: normalizeExternalSamples(parsed, repeat, spec.name),
	}
}

function externalRunnerCommand(
	spec: ExternalRunnerSpec,
	category: CompetitiveCategory,
	target: WorkbookTarget,
	options: { repeat?: number; warmup?: number; json?: boolean } = {},
): string[] {
	const command = [
		...resolveExternalRunnerCommand(spec.command),
		'--operation',
		category,
		'--file',
		target.path,
	]
	if (category === 'edit-roundtrip') {
		const edit = selectEditTarget(target)
		command.push(
			'--edit-sheet',
			edit.sheetName,
			'--edit-ref',
			edit.ref,
			'--edit-value-type',
			edit.valueType,
			'--edit-value',
			String(edit.newValue),
			'--edit-old-value',
			String(edit.oldValue),
		)
	}
	if (options.repeat !== undefined) command.push('--repeat', String(options.repeat))
	if (options.warmup !== undefined) command.push('--warmup', String(options.warmup))
	if (options.json) command.push('--json')
	return command
}

export function normalizeAssertions(
	value: unknown,
): Record<string, string | number | boolean | null> {
	const source =
		typeof value === 'object' &&
		value !== null &&
		'assertions' in value &&
		typeof value.assertions === 'object' &&
		value.assertions !== null
			? value.assertions
			: value
	if (typeof source !== 'object' || source === null || Array.isArray(source)) {
		throw new Error('External runner output must be a JSON object or { "assertions": object }')
	}
	const assertions: Record<string, string | number | boolean | null> = {}
	for (const [key, entry] of Object.entries(source)) {
		if (
			typeof entry === 'string' ||
			typeof entry === 'number' ||
			typeof entry === 'boolean' ||
			entry === null
		) {
			assertions[key] = entry
		} else {
			throw new Error(`External runner assertion "${key}" must be a primitive value`)
		}
	}
	return assertions
}

export function normalizeExternalSamples(
	value: unknown,
	expectedRepeat?: number,
	runnerName = 'external runner',
): readonly MetricSample[] | undefined {
	if (
		typeof value !== 'object' ||
		value === null ||
		!('samples' in value) ||
		!Array.isArray(value.samples)
	) {
		return undefined
	}
	const samples: MetricSample[] = []
	for (const [index, sample] of value.samples.entries()) {
		if (typeof sample !== 'object' || sample === null) {
			throw new Error(`${runnerName} sample ${index} must be an object`)
		}
		const durationMs = (sample as { durationMs?: unknown }).durationMs
		if (typeof durationMs !== 'number' || !Number.isFinite(durationMs) || durationMs <= 0) {
			throw new Error(`${runnerName} sample ${index} must provide a positive durationMs`)
		}
		samples.push({
			durationMs,
			...optionalSampleNumber(sample, 'throughputPerSec'),
			...optionalSampleNumber(sample, 'rssDeltaBytes'),
			...optionalSampleNumber(sample, 'retainedRssDeltaBytes'),
			...optionalSampleNumber(sample, 'rssAfterBytes'),
			...optionalSampleNumber(sample, 'rssAfterGcBytes'),
			...optionalSampleNumber(sample, 'peakRssBytes'),
			...optionalSampleNumber(sample, 'heapDeltaBytes'),
			...optionalSampleNumber(sample, 'heapUsedBytes'),
			...optionalSampleNumber(sample, 'heapTotalBytes'),
			...optionalSampleNumber(sample, 'heapAfterGcBytes'),
		})
	}
	if (expectedRepeat !== undefined && samples.length !== expectedRepeat) {
		throw new Error(
			`${runnerName} reported ${samples.length} samples but repeat requested ${expectedRepeat}`,
		)
	}
	return samples.length > 0 ? samples : undefined
}

export function normalizeExternalSampleAssertions(
	value: unknown,
	expectedRepeat: number,
	runnerName = 'external runner',
): readonly Record<string, string | number | boolean | null>[] | undefined {
	if (typeof value !== 'object' || value === null) return undefined
	const record = value as Record<string, unknown>
	const source = Array.isArray(record.assertionsBySample)
		? record.assertionsBySample
		: Array.isArray(record.samples) &&
				record.samples.some(
					(sample) => typeof sample === 'object' && sample !== null && 'assertions' in sample,
				)
			? record.samples.map((sample) =>
					typeof sample === 'object' && sample !== null
						? (sample as Record<string, unknown>).assertions
						: undefined,
				)
			: undefined
	if (source === undefined) return undefined
	if (source.length !== expectedRepeat) {
		throw new Error(
			`${runnerName} reported ${source.length} assertion samples but repeat requested ${expectedRepeat}`,
		)
	}
	return source.map((assertions, index) => {
		if (typeof assertions !== 'object' || assertions === null || Array.isArray(assertions)) {
			throw new Error(`${runnerName} assertion sample ${index} must be an object`)
		}
		return normalizeAssertions({ assertions })
	})
}

function optionalSampleNumber(
	sample: object,
	key: keyof Omit<MetricSample, 'durationMs'>,
): Partial<MetricSample> {
	const value = (sample as Record<string, unknown>)[key]
	if (value === undefined) return {}
	if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) {
		throw new Error(`External runner sample field "${key}" must be a non-negative finite number`)
	}
	return { [key]: value }
}

function skipReason(benchmarkCase: CompetitiveCase, target: WorkbookTarget): string | null {
	if (
		target.extension === '.xlsm' &&
		(benchmarkCase.category === 'roundtrip' || benchmarkCase.category === 'edit-roundtrip') &&
		benchmarkCase.capabilities?.xlsmRoundtrip !== true
	) {
		return `${benchmarkCase.category} for .xlsm requires explicit VBA/package preservation capability`
	}
	return null
}

function operationProfile(benchmarkCase: CompetitiveCase): string {
	if (benchmarkCase.category === 'edit-roundtrip') return 'edit-roundtrip'
	if (benchmarkCase.category === 'roundtrip') return 'no-op-roundtrip'
	if (benchmarkCase.capabilities?.valueOnlyRead) return 'read-values'
	if (benchmarkCase.name === 'ascend:read-values') return 'read-values'
	return 'read-formula-preserving'
}

export function coalesceRepeatCorrectnessStatus(statuses: readonly string[]): string {
	if (statuses.length === 0) return 'not-evaluated'
	const first = statuses[0] ?? 'not-evaluated'
	return statuses.every((status) => status === first) ? first : 'intermittent-mismatch'
}

function timingModel(benchmarkCase: CompetitiveCase): string {
	if (benchmarkCase.timingModel) return benchmarkCase.timingModel
	if (benchmarkCase.executionScope === 'external-process') {
		return benchmarkCase.capabilities?.internalTiming
			? 'external-internal-operation-timing'
			: 'external-process-wall-clock'
	}
	return benchmarkCase.runMeasured ? 'operation-only' : 'operation-plus-validation'
}

function validationModel(benchmarkCase: CompetitiveCase): string {
	if (benchmarkCase.validationModel) return benchmarkCase.validationModel
	if (benchmarkCase.executionScope === 'external-process') {
		return benchmarkCase.capabilities?.internalTiming
			? 'external-post-operation-assertions'
			: 'external-inline-assertions'
	}
	return benchmarkCase.runMeasured ? 'post-operation-assertions' : 'inline-assertions'
}

function timingLane(benchmarkCase: CompetitiveCase): string {
	if (benchmarkCase.executionScope === 'external-process') {
		if (benchmarkCase.timingModel) return benchmarkCase.timingModel
		return benchmarkCase.capabilities?.internalTiming
			? 'external-internal-file-path'
			: 'external-process-wall-file-path'
	}
	return 'in-process-preloaded-bytes'
}

async function runCompetitiveCase(
	benchmarkCase: CompetitiveCase,
	target: WorkbookTarget,
	repeat: number,
	warmup: number,
): Promise<BenchmarkCaseResult> {
	if (benchmarkCase.runBatched) {
		const profile = operationProfile(benchmarkCase)
		const result = await benchmarkCase.runBatched(target, repeat, warmup)
		const evaluatedBySample =
			result.assertionsBySample?.map((assertions) =>
				evaluateAssertions(benchmarkCase.category, target.expectedInfo, assertions, profile),
			) ?? []
		const evaluated =
			evaluatedBySample[0] ??
			evaluateAssertions(benchmarkCase.category, target.expectedInfo, result.assertions, profile)
		const samples =
			result.samples?.map((sample) => ({
				...sample,
				throughputPerSec:
					sample.throughputPerSec ??
					(target.expectedInfo.cellCount > 0 && sample.durationMs > 0
						? (target.expectedInfo.cellCount / sample.durationMs) * 1000
						: undefined),
			})) ?? []
		if (samples.length === 0) {
			throw new Error(`External runner "${benchmarkCase.library}" did not report samples`)
		}
		const correctnessStatuses =
			evaluatedBySample.length > 0 ? evaluatedBySample.map((entry) => entry.status) : []
		return buildCompetitiveResult(
			benchmarkCase,
			target,
			repeat,
			samples,
			evaluatedBySample.length > 0
				? coalesceRepeatCorrectnessStatus(correctnessStatuses)
				: evaluated.status,
			evaluated.assertions,
			correctnessStatuses,
		)
	}
	for (let i = 0; i < warmup; i++) {
		await benchmarkCase.run(target)
	}
	const samples: MetricSample[] = []
	let assertions: Record<string, string | number | boolean | null> | undefined
	const correctnessStatuses: string[] = []
	const measureParentMemory = benchmarkCase.executionScope !== 'external-process'
	const profile = operationProfile(benchmarkCase)
	for (let i = 0; i < repeat; i++) {
		runGc()
		const rssBefore = measureParentMemory ? getRssBytes() : undefined
		const heapBefore = measureParentMemory ? process.memoryUsage().heapUsed : undefined
		const start = performance.now()
		const result = benchmarkCase.runMeasured
			? await benchmarkCase.runMeasured(target)
			: await benchmarkCase.run(target)
		const durationMs =
			'durationMs' in result && typeof result.durationMs === 'number'
				? result.durationMs
				: performance.now() - start
		const memAfter = measureParentMemory ? process.memoryUsage() : undefined
		const rssAfter = measureParentMemory ? getRssBytes() : undefined
		runGc()
		const rssAfterGc = measureParentMemory ? getRssBytes() : undefined
		const heapAfterGc = measureParentMemory ? process.memoryUsage().heapUsed : undefined
		const evaluated = evaluateAssertions(
			benchmarkCase.category,
			target.expectedInfo,
			result.assertions,
			profile,
		)
		samples.push({
			durationMs,
			throughputPerSec:
				target.expectedInfo.cellCount > 0 && durationMs > 0
					? (target.expectedInfo.cellCount / durationMs) * 1000
					: undefined,
			rssDeltaBytes:
				rssBefore !== undefined && rssAfter !== undefined
					? Math.max(0, rssAfter - rssBefore)
					: undefined,
			retainedRssDeltaBytes:
				rssBefore !== undefined && rssAfterGc !== undefined
					? Math.max(0, rssAfterGc - rssBefore)
					: undefined,
			peakRssBytes: observedPeakRssBytes([rssBefore, rssAfter, rssAfterGc]),
			heapDeltaBytes:
				memAfter !== undefined && heapBefore !== undefined
					? Math.max(0, memAfter.heapUsed - heapBefore)
					: undefined,
			heapUsedBytes: memAfter?.heapUsed,
			heapTotalBytes: memAfter?.heapTotal,
			heapAfterGcBytes: heapAfterGc,
		})
		assertions ??= evaluated.assertions
		correctnessStatuses.push(evaluated.status)
	}
	const correctnessStatus = coalesceRepeatCorrectnessStatus(correctnessStatuses)
	return buildCompetitiveResult(
		benchmarkCase,
		target,
		repeat,
		samples,
		correctnessStatus,
		assertions,
		correctnessStatuses,
	)
}

function buildCompetitiveResult(
	benchmarkCase: CompetitiveCase,
	target: WorkbookTarget,
	repeat: number,
	samples: readonly MetricSample[],
	correctnessStatus: string,
	assertions: Record<string, string | number | boolean | null> | undefined,
	correctnessStatuses: readonly string[] = [],
): BenchmarkCaseResult {
	return {
		name: `${benchmarkCase.name}:${target.name}`,
		category: benchmarkCase.category,
		dimensions: {
			library: benchmarkCase.library,
			workload: 'real-workbook',
			file: target.name,
			extension: target.extension,
			bytes: target.sizeBytes,
			sheets: target.expectedInfo.sheetCount,
			cells: target.expectedInfo.cellCount,
			...(target.corpus
				? {
						corpusTier: target.corpus.benchmarkTier,
						corpusAssertionClass: target.corpus.assertionClass,
						corpusRiskClass: target.corpus.riskClass,
						corpusTags: target.corpus.featureTags.join(','),
					}
				: {}),
			operationProfile: operationProfile(benchmarkCase),
			repeat,
			executionScope: benchmarkCase.executionScope ?? 'in-process',
			timingLane: timingLane(benchmarkCase),
			timingModel: timingModel(benchmarkCase),
			validationModel: validationModel(benchmarkCase),
			...(benchmarkCase.memoryModel ? { memoryModel: benchmarkCase.memoryModel } : {}),
			...benchmarkProvenanceDimensions(assertions, benchmarkCase.runnerProvenance),
			correctnessStatus,
			rankingEligible: isRankingEligible(correctnessStatus),
			...(correctnessStatuses.length > 1
				? { repeatCorrectnessStatuses: correctnessStatuses.join(',') }
				: {}),
		},
		metrics: summarizeSamples(samples),
		...(repeat > 1 ? { samples } : {}),
		...(assertions ? { assertions } : {}),
	}
}

function buildNonRankingResult(input: {
	readonly benchmarkCase: CompetitiveCase
	readonly target: WorkbookTarget
	readonly repeat: number
	readonly status: string
	readonly reason: string
}): BenchmarkCaseResult {
	return {
		name: `${input.benchmarkCase.name}:${input.target.name}`,
		category: input.benchmarkCase.category,
		dimensions: {
			library: input.benchmarkCase.library,
			workload: 'real-workbook',
			file: input.target.name,
			extension: input.target.extension,
			bytes: input.target.sizeBytes,
			sheets: input.target.expectedInfo.sheetCount,
			cells: input.target.expectedInfo.cellCount,
			...(input.target.corpus
				? {
						corpusTier: input.target.corpus.benchmarkTier,
						corpusAssertionClass: input.target.corpus.assertionClass,
						corpusRiskClass: input.target.corpus.riskClass,
						corpusTags: input.target.corpus.featureTags.join(','),
					}
				: {}),
			operationProfile: operationProfile(input.benchmarkCase),
			repeat: input.repeat,
			executionScope: input.benchmarkCase.executionScope ?? 'in-process',
			timingLane: timingLane(input.benchmarkCase),
			timingModel: timingModel(input.benchmarkCase),
			validationModel: validationModel(input.benchmarkCase),
			...(input.benchmarkCase.memoryModel ? { memoryModel: input.benchmarkCase.memoryModel } : {}),
			...benchmarkProvenanceDimensions(undefined, input.benchmarkCase.runnerProvenance),
			correctnessStatus: input.status,
			rankingEligible: false,
			errorReason: input.reason,
		},
		metrics: summarizeSamples([{ durationMs: 0 }]),
		assertions: { errorReason: input.reason },
	}
}

function renderSummary(
	results: readonly BenchmarkCaseResult[],
	skipped: readonly unknown[],
): string {
	const headers = [
		'case',
		'category',
		'median-ms',
		'p95-ms',
		'throughput',
		'rss-delta',
		'heap-delta',
	]
	const rows = results.map((result) => [
		result.name,
		result.category,
		result.metrics.medianMs.toFixed(2),
		result.metrics.p95Ms.toFixed(2),
		result.metrics.throughputPerSec !== undefined
			? formatRate(result.metrics.throughputPerSec)
			: 'n/a',
		result.metrics.rssDeltaBytes !== undefined ? formatBytes(result.metrics.rssDeltaBytes) : 'n/a',
		result.metrics.heapDeltaBytes !== undefined
			? formatBytes(result.metrics.heapDeltaBytes)
			: 'n/a',
	])
	const widths = headers.map((header, index) =>
		Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)),
	)
	const pad = (value: string, width: number) =>
		value + ' '.repeat(Math.max(0, width - value.length))
	const line = (cells: readonly string[]) =>
		cells.map((cell, index) => pad(cell, widths[index] ?? 0)).join('  ')
	const lines = [
		line(headers),
		widths.map((width) => '-'.repeat(width)).join('--'),
		...rows.map(line),
	]
	if (skipped.length > 0) {
		lines.push('')
		lines.push(`Skipped cases: ${JSON.stringify(skipped)}`)
	}
	return lines.join('\n')
}

async function main(): Promise<void> {
	const repeat = readPositiveIntFlag('--repeat', 3)
	const warmup = readNonNegativeIntFlag('--warmup', 1)
	const categoryFilter = readCategoryFilter()
	const json = hasFlag('--json')
	const loadedCases = await loadCases()
	const competitor = competitorFilter()
	const libraryAllowlist = parseLibraryAllowlist(readFlag('--libraries'))
	const categoryCases = categoryFilter
		? loadedCases.cases.filter((entry) => entry.category === categoryFilter)
		: loadedCases.cases
	const competitorCases =
		competitor === 'all'
			? categoryCases
			: categoryCases.filter((entry) =>
					competitor === 'external'
						? entry.executionScope === 'external-process'
						: entry.executionScope !== 'external-process',
				)
	const cases = competitorCases.filter((entry) => libraryAllowed(entry.library, libraryAllowlist))
	const targets = await loadTargets({
		needAscendObservation: cases.some((entry) => entry.library === 'ascend'),
	})
	const skipped = loadedCases.skipped
	const results: BenchmarkCaseResult[] = []
	const failed: Array<{ case: string; file: string; reason: string }> = []
	const skippedCases: Array<{ case: string; file: string; library: string; reason: string }> = []
	for (const target of targets) {
		for (const benchmarkCase of cases) {
			const unsupported = skipReason(benchmarkCase, target)
			if (unsupported) {
				skippedCases.push({
					case: benchmarkCase.name,
					file: target.name,
					library: benchmarkCase.library,
					reason: unsupported,
				})
				results.push(
					buildNonRankingResult({
						benchmarkCase,
						target,
						repeat,
						status: 'unsupported-capability',
						reason: unsupported,
					}),
				)
				continue
			}
			try {
				results.push(await runCompetitiveCase(benchmarkCase, target, repeat, warmup))
				if (!json) console.log(`completed ${benchmarkCase.name}:${target.name}`)
			} catch (error) {
				const reason = error instanceof Error ? error.message : String(error)
				failed.push({
					case: benchmarkCase.name,
					file: target.name,
					reason,
				})
				results.push(
					buildNonRankingResult({
						benchmarkCase,
						target,
						repeat,
						status: 'error',
						reason,
					}),
				)
			}
		}
	}
	const suite = createBenchmarkSuite({
		suite: 'ascend-competitive-real-workbooks',
		kind: 'real-workbook',
		cases: results,
		metadata: {
			workload: 'real-workbook-read-roundtrip',
			inProcessInputMode:
				'preloaded-bytes; in-process JS adapters do not include filesystem read latency',
			targetMode: targetMode(),
			competitor,
			expectedShapeSidecar: readFlag('--expected-shape-sidecar') ?? null,
			...(hasCorpusTargetMode()
				? {
						corpusSelection: readCorpusSelection(),
						corpusManifest: resolve(readFlag('--corpus-manifest') ?? DEFAULT_CORPUS_MANIFEST),
						corpusRoot: resolve(readFlag('--corpus-root') ?? DEFAULT_CORPUS_ROOT),
					}
				: {}),
			files: targets.map((target) => ({
				path: target.path,
				name: target.name,
				extension: target.extension,
				bytes: target.sizeBytes,
				sha256: target.sha256,
				...(target.corpus ? { corpus: target.corpus } : {}),
				expected: target.expectedInfo,
				...(target.ascendInfo ? { ascendObserved: target.ascendInfo } : {}),
			})),
			repeat,
			warmup,
			category: categoryFilter ?? 'all',
			...(libraryAllowlist ? { libraries: [...libraryAllowlist] } : {}),
			tempDir: tmpdir(),
			externalRunnerProtocol: {
				manifestFlag: '--runner-manifest',
				manifestShape: [
					{
						name: 'openpyxl',
						command: ['python3', 'fixtures/benchmarks/runners/openpyxl_runner.py'],
						categories: ['read', 'roundtrip', 'edit-roundtrip'],
					},
				],
				invocation:
					'<command...> --operation <read|roundtrip|edit-roundtrip> --file <path> [--edit-sheet NAME --edit-ref A1 --edit-value-type number|string|boolean --edit-value VALUE] [--repeat N --warmup N] --json',
				output:
					'JSON object or { "assertions": object, "samples": [{"durationMs": number}] } with primitive assertion values',
				timingModel:
					'Manifests with capabilities.internalTiming=true must report operation-only samples and run validation after the timed region.',
				validationModel:
					'Correctness is evaluated in the harness from primitive assertions; repeated in-process runs must keep the same correctness status or become intermittent-mismatch.',
				timingNote:
					'External runner timings include process startup unless capabilities.internalTiming=true or the runner command is a long-lived wrapper.',
				memoryNote:
					'External-process runners may report peakRssBytes/rssAfterBytes; parent-process RSS and heap deltas are omitted for external-process runners. In-process runners report peakRssBytes as the maximum observed parent RSS at before/after/after-GC checkpoints.',
			},
			externalRunners: loadedCases.externalRunnerSpecs,
			skipped,
			skippedCases,
			failed,
		},
	})
	if (json) {
		console.log(JSON.stringify(suite, null, 2))
		return
	}
	console.log('')
	console.log(renderSummary(results, [...skipped, ...skippedCases, ...failed]))
}

if (import.meta.main) {
	await main()
}
