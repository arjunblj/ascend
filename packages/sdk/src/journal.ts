import {
	type AutoFilter,
	type Cell,
	type CellStyle,
	type ChartPartInfo,
	type ChartSeriesInfo,
	cloneCellStyle,
	clonePivotCacheInfo,
	DEFAULT_STYLE_ID,
	type DefinedName,
	type DefinedNameScope,
	indexToColumn,
	type PivotCacheInfo,
	parseA1,
	parseRange,
	type RangeRef,
	type Sheet,
	type SheetColDef,
	type SheetComment,
	type SheetConditionalFormat,
	type SheetConditionalFormatRule,
	type SheetConditionalFormatValueObject,
	type SheetDataValidation,
	type SheetDrawingObjectRef,
	type SheetFormatPr,
	type SheetHyperlink,
	type SheetImageAnchor,
	type SheetOutlinePr,
	type SheetPageMargins,
	type SheetPageSetup,
	type SheetProtection,
	type SheetRowDef,
	type SheetState,
	type SheetTabColor,
	type SheetThreadedComment,
	type Table,
	type TableColumn,
	type TableStyleInfo,
	toA1,
	toRangeString,
	type Workbook,
	type WorkbookDocumentProperties,
	type WorkbookProperties,
	type WorkbookProtection,
	type WorkbookThemeColor,
	type WorkbookThemeMetadata,
	type WorkbookView,
} from '@ascend/core'
import { applyOperation } from '@ascend/engine'
import {
	cachedParseFormula,
	compareValues,
	dateToSerial,
	extractRefs,
	type FormulaRef,
	normalizeFormulaInput,
	serialToDate,
} from '@ascend/formulas'
import type {
	CalcSettings,
	CellValue,
	ConditionalFormatRule,
	DataValidationRule,
	InputValue,
	Operation,
	ScalarCellValue,
	StyleInput,
} from '@ascend/schema'
import {
	EMPTY,
	validateExcelDefinedName,
	validateExcelTableName,
	validateExcelWorksheetName,
} from '@ascend/schema'

export type MutationJournalSurface =
	| 'cells'
	| 'formulas'
	| 'formula-bindings'
	| 'shared-formulas'
	| 'dynamic-arrays'
	| 'legacy-arrays'
	| 'data-tables'
	| 'spills'
	| 'tables'
	| 'defined-names'
	| 'comments'
	| 'hyperlinks'
	| 'data-validations'
	| 'conditional-formats'
	| 'auto-filters'
	| 'merged-cells'
	| 'row-layout'
	| 'column-layout'
	| 'page-setup'
	| 'sheet-layout'
	| 'x14-metadata'
	| 'drawings'
	| 'charts'
	| 'pivot-caches'
	| 'workbook-metadata'
	| 'package-parts'

export type MutationJournalExactness = 'exact' | 'conditional' | 'lossy'

export type MutationJournalPublicInverse = 'exact' | 'conditional' | 'none'

export const MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION = 1

export type MutationJournalReasonCode =
	| 'operation-unsupported'
	| 'value-unsupported'
	| 'journal-build-failed'
	| 'journal-unavailable'
	| 'formula-binding-metadata'
	| 'formula-cache-unsupported-value'
	| 'formula-reference-rewrite'
	| 'rich-text-unsupported-runs'
	| 'data-validation-default-attributes'
	| 'metadata-order'
	| 'metadata-duplicate'
	| 'metadata-collision'
	| 'merge-overlap'
	| 'x14-metadata'
	| 'auto-filter-column-metadata'
	| 'auto-filter-extension-metadata'
	| 'auto-filter-sort-metadata'
	| 'legacy-comment-drawing'
	| 'comment-author-removal'
	| 'threaded-comment-selector'
	| 'drawing-text-selector'
	| 'chart-series-unsettable'
	| 'pivot-cache-unsettable'
	| 'defined-name-metadata'
	| 'page-setup-unsettable'
	| 'page-margins-unsettable'
	| 'row-layout-created'
	| 'row-layout-custom-height'
	| 'column-layout-created'
	| 'column-layout-width-metadata'
	| 'table-metadata'
	| 'sheet-topology'
	| 'workbook-protection-absence'
	| 'package-part-preservation'
	| 'partial-workbook'

export const MUTATION_JOURNAL_REASON_DESCRIPTIONS: Record<MutationJournalReasonCode, string> = {
	'operation-unsupported': 'No public inverse operation exists for this edit.',
	'value-unsupported': 'The preimage value cannot be expressed through public cell operations.',
	'journal-build-failed': 'Journal construction failed before a stable inverse could be built.',
	'journal-unavailable': 'The workbook state cannot provide a journal for this edit.',
	'formula-binding-metadata':
		'Imported formulaInfo metadata cannot be reconstructed by public formula operations.',
	'formula-cache-unsupported-value':
		'The formula cache preimage cannot be restored by public cell operations.',
	'formula-reference-rewrite':
		'Formula reference rewrites changed metadata that public operations cannot fully reverse.',
	'rich-text-unsupported-runs': 'The rich text run preimage is not representable by setRichText.',
	'data-validation-default-attributes':
		'Materialized validation defaults cannot be restored exactly.',
	'metadata-order': 'The metadata list order cannot be restored exactly.',
	'metadata-duplicate': 'Duplicate metadata cannot be targeted exactly by public operations.',
	'metadata-collision': 'Transferred metadata collides with existing target metadata.',
	'merge-overlap': 'Merge metadata partially overlaps an edited range or target.',
	'x14-metadata': 'x14 extension metadata has no public inverse operation.',
	'auto-filter-column-metadata':
		'AutoFilter column metadata is not fully representable by public filter operations.',
	'auto-filter-extension-metadata':
		'AutoFilter extension metadata has no public inverse operation.',
	'auto-filter-sort-metadata':
		'AutoFilter sort metadata is not fully representable by public filter operations.',
	'legacy-comment-drawing': 'Legacy VML comment drawing metadata has no public inverse operation.',
	'comment-author-removal': 'Comment author metadata cannot be removed exactly.',
	'threaded-comment-selector':
		'The threaded comment selector is not stable enough for exact inverse.',
	'drawing-text-selector': 'The drawing selector is not stable enough for exact text inverse.',
	'chart-series-unsettable':
		'A chart series field would need to be unset through public operations.',
	'pivot-cache-unsettable': 'A pivot cache field would need to be unset through public operations.',
	'defined-name-metadata': 'Defined-name metadata exceeds what public name operations can restore.',
	'page-setup-unsettable': 'Page setup metadata would need to be unset through public operations.',
	'page-margins-unsettable':
		'Page margin metadata would need to be unset through public operations.',
	'row-layout-created': 'Created row layout metadata cannot be cleared exactly.',
	'row-layout-custom-height': 'Imported row customHeight metadata cannot be restored exactly.',
	'column-layout-created': 'Created column layout metadata cannot be cleared exactly.',
	'column-layout-width-metadata': 'Imported column width metadata cannot be restored exactly.',
	'table-metadata': 'Table metadata exceeds what public table operations can restore.',
	'sheet-topology': 'Sheet topology metadata cannot be reconstructed through public operations.',
	'workbook-protection-absence': 'Workbook or sheet protection absence cannot be restored exactly.',
	'package-part-preservation': 'Package-preserved metadata has no public inverse operation.',
	'partial-workbook': 'Partial workbook state cannot prove an exact inverse.',
}

export interface MutationJournalExactnessRule {
	readonly surface: MutationJournalSurface
	readonly exactness: MutationJournalExactness
	readonly publicInverse: MutationJournalPublicInverse
	readonly constraints: readonly string[]
	readonly lossReasons: readonly MutationJournalReasonCode[]
	readonly representativeOps: readonly string[]
}

export type MutationJournalOperationName = Operation['op']

export interface MutationJournalOperationSurfaceRule {
	readonly primarySurface: MutationJournalSurface
	readonly surfaces: readonly MutationJournalSurface[]
}

export const MUTATION_JOURNAL_EXACTNESS_MATRIX: readonly MutationJournalExactnessRule[] = [
	{
		surface: 'cells',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'scalar values restore with setCells',
			'representable rich text restores with setRichText',
			'unsupported cached value kinds make the inverse unsupported',
		],
		lossReasons: ['operation-unsupported', 'value-unsupported', 'rich-text-unsupported-runs'],
		representativeOps: ['setCells', 'clearRange', 'copyRange', 'moveRange', 'sortRange'],
	},
	{
		surface: 'formulas',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'formula text restores with setFormula',
			'formula cache restores only when the cached value is accepted by public cell writes',
			'invalid fillFormula text cannot be journaled exactly',
		],
		lossReasons: [
			'formula-cache-unsupported-value',
			'formula-reference-rewrite',
			'value-unsupported',
		],
		representativeOps: [
			'setFormula',
			'fillFormula',
			'setCells',
			'clearRange',
			'copyRange',
			'moveRange',
			'sortRange',
		],
	},
	{
		surface: 'formula-bindings',
		exactness: 'lossy',
		publicInverse: 'none',
		constraints: [
			'public operations can materialize formula text but cannot recreate imported formulaInfo bindings',
		],
		lossReasons: ['formula-binding-metadata'],
		representativeOps: [
			'setCells',
			'clearRange',
			'fillFormula',
			'copyRange',
			'moveRange',
			'sortRange',
			'renameSheet',
			'renameTable',
			'resizeTable',
			'deleteTable',
			'setTableColumn',
		],
	},
	{
		surface: 'shared-formulas',
		exactness: 'lossy',
		publicInverse: 'none',
		constraints: ['shared formula master/member metadata is formula binding metadata'],
		lossReasons: ['formula-binding-metadata'],
		representativeOps: [
			'setCells',
			'clearRange',
			'copyRange',
			'moveRange',
			'sortRange',
			'renameSheet',
			'renameTable',
			'resizeTable',
			'deleteTable',
			'setTableColumn',
		],
	},
	{
		surface: 'dynamic-arrays',
		exactness: 'lossy',
		publicInverse: 'none',
		constraints: ['dynamic array metadata is formula binding metadata'],
		lossReasons: ['formula-binding-metadata'],
		representativeOps: [
			'setCells',
			'clearRange',
			'copyRange',
			'moveRange',
			'sortRange',
			'renameSheet',
			'renameTable',
			'resizeTable',
			'deleteTable',
			'setTableColumn',
		],
	},
	{
		surface: 'legacy-arrays',
		exactness: 'lossy',
		publicInverse: 'none',
		constraints: ['legacy array formula metadata is formula binding metadata'],
		lossReasons: ['formula-binding-metadata'],
		representativeOps: [
			'setCells',
			'clearRange',
			'copyRange',
			'moveRange',
			'sortRange',
			'renameSheet',
			'renameTable',
			'resizeTable',
			'deleteTable',
			'setTableColumn',
		],
	},
	{
		surface: 'data-tables',
		exactness: 'lossy',
		publicInverse: 'none',
		constraints: ['data-table formula metadata is formula binding metadata'],
		lossReasons: ['formula-binding-metadata'],
		representativeOps: [
			'setCells',
			'clearRange',
			'copyRange',
			'moveRange',
			'sortRange',
			'renameSheet',
			'renameTable',
			'resizeTable',
			'deleteTable',
			'setTableColumn',
		],
	},
	{
		surface: 'spills',
		exactness: 'lossy',
		publicInverse: 'none',
		constraints: ['spill footprint metadata is formula binding metadata'],
		lossReasons: ['formula-binding-metadata'],
		representativeOps: [
			'setCells',
			'clearRange',
			'copyRange',
			'moveRange',
			'sortRange',
			'renameSheet',
			'renameTable',
			'resizeTable',
			'deleteTable',
			'setTableColumn',
		],
	},
	{
		surface: 'tables',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'table lifecycle and editable table fields restore when public operations can address the original table',
			'unsupported table metadata or missing selectors make rollback lossy',
		],
		lossReasons: ['operation-unsupported', 'table-metadata', 'value-unsupported'],
		representativeOps: [
			'createTable',
			'deleteTable',
			'renameTable',
			'resizeTable',
			'setTableColumn',
		],
	},
	{
		surface: 'defined-names',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'plain workbook and sheet scoped names restore with setDefinedName/deleteDefinedName',
			'print areas use the same defined-name inverse path',
			'deleteDefinedName requires a public name that exists in the requested scope',
			'extra imported defined-name metadata is not publicly settable',
		],
		lossReasons: ['defined-name-metadata', 'value-unsupported'],
		representativeOps: ['setDefinedName', 'deleteDefinedName', 'setPrintArea'],
	},
	{
		surface: 'comments',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'simple cell comments restore with setComment/deleteComment',
			'threaded comment selectors and author removals are exact only when public selectors can address the original metadata',
		],
		lossReasons: [
			'value-unsupported',
			'legacy-comment-drawing',
			'comment-author-removal',
			'threaded-comment-selector',
		],
		representativeOps: [
			'setComment',
			'deleteComment',
			'setThreadedComment',
			'copyRange',
			'moveRange',
		],
	},
	{
		surface: 'hyperlinks',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'public hyperlink fields restore with setHyperlink/deleteHyperlink',
			'invalid hyperlink updates without a public destination cannot be journaled exactly',
		],
		lossReasons: ['value-unsupported'],
		representativeOps: ['setHyperlink', 'deleteHyperlink', 'copyRange', 'moveRange'],
	},
	{
		surface: 'data-validations',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'standard validations restore with setDataValidation/deleteDataValidation',
			'ordering, duplicates, defaults, and x14 extension payloads can make rollback lossy',
		],
		lossReasons: [
			'value-unsupported',
			'data-validation-default-attributes',
			'metadata-order',
			'metadata-duplicate',
			'metadata-collision',
			'x14-metadata',
		],
		representativeOps: [
			'setDataValidation',
			'deleteDataValidation',
			'copyRange',
			'moveRange',
			'sortRange',
		],
	},
	{
		surface: 'conditional-formats',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'standard conditional formats restore with setConditionalFormat/deleteConditionalFormat',
			'rule ordering, duplicate ranges, collisions, and x14 extension payloads can make rollback lossy',
		],
		lossReasons: [
			'value-unsupported',
			'metadata-order',
			'metadata-duplicate',
			'metadata-collision',
			'x14-metadata',
		],
		representativeOps: [
			'setConditionalFormat',
			'deleteConditionalFormat',
			'copyRange',
			'moveRange',
			'sortRange',
		],
	},
	{
		surface: 'auto-filters',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'simple range filters and single-condition public sort state restore with setAutoFilter/clearAutoFilter',
			'advanced filter columns, extension metadata, and advanced sort metadata are lossy',
		],
		lossReasons: [
			'value-unsupported',
			'operation-unsupported',
			'auto-filter-column-metadata',
			'auto-filter-extension-metadata',
			'auto-filter-sort-metadata',
		],
		representativeOps: ['setAutoFilter', 'clearAutoFilter'],
	},
	{
		surface: 'merged-cells',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'whole merge add/remove operations restore exactly',
			'partial overlaps, duplicate metadata, and target collisions are lossy or unsupported',
		],
		lossReasons: ['value-unsupported', 'merge-overlap', 'metadata-duplicate', 'metadata-collision'],
		representativeOps: ['mergeCells', 'unmergeCells', 'copyRange', 'moveRange'],
	},
	{
		surface: 'row-layout',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'existing row height and hidden-state edits restore with public operations',
			'created row layout and customHeight=false metadata cannot be cleared exactly',
		],
		lossReasons: ['value-unsupported', 'row-layout-created', 'row-layout-custom-height'],
		representativeOps: [
			'setRowHeight',
			'hideRows',
			'groupRows',
			'insertRows',
			'deleteRows',
			'sortRange',
		],
	},
	{
		surface: 'column-layout',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'existing column width and hidden-state edits restore with public operations',
			'created column layout and unsupported width metadata cannot be cleared exactly',
		],
		lossReasons: ['value-unsupported', 'column-layout-created', 'column-layout-width-metadata'],
		representativeOps: ['setColWidth', 'hideCols', 'groupCols', 'insertCols', 'deleteCols'],
	},
	{
		surface: 'page-setup',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'representable page setup and margin fields restore with setPageSetup',
			'absence and imported-only page metadata cannot be restored exactly',
		],
		lossReasons: ['value-unsupported', 'page-setup-unsettable', 'page-margins-unsettable'],
		representativeOps: ['setPageSetup', 'setPrintArea'],
	},
	{
		surface: 'sheet-layout',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'sheet position, pane, visibility, tab color, and supported protection fields restore with public operations',
			'deleted sheets and unsupported sheet metadata are package-preservation risks',
		],
		lossReasons: ['value-unsupported', 'sheet-topology', 'workbook-protection-absence'],
		representativeOps: ['renameSheet', 'moveSheet', 'deleteSheet', 'freezePane', 'setTabColor'],
	},
	{
		surface: 'x14-metadata',
		exactness: 'lossy',
		publicInverse: 'none',
		constraints: [
			'x14 conditional-format, data-validation, and sparkline extension payloads are preserved package metadata with no public inverse operation',
		],
		lossReasons: ['operation-unsupported', 'x14-metadata'],
		representativeOps: [
			'copyRange',
			'moveRange',
			'sortRange',
			'insertRows',
			'deleteRows',
			'setSparklineGroup',
		],
	},
	{
		surface: 'drawings',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'drawing text restores when the public selector uniquely identifies an editable text-bearing drawing object',
			'other drawing payloads remain package-preserved metadata',
		],
		lossReasons: ['operation-unsupported', 'drawing-text-selector', 'package-part-preservation'],
		representativeOps: ['setDrawingText', 'insertImage', 'replaceImage', 'deleteImage'],
	},
	{
		surface: 'charts',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'chart series refs restore when every touched field had a prior public value',
			'unsetting chart refs and unrelated chart payloads remain package-preserved metadata',
		],
		lossReasons: ['chart-series-unsettable', 'package-part-preservation'],
		representativeOps: ['setChartSeriesSource'],
	},
	{
		surface: 'pivot-caches',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'pivot cache source and refresh fields restore when every touched field had a prior public value',
			'pivot cache edits require a public selector, at least one editable field, and valid source refs',
			'unsetting fields and pivot cache records remain package-preserved metadata',
		],
		lossReasons: [
			'operation-unsupported',
			'pivot-cache-unsettable',
			'package-part-preservation',
			'value-unsupported',
		],
		representativeOps: ['setPivotCache'],
	},
	{
		surface: 'workbook-metadata',
		exactness: 'conditional',
		publicInverse: 'conditional',
		constraints: [
			'public workbook properties, document properties, views, calc settings, theme colors, and protection restore when represented by public operations',
			'absence or unsupported imported metadata remains lossy',
		],
		lossReasons: ['workbook-protection-absence', 'value-unsupported'],
		representativeOps: [
			'setWorkbookProperties',
			'setDocumentProperties',
			'setWorkbookView',
			'setCalcSettings',
			'setTheme',
		],
	},
	{
		surface: 'package-parts',
		exactness: 'lossy',
		publicInverse: 'none',
		constraints: [
			'unknown or preserved package parts cannot be reconstructed through public workbook operations',
		],
		lossReasons: [
			'operation-unsupported',
			'journal-build-failed',
			'journal-unavailable',
			'package-part-preservation',
			'partial-workbook',
		],
		representativeOps: ['unsupported', 'preserved-package-part'],
	},
]

const MUTATION_JOURNAL_EXACTNESS_BY_SURFACE = new Map(
	MUTATION_JOURNAL_EXACTNESS_MATRIX.map((rule) => [rule.surface, rule]),
)

export const MUTATION_JOURNAL_SURFACES: readonly MutationJournalSurface[] =
	MUTATION_JOURNAL_EXACTNESS_MATRIX.map((rule) => rule.surface)

export const MUTATION_JOURNAL_REASON_CODES = Object.keys(
	MUTATION_JOURNAL_REASON_DESCRIPTIONS,
) as readonly MutationJournalReasonCode[]

export const MUTATION_JOURNAL_ISSUE_CODES = [
	'UNSUPPORTED_OPERATION',
	'LOSSY_INVERSE',
	'UNSUPPORTED_VALUE',
	'JOURNAL_UNAVAILABLE',
	'JOURNAL_BUILD_FAILED',
] as const

export const MUTATION_JOURNAL_ISSUE_SCHEMA = {
	$schema: 'https://json-schema.org/draft/2020-12/schema',
	$id: 'https://ascend.dev/schemas/mutation-journal-issue-v1.json',
	title: 'Ascend mutation journal issue v1',
	type: 'object',
	additionalProperties: false,
	required: ['code', 'message', 'surface', 'reason'],
	properties: {
		code: { type: 'string', enum: MUTATION_JOURNAL_ISSUE_CODES },
		message: { type: 'string' },
		surface: { type: 'string', enum: MUTATION_JOURNAL_SURFACES },
		reason: { type: 'string', enum: MUTATION_JOURNAL_REASON_CODES },
		refs: { type: 'array', items: { type: 'string' } },
	},
} as const

export const MUTATION_JOURNAL_OPERATION_SURFACE_RULES = {
	setCells: {
		primarySurface: 'cells',
		surfaces: [
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
		],
	},
	setFormula: {
		primarySurface: 'formulas',
		surfaces: [
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
		],
	},
	fillFormula: {
		primarySurface: 'formulas',
		surfaces: [
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
		],
	},
	clearRange: {
		primarySurface: 'cells',
		surfaces: [
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
		],
	},
	insertRows: {
		primarySurface: 'row-layout',
		surfaces: [
			'row-layout',
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'tables',
			'defined-names',
			'data-validations',
			'conditional-formats',
			'auto-filters',
			'merged-cells',
			'x14-metadata',
			'drawings',
			'charts',
		],
	},
	deleteRows: {
		primarySurface: 'row-layout',
		surfaces: [
			'row-layout',
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'tables',
			'defined-names',
			'data-validations',
			'conditional-formats',
			'auto-filters',
			'merged-cells',
			'x14-metadata',
			'drawings',
			'charts',
		],
	},
	insertCols: {
		primarySurface: 'column-layout',
		surfaces: [
			'column-layout',
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'tables',
			'defined-names',
			'data-validations',
			'conditional-formats',
			'auto-filters',
			'merged-cells',
			'x14-metadata',
			'drawings',
			'charts',
		],
	},
	deleteCols: {
		primarySurface: 'column-layout',
		surfaces: [
			'column-layout',
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'tables',
			'defined-names',
			'data-validations',
			'conditional-formats',
			'auto-filters',
			'merged-cells',
			'x14-metadata',
			'drawings',
			'charts',
		],
	},
	addSheet: { primarySurface: 'sheet-layout', surfaces: ['sheet-layout', 'package-parts'] },
	deleteSheet: {
		primarySurface: 'sheet-layout',
		surfaces: [
			'sheet-layout',
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'tables',
			'defined-names',
			'comments',
			'hyperlinks',
			'data-validations',
			'conditional-formats',
			'auto-filters',
			'merged-cells',
			'row-layout',
			'column-layout',
			'page-setup',
			'x14-metadata',
			'drawings',
			'charts',
			'pivot-caches',
			'package-parts',
		],
	},
	renameSheet: {
		primarySurface: 'sheet-layout',
		surfaces: [
			'sheet-layout',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'defined-names',
			'data-validations',
			'conditional-formats',
			'charts',
			'package-parts',
		],
	},
	moveSheet: { primarySurface: 'sheet-layout', surfaces: ['sheet-layout', 'package-parts'] },
	createTable: { primarySurface: 'tables', surfaces: ['tables'] },
	appendRows: { primarySurface: 'tables', surfaces: ['tables', 'cells', 'formulas'] },
	sortRange: {
		primarySurface: 'cells',
		surfaces: [
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'comments',
			'hyperlinks',
			'data-validations',
			'conditional-formats',
			'row-layout',
			'x14-metadata',
		],
	},
	mergeCells: { primarySurface: 'merged-cells', surfaces: ['merged-cells'] },
	unmergeCells: { primarySurface: 'merged-cells', surfaces: ['merged-cells'] },
	setColWidth: { primarySurface: 'column-layout', surfaces: ['column-layout'] },
	setRowHeight: { primarySurface: 'row-layout', surfaces: ['row-layout'] },
	setComment: { primarySurface: 'comments', surfaces: ['comments'] },
	setHyperlink: { primarySurface: 'hyperlinks', surfaces: ['hyperlinks'] },
	setNumberFormat: { primarySurface: 'cells', surfaces: ['cells'] },
	setDefinedName: {
		primarySurface: 'defined-names',
		surfaces: ['defined-names', 'package-parts'],
	},
	deleteDefinedName: {
		primarySurface: 'defined-names',
		surfaces: ['defined-names', 'package-parts'],
	},
	setStyle: { primarySurface: 'cells', surfaces: ['cells'] },
	freezePane: { primarySurface: 'sheet-layout', surfaces: ['sheet-layout'] },
	deleteComment: { primarySurface: 'comments', surfaces: ['comments'] },
	deleteHyperlink: { primarySurface: 'hyperlinks', surfaces: ['hyperlinks'] },
	setDataValidation: { primarySurface: 'data-validations', surfaces: ['data-validations'] },
	deleteDataValidation: { primarySurface: 'data-validations', surfaces: ['data-validations'] },
	setAutoFilter: { primarySurface: 'auto-filters', surfaces: ['auto-filters'] },
	clearAutoFilter: { primarySurface: 'auto-filters', surfaces: ['auto-filters'] },
	setSheetProtection: { primarySurface: 'sheet-layout', surfaces: ['sheet-layout'] },
	setTabColor: { primarySurface: 'sheet-layout', surfaces: ['sheet-layout'] },
	hideSheet: { primarySurface: 'sheet-layout', surfaces: ['sheet-layout'] },
	hideRows: { primarySurface: 'row-layout', surfaces: ['row-layout'] },
	hideCols: { primarySurface: 'column-layout', surfaces: ['column-layout'] },
	copySheet: {
		primarySurface: 'sheet-layout',
		surfaces: [
			'sheet-layout',
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'tables',
			'defined-names',
			'comments',
			'hyperlinks',
			'data-validations',
			'conditional-formats',
			'auto-filters',
			'merged-cells',
			'row-layout',
			'column-layout',
			'page-setup',
			'x14-metadata',
			'drawings',
			'charts',
			'package-parts',
		],
	},
	setConditionalFormat: {
		primarySurface: 'conditional-formats',
		surfaces: ['conditional-formats'],
	},
	deleteConditionalFormat: {
		primarySurface: 'conditional-formats',
		surfaces: ['conditional-formats'],
	},
	setPageSetup: { primarySurface: 'page-setup', surfaces: ['page-setup'] },
	setPrintArea: { primarySurface: 'defined-names', surfaces: ['defined-names', 'page-setup'] },
	copyRange: {
		primarySurface: 'cells',
		surfaces: [
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'comments',
			'hyperlinks',
			'data-validations',
			'conditional-formats',
			'merged-cells',
			'x14-metadata',
		],
	},
	moveRange: {
		primarySurface: 'cells',
		surfaces: [
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'defined-names',
			'comments',
			'hyperlinks',
			'data-validations',
			'conditional-formats',
			'merged-cells',
			'x14-metadata',
		],
	},
	groupRows: { primarySurface: 'row-layout', surfaces: ['row-layout'] },
	groupCols: { primarySurface: 'column-layout', surfaces: ['column-layout'] },
	setRichText: {
		primarySurface: 'cells',
		surfaces: [
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
		],
	},
	setWorkbookProperties: { primarySurface: 'workbook-metadata', surfaces: ['workbook-metadata'] },
	setDocumentProperties: {
		primarySurface: 'workbook-metadata',
		surfaces: ['workbook-metadata', 'package-parts'],
	},
	setWorkbookView: { primarySurface: 'workbook-metadata', surfaces: ['workbook-metadata'] },
	setCalcSettings: { primarySurface: 'workbook-metadata', surfaces: ['workbook-metadata'] },
	setTheme: { primarySurface: 'workbook-metadata', surfaces: ['workbook-metadata'] },
	setWorkbookProtection: { primarySurface: 'workbook-metadata', surfaces: ['workbook-metadata'] },
	deleteTable: {
		primarySurface: 'tables',
		surfaces: [
			'tables',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'defined-names',
		],
	},
	renameTable: {
		primarySurface: 'tables',
		surfaces: [
			'tables',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'defined-names',
		],
	},
	resizeTable: {
		primarySurface: 'tables',
		surfaces: [
			'tables',
			'cells',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'data-validations',
			'conditional-formats',
		],
	},
	setTableColumn: {
		primarySurface: 'tables',
		surfaces: [
			'tables',
			'formulas',
			'formula-bindings',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
		],
	},
	setTableStyle: { primarySurface: 'tables', surfaces: ['tables'] },
	replaceImage: { primarySurface: 'drawings', surfaces: ['drawings', 'package-parts'] },
	insertImage: { primarySurface: 'drawings', surfaces: ['drawings', 'package-parts'] },
	deleteImage: { primarySurface: 'drawings', surfaces: ['drawings', 'package-parts'] },
	setDrawingText: { primarySurface: 'drawings', surfaces: ['drawings'] },
	setThreadedComment: { primarySurface: 'comments', surfaces: ['comments'] },
	setChartSeriesSource: { primarySurface: 'charts', surfaces: ['charts'] },
	setPivotCache: { primarySurface: 'pivot-caches', surfaces: ['pivot-caches'] },
	setPivotFieldItem: { primarySurface: 'pivot-caches', surfaces: ['pivot-caches'] },
	setSlicerCacheItem: { primarySurface: 'pivot-caches', surfaces: ['pivot-caches'] },
	setTimelineRange: { primarySurface: 'pivot-caches', surfaces: ['pivot-caches'] },
	setSparklineGroup: { primarySurface: 'x14-metadata', surfaces: ['x14-metadata'] },
	setAdvancedFilter: { primarySurface: 'auto-filters', surfaces: ['auto-filters'] },
	setConnectionRefresh: { primarySurface: 'package-parts', surfaces: ['package-parts'] },
	rewriteExternalLink: { primarySurface: 'package-parts', surfaces: ['package-parts', 'formulas'] },
} as const satisfies Readonly<
	Record<MutationJournalOperationName, MutationJournalOperationSurfaceRule>
>

export function classifyMutationJournalOperationPrimarySurface(
	op: Operation | MutationJournalOperationName,
): MutationJournalSurface {
	const opName = typeof op === 'string' ? op : op.op
	return MUTATION_JOURNAL_OPERATION_SURFACE_RULES[opName].primarySurface
}

export function classifyMutationJournalOperationSurfaces(
	op: Operation | MutationJournalOperationName,
): readonly MutationJournalSurface[] {
	const opName = typeof op === 'string' ? op : op.op
	return MUTATION_JOURNAL_OPERATION_SURFACE_RULES[opName].surfaces
}

export function classifyMutationJournalSurface(
	surface: MutationJournalSurface,
): MutationJournalExactnessRule {
	const rule = MUTATION_JOURNAL_EXACTNESS_BY_SURFACE.get(surface)
	if (!rule) throw new Error(`Unknown mutation journal surface: ${surface}`)
	return rule
}

export interface MutationJournalIssue {
	readonly code:
		| 'UNSUPPORTED_OPERATION'
		| 'LOSSY_INVERSE'
		| 'UNSUPPORTED_VALUE'
		| 'JOURNAL_UNAVAILABLE'
		| 'JOURNAL_BUILD_FAILED'
	readonly message: string
	readonly surface?: MutationJournalSurface
	readonly reason?: MutationJournalReasonCode
	readonly refs?: readonly string[]
}

export type MutationJournalStructuredIssue = MutationJournalIssue & {
	readonly surface: MutationJournalSurface
	readonly reason: MutationJournalReasonCode
}

export interface MutationJournalIssueClassification {
	readonly surface: MutationJournalSurface
	readonly reason: MutationJournalReasonCode
	readonly exactness: MutationJournalExactness
	readonly publicInverse: MutationJournalPublicInverse
}

export function classifyMutationJournalIssue(
	issue: MutationJournalIssue,
): MutationJournalIssueClassification {
	const surface = issue.surface ?? inferMutationJournalIssueSurface(issue)
	const reason = issue.reason ?? inferMutationJournalIssueReason(issue, surface)
	const rule = classifyMutationJournalSurface(surface)
	return {
		surface,
		reason,
		exactness: rule.exactness,
		publicInverse: rule.publicInverse,
	}
}

export function classifyMutationJournalIssues(
	issues: readonly MutationJournalIssue[],
): readonly MutationJournalIssueClassification[] {
	return issues.map((issue) => classifyMutationJournalIssue(issue))
}

export function structureMutationJournalIssue(
	issue: MutationJournalIssue,
): MutationJournalStructuredIssue {
	const classification = classifyMutationJournalIssue(issue)
	return {
		...issue,
		surface: classification.surface,
		reason: classification.reason,
	}
}

function inferMutationJournalIssueSurface(issue: MutationJournalIssue): MutationJournalSurface {
	if (issue.code === 'JOURNAL_BUILD_FAILED' || issue.code === 'JOURNAL_UNAVAILABLE') {
		return 'package-parts'
	}
	if (issue.code === 'UNSUPPORTED_OPERATION') {
		return unsupportedOperationSurface(issue) ?? 'package-parts'
	}
	const text = journalIssueSearchText(issue)
	if (text.includes('x14')) return 'x14-metadata'
	if (text.includes('shared formula')) return 'shared-formulas'
	if (text.includes('dynamic array')) return 'dynamic-arrays'
	if (text.includes('legacy array')) return 'legacy-arrays'
	if (text.includes('data table')) return 'data-tables'
	if (text.includes('spill')) return 'spills'
	if (text.includes('formula binding') || text.includes('formulainfo')) return 'formula-bindings'
	if (text.includes('formula')) return 'formulas'
	if (text.includes('data validation') || text.includes('validation')) return 'data-validations'
	if (text.includes('conditional format') || text.includes('conditional-format')) {
		return 'conditional-formats'
	}
	if (text.includes('autofilter') || text.includes('auto filter')) return 'auto-filters'
	if (text.includes('merge')) return 'merged-cells'
	if (
		text.includes('row layout') ||
		text.includes('row height') ||
		text.includes('row metadata') ||
		text.includes('row hidden') ||
		text.includes(' rows')
	) {
		return 'row-layout'
	}
	if (
		text.includes('column layout') ||
		text.includes('column width') ||
		text.includes('column metadata') ||
		text.includes('column hidden') ||
		text.includes(' columns') ||
		text.includes(' col layout') ||
		text.includes(' cols')
	) {
		return 'column-layout'
	}
	if (text.includes('threaded comment') || text.includes('comment')) return 'comments'
	if (text.includes('hyperlink')) return 'hyperlinks'
	if (text.includes('drawing')) return 'drawings'
	if (text.includes('chart')) return 'charts'
	if (text.includes('pivot')) return 'pivot-caches'
	if (text.includes('table')) return 'tables'
	if (text.includes('defined name') || text.includes('print area') || text.includes('name:')) {
		return 'defined-names'
	}
	if (text.includes('page setup') || text.includes('page margins')) return 'page-setup'
	if (text.includes('workbook')) return 'workbook-metadata'
	if (text.includes('sheet')) return 'sheet-layout'
	if (issue.code === 'UNSUPPORTED_VALUE') return 'cells'
	return 'package-parts'
}

function inferMutationJournalIssueReason(
	issue: MutationJournalIssue,
	surface: MutationJournalSurface,
): MutationJournalReasonCode {
	const text = journalIssueSearchText(issue)
	if (issue.code === 'UNSUPPORTED_OPERATION') return 'operation-unsupported'
	if (issue.code === 'JOURNAL_BUILD_FAILED') return 'journal-build-failed'
	if (issue.code === 'JOURNAL_UNAVAILABLE') return 'journal-unavailable'
	if (text.includes('x14')) return 'x14-metadata'
	if (text.includes('formula binding') || text.includes('formulainfo')) {
		return 'formula-binding-metadata'
	}
	if (text.includes('formula reference rewrite') || text.includes('formula reference rewrites')) {
		return 'formula-reference-rewrite'
	}
	if (text.includes('formula cache')) return 'formula-cache-unsupported-value'
	if (text.includes('richtext')) return 'rich-text-unsupported-runs'
	if (text.includes('default attributes')) return 'data-validation-default-attributes'
	if (text.includes('autofilter column')) return 'auto-filter-column-metadata'
	if (text.includes('autofilter extension metadata')) return 'auto-filter-extension-metadata'
	if (text.includes('autofilter sort metadata')) return 'auto-filter-sort-metadata'
	if (text.includes('legacy comment drawing')) return 'legacy-comment-drawing'
	if (text.includes('threaded comment selector')) return 'threaded-comment-selector'
	if (text.includes('comment author')) return 'comment-author-removal'
	if (text.includes('drawing object selector')) return 'drawing-text-selector'
	if (text.includes('chart series selector')) return 'chart-series-unsettable'
	if (text.includes('pivot cache selector')) return 'pivot-cache-unsettable'
	if (text.includes('page setup')) return 'page-setup-unsettable'
	if (text.includes('page margins')) return 'page-margins-unsettable'
	if (text.includes('created row layout')) return 'row-layout-created'
	if (
		text.includes('row metadata') ||
		text.includes('row hidden') ||
		text.includes('grouped rows')
	) {
		return 'row-layout-created'
	}
	if (text.includes('customheight=false')) return 'row-layout-custom-height'
	if (text.includes('created col layout') || text.includes('created column layout')) {
		return 'column-layout-created'
	}
	if (
		text.includes('column metadata') ||
		text.includes('column hidden') ||
		text.includes('grouped columns')
	) {
		return 'column-layout-created'
	}
	if (text.includes('column width')) return 'column-layout-width-metadata'
	if (text.includes('rule ordering') || text.includes('order')) return 'metadata-order'
	if (text.includes('duplicate')) return 'metadata-duplicate'
	if (text.includes('collides') || text.includes('collision')) return 'metadata-collision'
	if (text.includes('partial') || text.includes('overlap')) return 'merge-overlap'
	if (text.includes('defined name') || text.includes('print area')) return 'defined-name-metadata'
	if (text.includes('table')) return 'table-metadata'
	if (text.includes('deleted sheet') || text.includes('sheet topology')) return 'sheet-topology'
	if (text.includes('workbook protection') || text.includes('sheet protection')) {
		return 'workbook-protection-absence'
	}
	if (surface === 'package-parts') return 'package-part-preservation'
	if (issue.code === 'UNSUPPORTED_VALUE') return 'value-unsupported'
	return surfaceDefaultLossReason(surface)
}

function surfaceDefaultLossReason(surface: MutationJournalSurface): MutationJournalReasonCode {
	switch (surface) {
		case 'cells':
			return 'value-unsupported'
		case 'formulas':
			return 'formula-cache-unsupported-value'
		case 'formula-bindings':
		case 'shared-formulas':
		case 'dynamic-arrays':
		case 'legacy-arrays':
		case 'data-tables':
		case 'spills':
			return 'formula-binding-metadata'
		case 'tables':
			return 'table-metadata'
		case 'defined-names':
			return 'defined-name-metadata'
		case 'comments':
			return 'threaded-comment-selector'
		case 'data-validations':
		case 'conditional-formats':
			return 'metadata-order'
		case 'auto-filters':
			return 'auto-filter-column-metadata'
		case 'merged-cells':
			return 'merge-overlap'
		case 'row-layout':
			return 'row-layout-created'
		case 'column-layout':
			return 'column-layout-created'
		case 'page-setup':
			return 'page-setup-unsettable'
		case 'x14-metadata':
			return 'x14-metadata'
		case 'drawings':
			return 'drawing-text-selector'
		case 'charts':
			return 'chart-series-unsettable'
		case 'pivot-caches':
			return 'pivot-cache-unsettable'
		case 'sheet-layout':
			return 'sheet-topology'
		case 'workbook-metadata':
			return 'value-unsupported'
		case 'package-parts':
			return 'package-part-preservation'
		case 'hyperlinks':
			return 'value-unsupported'
	}
}

function journalIssueSearchText(issue: MutationJournalIssue): string {
	return `${issue.message} ${(issue.refs ?? []).join(' ')}`.toLowerCase()
}

function unsupportedOperationSurface(issue: MutationJournalIssue): MutationJournalSurface | null {
	const match = /^No reversible journal support for ([A-Za-z0-9]+)(?:\b|$)/.exec(issue.message)
	if (!match) return null
	const opName = match[1]
	if (!opName || !isMutationJournalOperationName(opName)) return null
	return classifyMutationJournalOperationPrimarySurface(opName)
}

function isMutationJournalOperationName(name: string): name is MutationJournalOperationName {
	return Object.hasOwn(MUTATION_JOURNAL_OPERATION_SURFACE_RULES, name)
}

export interface MutationJournalCellPreimage {
	readonly sheet: string
	readonly ref: string
	readonly dateSystem?: '1900' | '1904'
	readonly existed: boolean
	readonly value: CellValue
	readonly formula: string | null
	readonly formulaInfo?: Cell['formulaInfo']
	readonly styleId: number
	readonly style: CellStyle
}

export interface MutationJournalCommentPreimage {
	readonly sheet: string
	readonly ref: string
	readonly comment: SheetComment | null
}

export interface MutationJournalHyperlinkPreimage {
	readonly sheet: string
	readonly ref: string
	readonly hyperlink: SheetHyperlink | null
}

export interface MutationJournalThreadedCommentPreimage {
	readonly sheet: string
	readonly commentIndex: number | null
	readonly threadedComment: SheetThreadedComment | null
}

export interface MutationJournalDrawingTextPreimage {
	readonly sheet: string
	readonly drawingObjectIndex: number | null
	readonly drawingObject: SheetDrawingObjectRef | null
}

export interface MutationJournalChartSeriesPreimage {
	readonly chartIndex: number | null
	readonly seriesIndex: number
	readonly chart: ChartPartInfo | null
	readonly series: ChartSeriesInfo | null
}

export interface MutationJournalPivotCachePreimage {
	readonly cacheIndex: number | null
	readonly cache: PivotCacheInfo | null
}

export interface MutationJournalPanePreimage {
	readonly sheet: string
	readonly frozenRows: number
	readonly frozenCols: number
}

export interface MutationJournalMergePreimage {
	readonly sheet: string
	readonly range: string
	readonly existed: boolean
}

export interface MutationJournalAutoFilterPreimage {
	readonly sheet: string
	readonly autoFilter: AutoFilter | null
}

export interface MutationJournalDataValidationPreimage {
	readonly sheet: string
	readonly range: string
	readonly validation: SheetDataValidation | null
}

export interface MutationJournalConditionalFormatPreimage {
	readonly sheet: string
	readonly range?: string
	readonly formats: readonly SheetConditionalFormat[]
}

export interface MutationJournalDefinedNamePreimage {
	readonly name: string
	readonly scope?: string
	readonly definedName: DefinedName | null
}

export interface MutationJournalPageSetupPreimage {
	readonly sheet: string
	readonly pageSetup: SheetPageSetup | null
	readonly pageMargins: SheetPageMargins | null
}

export interface MutationJournalTableRenamePreimage {
	readonly table: string
	readonly existed: boolean
}

export interface MutationJournalTableColumnPreimage {
	readonly table: string
	readonly column: string | number
	readonly columnIndex: number
	readonly columnState: TableColumn | null
}

export interface MutationJournalTablePreimage {
	readonly table: Table | null
	readonly sheet: string | null
	readonly ref: string | null
}

export interface MutationJournalTableStylePreimage {
	readonly table: string
	readonly style: TableStyleInfo | null
}

export interface MutationJournalSheetMovePreimage {
	readonly sheet: string
	readonly position: number | null
}

export interface MutationJournalSheetDeletePreimage {
	readonly sheet: string
	readonly position: number | null
	readonly existed: boolean
}

export interface MutationJournalSheetLayoutPreimage {
	readonly sheet: string
	readonly axis: 'row' | 'col'
	readonly index: number
	readonly value: number | null
}

export interface MutationJournalSheetTabColorPreimage {
	readonly sheet: string
	readonly tabColor: SheetTabColor | null
}

export interface MutationJournalSheetProtectionPreimage {
	readonly sheet: string
	readonly protection: SheetProtection | null
}

export interface MutationJournalSheetVisibilityPreimage {
	readonly sheet: string
	readonly state: SheetState | null
}

export interface MutationJournalRowsHiddenPreimage {
	readonly sheet: string
	readonly rows: readonly {
		readonly row: number
		readonly height: number | null
		readonly rowDef: SheetRowDef | null
	}[]
}

export interface MutationJournalColsHiddenPreimage {
	readonly sheet: string
	readonly cols: readonly {
		readonly col: number
		readonly colDef: SheetColDef | null
	}[]
}

export interface MutationJournalOutlinePreimage {
	readonly sheet: string
	readonly axis: 'row' | 'col'
	readonly from: number
	readonly to: number
	readonly outlinePr: SheetOutlinePr | null
	readonly sheetFormatPr: SheetFormatPr | null
	readonly rowDefs?: readonly {
		readonly row: number
		readonly rowDef: SheetRowDef | null
	}[]
	readonly colDefs?: readonly {
		readonly col: number
		readonly colDef: SheetColDef | null
	}[]
}

export interface MutationJournalStructuralPreimage {
	readonly sheet: string
	readonly axis: 'row' | 'col'
	readonly at: number
	readonly count: number
	readonly deletedCells: readonly MutationJournalCellPreimage[]
}

export interface MutationJournalWorkbookPropertiesPreimage {
	readonly properties: WorkbookProperties
}

export interface MutationJournalDocumentPropertiesPreimage {
	readonly properties: WorkbookDocumentProperties
}

export interface MutationJournalWorkbookViewPreimage {
	readonly index: number
	readonly view: WorkbookView | null
}

export interface MutationJournalCalcSettingsPreimage {
	readonly settings: CalcSettings
	readonly workbookProperties: WorkbookProperties
}

export interface MutationJournalWorkbookProtectionPreimage {
	readonly protection: WorkbookProtection | null
}

export interface MutationJournalThemePreimage {
	readonly metadata: WorkbookThemeMetadata
	readonly colors: readonly WorkbookThemeColor[]
}

export type MutationJournalPreimage =
	| { readonly kind: 'cells'; readonly cells: readonly MutationJournalCellPreimage[] }
	| { readonly kind: 'comment'; readonly comment: MutationJournalCommentPreimage }
	| { readonly kind: 'hyperlink'; readonly hyperlink: MutationJournalHyperlinkPreimage }
	| {
			readonly kind: 'threaded-comment'
			readonly threadedComment: MutationJournalThreadedCommentPreimage
	  }
	| { readonly kind: 'drawing-text'; readonly drawingText: MutationJournalDrawingTextPreimage }
	| { readonly kind: 'chart-series'; readonly chartSeries: MutationJournalChartSeriesPreimage }
	| { readonly kind: 'pivot-cache'; readonly pivotCache: MutationJournalPivotCachePreimage }
	| { readonly kind: 'pane'; readonly pane: MutationJournalPanePreimage }
	| { readonly kind: 'merge'; readonly merge: MutationJournalMergePreimage }
	| { readonly kind: 'auto-filter'; readonly autoFilter: MutationJournalAutoFilterPreimage }
	| {
			readonly kind: 'data-validations'
			readonly validations: readonly MutationJournalDataValidationPreimage[]
	  }
	| {
			readonly kind: 'conditional-formats'
			readonly conditionalFormats: MutationJournalConditionalFormatPreimage
	  }
	| { readonly kind: 'page-setup'; readonly pageSetup: MutationJournalPageSetupPreimage }
	| { readonly kind: 'defined-name'; readonly definedName: MutationJournalDefinedNamePreimage }
	| { readonly kind: 'table-rename'; readonly tableRename: MutationJournalTableRenamePreimage }
	| { readonly kind: 'table-column'; readonly tableColumn: MutationJournalTableColumnPreimage }
	| { readonly kind: 'table'; readonly table: MutationJournalTablePreimage }
	| { readonly kind: 'table-style'; readonly tableStyle: MutationJournalTableStylePreimage }
	| { readonly kind: 'sheet-move'; readonly sheetMove: MutationJournalSheetMovePreimage }
	| { readonly kind: 'sheet-delete'; readonly sheetDelete: MutationJournalSheetDeletePreimage }
	| { readonly kind: 'sheet-layout'; readonly sheetLayout: MutationJournalSheetLayoutPreimage }
	| {
			readonly kind: 'sheet-tab-color'
			readonly sheetTabColor: MutationJournalSheetTabColorPreimage
	  }
	| {
			readonly kind: 'sheet-protection'
			readonly sheetProtection: MutationJournalSheetProtectionPreimage
	  }
	| {
			readonly kind: 'sheet-visibility'
			readonly sheetVisibility: MutationJournalSheetVisibilityPreimage
	  }
	| { readonly kind: 'rows-hidden'; readonly rowsHidden: MutationJournalRowsHiddenPreimage }
	| { readonly kind: 'cols-hidden'; readonly colsHidden: MutationJournalColsHiddenPreimage }
	| { readonly kind: 'outline'; readonly outline: MutationJournalOutlinePreimage }
	| { readonly kind: 'structural'; readonly structural: MutationJournalStructuralPreimage }
	| {
			readonly kind: 'workbook-properties'
			readonly workbookProperties: MutationJournalWorkbookPropertiesPreimage
	  }
	| {
			readonly kind: 'document-properties'
			readonly documentProperties: MutationJournalDocumentPropertiesPreimage
	  }
	| { readonly kind: 'workbook-view'; readonly workbookView: MutationJournalWorkbookViewPreimage }
	| { readonly kind: 'calc-settings'; readonly calcSettings: MutationJournalCalcSettingsPreimage }
	| {
			readonly kind: 'workbook-protection'
			readonly workbookProtection: MutationJournalWorkbookProtectionPreimage
	  }
	| { readonly kind: 'theme'; readonly theme: MutationJournalThemePreimage }

export interface MutationJournalEntry {
	readonly opIndex: number
	readonly op: Operation
	readonly supported: boolean
	readonly exact: boolean
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalStructuredIssue[]
}

export interface MutationJournal {
	readonly schemaVersion: number
	readonly schemaId: string
	readonly entries: readonly MutationJournalEntry[]
	readonly inverseOps: readonly Operation[]
	readonly supported: boolean
	readonly exact: boolean
	readonly issues: readonly MutationJournalStructuredIssue[]
	readonly undoPolicy: MutationJournalUndoPolicy
}

export type MutationJournalUndoRiskLevel = 'none' | 'low' | 'medium' | 'high'

export type MutationJournalUndoPolicyReason =
	| 'exact'
	| 'lossy'
	| 'lossy-without-inverse'
	| 'unsupported'
	| 'unavailable'
	| 'build-failed'

export interface MutationJournalUndoPolicy {
	readonly undoable: boolean
	readonly exact: boolean
	readonly reason: MutationJournalUndoPolicyReason
	readonly userMessage: string
	readonly riskLevel: MutationJournalUndoRiskLevel
}

export interface MutationJournalClassifiedIssue extends MutationJournalIssueClassification {
	readonly code: MutationJournalIssue['code']
	readonly message: string
	readonly refs?: readonly string[]
	readonly allowedByMatrix: boolean
}

export interface MutationJournalExactnessAnalysis {
	readonly supported: boolean
	readonly exact: boolean
	readonly issueCount: number
	readonly issues: readonly MutationJournalClassifiedIssue[]
	readonly operationSurfaces: readonly MutationJournalSurface[]
	readonly primaryOperationSurfaces: readonly MutationJournalSurface[]
	readonly surfaces: readonly MutationJournalSurface[]
	readonly reasons: readonly MutationJournalReasonCode[]
	readonly hasLossyInverse: boolean
	readonly hasUnsupportedOperation: boolean
	readonly hasUnsupportedValue: boolean
	readonly hasUnavailableJournal: boolean
	readonly hasJournalBuildFailure: boolean
	readonly hasMatrixViolation: boolean
}

export function analyzeMutationJournalExactness(
	journal: MutationJournal,
): MutationJournalExactnessAnalysis {
	const issues = journal.issues.map((issue) => classifyMutationJournalIssueForAnalysis(issue))
	const operationSurfaces = uniqueSorted(
		journal.entries.flatMap((entry) => classifyMutationJournalOperationSurfaces(entry.op)),
	)
	const primaryOperationSurfaces = uniqueSorted(
		journal.entries.map((entry) => classifyMutationJournalOperationPrimarySurface(entry.op)),
	)
	return {
		supported: journal.supported,
		exact: journal.exact,
		issueCount: issues.length,
		issues,
		operationSurfaces,
		primaryOperationSurfaces,
		surfaces: uniqueSorted(issues.map((issue) => issue.surface)),
		reasons: uniqueSorted(issues.map((issue) => issue.reason)),
		hasLossyInverse: issues.some((issue) => issue.code === 'LOSSY_INVERSE'),
		hasUnsupportedOperation: issues.some((issue) => issue.code === 'UNSUPPORTED_OPERATION'),
		hasUnsupportedValue: issues.some((issue) => issue.code === 'UNSUPPORTED_VALUE'),
		hasUnavailableJournal: issues.some((issue) => issue.code === 'JOURNAL_UNAVAILABLE'),
		hasJournalBuildFailure: issues.some((issue) => issue.code === 'JOURNAL_BUILD_FAILED'),
		hasMatrixViolation: issues.some((issue) => !issue.allowedByMatrix),
	}
}

export function summarizeMutationJournalUndoPolicy(
	journal: Pick<MutationJournal, 'supported' | 'exact' | 'inverseOps' | 'issues'>,
): MutationJournalUndoPolicy {
	if (journal.supported && journal.exact) {
		return {
			undoable: true,
			exact: true,
			reason: 'exact',
			userMessage: 'Undo available.',
			riskLevel: 'none',
		}
	}
	if (journal.issues.some((issue) => issue.code === 'JOURNAL_BUILD_FAILED')) {
		return {
			undoable: false,
			exact: false,
			reason: 'build-failed',
			userMessage: 'Undo is unavailable because Ascend could not build a journal for this edit.',
			riskLevel: 'high',
		}
	}
	if (journal.issues.some((issue) => issue.code === 'JOURNAL_UNAVAILABLE')) {
		return {
			undoable: false,
			exact: false,
			reason: 'unavailable',
			userMessage: 'Undo is unavailable for this edit.',
			riskLevel: 'high',
		}
	}
	if (!journal.supported) {
		return {
			undoable: false,
			exact: false,
			reason: 'unsupported',
			userMessage: 'Undo is unavailable because this edit has no supported inverse operation.',
			riskLevel: 'high',
		}
	}
	if (journal.inverseOps.length === 0) {
		return {
			undoable: false,
			exact: false,
			reason: 'lossy-without-inverse',
			userMessage: 'Undo is unavailable because no public inverse operations were produced.',
			riskLevel: 'high',
		}
	}
	return {
		undoable: true,
		exact: false,
		reason: 'lossy',
		userMessage: 'Undo available, but it may not restore every workbook detail exactly.',
		riskLevel: 'medium',
	}
}

function classifyMutationJournalIssueForAnalysis(
	issue: MutationJournalIssue,
): MutationJournalClassifiedIssue {
	const classification = classifyMutationJournalIssue(issue)
	const rule = classifyMutationJournalSurface(classification.surface)
	const allowedByMatrix = rule.lossReasons.includes(classification.reason)
	return {
		code: issue.code,
		message: issue.message,
		...(issue.refs && issue.refs.length > 0 ? { refs: issue.refs } : {}),
		...classification,
		allowedByMatrix,
	}
}

function uniqueSorted<T extends string>(values: readonly T[]): readonly T[] {
	return [...new Set(values)].sort()
}

interface DraftJournalEntry {
	readonly opIndex: number
	readonly op: Operation
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
}

export function buildMutationJournal(
	workbook: Workbook,
	ops: readonly Operation[],
): MutationJournal {
	if (ops.length === 1) {
		const op = ops[0]
		if (!op) return emptyMutationJournal()
		return mutationJournalFromEntries(
			savedSourceRecalcPackageStateEntries(workbook, [buildJournalEntry(workbook, op, 0)]),
		)
	}
	const journalWorkbook = workbook.clone()
	const entries: MutationJournalEntry[] = []
	for (let opIndex = 0; opIndex < ops.length; opIndex++) {
		const op = ops[opIndex]
		if (!op) continue
		entries.push(buildJournalEntry(journalWorkbook, op, opIndex))
		const result = applyOperation(journalWorkbook, op)
		if (!result.ok) break
	}
	return mutationJournalFromEntries(savedSourceRecalcPackageStateEntries(workbook, entries))
}

function savedSourceRecalcPackageStateEntries(
	workbook: Workbook,
	entries: readonly MutationJournalEntry[],
): readonly MutationJournalEntry[] {
	if (!workbook.sourceArchiveBytes || entries.length === 0) return entries
	const journalWorkbook = workbook.clone()
	let changed = false
	const updated: MutationJournalEntry[] = []
	for (const entry of entries) {
		const result = applyOperation(journalWorkbook, entry.op)
		const additions =
			result.ok && result.value.recalcRequired
				? savedSourcePackageStateIssues(
						workbook,
						entry.op.op,
						savedSourcePackageStateRefsForOp(entry.op),
					).map(structureMutationJournalIssue)
				: []
		if (additions.length === 0) {
			updated.push(entry)
			if (!result.ok) break
			continue
		}
		const issues = appendMutationJournalIssues(entry.issues, additions)
		changed ||= issues.length !== entry.issues.length
		updated.push({
			...entry,
			issues,
			supported: issues.every((issue) => issue.code !== 'UNSUPPORTED_OPERATION'),
			exact: issues.length === 0,
		})
		if (!result.ok) break
	}
	return changed ? updated : entries
}

function appendMutationJournalIssues(
	issues: readonly MutationJournalStructuredIssue[],
	additions: readonly MutationJournalStructuredIssue[],
): readonly MutationJournalStructuredIssue[] {
	const seen = new Set(issues.map(mutationJournalIssueKey))
	const next = [...issues]
	for (const issue of additions) {
		const key = mutationJournalIssueKey(issue)
		if (seen.has(key)) continue
		seen.add(key)
		next.push(issue)
	}
	return next
}

function mutationJournalIssueKey(issue: MutationJournalIssue): string {
	return JSON.stringify([
		issue.code,
		issue.message,
		issue.surface ?? null,
		issue.reason ?? null,
		issue.refs ?? [],
	])
}

function savedSourcePackageStateRefsForOp(op: Operation): readonly string[] {
	try {
		switch (op.op) {
			case 'setCells':
				return op.updates.map((update) => sheetRef(op.sheet, update.ref))
			case 'setFormula':
			case 'setRichText':
				return [sheetRef(op.sheet, op.ref)]
			case 'fillFormula':
			case 'clearRange':
			case 'sortRange':
			case 'setNumberFormat':
			case 'setStyle':
				return [sheetRef(op.sheet, op.range)]
			case 'insertRows':
			case 'deleteRows':
				return [`${op.sheet}!${op.at + 1}:${op.at + op.count}`]
			case 'hideRows':
				return [`${op.sheet}!${op.at + 1}:${op.at + op.count}`]
			case 'groupRows':
				return [`${op.sheet}!${op.from + 1}:${op.to + 1}`]
			case 'insertCols':
			case 'deleteCols':
				return [`${op.sheet}!${indexToColumn(op.at)}:${indexToColumn(op.at + op.count - 1)}`]
			case 'hideCols':
				return [`${op.sheet}!${indexToColumn(op.at)}:${indexToColumn(op.at + op.count - 1)}`]
			case 'groupCols':
				return [`${op.sheet}!${indexToColumn(op.from)}:${indexToColumn(op.to)}`]
			case 'copyRange':
			case 'moveRange': {
				const targetSheet = op.targetSheet ?? op.sheet
				return [
					sheetRef(op.sheet, op.source),
					sheetRef(targetSheet, rangeToA1(transferTargetRange(op.source, op.target))),
				]
			}
			case 'createTable':
				return [`table:${op.name}`, sheetRef(op.sheet, op.ref)]
			case 'appendRows':
				return [`table:${op.table}`]
			case 'copySheet':
				return [`sheet:${op.sheet}`, `sheet:${op.newName}`]
			case 'renameTable':
				return [`table:${op.table}`, `table:${op.newName}`]
			case 'deleteTable':
			case 'resizeTable':
			case 'setTableColumn':
			case 'setTableStyle':
				return [`table:${op.table}`]
			case 'setWorkbookProperties':
				return ['workbook:properties']
			case 'setPivotCache':
				return [`pivot:${op.pivotTable}`]
			case 'setPivotFieldItem':
				return [`pivot-field:${op.fieldIndex}`, `pivot-item:${op.itemIndex}`]
			case 'setTimelineRange':
				return [`timeline:${op.timelineCache ?? op.partPath ?? 'unknown'}`]
			case 'setConnectionRefresh':
				return [`connection:${op.connectionId}`]
			case 'rewriteExternalLink':
				return [`external-link:${op.relId}`]
			default:
				return []
		}
	} catch {
		return []
	}
}

function sheetRef(sheet: string, ref: string): string {
	return `${sheet}!${ref}`
}

export function emptyMutationJournal(): MutationJournal {
	return withUndoPolicy({
		entries: [],
		inverseOps: [],
		supported: true,
		exact: true,
		issues: [],
	})
}

export function failedMutationJournal(error: unknown): MutationJournal {
	const detail = error instanceof Error && error.message ? `: ${error.message}` : ''
	return withUndoPolicy({
		entries: [],
		inverseOps: [],
		supported: false,
		exact: false,
		issues: [
			{
				code: 'JOURNAL_BUILD_FAILED',
				message: `Mutation journal build failed${detail}`,
				surface: 'package-parts',
				reason: 'journal-build-failed',
			},
		],
	})
}

export function unavailableMutationJournal(
	message: string,
	refs?: readonly string[],
	classification?: {
		readonly surface?: MutationJournalSurface
		readonly reason?: MutationJournalReasonCode
	},
): MutationJournal {
	return withUndoPolicy({
		entries: [],
		inverseOps: [],
		supported: false,
		exact: false,
		issues: [
			{
				code: 'JOURNAL_UNAVAILABLE',
				message,
				surface: classification?.surface ?? 'package-parts',
				reason: classification?.reason ?? 'journal-unavailable',
				...(refs && refs.length > 0 ? { refs } : {}),
			},
		],
	})
}

function mutationJournalFromEntries(entries: readonly MutationJournalEntry[]): MutationJournal {
	const inverseOps = [...entries].reverse().flatMap((entry) => entry.inverseOps)
	const issues = entries.flatMap((entry) => entry.issues)
	return withUndoPolicy({
		entries,
		inverseOps,
		supported: entries.every((entry) => entry.supported),
		exact: entries.every((entry) => entry.exact),
		issues,
	})
}

function withUndoPolicy(
	journal: Omit<MutationJournal, 'schemaVersion' | 'schemaId' | 'undoPolicy'>,
): MutationJournal {
	return {
		schemaVersion: MUTATION_JOURNAL_ISSUE_SCHEMA_VERSION,
		schemaId: MUTATION_JOURNAL_ISSUE_SCHEMA.$id,
		...journal,
		undoPolicy: summarizeMutationJournalUndoPolicy(journal),
	}
}

function buildJournalEntry(
	workbook: Workbook,
	op: Operation,
	opIndex: number,
): MutationJournalEntry {
	const missingSheet = missingJournalOperationSheet(workbook, op)
	if (missingSheet) {
		const issues = [
			structureMutationJournalIssue(
				missingSheetTopologyIssue(
					missingSheet,
					`Cannot build exact rollback journal for ${op.op} because sheet ${missingSheet} was not found`,
				),
			),
		]
		return {
			opIndex,
			op,
			supported: true,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues,
		}
	}
	const rangeIssue = journalOperationRangeValueIssue(op)
	if (rangeIssue) {
		const issues = [structureMutationJournalIssue(rangeIssue)]
		return {
			opIndex,
			op,
			supported: true,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues,
		}
	}
	const cellRefIssue = journalOperationCellRefValueIssue(op)
	if (cellRefIssue) {
		const issues = [structureMutationJournalIssue(cellRefIssue)]
		return {
			opIndex,
			op,
			supported: true,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues,
		}
	}
	const axisIssue = journalOperationAxisValueIssue(op)
	if (axisIssue) {
		const issues = [structureMutationJournalIssue(axisIssue)]
		return {
			opIndex,
			op,
			supported: true,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues,
		}
	}
	const tableIssue = journalOperationTableTopologyIssue(workbook, op)
	if (tableIssue) {
		const issues = [structureMutationJournalIssue(tableIssue)]
		return {
			opIndex,
			op,
			supported: true,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues,
		}
	}
	const duplicateSheet = duplicateJournalOperationTargetSheet(workbook, op)
	if (duplicateSheet) {
		const issues = [
			structureMutationJournalIssue(
				sheetTopologyJournalIssue(
					duplicateSheet,
					`Cannot build exact rollback journal for ${op.op} because target sheet ${duplicateSheet} already exists`,
				),
			),
		]
		return {
			opIndex,
			op,
			supported: true,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues,
		}
	}
	const invalidSheetName = invalidJournalOperationTargetSheetName(op)
	if (invalidSheetName) {
		const issues = [
			structureMutationJournalIssue(
				sheetTopologyJournalIssue(
					invalidSheetName,
					`Cannot build exact rollback journal for ${op.op} because target sheet name ${invalidSheetName} is invalid`,
				),
			),
		]
		return {
			opIndex,
			op,
			supported: true,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues,
		}
	}
	const sheetLayoutIssue = journalOperationSheetLayoutValueIssue(workbook, op)
	if (sheetLayoutIssue) {
		const issues = [structureMutationJournalIssue(sheetLayoutIssue)]
		return {
			opIndex,
			op,
			supported: true,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues,
		}
	}
	const pageSetupIssue = journalOperationPageSetupValueIssue(op)
	if (pageSetupIssue) {
		const issues = [structureMutationJournalIssue(pageSetupIssue)]
		return {
			opIndex,
			op,
			supported: true,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues,
		}
	}
	const draft = buildSupportedJournalEntry(workbook, op, opIndex)
	if (!draft) {
		const issues: MutationJournalStructuredIssue[] = [
			{
				code: 'UNSUPPORTED_OPERATION',
				message: `No reversible journal support for ${op.op}`,
				surface: classifyMutationJournalOperationPrimarySurface(op),
				reason: 'operation-unsupported',
			},
		]
		return {
			opIndex,
			op,
			supported: false,
			exact: false,
			inverseOps: [],
			preimages: [],
			issues,
		}
	}
	const issues = draft.issues.map(structureMutationJournalIssue)
	return {
		...draft,
		issues,
		supported: issues.every((issue) => issue.code !== 'UNSUPPORTED_OPERATION'),
		exact: issues.length === 0,
	}
}

function missingJournalOperationSheet(workbook: Workbook, op: Operation): string | null {
	const sheet = journalOperationRequiredSheet(op)
	if (sheet !== null && !workbook.getSheet(sheet)) return sheet
	const targetSheet = journalOperationRequiredTargetSheet(op)
	if (targetSheet !== null && !workbook.getSheet(targetSheet)) return targetSheet
	return null
}

function journalOperationRequiredSheet(op: Operation): string | null {
	switch (op.op) {
		case 'setCells':
		case 'setFormula':
		case 'fillFormula':
		case 'setRichText':
		case 'clearRange':
		case 'insertRows':
		case 'insertCols':
		case 'deleteRows':
		case 'deleteCols':
		case 'setNumberFormat':
		case 'setStyle':
		case 'mergeCells':
		case 'unmergeCells':
		case 'setDataValidation':
		case 'deleteDataValidation':
		case 'setAutoFilter':
		case 'clearAutoFilter':
		case 'setConditionalFormat':
		case 'deleteConditionalFormat':
		case 'sortRange':
		case 'copyRange':
		case 'moveRange':
		case 'createTable':
		case 'setComment':
		case 'deleteComment':
		case 'setHyperlink':
		case 'deleteHyperlink':
		case 'setThreadedComment':
		case 'setDrawingText':
		case 'freezePane':
		case 'copySheet':
			return op.sheet
		default:
			return null
	}
}

function journalOperationRequiredTargetSheet(op: Operation): string | null {
	switch (op.op) {
		case 'copyRange':
		case 'moveRange':
			return op.targetSheet ?? null
		default:
			return null
	}
}

function journalOperationRangeValueIssue(op: Operation): MutationJournalIssue | null {
	const range = journalOperationRange(op)
	if (!range) return null
	try {
		if (range.kind === 'a1') parseA1(range.range)
		else parseRange(range.range)
		return null
	} catch {
		return {
			code: 'UNSUPPORTED_VALUE',
			message: `Cannot build exact rollback journal for ${op.op} because range ${range.range} is invalid`,
			surface: range.surface,
			reason: 'value-unsupported',
			refs: [`${range.sheet}!${range.range}`],
		}
	}
}

function journalOperationRange(op: Operation): {
	readonly kind?: 'range' | 'a1'
	readonly surface: MutationJournalSurface
	readonly sheet: string
	readonly range: string
} | null {
	switch (op.op) {
		case 'fillFormula':
			return { surface: 'formulas', sheet: op.sheet, range: op.range }
		case 'clearRange':
		case 'setNumberFormat':
		case 'setStyle':
		case 'sortRange':
			return { surface: 'cells', sheet: op.sheet, range: op.range }
		case 'copyRange':
		case 'moveRange':
			try {
				parseRange(op.source)
			} catch {
				return { surface: 'cells', sheet: op.sheet, range: op.source }
			}
			return {
				kind: 'a1',
				surface: 'cells',
				sheet: op.targetSheet ?? op.sheet,
				range: op.target,
			}
		case 'mergeCells':
		case 'unmergeCells':
			return { surface: 'merged-cells', sheet: op.sheet, range: op.range }
		case 'setDataValidation':
		case 'deleteDataValidation':
			return { surface: 'data-validations', sheet: op.sheet, range: op.range }
		case 'setConditionalFormat':
			return { surface: 'conditional-formats', sheet: op.sheet, range: op.range }
		case 'deleteConditionalFormat':
			return op.range === undefined
				? null
				: { surface: 'conditional-formats', sheet: op.sheet, range: op.range }
		case 'setAutoFilter':
			return { surface: 'auto-filters', sheet: op.sheet, range: op.range }
		default:
			return null
	}
}

function journalOperationCellRefValueIssue(op: Operation): MutationJournalIssue | null {
	const target = journalOperationCellRefTarget(op)
	if (!target) return null
	const invalidRefs = target.refs.filter((ref) => {
		try {
			parseA1(ref)
			return false
		} catch {
			return true
		}
	})
	if (invalidRefs.length === 0) return null
	return {
		code: 'UNSUPPORTED_VALUE',
		message: `Cannot build exact rollback journal for ${op.op} because cell reference ${invalidRefs[0]} is invalid`,
		surface: target.surface,
		reason: 'value-unsupported',
		refs: invalidRefs.map((ref) => `${target.sheet}!${ref}`),
	}
}

function journalOperationCellRefTarget(op: Operation): {
	readonly surface: MutationJournalSurface
	readonly sheet: string
	readonly refs: readonly string[]
} | null {
	switch (op.op) {
		case 'setCells':
			return { surface: 'cells', sheet: op.sheet, refs: op.updates.map((update) => update.ref) }
		case 'setFormula':
			return { surface: 'formulas', sheet: op.sheet, refs: [op.ref] }
		case 'setRichText':
			return { surface: 'cells', sheet: op.sheet, refs: [op.ref] }
		case 'setComment':
		case 'deleteComment':
			return { surface: 'comments', sheet: op.sheet, refs: [op.ref] }
		case 'setThreadedComment':
			return op.ref === undefined ? null : { surface: 'comments', sheet: op.sheet, refs: [op.ref] }
		case 'setHyperlink':
		case 'deleteHyperlink':
			return { surface: 'hyperlinks', sheet: op.sheet, refs: [op.ref] }
		default:
			return null
	}
}

function journalOperationAxisValueIssue(op: Operation): MutationJournalIssue | null {
	const target = journalOperationAxisTarget(op)
	if (!target) return null
	if (target.valid) return null
	return {
		code: 'UNSUPPORTED_VALUE',
		message: `Cannot build exact rollback journal for ${op.op} because ${target.label} ${target.value} is invalid`,
		surface: target.surface,
		reason: 'value-unsupported',
		refs: [target.ref],
	}
}

function journalOperationAxisTarget(op: Operation):
	| {
			readonly valid: true
	  }
	| {
			readonly valid: false
			readonly surface: MutationJournalSurface
			readonly label: string
			readonly value: number
			readonly ref: string
	  }
	| null {
	switch (op.op) {
		case 'insertRows':
		case 'deleteRows':
		case 'hideRows':
			return journalAxisSpanTarget('row-layout', op.sheet, 'row', op.at, op.count)
		case 'insertCols':
		case 'deleteCols':
		case 'hideCols':
			return journalAxisSpanTarget('column-layout', op.sheet, 'column', op.at, op.count)
		case 'setRowHeight':
			return journalAxisScalarTarget('row-layout', op.sheet, 'row', op.row, op.height, 'height')
		case 'setColWidth':
			return journalAxisScalarTarget('column-layout', op.sheet, 'column', op.col, op.width, 'width')
		case 'groupRows':
			return journalAxisBandTarget('row-layout', op.sheet, 'row', op.from, op.to)
		case 'groupCols':
			return journalAxisBandTarget('column-layout', op.sheet, 'column', op.from, op.to)
		default:
			return null
	}
}

function journalAxisSpanTarget(
	surface: MutationJournalSurface,
	sheet: string,
	label: 'row' | 'column',
	at: number,
	count: number,
) {
	if (!Number.isInteger(at) || at < 0) {
		return { valid: false as const, surface, label, value: at, ref: `${sheet}!${label}:${at}` }
	}
	if (!Number.isInteger(count) || count <= 0) {
		return {
			valid: false as const,
			surface,
			label: `${label} count`,
			value: count,
			ref: `${sheet}!${label}-count:${count}`,
		}
	}
	return { valid: true as const }
}

function journalAxisScalarTarget(
	surface: MutationJournalSurface,
	sheet: string,
	label: 'row' | 'column',
	index: number,
	value: number,
	valueLabel: 'height' | 'width',
) {
	if (!Number.isInteger(index) || index < 0) {
		return {
			valid: false as const,
			surface,
			label,
			value: index,
			ref: `${sheet}!${label}:${index}`,
		}
	}
	if (!Number.isFinite(value) || value < 0) {
		return {
			valid: false as const,
			surface,
			label: `${label} ${valueLabel}`,
			value,
			ref: `${sheet}!${label}-${valueLabel}:${value}`,
		}
	}
	return { valid: true as const }
}

function journalAxisBandTarget(
	surface: MutationJournalSurface,
	sheet: string,
	label: 'row' | 'column',
	from: number,
	to: number,
) {
	if (!Number.isInteger(from) || from < 0) {
		return { valid: false as const, surface, label, value: from, ref: `${sheet}!${label}:${from}` }
	}
	if (!Number.isInteger(to) || to < from) {
		return { valid: false as const, surface, label, value: to, ref: `${sheet}!${label}:${to}` }
	}
	return { valid: true as const }
}

function journalOperationSheetLayoutValueIssue(
	workbook: Workbook,
	op: Operation,
): MutationJournalIssue | null {
	const target = journalOperationSheetLayoutTarget(workbook, op)
	if (!target) return null
	if (target.valid) return null
	return {
		code: 'UNSUPPORTED_VALUE',
		message: `Cannot build exact rollback journal for ${op.op} because ${target.label} ${target.value} is invalid`,
		surface: 'sheet-layout',
		reason: 'value-unsupported',
		refs: [target.ref],
	}
}

function journalOperationSheetLayoutTarget(
	workbook: Workbook,
	op: Operation,
):
	| {
			readonly valid: true
	  }
	| {
			readonly valid: false
			readonly label: string
			readonly value: number | string
			readonly ref: string
	  }
	| null {
	switch (op.op) {
		case 'addSheet':
			return op.position === undefined
				? null
				: journalSheetPositionTarget('sheet-position', op.position, workbook.sheets.length)
		case 'copySheet':
			return op.position === undefined
				? null
				: journalSheetPositionTarget('sheet-position', op.position, workbook.sheets.length)
		case 'moveSheet':
			return journalSheetPositionTarget('sheet-position', op.position, workbook.sheets.length - 1)
		case 'freezePane':
			return journalFreezePaneTarget(op)
		case 'setTabColor':
			return /^(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$/.test(op.color)
				? { valid: true as const }
				: {
						valid: false as const,
						label: 'tab color',
						value: op.color,
						ref: `sheet:${op.sheet}:tabColor:${op.color}`,
					}
		default:
			return null
	}
}

function journalSheetPositionTarget(label: string, position: number, maxInclusive: number) {
	if (Number.isInteger(position) && position >= 0 && position <= maxInclusive) {
		return { valid: true as const }
	}
	return {
		valid: false as const,
		label,
		value: position,
		ref: `sheet-position:${position}`,
	}
}

function journalFreezePaneTarget(op: Extract<Operation, { op: 'freezePane' }>) {
	if (!Number.isInteger(op.row) || op.row < 0) {
		return {
			valid: false as const,
			label: 'freeze row',
			value: op.row,
			ref: `${op.sheet}!freeze-row:${op.row}`,
		}
	}
	if (!Number.isInteger(op.col) || op.col < 0) {
		return {
			valid: false as const,
			label: 'freeze column',
			value: op.col,
			ref: `${op.sheet}!freeze-column:${op.col}`,
		}
	}
	return { valid: true as const }
}

function journalOperationPageSetupValueIssue(op: Operation): MutationJournalIssue | null {
	if (op.op !== 'setPageSetup') return null
	const issue = journalPageSetupValueTarget(op)
	if (!issue) return null
	return {
		code: 'UNSUPPORTED_VALUE',
		message: `Cannot build exact rollback journal for setPageSetup because ${issue.label} is invalid`,
		surface: 'page-setup',
		reason: 'value-unsupported',
		refs: [issue.ref],
	}
}

function journalPageSetupValueTarget(op: Extract<Operation, { op: 'setPageSetup' }>) {
	const setup = op.setup
	if (
		setup.orientation !== undefined &&
		setup.orientation !== 'portrait' &&
		setup.orientation !== 'landscape'
	) {
		return { label: 'orientation', ref: `${op.sheet}!pageSetup:orientation` }
	}
	for (const [field, value] of [
		['paperSize', setup.paperSize],
		['fitToWidth', setup.fitToWidth],
		['fitToHeight', setup.fitToHeight],
	] as const) {
		if (value !== undefined && (!Number.isInteger(value) || value < 0)) {
			return { label: field, ref: `${op.sheet}!pageSetup:${field}` }
		}
	}
	if (setup.scale !== undefined && (!Number.isInteger(setup.scale) || setup.scale <= 0)) {
		return { label: 'scale', ref: `${op.sheet}!pageSetup:scale` }
	}
	if (setup.margins) {
		for (const [field, value] of Object.entries(setup.margins)) {
			if (value !== undefined && (!Number.isFinite(value) || value < 0)) {
				return { label: `margin ${field}`, ref: `${op.sheet}!pageMargins:${field}` }
			}
		}
	}
	return null
}

function journalOperationTableTopologyIssue(
	workbook: Workbook,
	op: Operation,
): MutationJournalIssue | null {
	const requiredTable = journalOperationRequiredTable(op)
	if (requiredTable !== null) {
		const matches = findTableMatches(workbook, requiredTable)
		if (matches.length === 0) {
			return tableTopologyJournalIssue(
				requiredTable,
				`Cannot build exact rollback journal for ${op.op} because table ${requiredTable} was not found`,
			)
		}
		if (matches.length > 1) {
			return tableTopologyJournalIssue(
				requiredTable,
				`Cannot build exact rollback journal for ${op.op} because table ${requiredTable} is ambiguous`,
			)
		}
	}
	const metadataIssue = journalOperationTableMetadataIssue(workbook, op)
	if (metadataIssue) return metadataIssue
	const createdTable = journalOperationCreatedTable(op)
	if (createdTable !== null) {
		if (validateExcelTableName(createdTable)) {
			return tableUnsupportedValueIssue(
				createdTable,
				`Cannot build exact rollback journal for ${op.op} because target table name ${createdTable} is invalid`,
			)
		}
		const sourceTable =
			op.op === 'renameTable' ? findTableMatches(workbook, op.table)[0]?.table : undefined
		const collision = findTableMatches(workbook, createdTable).some(
			(match) => match.table.id !== sourceTable?.id,
		)
		if (collision) {
			return tableTopologyJournalIssue(
				createdTable,
				`Cannot build exact rollback journal for ${op.op} because target table ${createdTable} already exists`,
			)
		}
	}
	const rangeIssue = journalOperationTableRangeValueIssue(op)
	if (rangeIssue) return rangeIssue
	const rangeCollision = journalOperationTableRangeCollisionIssue(workbook, op)
	if (rangeCollision) return rangeCollision
	return null
}

function journalOperationRequiredTable(op: Operation): string | null {
	switch (op.op) {
		case 'deleteTable':
		case 'renameTable':
		case 'resizeTable':
		case 'setTableColumn':
		case 'setTableStyle':
			return op.table
		default:
			return null
	}
}

function journalOperationCreatedTable(op: Operation): string | null {
	switch (op.op) {
		case 'createTable':
			return op.name
		case 'renameTable':
			return op.newName
		default:
			return null
	}
}

function journalOperationTableMetadataIssue(
	workbook: Workbook,
	op: Operation,
): MutationJournalIssue | null {
	if (op.op !== 'setTableColumn' || op.newName === undefined) return null
	const table = findTableMatches(workbook, op.table)[0]?.table
	if (!table) return null
	const sourceIndex = tableColumnIndex(table, op.column)
	const targetIndex = table.columns.findIndex(
		(column) => column.name.toLowerCase() === op.newName?.toLowerCase(),
	)
	if (sourceIndex >= 0 && targetIndex >= 0 && targetIndex !== sourceIndex) {
		return tableTopologyJournalIssue(
			op.table,
			`Cannot build exact rollback journal for ${op.op} because target column ${op.newName} already exists in table ${op.table}`,
		)
	}
	if (table.queryTable) {
		return {
			code: 'UNSUPPORTED_VALUE',
			message: `Cannot build exact rollback journal for ${op.op} because queryTable-backed table ${op.table} cannot safely rename columns`,
			surface: 'tables',
			reason: 'table-metadata',
			refs: [`table:${op.table}`],
		}
	}
	return null
}

function journalOperationTableRangeValueIssue(op: Operation): MutationJournalIssue | null {
	if (op.op !== 'createTable' && op.op !== 'resizeTable') return null
	try {
		parseRange(op.ref)
		return null
	} catch {
		const tableName = op.op === 'createTable' ? op.name : op.table
		return tableUnsupportedValueIssue(
			tableName,
			`Cannot build exact rollback journal for ${op.op} because range ${op.ref} is invalid`,
		)
	}
}

function journalOperationTableRangeCollisionIssue(
	workbook: Workbook,
	op: Operation,
): MutationJournalIssue | null {
	if (op.op !== 'createTable' && op.op !== 'resizeTable') return null
	const sheet =
		op.op === 'createTable'
			? workbook.getSheet(op.sheet)
			: findTableMatches(workbook, op.table)[0]?.sheet
	if (!sheet) return null
	const sourceTable =
		op.op === 'resizeTable' ? findTableMatches(workbook, op.table)[0]?.table : null
	const targetRange = parseRange(op.ref)
	const overlappingTable = sheet.tables.find(
		(table) => table.id !== sourceTable?.id && rangesOverlap(table.ref, targetRange),
	)
	if (!overlappingTable) return null
	const tableName = op.op === 'createTable' ? op.name : op.table
	return tableTopologyJournalIssue(
		tableName,
		`Cannot build exact rollback journal for ${op.op} because range ${sheet.name}!${op.ref} overlaps table ${overlappingTable.name}`,
	)
}

function duplicateJournalOperationTargetSheet(workbook: Workbook, op: Operation): string | null {
	const targetSheet = journalOperationCreatedSheet(op)
	if (targetSheet !== null && workbook.getSheet(targetSheet)) return targetSheet
	return null
}

function journalOperationCreatedSheet(op: Operation): string | null {
	switch (op.op) {
		case 'addSheet':
			return op.name
		case 'copySheet':
		case 'renameSheet':
			return op.newName
		default:
			return null
	}
}

function invalidJournalOperationTargetSheetName(op: Operation): string | null {
	const targetSheet = journalOperationCreatedSheet(op)
	if (targetSheet === null) return null
	return validateExcelWorksheetName(targetSheet) ? targetSheet : null
}

function buildSupportedJournalEntry(
	workbook: Workbook,
	op: Operation,
	opIndex: number,
): DraftJournalEntry | null {
	switch (op.op) {
		case 'setCells':
			return journalSetCells(workbook, op, opIndex)
		case 'setFormula':
			return journalSetFormula(workbook, op, opIndex)
		case 'fillFormula':
			return journalFillFormula(workbook, op, opIndex)
		case 'setRichText':
			return journalSetRichText(workbook, op, opIndex)
		case 'clearRange':
			return journalClearRange(workbook, op, opIndex)
		case 'insertRows':
			return journalInsertAxis(op, opIndex, 'row')
		case 'insertCols':
			return journalInsertAxis(op, opIndex, 'col')
		case 'deleteRows':
			return journalDeleteAxis(workbook, op, opIndex, 'row')
		case 'deleteCols':
			return journalDeleteAxis(workbook, op, opIndex, 'col')
		case 'setNumberFormat':
		case 'setStyle':
			return journalStyleRange(workbook, op, opIndex)
		case 'mergeCells':
			return journalMergeCells(workbook, op, opIndex)
		case 'unmergeCells':
			return journalUnmergeCells(workbook, op, opIndex)
		case 'setDataValidation':
			return journalSetDataValidation(workbook, op, opIndex)
		case 'deleteDataValidation':
			return journalDeleteDataValidation(workbook, op, opIndex)
		case 'setAutoFilter':
			return journalSetAutoFilter(workbook, op, opIndex)
		case 'clearAutoFilter':
			return journalClearAutoFilter(workbook, op, opIndex)
		case 'setConditionalFormat':
			return journalSetConditionalFormat(workbook, op, opIndex)
		case 'deleteConditionalFormat':
			return journalDeleteConditionalFormat(workbook, op, opIndex)
		case 'setPageSetup':
			return journalSetPageSetup(workbook, op, opIndex)
		case 'setPrintArea':
			return journalSetPrintArea(workbook, op, opIndex)
		case 'setDefinedName':
			return journalSetDefinedName(workbook, op, opIndex)
		case 'deleteDefinedName':
			return journalDeleteDefinedName(workbook, op, opIndex)
		case 'renameTable':
			return journalRenameTable(workbook, op, opIndex)
		case 'createTable':
			return journalCreateTable(op, opIndex)
		case 'sortRange':
			return journalSortRange(workbook, op, opIndex)
		case 'copyRange':
			return journalCopyRange(workbook, op, opIndex)
		case 'moveRange':
			return journalMoveRange(workbook, op, opIndex)
		case 'deleteTable':
			return journalDeleteTable(workbook, op, opIndex)
		case 'resizeTable':
			return journalResizeTable(workbook, op, opIndex)
		case 'setTableColumn':
			return journalSetTableColumn(workbook, op, opIndex)
		case 'setTableStyle':
			return journalSetTableStyle(workbook, op, opIndex)
		case 'setComment':
			return journalSetComment(workbook, op, opIndex)
		case 'deleteComment':
			return journalDeleteComment(workbook, op, opIndex)
		case 'setHyperlink':
			return journalSetHyperlink(workbook, op, opIndex)
		case 'deleteHyperlink':
			return journalDeleteHyperlink(workbook, op, opIndex)
		case 'setThreadedComment':
			return journalSetThreadedComment(workbook, op, opIndex)
		case 'setDrawingText':
			return journalSetDrawingText(workbook, op, opIndex)
		case 'setChartSeriesSource':
			return journalSetChartSeriesSource(workbook, op, opIndex)
		case 'setPivotCache':
			return journalSetPivotCache(workbook, op, opIndex)
		case 'freezePane':
			return journalFreezePane(workbook, op, opIndex)
		case 'setWorkbookProperties':
			return journalSetWorkbookProperties(workbook, op, opIndex)
		case 'setDocumentProperties':
			return journalSetDocumentProperties(workbook, op, opIndex)
		case 'setWorkbookView':
			return journalSetWorkbookView(workbook, op, opIndex)
		case 'setCalcSettings':
			return journalSetCalcSettings(workbook, op, opIndex)
		case 'setWorkbookProtection':
			return journalSetWorkbookProtection(workbook, op, opIndex)
		case 'setTheme':
			return journalSetTheme(workbook, op, opIndex)
		case 'renameSheet':
			return journalRenameSheet(workbook, op, opIndex)
		case 'moveSheet':
			return journalMoveSheet(workbook, op, opIndex)
		case 'addSheet':
			return {
				opIndex,
				op,
				inverseOps: [{ op: 'deleteSheet', sheet: op.name }],
				preimages: [],
				issues: savedSourcePackageStateIssues(workbook, op.op, [`sheet:${op.name}`]),
			}
		case 'deleteSheet':
			return journalDeleteSheet(workbook, op, opIndex)
		case 'copySheet':
			return {
				opIndex,
				op,
				inverseOps: [{ op: 'deleteSheet', sheet: op.newName }],
				preimages: [],
				issues: savedSourcePackageStateIssues(workbook, op.op, [
					`sheet:${op.sheet}`,
					`sheet:${op.newName}`,
				]),
			}
		case 'setRowHeight':
			return journalSetSheetLayout(workbook, op, opIndex, 'row')
		case 'setColWidth':
			return journalSetSheetLayout(workbook, op, opIndex, 'col')
		case 'setTabColor':
			return journalSetTabColor(workbook, op, opIndex)
		case 'setSheetProtection':
			return journalSetSheetProtection(workbook, op, opIndex)
		case 'hideSheet':
			return journalHideSheet(workbook, op, opIndex)
		case 'hideRows':
			return journalHideRows(workbook, op, opIndex)
		case 'hideCols':
			return journalHideCols(workbook, op, opIndex)
		case 'groupRows':
			return journalGroupOutline(workbook, op, opIndex, 'row')
		case 'groupCols':
			return journalGroupOutline(workbook, op, opIndex, 'col')
		default:
			return null
	}
}

function journalSetTabColor(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTabColor' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const tabColor = sheet?.tabColor ? cloneSheetTabColor(sheet.tabColor) : null
	const preimage: MutationJournalSheetTabColorPreimage = {
		sheet: op.sheet,
		tabColor,
	}
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-tab-color', sheetTabColor: preimage }],
			issues: [
				missingSheetTopologyIssue(
					op.sheet,
					`Cannot restore tab color for ${op.sheet} because the sheet was not found`,
				),
			],
		}
	}
	const inverseOps: Operation[] = tabColor?.rgb
		? [{ op: 'setTabColor', sheet: op.sheet, color: tabColor.rgb }]
		: []
	const issues = tabColorIsPublicRgb(tabColor)
		? []
		: [
				{
					code: 'LOSSY_INVERSE',
					message: tabColor
						? `Sheet tab color for ${op.sheet} uses unsupported color metadata and cannot be fully restored with public operations`
						: `Sheet tab color absence for ${op.sheet} cannot be restored with public operations`,
					surface: 'sheet-layout',
					reason: 'sheet-topology',
					refs: [`sheet:${op.sheet}:tabColor`],
				} satisfies MutationJournalIssue,
			]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'sheet-tab-color', sheetTabColor: preimage }],
		issues,
	}
}

function journalSetSheetProtection(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setSheetProtection' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const protection = sheet?.protection ? cloneSheetProtection(sheet.protection) : null
	const preimage: MutationJournalSheetProtectionPreimage = {
		sheet: op.sheet,
		protection,
	}
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-protection', sheetProtection: preimage }],
			issues: [
				missingSheetTopologyIssue(
					op.sheet,
					`Cannot restore sheet protection for ${op.sheet} because the sheet was not found`,
				),
			],
		}
	}
	const valueIssues = sheetProtectionValueIssues(op)
	if (valueIssues.length > 0) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-protection', sheetProtection: preimage }],
			issues: valueIssues,
		}
	}
	const inverseOp = protection ? sheetProtectionInverseOp(op.sheet, protection) : null
	const unsupportedKeys = protection ? unsupportedSheetProtectionKeys(protection) : []
	const issues: MutationJournalIssue[] = []
	if (!protection) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Sheet protection absence for ${op.sheet} cannot be restored with public operations`,
			surface: 'sheet-layout',
			reason: 'workbook-protection-absence',
			refs: [`sheet:${op.sheet}:protection`],
		})
	} else if (protection.sheet !== true || unsupportedKeys.length > 0) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Sheet protection for ${op.sheet} contains metadata that cannot be fully restored with public operations`,
			surface: 'sheet-layout',
			reason: 'workbook-protection-absence',
			refs: unsupportedKeys.map((key) => `sheet:${op.sheet}:protection:${key}`),
		})
	}
	return {
		opIndex,
		op,
		inverseOps: inverseOp ? [inverseOp] : [],
		preimages: [{ kind: 'sheet-protection', sheetProtection: preimage }],
		issues,
	}
}

function sheetProtectionValueIssues(
	op: Extract<Operation, { op: 'setSheetProtection' }>,
): MutationJournalIssue[] {
	const issues: MutationJournalIssue[] = []
	if (op.password !== undefined && typeof op.password !== 'string') {
		issues.push(sheetProtectionUnsupportedValueIssue(op.sheet, 'password'))
	}
	const options = op.options as unknown
	if (options !== undefined && !isPlainJournalObject(options)) {
		issues.push(sheetProtectionUnsupportedValueIssue(op.sheet, 'options'))
		return issues
	}
	if (isPlainJournalObject(options)) {
		for (const field of SHEET_PROTECTION_OPTION_KEYS) {
			const value = options[field]
			if (value !== undefined && typeof value !== 'boolean') {
				issues.push(sheetProtectionUnsupportedValueIssue(op.sheet, field))
			}
		}
	}
	return issues
}

function sheetProtectionUnsupportedValueIssue(sheet: string, field: string): MutationJournalIssue {
	return {
		code: 'UNSUPPORTED_VALUE',
		message: `Cannot build exact rollback journal for setSheetProtection because ${field} is not supported by sheet protection validation`,
		surface: 'sheet-layout',
		reason: 'value-unsupported',
		refs: [`sheet:${sheet}:protection:${field}`],
	}
}

function savedSourcePackageStateIssues(
	workbook: Workbook,
	opName: MutationJournalOperationName,
	refs: readonly string[] = [],
): MutationJournalIssue[] {
	if (!workbook.sourceArchiveBytes) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `${opName} changes saved package state that public inverse operations cannot restore byte-for-byte`,
			surface: 'package-parts',
			reason: 'package-part-preservation',
			...(refs.length > 0 ? { refs } : {}),
		},
	]
}

function savedSourceDefinedNamePackageStateIssues(
	workbook: Workbook,
	opName: Extract<Operation, { op: 'setDefinedName' | 'deleteDefinedName' }>['op'],
	name: string,
): MutationJournalIssue[] {
	if (name === '_xlnm.Print_Area') return []
	return savedSourcePackageStateIssues(workbook, opName, [`name:${name}`])
}

function missingSheetTopologyIssue(sheet: string, message: string): MutationJournalIssue {
	return sheetTopologyJournalIssue(sheet, message)
}

function sheetTopologyJournalIssue(sheet: string, message: string): MutationJournalIssue {
	return {
		code: 'UNSUPPORTED_VALUE',
		message,
		surface: 'sheet-layout',
		reason: 'sheet-topology',
		refs: [`sheet:${sheet}`],
	}
}

function tableTopologyJournalIssue(table: string, message: string): MutationJournalIssue {
	return {
		code: 'UNSUPPORTED_VALUE',
		message,
		surface: 'tables',
		reason: 'operation-unsupported',
		refs: [`table:${table}`],
	}
}

function tableUnsupportedValueIssue(table: string, message: string): MutationJournalIssue {
	return {
		code: 'UNSUPPORTED_VALUE',
		message,
		surface: 'tables',
		reason: 'value-unsupported',
		refs: [`table:${table}`],
	}
}

function journalDeleteSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteSheet' }>,
	opIndex: number,
): DraftJournalEntry {
	const position = workbook.sheets.findIndex((sheet) => sheet.name === op.sheet)
	const sheet = position >= 0 ? workbook.sheets[position] : undefined
	const preimage: MutationJournalSheetDeletePreimage = {
		sheet: op.sheet,
		position: position >= 0 ? position : null,
		existed: sheet !== undefined,
	}
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-delete', sheetDelete: preimage }],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore deleted sheet ${op.sheet} because it was not found`,
					surface: 'sheet-layout',
					reason: 'sheet-topology',
					refs: [`sheet:${op.sheet}`],
				},
			],
		}
	}
	const refs = sheetDeleteLossRefs(workbook, sheet)
	const issues: MutationJournalIssue[] =
		refs.length === 0
			? []
			: [
					{
						code: 'LOSSY_INVERSE',
						message: `Deleted sheet ${op.sheet} cannot be fully restored with public operations`,
						surface: 'sheet-layout',
						reason: 'sheet-topology',
						refs,
					},
				]
	issues.push(...savedSourcePackageStateIssues(workbook, op.op, [`sheet:${op.sheet}`]))
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'addSheet', name: op.sheet, position }],
		preimages: [{ kind: 'sheet-delete', sheetDelete: preimage }],
		issues,
	}
}

function journalMoveSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'moveSheet' }>,
	opIndex: number,
): DraftJournalEntry {
	const position = workbook.sheets.findIndex((sheet) => sheet.name === op.sheet)
	const preimage: MutationJournalSheetMovePreimage = {
		sheet: op.sheet,
		position: position >= 0 ? position : null,
	}
	const issues: MutationJournalIssue[] =
		position >= 0
			? [...savedSourcePackageStateIssues(workbook, op.op, [`sheet:${op.sheet}`])]
			: [
					missingSheetTopologyIssue(
						op.sheet,
						`Cannot restore sheet move for ${op.sheet} because the sheet was not found`,
					),
				]
	return {
		opIndex,
		op,
		inverseOps: position >= 0 ? [{ op: 'moveSheet', sheet: op.sheet, position }] : [],
		preimages: [{ kind: 'sheet-move', sheetMove: preimage }],
		issues,
	}
}

function journalSetSheetLayout(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setRowHeight' | 'setColWidth' }>,
	opIndex: number,
	axis: 'row' | 'col',
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const index =
		axis === 'row' && op.op === 'setRowHeight' ? op.row : op.op === 'setColWidth' ? op.col : -1
	const rowDef = sheet && axis === 'row' ? sheet.rowDefs.get(index) : undefined
	const colDef =
		sheet && axis === 'col'
			? sheet.colDefs.find((def) => def.min <= index && def.max >= index)
			: undefined
	const colWidth = sheet && axis === 'col' ? sheet.colWidths.get(index) : undefined
	const value = sheet
		? axis === 'row'
			? sheet.rowHeights.get(index)
			: (colWidth ?? colDef?.width)
		: undefined
	const preimage: MutationJournalSheetLayoutPreimage = {
		sheet: op.sheet,
		axis,
		index,
		value: value ?? null,
	}
	const ref = layoutRef(op.sheet, axis, index)
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-layout', sheetLayout: preimage }],
			issues: [
				missingSheetTopologyIssue(
					op.sheet,
					`Cannot restore ${axis} layout at ${ref} because the sheet was not found`,
				),
			],
		}
	}
	if (value === undefined) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-layout', sheetLayout: preimage }],
			issues: [
				{
					code: 'LOSSY_INVERSE',
					message: `Created ${axis} layout at ${ref} cannot be cleared with public operations`,
					surface: axis === 'row' ? 'row-layout' : 'column-layout',
					reason: axis === 'row' ? 'row-layout-created' : 'column-layout-created',
					refs: [ref],
				},
			],
		}
	}
	const issues: MutationJournalIssue[] = []
	if (axis === 'row' && rowDef?.customHeight === false) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Row height metadata at ${ref} has customHeight=false and cannot be restored exactly with public operations`,
			surface: 'row-layout',
			reason: 'row-layout-custom-height',
			refs: [ref],
		})
	}
	if (
		axis === 'col' &&
		colWidth !== undefined &&
		colDef?.width !== undefined &&
		colWidth !== colDef.width
	) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Column width map and column metadata disagree at ${ref} and cannot be restored exactly with public operations`,
			surface: 'column-layout',
			reason: 'column-layout-width-metadata',
			refs: [ref],
		})
	}
	if (
		axis === 'col' &&
		colWidth === undefined &&
		colDef?.width !== undefined &&
		colDef.customWidth !== true
	) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Column width metadata at ${ref} has customWidth=${String(colDef.customWidth)} and cannot be restored exactly with public operations`,
			surface: 'column-layout',
			reason: 'column-layout-width-metadata',
			refs: [ref],
		})
	}
	const inverseOps: Operation[] =
		axis === 'row'
			? [{ op: 'setRowHeight', sheet: op.sheet, row: index, height: value }]
			: [{ op: 'setColWidth', sheet: op.sheet, col: index, width: value }]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'sheet-layout', sheetLayout: preimage }],
		issues,
	}
}

function journalHideSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideSheet' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const preimage: MutationJournalSheetVisibilityPreimage = {
		sheet: op.sheet,
		state: sheet?.state ?? null,
	}
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'sheet-visibility', sheetVisibility: preimage }],
			issues: [
				missingSheetTopologyIssue(
					op.sheet,
					`Cannot restore sheet visibility for ${op.sheet} because the sheet was not found`,
				),
			],
		}
	}
	const issues: MutationJournalIssue[] =
		sheet.state === 'veryHidden'
			? [
					{
						code: 'LOSSY_INVERSE',
						message: `Sheet visibility for ${op.sheet} was veryHidden and cannot be restored with public operations`,
						surface: 'sheet-layout',
						reason: 'sheet-topology',
						refs: [`sheet:${op.sheet}:state:veryHidden`],
					},
				]
			: []
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'hideSheet', sheet: op.sheet, hidden: sheet.state !== 'visible' }],
		preimages: [{ kind: 'sheet-visibility', sheetVisibility: preimage }],
		issues,
	}
}

function journalHideRows(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideRows' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const rows = Array.from({ length: op.count }, (_, offset) => {
		const row = op.at + offset
		const height = sheet?.rowHeights.get(row)
		const rowDef = sheet?.rowDefs.get(row)
		return { row, height: height ?? null, rowDef: rowDef ? cloneSheetRowDef(rowDef) : null }
	})
	const preimage: MutationJournalRowsHiddenPreimage = { sheet: op.sheet, rows }
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'rows-hidden', rowsHidden: preimage }],
			issues: [
				missingSheetTopologyIssue(
					op.sheet,
					`Cannot restore row visibility for ${op.sheet} because the sheet was not found`,
				),
			],
		}
	}
	const hidden = op.hidden ?? true
	const inverseOps = rows.flatMap((row): Operation[] => {
		if (hidden) {
			return row.rowDef?.hidden === true
				? []
				: [{ op: 'hideRows', sheet: op.sheet, at: row.row, count: 1, hidden: false }]
		}
		return row.rowDef?.hidden === true
			? [{ op: 'hideRows', sheet: op.sheet, at: row.row, count: 1, hidden: true }]
			: []
	})
	const refs = rows
		.filter((row) => row.rowDef?.hidden === false)
		.map((row) => layoutRef(op.sheet, 'row', row.row))
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'rows-hidden', rowsHidden: preimage }],
		issues:
			refs.length === 0
				? []
				: [
						{
							code: 'LOSSY_INVERSE',
							message: `Explicit row hidden=false metadata cannot be restored exactly with public operations`,
							surface: 'row-layout',
							reason: 'row-layout-created',
							refs,
						},
					],
	}
}

function journalHideCols(
	workbook: Workbook,
	op: Extract<Operation, { op: 'hideCols' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const cols = Array.from({ length: op.count }, (_, offset) => {
		const col = op.at + offset
		const colDef = sheet?.colDefs.find((def) => def.min <= col && def.max >= col)
		return { col, colDef: colDef ? cloneSheetColDef(colDef) : null }
	})
	const preimage: MutationJournalColsHiddenPreimage = { sheet: op.sheet, cols }
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'cols-hidden', colsHidden: preimage }],
			issues: [
				missingSheetTopologyIssue(
					op.sheet,
					`Cannot restore column visibility for ${op.sheet} because the sheet was not found`,
				),
			],
		}
	}
	const hidden = op.hidden ?? true
	const inverseOps = cols.flatMap((col): Operation[] => {
		if (hidden) {
			return col.colDef?.hidden === true
				? []
				: [{ op: 'hideCols', sheet: op.sheet, at: col.col, count: 1, hidden: false }]
		}
		return col.colDef?.hidden === true
			? [{ op: 'hideCols', sheet: op.sheet, at: col.col, count: 1, hidden: true }]
			: []
	})
	const refs = cols
		.filter((col) => col.colDef?.hidden === false)
		.map((col) => layoutRef(op.sheet, 'col', col.col))
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'cols-hidden', colsHidden: preimage }],
		issues:
			refs.length === 0
				? []
				: [
						{
							code: 'LOSSY_INVERSE',
							message: `Explicit column hidden=false metadata cannot be restored exactly with public operations`,
							surface: 'column-layout',
							reason: 'column-layout-created',
							refs,
						},
					],
	}
}

function journalGroupOutline(
	workbook: Workbook,
	op: Extract<Operation, { op: 'groupRows' | 'groupCols' }>,
	opIndex: number,
	axis: 'row' | 'col',
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const preimage: MutationJournalOutlinePreimage = {
		sheet: op.sheet,
		axis,
		from: op.from,
		to: op.to,
		outlinePr: sheet?.outlinePr ? cloneSheetOutlinePr(sheet.outlinePr) : null,
		sheetFormatPr: sheet?.sheetFormatPr ? cloneSheetFormatPr(sheet.sheetFormatPr) : null,
		...(axis === 'row'
			? {
					rowDefs: outlineIndexes(op).map((row) => {
						const rowDef = sheet?.rowDefs.get(row)
						return {
							row,
							rowDef: rowDef ? cloneSheetRowDef(rowDef) : null,
						}
					}),
				}
			: {
					colDefs: outlineIndexes(op).map((col) => ({
						col,
						colDef: cloneMatchingSheetColDef(sheet, col),
					})),
				}),
	}
	if (!sheet) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [{ kind: 'outline', outline: preimage }],
			issues: [
				missingSheetTopologyIssue(
					op.sheet,
					`Cannot restore ${axis} outline grouping for ${op.sheet} because the sheet was not found`,
				),
			],
		}
	}
	const outlineRef = axis === 'row' ? 'summaryBelow' : 'summaryRight'
	const formatRef = axis === 'row' ? 'outlineLevelRow' : 'outlineLevelCol'
	const refs = [
		...outlineIndexes(op).map((index) => layoutRef(op.sheet, axis, index)),
		`sheet:${op.sheet}:outlinePr:${outlineRef}`,
		`sheet:${op.sheet}:sheetFormatPr:${formatRef}`,
	]
	return {
		opIndex,
		op,
		inverseOps: [],
		preimages: [{ kind: 'outline', outline: preimage }],
		issues: [
			{
				code: 'LOSSY_INVERSE',
				message: `Grouped ${axis === 'row' ? 'rows' : 'columns'} for ${op.sheet} cannot be restored with public operations`,
				surface: axis === 'row' ? 'row-layout' : 'column-layout',
				reason: axis === 'row' ? 'row-layout-created' : 'column-layout-created',
				refs: [...new Set(refs)],
			},
		],
	}
}

function cloneSheetColDef(def: SheetColDef): SheetColDef {
	return { ...def }
}

function cloneMatchingSheetColDef(sheet: Sheet | undefined, col: number): SheetColDef | null {
	const colDef = sheet?.colDefs.find((def) => def.min === col && def.max === col)
	return colDef ? cloneSheetColDef(colDef) : null
}

function cloneSheetRowDef(def: SheetRowDef): SheetRowDef {
	return { ...def }
}

function cloneSheetOutlinePr(outlinePr: SheetOutlinePr): SheetOutlinePr {
	return { ...outlinePr }
}

function cloneSheetFormatPr(sheetFormatPr: SheetFormatPr): SheetFormatPr {
	return { ...sheetFormatPr }
}

function cloneSheetTabColor(tabColor: SheetTabColor): SheetTabColor {
	return { ...tabColor }
}

function cloneSheetProtection(protection: SheetProtection): SheetProtection {
	return { ...protection }
}

function tabColorIsPublicRgb(tabColor: SheetTabColor | null): boolean {
	if (!tabColor?.rgb) return false
	return (
		tabColor.theme === undefined && tabColor.tint === undefined && tabColor.indexed === undefined
	)
}

function sheetProtectionInverseOp(
	sheet: string,
	protection: SheetProtection,
): Extract<Operation, { op: 'setSheetProtection' }> {
	type SheetProtectionInverseOptions = {
		-readonly [Key in keyof NonNullable<
			Extract<Operation, { op: 'setSheetProtection' }>['options']
		>]?: boolean
	}
	const options: SheetProtectionInverseOptions = {}
	for (const key of SHEET_PROTECTION_OPTION_KEYS) {
		const value = protection[key]
		if (value !== undefined) options[key] = value
	}
	return {
		op: 'setSheetProtection',
		sheet,
		...(protection.password ? { password: protection.password } : {}),
		...(Object.keys(options).length > 0 ? { options } : {}),
	}
}

const SHEET_PROTECTION_OPTION_KEYS = [
	'formatCells',
	'formatColumns',
	'formatRows',
	'insertColumns',
	'insertRows',
	'deleteColumns',
	'deleteRows',
	'sort',
	'autoFilter',
] as const satisfies readonly (keyof SheetProtection)[]

function unsupportedSheetProtectionKeys(protection: SheetProtection): string[] {
	const supported = new Set<string>(['sheet', 'password', ...SHEET_PROTECTION_OPTION_KEYS])
	return Object.keys(protection).filter((key) => !supported.has(key))
}

function outlineIndexes(op: Extract<Operation, { op: 'groupRows' | 'groupCols' }>): number[] {
	const indexes: number[] = []
	if (op.from <= op.to) {
		for (let index = op.from; index <= op.to; index++) indexes.push(index)
	}
	if (op.collapsed) {
		const summaryAfter =
			op.op === 'groupRows' ? (op.summaryBelow ?? true) : (op.summaryRight ?? true)
		const boundary = summaryAfter ? op.to + 1 : op.from - 1
		if (boundary >= 0) indexes.push(boundary)
	}
	return indexes
}

function layoutRef(sheet: string, axis: 'row' | 'col', index: number): string {
	return axis === 'row' ? `${sheet}!${index + 1}` : `${sheet}!${indexToColumn(index)}`
}

function sheetDeleteLossRefs(workbook: Workbook, sheet: Sheet): string[] {
	const refs: string[] = []
	const cellCount = sheet.cells.storageStats().cellCount
	if (cellCount > 0) refs.push(`${sheet.name}!cells:${cellCount}`)
	if (sheet.merges.length > 0) refs.push(`${sheet.name}!merges:${sheet.merges.length}`)
	if (sheet.tables.length > 0) refs.push(`${sheet.name}!tables:${sheet.tables.length}`)
	if (sheet.comments.size > 0) refs.push(`${sheet.name}!comments:${sheet.comments.size}`)
	if (sheet.threadedComments.length > 0) {
		refs.push(`${sheet.name}!threadedComments:${sheet.threadedComments.length}`)
	}
	if (sheet.hyperlinks.size > 0) refs.push(`${sheet.name}!hyperlinks:${sheet.hyperlinks.size}`)
	if (sheet.dataValidations.length > 0) {
		refs.push(`${sheet.name}!validations:${sheet.dataValidations.length}`)
	}
	if (sheet.conditionalFormats.length > 0) {
		refs.push(`${sheet.name}!conditionalFormats:${sheet.conditionalFormats.length}`)
	}
	if (sheet.autoFilter) refs.push(`${sheet.name}!autoFilter:${sheet.autoFilter.ref}`)
	if (sheet.rowHeights.size > 0) refs.push(`${sheet.name}!rowHeights:${sheet.rowHeights.size}`)
	if (sheet.colWidths.size > 0) refs.push(`${sheet.name}!colWidths:${sheet.colWidths.size}`)
	if (sheet.frozenRows > 0 || sheet.frozenCols > 0) refs.push(`${sheet.name}!pane`)
	if (sheet.state !== 'visible') refs.push(`${sheet.name}!state:${sheet.state}`)
	if (sheet.imageRefs.length > 0) refs.push(`${sheet.name}!images:${sheet.imageRefs.length}`)
	if (sheet.drawingObjectRefs.length > 0) {
		refs.push(`${sheet.name}!drawings:${sheet.drawingObjectRefs.length}`)
	}
	if (sheet.x14DataValidations.length > 0) {
		refs.push(`${sheet.name}!x14Validations:${sheet.x14DataValidations.length}`)
	}
	if (sheet.x14ConditionalFormats.length > 0) {
		refs.push(`${sheet.name}!x14ConditionalFormats:${sheet.x14ConditionalFormats.length}`)
	}
	if (sheet.advancedFilters.length > 0) {
		refs.push(`${sheet.name}!advancedFilters:${sheet.advancedFilters.length}`)
	}
	for (const entry of workbook.definedNames.list()) {
		const scopeSheet =
			entry.scope.kind === 'sheet' ? sheetNameForId(workbook, entry.scope.sheetId) : undefined
		if (scopeSheet !== undefined && sameSheetName(scopeSheet, sheet.name)) {
			refs.push(definedNameJournalKey(workbook, entry))
		}
		if (formulaReferencesSheet(workbook, entry.formula, scopeSheet ?? sheet.name, sheet.name)) {
			refs.push(definedNameJournalKey(workbook, entry))
		}
	}
	for (const workbookSheet of workbook.sheets) {
		if (sameSheetName(workbookSheet.name, sheet.name)) continue
		for (const [row, col, cell] of workbookSheet.cells.iterate()) {
			if (
				cell.formula &&
				formulaReferencesSheet(workbook, cell.formula, workbookSheet.name, sheet.name)
			) {
				refs.push(`${workbookSheet.name}!${toA1({ row, col })}`)
			}
		}
	}
	for (const chart of workbook.chartParts) {
		if (chart.sheetName !== undefined && sameSheetName(chart.sheetName, sheet.name)) {
			refs.push(`chart:${chart.partPath}`)
		}
		chart.series.forEach((series, seriesIndex) => {
			for (const [field, formula] of [
				['nameRef', series.nameRef],
				['categoryRef', series.categoryRef],
				['valueRef', series.valueRef],
			] as const) {
				if (
					formula &&
					formulaReferencesSheet(workbook, formula, chart.sheetName ?? sheet.name, sheet.name)
				) {
					refs.push(`chart:${chart.partPath}:series:${seriesIndex}:${field}`)
				}
			}
		})
	}
	for (const pivot of workbook.pivotTables) {
		if (sameSheetName(pivot.sheetName, sheet.name)) refs.push(`pivotTable:${pivot.partPath}`)
	}
	for (const pivotCache of workbook.pivotCaches) {
		if (pivotCache.sourceSheet !== undefined && sameSheetName(pivotCache.sourceSheet, sheet.name)) {
			refs.push(`pivotCache:${pivotCache.partPath}`)
		}
	}
	return [...new Set(refs)]
}

function formulaReferencesSheet(
	workbook: Workbook,
	formula: string,
	ownerSheet: string,
	sheetName: string,
): boolean {
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return false
	return extractRefs(parsed.value).some((ref) =>
		formulaRefReferencesSheet(workbook, ref, ownerSheet, sheetName),
	)
}

function formulaRefReferencesSheet(
	workbook: Workbook,
	ref: FormulaRef,
	ownerSheet: string,
	sheetName: string,
): boolean {
	if (ref.kind === 'sheetSpan') {
		return (
			sheetSpanIncludes(workbook, ref.startSheet, ref.endSheet, sheetName) ||
			formulaRefReferencesSheet(workbook, ref.target, ownerSheet, sheetName)
		)
	}
	const refSheet = 'sheet' in ref && ref.sheet !== undefined ? ref.sheet : ownerSheet
	return sameSheetName(refSheet, sheetName)
}

function journalSetCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setCells' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellEditPreimages(
		workbook,
		op.sheet,
		op.updates.map((update) => update.ref),
	)
	const { inverseOps, issues } = inverseCellOps(cells)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'cells', cells }],
		issues,
	}
}

function journalSetFormula(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setFormula' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellEditPreimages(workbook, op.sheet, [op.ref])
	const formulaIssues = formulaParseIssues(op.op, op.sheet, op.ref, op.formula)
	const { inverseOps, issues } = inverseCellOps(cells)
	return {
		opIndex,
		op,
		inverseOps: formulaIssues.length === 0 ? inverseOps : [],
		preimages: [{ kind: 'cells', cells }],
		issues: [...formulaIssues, ...issues],
	}
}

function journalFillFormula(
	workbook: Workbook,
	op: Extract<Operation, { op: 'fillFormula' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellEditPreimages(workbook, op.sheet, refsInRange(op.range))
	const formulaIssues = formulaParseIssues(op.op, op.sheet, op.range, op.formula)
	const { inverseOps, issues } = inverseCellOps(cells)
	return {
		opIndex,
		op,
		inverseOps: formulaIssues.length === 0 ? inverseOps : [],
		preimages: [{ kind: 'cells', cells }],
		issues: [...formulaIssues, ...issues],
	}
}

function formulaParseIssues(
	opName: string,
	sheetName: string,
	ref: string,
	formula: string,
): MutationJournalIssue[] {
	const parsed = cachedParseFormula(normalizeFormulaInput(formula))
	if (parsed.ok) return []
	return [
		{
			code: 'UNSUPPORTED_VALUE',
			message: `Cannot build exact rollback journal for ${opName} because ${formula} is not a valid formula`,
			surface: 'formulas',
			reason: 'value-unsupported',
			refs: [`${sheetName}!${ref}`],
		},
	]
}

function journalSetRichText(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setRichText' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellEditPreimages(workbook, op.sheet, [op.ref])
	const { inverseOps, issues } = inverseCellOps(cells)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'cells', cells }],
		issues,
	}
}

function journalClearRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'clearRange' }>,
	opIndex: number,
): DraftJournalEntry {
	const refs = refsInRange(op.range)
	if (op.what === 'styles') {
		const cells = cellPreimages(workbook, op.sheet, refs)
		return {
			opIndex,
			op,
			inverseOps: styleInverseOps(cells),
			preimages: [{ kind: 'cells', cells }],
			issues: [],
		}
	}
	const cells = cellEditPreimages(workbook, op.sheet, refs, {
		blockedSpillBlockers: op.what !== 'formulas',
	})
	const { inverseOps: cellInverseOps, issues } = inverseCellOps(cells)
	const inverseOps =
		op.what === 'all' ? [...cellInverseOps, ...styleInverseOps(cells)] : cellInverseOps
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'cells', cells }],
		issues,
	}
}

function journalStyleRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setNumberFormat' | 'setStyle' }>,
	opIndex: number,
): DraftJournalEntry {
	const cells = cellPreimages(workbook, op.sheet, refsInRange(op.range))
	const issues = styleRegistryGrowthIssues(workbook, op, cells)
	return {
		opIndex,
		op,
		inverseOps: styleInverseOps(cells),
		preimages: [{ kind: 'cells', cells }],
		issues,
	}
}

function styleRegistryGrowthIssues(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setNumberFormat' | 'setStyle' }>,
	cells: readonly MutationJournalCellPreimage[],
): MutationJournalIssue[] {
	const styles = workbook.styles.clone()
	const originalSize = workbook.styles.size
	const createdStyleIds = new Set<number>()
	const refs: string[] = []
	for (const cell of cells) {
		const nextStyle =
			op.op === 'setNumberFormat'
				? { ...cell.style, numberFormat: op.format }
				: mergeStyleInputForJournal(cell.style, op.style)
		const styleId = styles.register(nextStyle)
		if (styleId < originalSize || styleId === cell.styleId) continue
		createdStyleIds.add(styleId)
		refs.push(`${cell.sheet}!${cell.ref}`)
	}
	if (createdStyleIds.size === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `${op.op} creates ${createdStyleIds.size} style registry ${
				createdStyleIds.size === 1 ? 'entry' : 'entries'
			} that cannot be removed with public operations`,
			surface: 'package-parts',
			reason: 'package-part-preservation',
			refs: [...new Set(refs)],
		},
	]
}

function mergeStyleInputForJournal(current: CellStyle, input: StyleInput): CellStyle {
	return {
		...current,
		...(input.font && { font: { ...current.font, ...input.font } }),
		...(input.fill && { fill: { ...current.fill, ...input.fill } }),
		...(input.border && { border: { ...current.border, ...input.border } }),
		...(input.alignment && { alignment: { ...current.alignment, ...input.alignment } }),
		...(input.numberFormat !== undefined && { numberFormat: input.numberFormat }),
	}
}

function journalInsertAxis(
	op: Extract<Operation, { op: 'insertRows' | 'insertCols' }>,
	opIndex: number,
	axis: 'row' | 'col',
): DraftJournalEntry {
	const inverseOps: Operation[] =
		axis === 'row'
			? [{ op: 'deleteRows', sheet: op.sheet, at: op.at, count: op.count }]
			: [{ op: 'deleteCols', sheet: op.sheet, at: op.at, count: op.count }]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [
			{
				kind: 'structural',
				structural: { sheet: op.sheet, axis, at: op.at, count: op.count, deletedCells: [] },
			},
		],
		issues: [],
	}
}

function journalDeleteAxis(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteRows' | 'deleteCols' }>,
	opIndex: number,
	axis: 'row' | 'col',
): DraftJournalEntry {
	const preimage = structuralDeletePreimage(workbook, op.sheet, axis, op.at, op.count)
	const { inverseOps: cellInverseOps, issues: cellIssues } = inverseCellOps(preimage.deletedCells)
	const inverseOps: Operation[] =
		axis === 'row'
			? [{ op: 'insertRows', sheet: op.sheet, at: op.at, count: op.count }]
			: [{ op: 'insertCols', sheet: op.sheet, at: op.at, count: op.count }]
	inverseOps.push(...cellInverseOps, ...styleInverseOps(preimage.deletedCells))
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'structural', structural: preimage }],
		issues: [...cellIssues, ...structuralDeleteIssues(workbook, preimage)],
	}
}

function journalMergeCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'mergeCells' }>,
	opIndex: number,
): DraftJournalEntry {
	const merge = mergePreimage(workbook, op.sheet, op.range)
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'unmergeCells', sheet: op.sheet, range: merge.range }],
		preimages: [{ kind: 'merge', merge }],
		issues: [],
	}
}

function journalUnmergeCells(
	workbook: Workbook,
	op: Extract<Operation, { op: 'unmergeCells' }>,
	opIndex: number,
): DraftJournalEntry {
	const merge = mergePreimage(workbook, op.sheet, op.range)
	return {
		opIndex,
		op,
		inverseOps: merge.existed ? [{ op: 'mergeCells', sheet: op.sheet, range: merge.range }] : [],
		preimages: [{ kind: 'merge', merge }],
		issues: [],
	}
}

function journalSetDataValidation(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDataValidation' }>,
	opIndex: number,
): DraftJournalEntry {
	const validation = dataValidationPreimage(workbook, op.sheet, op.range)
	const { inverseOps, issues } = restoreDataValidationOps(validation)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'data-validations', validations: [validation] }],
		issues,
	}
}

function journalDeleteDataValidation(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteDataValidation' }>,
	opIndex: number,
): DraftJournalEntry {
	const validation = dataValidationPreimage(workbook, op.sheet, op.range)
	const { inverseOps, issues } = validation.validation
		? restoreDataValidationOps(validation)
		: { inverseOps: [], issues: [] }
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'data-validations', validations: [validation] }],
		issues: [
			...issues,
			...(validation.validation
				? dataValidationRestoreOrderIssues(workbook, op.sheet, [validation])
				: []),
		],
	}
}

function journalSetAutoFilter(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setAutoFilter' }>,
	opIndex: number,
): DraftJournalEntry {
	const autoFilter = autoFilterPreimage(workbook, op.sheet)
	const valueIssues = autoFilterValueIssues(op)
	const { inverseOps, issues } = restoreAutoFilterOps(autoFilter)
	return {
		opIndex,
		op,
		inverseOps: valueIssues.length > 0 ? [] : inverseOps,
		preimages: [{ kind: 'auto-filter', autoFilter }],
		issues: valueIssues.length > 0 ? valueIssues : issues,
	}
}

function autoFilterValueIssues(
	op: Extract<Operation, { op: 'setAutoFilter' }>,
): MutationJournalIssue[] {
	for (const field of ['sortRef', 'sortBy'] as const) {
		const value = op[field]
		if (value === undefined) continue
		try {
			parseRange(value)
		} catch {
			return [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot build exact rollback journal for setAutoFilter because ${field} ${value} is invalid`,
					surface: 'auto-filters',
					reason: 'value-unsupported',
					refs: [`${op.sheet}!${value}`],
				},
			]
		}
	}
	if (op.descending !== undefined && typeof op.descending !== 'boolean') {
		return [
			{
				code: 'UNSUPPORTED_VALUE',
				message:
					'Cannot build exact rollback journal for setAutoFilter because descending is not boolean',
				surface: 'auto-filters',
				reason: 'value-unsupported',
				refs: [`${op.sheet}!${op.range}`],
			},
		]
	}
	return []
}

function journalClearAutoFilter(
	workbook: Workbook,
	op: Extract<Operation, { op: 'clearAutoFilter' }>,
	opIndex: number,
): DraftJournalEntry {
	const autoFilter = autoFilterPreimage(workbook, op.sheet)
	const { inverseOps, issues } = autoFilter.autoFilter
		? restoreAutoFilterOps(autoFilter)
		: { inverseOps: [], issues: [] }
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'auto-filter', autoFilter }],
		issues,
	}
}

function journalSetConditionalFormat(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setConditionalFormat' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = conditionalFormatPreimage(workbook, op.sheet, op.range)
	const { inverseOps, issues } =
		preimage.formats.length > 0
			? restoreConditionalFormatOps(op.sheet, preimage.formats)
			: {
					inverseOps: [
						{ op: 'deleteConditionalFormat', sheet: op.sheet, range: op.range },
					] satisfies Operation[],
					issues: [],
				}
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'conditional-formats', conditionalFormats: preimage }],
		issues: [
			...issues,
			...conditionalFormatRestoreOrderIssues(workbook, op.sheet, preimage.formats),
		],
	}
}

function journalDeleteConditionalFormat(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteConditionalFormat' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = conditionalFormatPreimage(workbook, op.sheet, op.range)
	const { inverseOps, issues } = restoreConditionalFormatOps(op.sheet, preimage.formats)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'conditional-formats', conditionalFormats: preimage }],
		issues: [
			...issues,
			...(op.range === undefined
				? []
				: conditionalFormatRestoreOrderIssues(workbook, op.sheet, preimage.formats)),
			...(op.range === undefined
				? [
						{
							code: 'LOSSY_INVERSE' as const,
							message:
								'Conditional-format deletion without a range may not restore original rule ordering exactly',
							surface: 'conditional-formats' as const,
							reason: 'metadata-order' as const,
						},
					]
				: []),
		],
	}
}

function journalSetPageSetup(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPageSetup' }>,
	opIndex: number,
): DraftJournalEntry {
	if (!workbook.getSheet(op.sheet)) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [],
			issues: [
				missingSheetTopologyIssue(
					op.sheet,
					`Cannot restore page setup for ${op.sheet} because the sheet was not found`,
				),
			],
		}
	}
	const preimage = pageSetupPreimage(workbook, op.sheet)
	const { inverseOps, issues } = restorePageSetupOps(preimage, op.setup)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'page-setup', pageSetup: preimage }],
		issues,
	}
}

function journalSetPrintArea(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPrintArea' }>,
	opIndex: number,
): DraftJournalEntry {
	if (!workbook.getSheet(op.sheet)) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [],
			issues: [
				missingSheetTopologyIssue(
					op.sheet,
					`Cannot restore print area for ${op.sheet} because the sheet was not found`,
				),
			],
		}
	}
	const rangeIssues = printAreaRangeIssues(op)
	const preimage = definedNamePreimage(workbook, '_xlnm.Print_Area', op.sheet)
	const { inverseOps, issues } = restoreDefinedNameOps(workbook, preimage)
	return {
		opIndex,
		op,
		inverseOps: rangeIssues.length === 0 ? inverseOps : [],
		preimages: [{ kind: 'defined-name', definedName: preimage }],
		issues: [...rangeIssues, ...issues],
	}
}

function printAreaRangeIssues(
	op: Extract<Operation, { op: 'setPrintArea' }>,
): MutationJournalIssue[] {
	try {
		parseRange(op.range)
		return []
	} catch {
		return [
			{
				code: 'UNSUPPORTED_VALUE',
				message: `Cannot build exact rollback journal for setPrintArea because ${op.range} is not a valid range`,
				surface: 'defined-names',
				reason: 'value-unsupported',
				refs: [`${op.sheet}!${op.range}`],
			},
		]
	}
}

function journalSetDefinedName(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDefinedName' }>,
	opIndex: number,
): DraftJournalEntry {
	if (op.scope !== undefined && !workbook.getSheet(op.scope)) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [],
			issues: [
				missingSheetTopologyIssue(
					op.scope,
					`Cannot restore defined name ${op.name} because sheet scope ${op.scope} was not found`,
				),
			],
		}
	}
	const nameIssues = definedNameValueIssues(op)
	const preimage = definedNamePreimage(workbook, op.name, op.scope)
	const { inverseOps, issues } = restoreDefinedNameOps(workbook, preimage)
	const allIssues = [
		...nameIssues,
		...issues,
		...savedSourceDefinedNamePackageStateIssues(workbook, op.op, op.name),
	]
	return {
		opIndex,
		op,
		inverseOps: nameIssues.length === 0 ? inverseOps : [],
		preimages: [{ kind: 'defined-name', definedName: preimage }],
		issues: allIssues,
	}
}

function definedNameValueIssues(
	op: Extract<Operation, { op: 'setDefinedName' }>,
): MutationJournalIssue[] {
	if (!validateExcelDefinedName(op.name)) return []
	return [
		{
			code: 'UNSUPPORTED_VALUE',
			message: `Cannot build exact rollback journal for setDefinedName because ${op.name} is not a valid defined name`,
			surface: 'defined-names',
			reason: 'value-unsupported',
			refs: [op.scope ? `${op.scope}!${op.name}` : op.name],
		},
	]
}

function journalDeleteDefinedName(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteDefinedName' }>,
	opIndex: number,
): DraftJournalEntry {
	if (op.scope !== undefined && !workbook.getSheet(op.scope)) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [],
			issues: [
				missingSheetTopologyIssue(
					op.scope,
					`Cannot restore defined name ${op.name} because sheet scope ${op.scope} was not found`,
				),
			],
		}
	}
	const preimage = definedNamePreimage(workbook, op.name, op.scope)
	const { inverseOps, issues } = preimage.definedName
		? restoreDefinedNameOps(workbook, preimage)
		: {
				inverseOps: [],
				issues: [
					{
						code: 'UNSUPPORTED_VALUE',
						message: `Cannot build exact rollback journal for deleteDefinedName because ${op.name} does not exist`,
						surface: 'defined-names',
						reason: 'value-unsupported',
						refs: [op.scope ? `${op.scope}!${op.name}` : op.name],
					} satisfies MutationJournalIssue,
				],
			}
	const allIssues = preimage.definedName
		? [...issues, ...savedSourceDefinedNamePackageStateIssues(workbook, op.op, op.name)]
		: issues
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'defined-name', definedName: preimage }],
		issues: allIssues,
	}
}

function journalRenameSheet(
	workbook: Workbook,
	op: Extract<Operation, { op: 'renameSheet' }>,
	opIndex: number,
): DraftJournalEntry {
	const formulaBindings = formulaBindingJournalAddendum(
		workbookFormulaBindingCellPreimages(workbook),
	)
	const issues = [
		...formulaBindings.issues,
		...savedSourcePackageStateIssues(workbook, op.op, [`sheet:${op.sheet}`, `sheet:${op.newName}`]),
	]
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'renameSheet', sheet: op.newName, newName: op.sheet }],
		preimages: formulaBindings.preimages,
		issues,
	}
}

function journalRenameTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'renameTable' }>,
	opIndex: number,
): DraftJournalEntry {
	const tableRename = { table: op.table, existed: true }
	const formulaBindings = formulaBindingJournalAddendum(
		workbookFormulaBindingCellPreimages(workbook),
	)
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'renameTable', table: op.newName, newName: op.table }],
		preimages: [{ kind: 'table-rename', tableRename }, ...formulaBindings.preimages],
		issues: formulaBindings.issues,
	}
}

function journalCreateTable(
	op: Extract<Operation, { op: 'createTable' }>,
	opIndex: number,
): DraftJournalEntry {
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'deleteTable', table: op.name }],
		preimages: [
			{
				kind: 'table',
				table: { table: null, sheet: op.sheet, ref: op.ref },
			},
		],
		issues: [],
	}
}

function journalSortRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'sortRange' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const dataRange = sheet ? sortRangeDataRange(sheet, op.range, op.by) : null
	const refs = dataRange ? refsInParsedRange(dataRange) : []
	const rangeCells = cellPreimages(workbook, op.sheet, refs)
	const formulaBindings = formulaBindingOnlyPreimages(workbook, op.sheet, refs)
	const cells = uniqueCellPreimages([...rangeCells, ...formulaBindings])
	const { inverseOps: cellInverseOps, issues: cellIssues } = inverseCellOps(cells)
	const comments =
		sheet && dataRange && hasMapRefInRange(sheet.comments, dataRange)
			? commentPreimages(workbook, op.sheet, refs)
			: []
	const hyperlinks =
		sheet && dataRange && hasMapRefInRange(sheet.hyperlinks, dataRange)
			? hyperlinkPreimages(workbook, op.sheet, refs)
			: []
	const validations =
		sheet && dataRange
			? sortRangeDataValidationRestoration(sheet, op, dataRange)
			: EMPTY_METADATA_RESTORATION
	const conditionalFormats =
		sheet && dataRange
			? sortRangeConditionalFormatRestoration(sheet, op, dataRange)
			: EMPTY_METADATA_RESTORATION
	return {
		opIndex,
		op,
		inverseOps: [
			...cellInverseOps,
			...styleInverseOps(rangeCells),
			...restoreCommentOps(comments),
			...restoreHyperlinkOps(hyperlinks),
			...validations.inverseOps,
			...conditionalFormats.inverseOps,
		],
		preimages: [
			{ kind: 'cells', cells },
			...comments.map((comment) => ({ kind: 'comment' as const, comment })),
			...hyperlinks.map((hyperlink) => ({ kind: 'hyperlink' as const, hyperlink })),
			...validations.preimages,
			...conditionalFormats.preimages,
		],
		issues: [
			...cellIssues,
			...validations.issues,
			...conditionalFormats.issues,
			...sortRangeMetadataIssues(workbook, op),
			...commentRestoreIssues(comments),
		],
	}
}

function journalCopyRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'copyRange' }>,
	opIndex: number,
): DraftJournalEntry {
	const mode = op.mode ?? 'all'
	if (!copyRangeCellModeSupported(mode)) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [],
			issues: [
				{
					code: 'UNSUPPORTED_OPERATION',
					message: `No reversible journal support for copyRange mode ${mode}`,
					surface: classifyMutationJournalOperationPrimarySurface(op),
					reason: 'operation-unsupported',
				},
			],
		}
	}
	const targetSheet = op.targetSheet ?? op.sheet
	const targetRange = transferTargetRange(op.source, op.target)
	if (mode === 'comments') {
		const targetComments = commentPreimages(workbook, targetSheet, refsInParsedRange(targetRange))
		return {
			opIndex,
			op,
			inverseOps: restoreCommentOps(targetComments),
			preimages: targetComments.map((comment) => ({ kind: 'comment', comment })),
			issues: commentTransferIssues(workbook, op.sheet, op.source, targetSheet, targetRange),
		}
	}
	if (mode === 'hyperlinks') {
		const hyperlinks = hyperlinkPreimages(workbook, targetSheet, refsInParsedRange(targetRange))
		return {
			opIndex,
			op,
			inverseOps: restoreHyperlinkOps(hyperlinks),
			preimages: hyperlinks.map((hyperlink) => ({ kind: 'hyperlink', hyperlink })),
			issues: [],
		}
	}
	if (mode === 'validations') {
		const { inverseOps, preimages, issues } = validationTransferRestoration(
			workbook,
			op.sheet,
			parseRange(op.source),
			targetSheet,
			targetRange,
			false,
		)
		return {
			opIndex,
			op,
			inverseOps,
			preimages,
			issues,
		}
	}
	const targetRefs = refsInParsedRange(targetRange)
	const cells = copyRangeOverwritesFormulas(mode)
		? cellEditPreimages(workbook, targetSheet, targetRefs)
		: cellPreimages(workbook, targetSheet, targetRefs)
	const { inverseOps: cellInverseOps, issues: cellIssues } = copyRangeRestoration(cells, mode)
	const metadataRestoration =
		mode === 'all'
			? copyRangeAllMetadataRestoration(
					workbook,
					op.sheet,
					parseRange(op.source),
					targetSheet,
					targetRange,
				)
			: EMPTY_METADATA_RESTORATION
	const conditionalFormats = copyRangeTransfersConditionalFormats(mode)
		? conditionalFormatTransferRestoration(
				workbook,
				op.sheet,
				parseRange(op.source),
				targetSheet,
				targetRange,
				false,
			)
		: EMPTY_METADATA_RESTORATION
	const merges = copyRangeTransfersMerges(mode)
		? mergeTransferRestoration(
				workbook,
				op.sheet,
				parseRange(op.source),
				targetSheet,
				targetRange,
				false,
			)
		: EMPTY_METADATA_RESTORATION
	return {
		opIndex,
		op,
		inverseOps: [
			...cellInverseOps,
			...metadataRestoration.inverseOps,
			...conditionalFormats.inverseOps,
			...merges.inverseOps,
		],
		preimages: [
			{ kind: 'cells', cells },
			...metadataRestoration.preimages,
			...conditionalFormats.preimages,
			...merges.preimages,
		],
		issues: [
			...cellIssues,
			...metadataRestoration.issues,
			...conditionalFormats.issues,
			...merges.issues,
		],
	}
}

function journalMoveRange(
	workbook: Workbook,
	op: Extract<Operation, { op: 'moveRange' }>,
	opIndex: number,
): DraftJournalEntry {
	const mode = op.mode ?? 'all'
	if (!copyRangeCellModeSupported(mode)) {
		return {
			opIndex,
			op,
			inverseOps: [],
			preimages: [],
			issues: [
				{
					code: 'UNSUPPORTED_OPERATION',
					message: `No reversible journal support for moveRange mode ${mode}`,
					surface: classifyMutationJournalOperationPrimarySurface(op),
					reason: 'operation-unsupported',
				},
			],
		}
	}
	const sourceRange = parseRange(op.source)
	const targetSheet = op.targetSheet ?? op.sheet
	const targetRange = transferTargetRange(op.source, op.target)
	const sourceRefs = refsInParsedRange(sourceRange)
	const targetRefs = refsInParsedRange(targetRange)
	if (mode === 'comments') {
		const targetComments = commentPreimages(workbook, targetSheet, targetRefs)
		const sourceComments = commentPreimages(workbook, op.sheet, sourceRefs)
		return {
			opIndex,
			op,
			inverseOps: moveRangeCommentRestoreOps(sourceComments, targetComments),
			preimages: [...targetComments, ...sourceComments].map((comment) => ({
				kind: 'comment',
				comment,
			})),
			issues: commentTransferIssues(workbook, op.sheet, op.source, targetSheet, targetRange, {
				move: true,
			}),
		}
	}
	if (mode === 'hyperlinks') {
		const targetHyperlinks = hyperlinkPreimages(workbook, targetSheet, targetRefs)
		const sourceHyperlinks = hyperlinkPreimages(workbook, op.sheet, sourceRefs)
		return {
			opIndex,
			op,
			inverseOps: moveRangeHyperlinkRestoreOps(sourceHyperlinks, targetHyperlinks),
			preimages: [...targetHyperlinks, ...sourceHyperlinks].map((hyperlink) => ({
				kind: 'hyperlink',
				hyperlink,
			})),
			issues: [],
		}
	}
	if (mode === 'validations') {
		const { inverseOps, preimages, issues } = validationTransferRestoration(
			workbook,
			op.sheet,
			sourceRange,
			targetSheet,
			targetRange,
			true,
		)
		return {
			opIndex,
			op,
			inverseOps,
			preimages,
			issues,
		}
	}
	const sourceCells = copyRangeOverwritesFormulas(mode)
		? cellEditPreimages(workbook, op.sheet, sourceRefs)
		: cellPreimages(workbook, op.sheet, sourceRefs)
	const targetCells = copyRangeOverwritesFormulas(mode)
		? cellEditPreimages(workbook, targetSheet, targetRefs)
		: cellPreimages(workbook, targetSheet, targetRefs)
	const { inverseOps: sourceInverseOps, issues: sourceIssues } = moveRangeSourceRestoration(
		sourceCells,
		mode,
	)
	const { inverseOps: targetInverseOps, issues: targetIssues } = copyRangeRestoration(
		targetCells,
		mode,
	)
	const metadataRestoration =
		mode === 'all'
			? moveRangeAllMetadataRestoration(workbook, op.sheet, sourceRange, targetSheet, targetRange)
			: EMPTY_METADATA_RESTORATION
	const conditionalFormats = copyRangeTransfersConditionalFormats(mode)
		? conditionalFormatTransferRestoration(
				workbook,
				op.sheet,
				sourceRange,
				targetSheet,
				targetRange,
				true,
			)
		: EMPTY_METADATA_RESTORATION
	const merges = copyRangeTransfersMerges(mode)
		? mergeTransferRestoration(workbook, op.sheet, sourceRange, targetSheet, targetRange, true)
		: EMPTY_METADATA_RESTORATION
	const formulaSurfaces = copyRangeOverwritesFormulas(mode)
		? moveRangeFormulaSurfaceRestoration(workbook, op, sourceRange, targetRange)
		: EMPTY_METADATA_RESTORATION
	const cells = uniqueCellPreimages([...targetCells, ...sourceCells])
	return {
		opIndex,
		op,
		inverseOps: [
			...targetInverseOps,
			...sourceInverseOps,
			...metadataRestoration.inverseOps,
			...conditionalFormats.inverseOps,
			...merges.inverseOps,
			...formulaSurfaces.inverseOps,
		],
		preimages: [
			{ kind: 'cells', cells },
			...metadataRestoration.preimages,
			...conditionalFormats.preimages,
			...merges.preimages,
			...formulaSurfaces.preimages,
		],
		issues: [
			...targetIssues,
			...sourceIssues,
			...metadataRestoration.issues,
			...conditionalFormats.issues,
			...merges.issues,
			...moveRangeOverlapIssues(op, sourceRange, targetRange),
			...formulaSurfaces.issues,
		],
	}
}

function journalDeleteTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteTable' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = tablePreimage(workbook, op.table)
	const { inverseOps, issues } = restoreTableOps(preimage)
	const formulaBindings = formulaBindingJournalAddendum(
		workbookFormulaBindingCellPreimages(workbook),
	)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'table', table: preimage }, ...formulaBindings.preimages],
		issues: [...issues, ...formulaBindings.issues],
	}
}

function journalResizeTable(
	workbook: Workbook,
	op: Extract<Operation, { op: 'resizeTable' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = tablePreimage(workbook, op.table)
	const { inverseOps, issues } = preimage.table
		? restoreExistingTableOps(preimage.table, preimage.sheet ?? undefined, {
				includeCreate: false,
				tableName: op.table,
			})
		: { inverseOps: [], issues: missingTableIssues(op.table) }
	const formulaBindings = formulaBindingJournalAddendum(
		workbookFormulaBindingCellPreimages(workbook),
	)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'table', table: preimage }, ...formulaBindings.preimages],
		issues: [...issues, ...formulaBindings.issues],
	}
}

function journalSetTableColumn(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTableColumn' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = tableColumnPreimage(workbook, op.table, op.column)
	const { inverseOps, issues } = restoreTableColumnOps(op, preimage)
	const formulaBindings = formulaBindingJournalAddendum(
		tableColumnFormulaBindingCellPreimages(workbook, op),
	)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'table-column', tableColumn: preimage }, ...formulaBindings.preimages],
		issues: [...issues, ...formulaBindings.issues],
	}
}

function journalSetTableStyle(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTableStyle' }>,
	opIndex: number,
): DraftJournalEntry {
	const style = tableStylePreimage(workbook, op.table)
	const issues =
		op.styleName === undefined &&
		op.showFirstColumn === undefined &&
		op.showLastColumn === undefined &&
		op.showRowStripes === undefined &&
		op.showColumnStripes === undefined
			? [
					{
						code: 'UNSUPPORTED_VALUE',
						message: `Cannot build exact rollback journal for setTableStyle because no style field was provided for table ${op.table}`,
						surface: 'tables',
						reason: 'value-unsupported',
						refs: [`table:${op.table}`],
					} satisfies MutationJournalIssue,
				]
			: tableStyleRestoreIssues(op, style.style)
	return {
		opIndex,
		op,
		inverseOps: issues.length === 0 ? [tableStyleSetOperation(op.table, style.style)] : [],
		preimages: [{ kind: 'table-style', tableStyle: style }],
		issues,
	}
}

function journalSetComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setComment' }>,
	opIndex: number,
): DraftJournalEntry {
	const comment = commentPreimage(workbook, op.sheet, op.ref)
	const threadedComments = threadedCommentPreimagesAtRef(workbook, op.sheet, op.ref)
	const { inverseOps, issues } = restoreSetCommentOps(op, comment, threadedComments)
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [
			{ kind: 'comment', comment },
			...threadedComments.map((threadedComment) => ({
				kind: 'threaded-comment' as const,
				threadedComment,
			})),
		],
		issues: [...issues, ...commentRestoreIssues([comment])],
	}
}

function journalDeleteComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteComment' }>,
	opIndex: number,
): DraftJournalEntry {
	const comment = commentPreimage(workbook, op.sheet, op.ref)
	const threadedComments = threadedCommentPreimagesAtRef(workbook, op.sheet, op.ref)
	const inverseOps: Operation[] = comment.comment
		? [
				{
					op: 'setComment',
					sheet: op.sheet,
					ref: comment.ref,
					text: comment.comment.text,
					...(comment.comment.author !== undefined ? { author: comment.comment.author } : {}),
				},
			]
		: []
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [
			{ kind: 'comment', comment },
			...threadedComments.map((threadedComment) => ({
				kind: 'threaded-comment' as const,
				threadedComment,
			})),
		],
		issues: [
			...commentRestoreIssues([comment]),
			...deleteCommentThreadedCommentIssues(op.sheet, op.ref, threadedComments),
		],
	}
}

function journalSetHyperlink(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setHyperlink' }>,
	opIndex: number,
): DraftJournalEntry {
	const hyperlink = hyperlinkPreimage(workbook, op.sheet, op.ref)
	const issues: MutationJournalIssue[] =
		hasJournalHyperlinkDestination(op.url) || hasJournalHyperlinkDestination(op.location)
			? []
			: [
					{
						code: 'UNSUPPORTED_VALUE',
						message: `Cannot build exact rollback journal for setHyperlink because ${op.sheet}!${op.ref} has no hyperlink destination`,
						surface: 'hyperlinks',
						reason: 'value-unsupported',
						refs: [`${op.sheet}!${op.ref}`],
					},
				]
	const inverseOps: Operation[] = hyperlink.hyperlink
		? [setHyperlinkInverse(op.sheet, hyperlink.ref, hyperlink.hyperlink)]
		: [{ op: 'deleteHyperlink', sheet: op.sheet, ref: hyperlink.ref }]
	return {
		opIndex,
		op,
		inverseOps: issues.length === 0 ? inverseOps : [],
		preimages: [{ kind: 'hyperlink', hyperlink }],
		issues,
	}
}

function hasJournalHyperlinkDestination(value: string | undefined): boolean {
	return typeof value === 'string' && value.trim().length > 0
}

function workbookMetadataUnsupportedValueIssue(
	opName: string,
	field: string,
	refs: readonly string[],
): MutationJournalIssue {
	return {
		code: 'UNSUPPORTED_VALUE',
		message: `Cannot build exact rollback journal for ${opName} because ${field} is not supported by workbook metadata validation`,
		surface: 'workbook-metadata',
		reason: 'value-unsupported',
		refs,
	}
}

function isJournalMode(mode: unknown): boolean {
	return mode === undefined || mode === 'merge' || mode === 'replace'
}

function isPlainJournalObject(value: unknown): value is Record<string, unknown> {
	return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function hasInvalidWorkbookViewValue(view: unknown): boolean {
	if (!isPlainJournalObject(view)) return true
	for (const key of ['activeTab', 'firstSheet', 'tabRatio'] as const) {
		const value = view[key]
		if (value === undefined || value === null) continue
		if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) return true
	}
	return false
}

const JOURNAL_THEME_COLOR_SLOTS = new Set([
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
])

function journalDeleteHyperlink(
	workbook: Workbook,
	op: Extract<Operation, { op: 'deleteHyperlink' }>,
	opIndex: number,
): DraftJournalEntry {
	const hyperlink = hyperlinkPreimage(workbook, op.sheet, op.ref)
	const inverseOps: Operation[] = hyperlink.hyperlink
		? [setHyperlinkInverse(op.sheet, hyperlink.ref, hyperlink.hyperlink)]
		: []
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'hyperlink', hyperlink }],
		issues: [],
	}
}

function journalSetThreadedComment(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setThreadedComment' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = threadedCommentPreimage(workbook, op)
	const inverseOps: Operation[] = preimage.threadedComment
		? [
				{
					op: 'setThreadedComment',
					sheet: preimage.sheet,
					text: preimage.threadedComment.text,
					...threadedCommentStableSelector(preimage),
				},
			]
		: []
	const issues: MutationJournalIssue[] = preimage.threadedComment
		? []
		: [
				{
					code: 'LOSSY_INVERSE',
					message: `Threaded comment selector on ${op.sheet} cannot be resolved exactly`,
					surface: 'comments',
					reason: 'threaded-comment-selector',
				},
			]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'threaded-comment', threadedComment: preimage }],
		issues,
	}
}

function journalSetDrawingText(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDrawingText' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = drawingTextPreimage(workbook, op)
	const inverseOps: Operation[] =
		preimage.drawingObject?.text !== undefined
			? [
					{
						op: 'setDrawingText',
						sheet: preimage.sheet,
						text: preimage.drawingObject.text,
						...drawingObjectStableSelector(preimage),
					},
				]
			: []
	const issues: MutationJournalIssue[] =
		preimage.drawingObject?.text !== undefined
			? []
			: [
					{
						code: 'LOSSY_INVERSE',
						message: `Drawing object selector on ${op.sheet} cannot be resolved to editable text exactly`,
						surface: 'drawings',
						reason: 'drawing-text-selector',
					},
				]
	return {
		opIndex,
		op,
		inverseOps,
		preimages: [{ kind: 'drawing-text', drawingText: preimage }],
		issues,
	}
}

function journalSetChartSeriesSource(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setChartSeriesSource' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = chartSeriesPreimage(workbook, op)
	const inverse = chartSeriesInverseOperation(op, preimage)
	const issues: MutationJournalIssue[] = inverse.exact
		? []
		: [
				{
					code: 'LOSSY_INVERSE',
					message: `Chart series selector cannot be restored exactly for series ${op.seriesIndex}`,
					surface: 'charts',
					reason: 'chart-series-unsettable',
				},
			]
	return {
		opIndex,
		op,
		inverseOps: inverse.op ? [inverse.op] : [],
		preimages: [{ kind: 'chart-series', chartSeries: preimage }],
		issues,
	}
}

function journalSetPivotCache(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPivotCache' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = pivotCachePreimage(workbook, op)
	const inverse = pivotCacheInverseOperation(op, preimage)
	const valueIssues = pivotCacheValueIssues(op)
	const issues: MutationJournalIssue[] =
		valueIssues.length > 0
			? valueIssues
			: inverse.exact
				? []
				: [
						{
							code: 'LOSSY_INVERSE',
							message: 'Pivot cache selector cannot be restored exactly',
							surface: 'pivot-caches',
							reason: 'pivot-cache-unsettable',
						},
					]
	return {
		opIndex,
		op,
		inverseOps: valueIssues.length === 0 && inverse.op ? [inverse.op] : [],
		preimages: [{ kind: 'pivot-cache', pivotCache: preimage }],
		issues,
	}
}

function pivotCacheValueIssues(
	op: Extract<Operation, { op: 'setPivotCache' }>,
): MutationJournalIssue[] {
	if (op.cacheId === undefined && op.partPath === undefined && op.pivotTable === undefined) {
		return [pivotCacheUnsupportedValueIssue('selector')]
	}
	if (!hasPivotCacheUpdate(op)) {
		return [pivotCacheUnsupportedValueIssue('update')]
	}
	if (op.sourceSheet !== undefined && typeof op.sourceSheet !== 'string') {
		return [pivotCacheUnsupportedValueIssue('sourceSheet')]
	}
	if (op.sourceRef !== undefined && !isValidJournalPivotSourceRef(op.sourceRef)) {
		return [pivotCacheUnsupportedValueIssue('sourceRef')]
	}
	for (const field of ['refreshOnLoad', 'enableRefresh', 'invalid', 'saveData'] as const) {
		if (op[field] !== undefined && typeof op[field] !== 'boolean') {
			return [pivotCacheUnsupportedValueIssue(field)]
		}
	}
	return []
}

function pivotCacheUnsupportedValueIssue(field: string): MutationJournalIssue {
	return {
		code: 'UNSUPPORTED_VALUE',
		message: `Cannot build exact rollback journal for setPivotCache because ${field} is not supported by pivot cache validation`,
		surface: 'pivot-caches',
		reason: 'value-unsupported',
		refs: ['pivot-cache'],
	}
}

function isValidJournalPivotSourceRef(sourceRef: string): boolean {
	try {
		const body = sourceRef.includes('!') ? sourceRef.slice(sourceRef.indexOf('!') + 1) : sourceRef
		if (body.split(':').length > 2) return false
		parseRange(sourceRef)
		return true
	} catch {
		return false
	}
}

function journalFreezePane(
	workbook: Workbook,
	op: Extract<Operation, { op: 'freezePane' }>,
	opIndex: number,
): DraftJournalEntry {
	const sheet = workbook.getSheet(op.sheet)
	const pane = {
		sheet: op.sheet,
		frozenRows: sheet?.frozenRows ?? 0,
		frozenCols: sheet?.frozenCols ?? 0,
	}
	return {
		opIndex,
		op,
		inverseOps: [{ op: 'freezePane', sheet: op.sheet, row: pane.frozenRows, col: pane.frozenCols }],
		preimages: [{ kind: 'pane', pane }],
		issues: [],
	}
}

function journalSetWorkbookProperties(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setWorkbookProperties' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = { properties: { ...workbook.workbookProperties } }
	const issues = workbookPropertiesValueIssues(op)
	return {
		opIndex,
		op,
		inverseOps:
			issues.length === 0
				? [{ op: 'setWorkbookProperties', properties: preimage.properties, mode: 'replace' }]
				: [],
		preimages: [{ kind: 'workbook-properties', workbookProperties: preimage }],
		issues,
	}
}

function journalSetDocumentProperties(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDocumentProperties' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = {
		properties: clonePlain(workbook.documentProperties) as WorkbookDocumentProperties,
	}
	const issues = documentPropertiesValueIssues(op)
	return {
		opIndex,
		op,
		inverseOps:
			issues.length === 0
				? [{ op: 'setDocumentProperties', properties: preimage.properties, mode: 'replace' }]
				: [],
		preimages: [{ kind: 'document-properties', documentProperties: preimage }],
		issues:
			issues.length > 0
				? issues
				: savedSourcePackageStateIssues(workbook, op.op, ['workbook:documentProperties']),
	}
}

function workbookPropertiesValueIssues(
	op: Extract<Operation, { op: 'setWorkbookProperties' }>,
): MutationJournalIssue[] {
	const issues: MutationJournalIssue[] = []
	if (!isJournalMode(op.mode)) {
		issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'mode', ['workbook:properties']))
	}
	const properties = (op as { readonly properties?: unknown }).properties
	if (!isPlainJournalObject(properties)) {
		issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'properties', ['workbook:properties']))
		return issues
	}
	const codeName = (properties as { readonly codeName?: unknown }).codeName
	if (codeName !== undefined && codeName !== null && typeof codeName !== 'string') {
		issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'codeName', ['workbook:properties']))
	}
	if (typeof codeName === 'string' && codeName.trim() === '') {
		issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'codeName', ['workbook:properties']))
	}
	const defaultThemeVersion = (properties as { readonly defaultThemeVersion?: unknown })
		.defaultThemeVersion
	if (
		defaultThemeVersion !== undefined &&
		defaultThemeVersion !== null &&
		(typeof defaultThemeVersion !== 'number' ||
			!Number.isInteger(defaultThemeVersion) ||
			defaultThemeVersion < 0)
	) {
		issues.push(
			workbookMetadataUnsupportedValueIssue(op.op, 'defaultThemeVersion', ['workbook:properties']),
		)
	}
	for (const [field, value] of [
		['filterPrivacy', (properties as { readonly filterPrivacy?: unknown }).filterPrivacy],
		['date1904', (properties as { readonly date1904?: unknown }).date1904],
	] as const) {
		if (value !== undefined && value !== null && typeof value !== 'boolean') {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, field, ['workbook:properties']))
		}
	}
	return issues
}

function documentPropertiesValueIssues(
	op: Extract<Operation, { op: 'setDocumentProperties' }>,
): MutationJournalIssue[] {
	const issues: MutationJournalIssue[] = []
	if (!isJournalMode(op.mode)) {
		issues.push(
			workbookMetadataUnsupportedValueIssue(op.op, 'mode', ['workbook:documentProperties']),
		)
	}
	const properties = (op as { readonly properties?: unknown }).properties
	if (!isPlainJournalObject(properties)) {
		issues.push(
			workbookMetadataUnsupportedValueIssue(op.op, 'properties', ['workbook:documentProperties']),
		)
		return issues
	}
	const core = (properties as { readonly core?: unknown }).core
	if (core !== undefined && core !== null && !isPlainJournalObject(core)) {
		issues.push(
			workbookMetadataUnsupportedValueIssue(op.op, 'core', ['workbook:documentProperties']),
		)
	} else if (isPlainJournalObject(core)) {
		for (const [key, value] of Object.entries(core)) {
			if (value !== null && value !== undefined && typeof value !== 'string') {
				issues.push(
					workbookMetadataUnsupportedValueIssue(op.op, `core.${key}`, [
						'workbook:documentProperties',
					]),
				)
			}
		}
	}
	const app = (properties as { readonly app?: unknown }).app
	if (app !== undefined && app !== null && !isPlainJournalObject(app)) {
		issues.push(
			workbookMetadataUnsupportedValueIssue(op.op, 'app', ['workbook:documentProperties']),
		)
	} else if (isPlainJournalObject(app)) {
		for (const [key, value] of Object.entries(app)) {
			if (value !== null && !isJournalDocumentPropertyAppValue(value)) {
				issues.push(
					workbookMetadataUnsupportedValueIssue(op.op, `app.${key}`, [
						'workbook:documentProperties',
					]),
				)
			}
		}
	}
	const custom = (properties as { readonly custom?: unknown }).custom
	if (custom !== undefined && custom !== null && !Array.isArray(custom)) {
		issues.push(
			workbookMetadataUnsupportedValueIssue(op.op, 'custom', ['workbook:documentProperties']),
		)
	} else if (Array.isArray(custom)) {
		for (const [index, property] of custom.entries()) {
			if (!isPlainJournalObject(property)) {
				issues.push(
					workbookMetadataUnsupportedValueIssue(op.op, `custom.${index}`, [
						'workbook:documentProperties',
					]),
				)
				continue
			}
			const name = (property as { readonly name?: unknown }).name
			if (typeof name !== 'string' || name.trim() === '') {
				issues.push(
					workbookMetadataUnsupportedValueIssue(op.op, `custom.${index}.name`, [
						'workbook:documentProperties',
					]),
				)
			}
			const value = (property as { readonly value?: unknown }).value
			if (
				typeof value !== 'string' &&
				typeof value !== 'boolean' &&
				!(typeof value === 'number' && Number.isFinite(value))
			) {
				issues.push(
					workbookMetadataUnsupportedValueIssue(op.op, `custom.${index}.value`, [
						'workbook:documentProperties',
					]),
				)
			}
		}
	}
	return issues
}

function isJournalDocumentPropertyAppValue(value: unknown): boolean {
	return (
		typeof value === 'string' ||
		typeof value === 'boolean' ||
		(typeof value === 'number' && Number.isFinite(value)) ||
		isJournalScalarArray(value)
	)
}

function isJournalScalarArray(value: unknown): boolean {
	return (
		Array.isArray(value) &&
		value.every(
			(entry) =>
				typeof entry === 'string' ||
				typeof entry === 'boolean' ||
				(typeof entry === 'number' && Number.isFinite(entry)),
		)
	)
}

function journalSetWorkbookView(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setWorkbookView' }>,
	opIndex: number,
): DraftJournalEntry {
	const index = op.index ?? 0
	const view = workbook.workbookViews[index]
	const preimage = { index, view: view ? { ...view } : null }
	const invalidIndex = !Number.isInteger(index) || index < 0
	const deletesMissingView = op.view === null && view === undefined
	const skipsViewSlot = op.view !== null && index > workbook.workbookViews.length
	const invalidMode = op.view !== null && !isJournalMode(op.mode)
	const invalidViewValue = op.view !== null && hasInvalidWorkbookViewValue(op.view)
	const issues: MutationJournalIssue[] =
		invalidIndex || deletesMissingView || skipsViewSlot || invalidMode || invalidViewValue
			? [
					{
						code: 'UNSUPPORTED_VALUE',
						message: `Cannot build exact rollback journal for setWorkbookView because workbook view index ${index} is invalid`,
						surface: 'workbook-metadata',
						reason: 'value-unsupported',
						refs: [`workbook:view:${index}`],
					},
				]
			: []
	return {
		opIndex,
		op,
		inverseOps:
			issues.length === 0
				? [
						preimage.view
							? { op: 'setWorkbookView', index, view: preimage.view, mode: 'replace' }
							: { op: 'setWorkbookView', index, view: null },
					]
				: [],
		preimages: [{ kind: 'workbook-view', workbookView: preimage }],
		issues,
	}
}

function journalSetCalcSettings(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setCalcSettings' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = {
		settings: clonePlain(workbook.calcSettings) as CalcSettings,
		workbookProperties: { ...workbook.workbookProperties },
	}
	const issues = calcSettingsValueIssues(op)
	return {
		opIndex,
		op,
		inverseOps:
			issues.length === 0
				? [
						{ op: 'setCalcSettings', settings: calcSettingsInverseInput(preimage.settings) },
						{
							op: 'setWorkbookProperties',
							properties: preimage.workbookProperties,
							mode: 'replace',
						},
					]
				: [],
		preimages: [{ kind: 'calc-settings', calcSettings: preimage }],
		issues,
	}
}

function calcSettingsValueIssues(
	op: Extract<Operation, { op: 'setCalcSettings' }>,
): MutationJournalIssue[] {
	const settings = (op as { readonly settings?: unknown }).settings
	if (!isPlainJournalObject(settings)) {
		return [workbookMetadataUnsupportedValueIssue(op.op, 'settings', ['workbook:calcSettings'])]
	}
	const issues: MutationJournalIssue[] = []
	const calcMode = (settings as { readonly calcMode?: unknown }).calcMode
	if (
		calcMode !== undefined &&
		calcMode !== 'auto' &&
		calcMode !== 'manual' &&
		calcMode !== 'autoNoTable'
	) {
		issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'calcMode', ['workbook:calcSettings']))
	}
	for (const [field, value] of [
		['fullCalcOnLoad', (settings as { readonly fullCalcOnLoad?: unknown }).fullCalcOnLoad],
		['calcCompleted', (settings as { readonly calcCompleted?: unknown }).calcCompleted],
		['calcOnSave', (settings as { readonly calcOnSave?: unknown }).calcOnSave],
		['forceFullCalc', (settings as { readonly forceFullCalc?: unknown }).forceFullCalc],
	] as const) {
		if (value !== undefined && value !== null && typeof value !== 'boolean') {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, field, ['workbook:calcSettings']))
		}
	}
	const calcId = (settings as { readonly calcId?: unknown }).calcId
	if (
		calcId !== undefined &&
		calcId !== null &&
		(typeof calcId !== 'number' || !Number.isInteger(calcId) || calcId < 0)
	) {
		issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'calcId', ['workbook:calcSettings']))
	}
	const dateSystem = (settings as { readonly dateSystem?: unknown }).dateSystem
	if (dateSystem !== undefined && dateSystem !== '1900' && dateSystem !== '1904') {
		issues.push(
			workbookMetadataUnsupportedValueIssue(op.op, 'dateSystem', ['workbook:calcSettings']),
		)
	}
	const iterative = (settings as { readonly iterativeCalc?: unknown }).iterativeCalc
	if (iterative !== undefined && iterative !== null && !isPlainJournalObject(iterative)) {
		issues.push(
			workbookMetadataUnsupportedValueIssue(op.op, 'iterativeCalc', ['workbook:calcSettings']),
		)
	} else if (isPlainJournalObject(iterative)) {
		const enabled = (iterative as { readonly enabled?: unknown }).enabled
		if (enabled !== undefined && typeof enabled !== 'boolean') {
			issues.push(
				workbookMetadataUnsupportedValueIssue(op.op, 'iterativeCalc.enabled', [
					'workbook:calcSettings',
				]),
			)
		}
		const maxIterations = (iterative as { readonly maxIterations?: unknown }).maxIterations
		if (
			maxIterations !== undefined &&
			(typeof maxIterations !== 'number' || !Number.isInteger(maxIterations) || maxIterations < 1)
		) {
			issues.push(
				workbookMetadataUnsupportedValueIssue(op.op, 'iterativeCalc.maxIterations', [
					'workbook:calcSettings',
				]),
			)
		}
		const maxChange = (iterative as { readonly maxChange?: unknown }).maxChange
		if (
			maxChange !== undefined &&
			(typeof maxChange !== 'number' || !Number.isFinite(maxChange) || maxChange < 0)
		) {
			issues.push(
				workbookMetadataUnsupportedValueIssue(op.op, 'iterativeCalc.maxChange', [
					'workbook:calcSettings',
				]),
			)
		}
	}
	return issues
}

function calcSettingsInverseInput(
	settings: CalcSettings,
): Extract<Operation, { op: 'setCalcSettings' }>['settings'] {
	return {
		calcMode: settings.calcMode,
		fullCalcOnLoad: settings.fullCalcOnLoad,
		calcCompleted: settings.calcCompleted ?? null,
		calcOnSave: settings.calcOnSave ?? null,
		forceFullCalc: settings.forceFullCalc ?? null,
		calcId: settings.calcId ?? null,
		dateSystem: settings.dateSystem,
		iterativeCalc: { ...settings.iterativeCalc },
	}
}

function journalSetWorkbookProtection(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setWorkbookProtection' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = {
		protection: workbook.workbookProtection ? { ...workbook.workbookProtection } : null,
	}
	const valueIssues = workbookProtectionValueIssues(op)
	const issues: MutationJournalIssue[] =
		valueIssues.length > 0
			? valueIssues
			: preimage.protection
				? []
				: [
						{
							code: 'LOSSY_INVERSE',
							message:
								'Workbook protection absence cannot be restored exactly with public operations',
							surface: 'workbook-metadata',
							reason: 'workbook-protection-absence',
						},
					]
	return {
		opIndex,
		op,
		inverseOps:
			valueIssues.length > 0
				? []
				: [{ op: 'setWorkbookProtection', protection: preimage.protection ?? {} }],
		preimages: [{ kind: 'workbook-protection', workbookProtection: preimage }],
		issues,
	}
}

function workbookProtectionValueIssues(
	op: Extract<Operation, { op: 'setWorkbookProtection' }>,
): MutationJournalIssue[] {
	const protection = op.protection as unknown
	if (!isPlainJournalObject(protection)) {
		return [workbookMetadataUnsupportedValueIssue(op.op, 'protection', ['workbook:protection'])]
	}
	const issues: MutationJournalIssue[] = []
	for (const field of ['lockStructure', 'lockWindows', 'lockRevision'] as const) {
		const value = protection[field]
		if (value !== undefined && typeof value !== 'boolean') {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, field, ['workbook:protection']))
		}
	}
	for (const field of [
		'workbookPassword',
		'revisionsPassword',
		'workbookAlgorithmName',
		'workbookHashValue',
		'workbookSaltValue',
		'revisionsAlgorithmName',
		'revisionsHashValue',
		'revisionsSaltValue',
	] as const) {
		const value = protection[field]
		if (value !== undefined && typeof value !== 'string') {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, field, ['workbook:protection']))
		}
	}
	for (const field of ['workbookSpinCount', 'revisionsSpinCount'] as const) {
		const value = protection[field]
		if (
			value !== undefined &&
			(typeof value !== 'number' || !Number.isInteger(value) || value < 0)
		) {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, field, ['workbook:protection']))
		}
	}
	return issues
}

function journalSetTheme(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTheme' }>,
	opIndex: number,
): DraftJournalEntry {
	const preimage = {
		metadata: { ...workbook.themeMetadata },
		colors: workbook.themeColors.map((color) => ({ ...color })),
	}
	const inverse = themeInverseOperation(op, preimage)
	const valueIssues = themeValueIssues(op)
	const issues: readonly MutationJournalIssue[] =
		valueIssues.length > 0 ? valueIssues : inverse.issues
	return {
		opIndex,
		op,
		inverseOps: issues.length === 0 && inverse.op ? [inverse.op] : [],
		preimages: [{ kind: 'theme', theme: preimage }],
		issues,
	}
}

function themeValueIssues(op: Extract<Operation, { op: 'setTheme' }>): MutationJournalIssue[] {
	const issues: MutationJournalIssue[] = []
	if (
		op.themeName === undefined &&
		op.colorSchemeName === undefined &&
		op.majorFontLatin === undefined &&
		op.minorFontLatin === undefined &&
		op.themeColors === undefined
	) {
		issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'theme', ['workbook:theme']))
	}
	for (const [field, value] of [
		['themeName', op.themeName],
		['colorSchemeName', op.colorSchemeName],
		['majorFontLatin', op.majorFontLatin],
		['minorFontLatin', op.minorFontLatin],
	] as const) {
		if (typeof value === 'string' && value.trim() === '') {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, field, ['workbook:theme']))
		}
	}
	if (op.themeColors === undefined) return issues
	if (op.themeColors.length === 0) {
		issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'themeColors', ['workbook:theme']))
		return issues
	}
	const seen = new Set<string>()
	for (const color of op.themeColors) {
		const refs = [`workbook:themeColor:${color.slot}`]
		if (!JOURNAL_THEME_COLOR_SLOTS.has(color.slot)) {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'themeColors.slot', refs))
		}
		if (seen.has(color.slot)) {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'themeColors.slot', refs))
		}
		seen.add(color.slot)
		if (color.rgb === undefined && color.systemColor === undefined) {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'themeColors.color', refs))
		}
		if (color.rgb !== undefined && !/^[0-9A-Fa-f]{6}$/.test(color.rgb)) {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'themeColors.rgb', refs))
		}
		if (color.lastColor !== undefined && !/^[0-9A-Fa-f]{6}$/.test(color.lastColor)) {
			issues.push(workbookMetadataUnsupportedValueIssue(op.op, 'themeColors.lastColor', refs))
		}
	}
	return issues
}

type SetThemeOperation = Extract<Operation, { op: 'setTheme' }>
type MutableSetThemeOperation = {
	-readonly [K in keyof SetThemeOperation]: SetThemeOperation[K]
}

function themeInverseOperation(
	op: SetThemeOperation,
	preimage: MutationJournalThemePreimage,
): {
	readonly op: SetThemeOperation | null
	readonly issues: readonly MutationJournalIssue[]
} {
	const inverse: Partial<MutableSetThemeOperation> = { op: 'setTheme' }
	const issues: MutationJournalIssue[] = []
	for (const [opKey, metadataKey] of [
		['themeName', 'name'],
		['colorSchemeName', 'colorSchemeName'],
		['majorFontLatin', 'majorFontLatin'],
		['minorFontLatin', 'minorFontLatin'],
	] as const) {
		if (op[opKey] === undefined) continue
		const previous = preimage.metadata[metadataKey]
		if (previous !== undefined) {
			inverse[opKey] = previous
			continue
		}
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Theme metadata field ${opKey} cannot be removed with public operations`,
			surface: 'package-parts',
			reason: 'package-part-preservation',
		})
	}
	if (op.themeColors !== undefined) {
		const colorsBySlot = new Map(preimage.colors.map((color) => [color.slot, color]))
		const inverseColors: WorkbookThemeColor[] = []
		for (const color of op.themeColors) {
			const previous = colorsBySlot.get(color.slot)
			if (!previous) {
				issues.push({
					code: 'LOSSY_INVERSE',
					message: `Theme color slot ${color.slot} cannot be removed with public operations`,
					surface: 'package-parts',
					reason: 'package-part-preservation',
				})
				continue
			}
			if (previous.rgb === undefined && previous.systemColor === undefined) {
				issues.push({
					code: 'LOSSY_INVERSE',
					message: `Theme color slot ${color.slot} cannot be restored with public operations`,
					surface: 'package-parts',
					reason: 'package-part-preservation',
				})
				continue
			}
			inverseColors.push({ ...previous })
		}
		if (inverseColors.length > 0) inverse.themeColors = inverseColors
	}
	const hasInverse =
		inverse.themeName !== undefined ||
		inverse.colorSchemeName !== undefined ||
		inverse.majorFontLatin !== undefined ||
		inverse.minorFontLatin !== undefined ||
		inverse.themeColors !== undefined
	return {
		op: hasInverse ? (inverse as SetThemeOperation) : null,
		issues,
	}
}

function mergePreimage(
	workbook: Workbook,
	sheetName: string,
	rangeText: string,
): MutationJournalMergePreimage {
	const target = parseRange(rangeText)
	const existed =
		workbook.getSheet(sheetName)?.merges.some((merge) => sameRange(merge, target)) ?? false
	return {
		sheet: sheetName,
		range: toRangeString(target),
		existed,
	}
}

function dataValidationPreimage(
	workbook: Workbook,
	sheetName: string,
	range: string,
): MutationJournalDataValidationPreimage {
	const validation = workbook.getSheet(sheetName)?.dataValidations.find((dv) => dv.sqref === range)
	return {
		sheet: sheetName,
		range,
		validation: validation ? { ...validation } : null,
	}
}

function dataValidationRestoreOrderIssues(
	workbook: Workbook,
	sheetName: string,
	validations: readonly MutationJournalDataValidationPreimage[],
): readonly MutationJournalIssue[] {
	if (validations.length === 0) return []
	const sheet = workbook.getSheet(sheetName)
	if (!sheet) return []
	const ranges = new Set(validations.map((validation) => validation.range))
	const duplicateRanges = [...ranges].filter(
		(range) => sheet.dataValidations.filter((validation) => validation.sqref === range).length > 1,
	)
	if (duplicateRanges.length > 0) {
		return [
			{
				code: 'LOSSY_INVERSE',
				message: `Duplicate data validation metadata on ${sheetName} cannot be restored exactly with public operations`,
				surface: 'data-validations',
				reason: 'metadata-duplicate',
				refs: duplicateRanges.map((range) => `${sheetName}!${range}`),
			},
		]
	}
	const indexes = sheet.dataValidations
		.map((validation, index) => (ranges.has(validation.sqref) ? index : -1))
		.filter((index) => index >= 0)
	if (indexes.length !== validations.length) return []
	const suffixStart = sheet.dataValidations.length - indexes.length
	const restoresSuffix = indexes.every((index, offset) => index === suffixStart + offset)
	if (restoresSuffix) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Data-validation order on ${sheetName} cannot be restored exactly with public operations`,
			surface: 'data-validations',
			reason: 'metadata-order',
			refs: validations.map((validation) => `${sheetName}!${validation.range}`),
		},
	]
}

function restoreDataValidationOps(validation: MutationJournalDataValidationPreimage): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	if (!validation.validation) {
		return {
			inverseOps: [
				{ op: 'deleteDataValidation', sheet: validation.sheet, range: validation.range },
			],
			issues: [],
		}
	}
	const { rule, issues } = dataValidationRuleFromSheet(validation)
	return {
		inverseOps: rule
			? [{ op: 'setDataValidation', sheet: validation.sheet, range: validation.range, rule }]
			: [],
		issues,
	}
}

interface DataValidationTransfer {
	readonly sourceIndex: number
	readonly source: MutationJournalDataValidationPreimage
	readonly target: MutationJournalDataValidationPreimage
	readonly retainedSourceRange: string | null
	readonly targetCollision: boolean
	readonly retainedCollision: boolean
}

interface ConditionalFormatTransfer {
	readonly sourceIndex: number
	readonly source: SheetConditionalFormat
	readonly target: MutationJournalConditionalFormatPreimage
	readonly retainedSourceRange: string | null
	readonly targetCollision: boolean
	readonly retainedCollision: boolean
}

function validationTransferRestoration(
	workbook: Workbook,
	sourceSheetName: string,
	sourceRange: RangeRef,
	targetSheetName: string,
	targetRange: RangeRef,
	move: boolean,
): {
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const sourceSheet = workbook.getSheet(sourceSheetName)
	const targetSheet = workbook.getSheet(targetSheetName)
	if (!sourceSheet || !targetSheet) {
		return { inverseOps: [], preimages: [], issues: [] }
	}
	const rowDelta = targetRange.start.row - sourceRange.start.row
	const colDelta = targetRange.start.col - sourceRange.start.col
	const transfers = dataValidationTransfers(
		workbook,
		sourceSheet,
		sourceRange,
		targetSheet,
		rowDelta,
		colDelta,
		move,
	)
	const preimageMap = new Map<string, MutationJournalDataValidationPreimage>()
	for (const transfer of transfers) {
		preimageMap.set(`${transfer.target.sheet}!${transfer.target.range}`, transfer.target)
		if (move) {
			preimageMap.set(`${transfer.source.sheet}!${transfer.source.range}`, transfer.source)
		}
	}

	const issues: MutationJournalIssue[] = [
		...dataValidationTransferCollisionIssues(transfers, move),
		...(move
			? dataValidationTransferDuplicateIssues(sourceSheet.name, sourceRange, transfers)
			: []),
		...x14DataValidationTransferIssues(sourceSheet, sourceRange),
	]
	const inverseOps: Operation[] = []
	for (const transfer of transfers) {
		if (!transfer.targetCollision) {
			inverseOps.push({
				op: 'deleteDataValidation',
				sheet: transfer.target.sheet,
				range: transfer.target.range,
			})
		}
	}
	if (move) {
		for (const transfer of transfers) {
			if (transfer.retainedSourceRange && !transfer.retainedCollision) {
				inverseOps.push({
					op: 'deleteDataValidation',
					sheet: transfer.source.sheet,
					range: transfer.retainedSourceRange,
				})
			}
		}
		for (const transfer of transfers) {
			const restored = restoreDataValidationOps(transfer.source)
			inverseOps.push(...restored.inverseOps)
			issues.push(...restored.issues)
		}
		issues.push(...dataValidationMoveOrderIssues(sourceSheet, sourceRange, transfers))
	}

	return {
		inverseOps: dedupeDataValidationInverseOps(inverseOps),
		preimages: [
			{
				kind: 'data-validations',
				validations: [...preimageMap.values()],
			},
		],
		issues,
	}
}

function dataValidationTransfers(
	workbook: Workbook,
	sourceSheet: Sheet,
	sourceRange: RangeRef,
	targetSheet: Sheet,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): DataValidationTransfer[] {
	const transfers: DataValidationTransfer[] = []
	for (let index = 0; index < sourceSheet.dataValidations.length; index++) {
		const validation = sourceSheet.dataValidations[index]
		if (!validation) continue
		const ranges = parseSqrefRanges(validation.sqref)
		if (ranges.length === 0) continue
		const copiedRanges = ranges.filter((range) => rangeContainsRange(sourceRange, range))
		if (copiedRanges.length === 0) continue
		const targetSqref = rangesToSqref(
			copiedRanges.map((range) => shiftRange(range, rowDelta, colDelta)),
		)
		const keptRanges = ranges.filter((range) => !rangeContainsRange(sourceRange, range))
		const retainedSourceRange = move && keptRanges.length > 0 ? rangesToSqref(keptRanges) : null
		const source: MutationJournalDataValidationPreimage = {
			sheet: sourceSheet.name,
			range: validation.sqref,
			validation: { ...validation },
		}
		const target = dataValidationPreimage(workbook, targetSheet.name, targetSqref)
		transfers.push({
			sourceIndex: index,
			source,
			target,
			retainedSourceRange,
			targetCollision: dataValidationTargetCollision(
				sourceSheet,
				targetSheet,
				validation,
				targetSqref,
				move,
			),
			retainedCollision: retainedSourceRange
				? dataValidationSqrefCount(sourceSheet, retainedSourceRange) > 0 ||
					transfers.some((transfer) => transfer.retainedSourceRange === retainedSourceRange)
				: false,
		})
	}
	return transfers
}

function dataValidationTargetCollision(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	validation: SheetDataValidation,
	targetSqref: string,
	move: boolean,
): boolean {
	const count = dataValidationSqrefCount(targetSheet, targetSqref)
	if (count === 0) return false
	if (!move || sourceSheet !== targetSheet || validation.sqref !== targetSqref) return true
	return count > 1
}

function dataValidationSqrefCount(sheet: Sheet, sqref: string): number {
	return sheet.dataValidations.filter((validation) => validation.sqref === sqref).length
}

function dataValidationTransferCollisionIssues(
	transfers: readonly DataValidationTransfer[],
	move: boolean,
): readonly MutationJournalIssue[] {
	const issues: MutationJournalIssue[] = []
	for (const transfer of transfers) {
		if (transfer.targetCollision) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Copied data validation at ${transfer.target.sheet}!${transfer.target.range} collides with existing validation metadata`,
				surface: 'data-validations',
				reason: 'metadata-collision',
				refs: [
					`${transfer.source.sheet}!${transfer.source.range}`,
					`${transfer.target.sheet}!${transfer.target.range}`,
				],
			})
		}
		if (move && transfer.retainedCollision && transfer.retainedSourceRange) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Retained data validation at ${transfer.source.sheet}!${transfer.retainedSourceRange} cannot be removed without touching unrelated validation metadata`,
				surface: 'data-validations',
				reason: 'metadata-collision',
				refs: [
					`${transfer.source.sheet}!${transfer.source.range}`,
					`${transfer.source.sheet}!${transfer.retainedSourceRange}`,
				],
			})
		}
	}
	return issues
}

function dataValidationTransferDuplicateIssues(
	sourceSheetName: string,
	sourceRange: RangeRef,
	transfers: readonly DataValidationTransfer[],
): readonly MutationJournalIssue[] {
	const duplicates = duplicateStrings(transfers.map((transfer) => transfer.source.range))
	if (duplicates.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Moved duplicate data validations on ${sourceSheetName}!${rangeToA1(sourceRange)} cannot be restored exactly with public operations`,
			surface: 'data-validations',
			reason: 'metadata-duplicate',
			refs: duplicates.map((range) => `${sourceSheetName}!${range}`),
		},
	]
}

function x14DataValidationTransferIssues(
	sourceSheet: Sheet,
	sourceRange: RangeRef,
): readonly MutationJournalIssue[] {
	const refs: string[] = []
	for (const validation of sourceSheet.x14DataValidations) {
		if (validation.deleted) continue
		const ranges = parseSqrefRanges(validation.sqref)
		if (ranges.some((range) => rangeContainsRange(sourceRange, range))) {
			refs.push(`${sourceSheet.name}!x14Validation:${validation.sqref}:${validation.index}`)
		}
	}
	if (refs.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Transferred x14 data validation metadata on ${sourceSheet.name}!${rangeToA1(sourceRange)} cannot be restored with public operations`,
			surface: 'x14-metadata',
			reason: 'x14-metadata',
			refs,
		},
	]
}

function dataValidationMoveOrderIssues(
	sourceSheet: Sheet,
	sourceRange: RangeRef,
	transfers: readonly DataValidationTransfer[],
): readonly MutationJournalIssue[] {
	if (transfers.length === 0) return []
	const moved = new Set(transfers.map((transfer) => transfer.sourceIndex))
	const firstMoved = Math.min(...moved)
	for (let index = firstMoved; index < sourceSheet.dataValidations.length; index++) {
		if (!moved.has(index)) {
			return [
				{
					code: 'LOSSY_INVERSE',
					message: `Moved data validation order on ${sourceSheet.name}!${rangeToA1(sourceRange)} cannot be restored exactly with public operations`,
					surface: 'data-validations',
					reason: 'metadata-order',
					refs: [`${sourceSheet.name}!${rangeToA1(sourceRange)}`],
				},
			]
		}
	}
	return []
}

function dedupeDataValidationInverseOps(ops: readonly Operation[]): Operation[] {
	const seen = new Set<string>()
	const deduped: Operation[] = []
	for (const op of ops) {
		const key =
			op.op === 'deleteDataValidation' ? `${op.op}:${op.sheet}:${op.range}` : JSON.stringify(op)
		if (seen.has(key)) continue
		seen.add(key)
		deduped.push(op)
	}
	return deduped
}

function conditionalFormatTransferRestoration(
	workbook: Workbook,
	sourceSheetName: string,
	sourceRange: RangeRef,
	targetSheetName: string,
	targetRange: RangeRef,
	move: boolean,
): {
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const sourceSheet = workbook.getSheet(sourceSheetName)
	const targetSheet = workbook.getSheet(targetSheetName)
	if (!sourceSheet || !targetSheet) {
		return { inverseOps: [], preimages: [], issues: [] }
	}
	const rowDelta = targetRange.start.row - sourceRange.start.row
	const colDelta = targetRange.start.col - sourceRange.start.col
	const transfers = conditionalFormatTransfers(
		workbook,
		sourceSheet,
		sourceRange,
		targetSheet,
		rowDelta,
		colDelta,
		move,
	)
	const issues: MutationJournalIssue[] = [
		...conditionalFormatTransferCollisionIssues(sourceSheet.name, transfers, move),
		...(move
			? conditionalFormatTransferDuplicateIssues(sourceSheet.name, sourceRange, transfers)
			: []),
		...x14ConditionalFormatTransferIssues(sourceSheet, sourceRange),
	]
	const inverseOps: Operation[] = []
	for (const transfer of transfers) {
		const targetRange = transfer.target.range
		if (!transfer.targetCollision && targetRange !== undefined) {
			inverseOps.push({
				op: 'deleteConditionalFormat',
				sheet: transfer.target.sheet,
				range: targetRange,
			})
		}
	}
	if (move) {
		for (const transfer of transfers) {
			if (transfer.retainedSourceRange && !transfer.retainedCollision) {
				inverseOps.push({
					op: 'deleteConditionalFormat',
					sheet: sourceSheet.name,
					range: transfer.retainedSourceRange,
				})
			}
		}
		const restored = restoreConditionalFormatOps(
			sourceSheet.name,
			transfers.map((transfer) => transfer.source),
		)
		inverseOps.push(...restored.inverseOps)
		issues.push(...restored.issues)
		issues.push(...conditionalFormatMoveOrderIssues(sourceSheet, sourceRange, transfers))
	}

	return {
		inverseOps: dedupeConditionalFormatInverseOps(inverseOps),
		preimages: conditionalFormatTransferPreimages(sourceSheet.name, transfers, move),
		issues,
	}
}

function conditionalFormatTransfers(
	workbook: Workbook,
	sourceSheet: Sheet,
	sourceRange: RangeRef,
	targetSheet: Sheet,
	rowDelta: number,
	colDelta: number,
	move: boolean,
): ConditionalFormatTransfer[] {
	const transfers: ConditionalFormatTransfer[] = []
	for (let index = 0; index < sourceSheet.conditionalFormats.length; index++) {
		const format = sourceSheet.conditionalFormats[index]
		if (!format) continue
		const ranges = parseSqrefRanges(format.sqref)
		if (ranges.length === 0) continue
		const copiedRanges = ranges.filter((range) => rangeContainsRange(sourceRange, range))
		if (copiedRanges.length === 0) continue
		const targetSqref = rangesToSqref(
			copiedRanges.map((range) => shiftRange(range, rowDelta, colDelta)),
		)
		const keptRanges = ranges.filter((range) => !rangeContainsRange(sourceRange, range))
		const retainedSourceRange = move && keptRanges.length > 0 ? rangesToSqref(keptRanges) : null
		transfers.push({
			sourceIndex: index,
			source: cloneConditionalFormat(format),
			target: conditionalFormatPreimage(workbook, targetSheet.name, targetSqref),
			retainedSourceRange,
			targetCollision: conditionalFormatTargetCollision(
				sourceSheet,
				targetSheet,
				format,
				targetSqref,
				move,
			),
			retainedCollision: retainedSourceRange
				? conditionalFormatSqrefCount(sourceSheet, retainedSourceRange) > 0 ||
					transfers.some((transfer) => transfer.retainedSourceRange === retainedSourceRange)
				: false,
		})
	}
	return transfers
}

function conditionalFormatTargetCollision(
	sourceSheet: Sheet,
	targetSheet: Sheet,
	format: SheetConditionalFormat,
	targetSqref: string,
	move: boolean,
): boolean {
	const count = conditionalFormatSqrefCount(targetSheet, targetSqref)
	if (count === 0) return false
	if (!move || sourceSheet !== targetSheet || format.sqref !== targetSqref) return true
	return count > 1
}

function conditionalFormatSqrefCount(sheet: Sheet, sqref: string): number {
	return sheet.conditionalFormats.filter((format) => format.sqref === sqref).length
}

function conditionalFormatTransferCollisionIssues(
	sourceSheetName: string,
	transfers: readonly ConditionalFormatTransfer[],
	move: boolean,
): readonly MutationJournalIssue[] {
	const issues: MutationJournalIssue[] = []
	for (const transfer of transfers) {
		if (transfer.targetCollision) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Copied conditional format at ${transfer.target.sheet}!${transfer.target.range} collides with existing conditional format metadata`,
				surface: 'conditional-formats',
				reason: 'metadata-collision',
				refs: [
					`${sourceSheetName}!${transfer.source.sqref}`,
					`${transfer.target.sheet}!${transfer.target.range}`,
				],
			})
		}
		if (move && transfer.retainedCollision && transfer.retainedSourceRange) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Retained conditional format at ${sourceSheetName}!${transfer.retainedSourceRange} cannot be removed without touching unrelated conditional format metadata`,
				surface: 'conditional-formats',
				reason: 'metadata-collision',
				refs: [
					`${sourceSheetName}!${transfer.source.sqref}`,
					`${sourceSheetName}!${transfer.retainedSourceRange}`,
				],
			})
		}
	}
	return issues
}

function conditionalFormatTransferDuplicateIssues(
	sourceSheetName: string,
	sourceRange: RangeRef,
	transfers: readonly ConditionalFormatTransfer[],
): readonly MutationJournalIssue[] {
	const duplicates = duplicateStrings(transfers.map((transfer) => transfer.source.sqref))
	if (duplicates.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Moved duplicate conditional formats on ${sourceSheetName}!${rangeToA1(sourceRange)} cannot be restored exactly with public operations`,
			surface: 'conditional-formats',
			reason: 'metadata-duplicate',
			refs: duplicates.map((range) => `${sourceSheetName}!${range}`),
		},
	]
}

function x14ConditionalFormatTransferIssues(
	sourceSheet: Sheet,
	sourceRange: RangeRef,
): readonly MutationJournalIssue[] {
	const refs: string[] = []
	for (const format of sourceSheet.x14ConditionalFormats) {
		if (format.deleted) continue
		const ranges = parseSqrefRanges(format.sqref)
		if (ranges.some((range) => rangeContainsRange(sourceRange, range))) {
			refs.push(`${sourceSheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}`)
		}
	}
	if (refs.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Transferred x14 conditional format metadata on ${sourceSheet.name}!${rangeToA1(sourceRange)} cannot be restored with public operations`,
			surface: 'x14-metadata',
			reason: 'x14-metadata',
			refs,
		},
	]
}

function conditionalFormatMoveOrderIssues(
	sourceSheet: Sheet,
	sourceRange: RangeRef,
	transfers: readonly ConditionalFormatTransfer[],
): readonly MutationJournalIssue[] {
	if (transfers.length === 0) return []
	const moved = new Set(transfers.map((transfer) => transfer.sourceIndex))
	const firstMoved = Math.min(...moved)
	for (let index = firstMoved; index < sourceSheet.conditionalFormats.length; index++) {
		if (!moved.has(index)) {
			return [
				{
					code: 'LOSSY_INVERSE',
					message: `Moved conditional format order on ${sourceSheet.name}!${rangeToA1(sourceRange)} cannot be restored exactly with public operations`,
					surface: 'conditional-formats',
					reason: 'metadata-order',
					refs: [`${sourceSheet.name}!${rangeToA1(sourceRange)}`],
				},
			]
		}
	}
	return []
}

function conditionalFormatTransferPreimages(
	sourceSheetName: string,
	transfers: readonly ConditionalFormatTransfer[],
	move: boolean,
): MutationJournalPreimage[] {
	if (transfers.length === 0) return []
	const preimages = new Map<string, MutationJournalConditionalFormatPreimage>()
	for (const transfer of transfers) {
		preimages.set(`${transfer.target.sheet}!${transfer.target.range}`, transfer.target)
	}
	if (move) {
		preimages.set(`${sourceSheetName}!<source>`, {
			sheet: sourceSheetName,
			formats: transfers.map((transfer) => cloneConditionalFormat(transfer.source)),
		})
	}
	return [...preimages.values()].map((conditionalFormats) => ({
		kind: 'conditional-formats' as const,
		conditionalFormats,
	}))
}

function dedupeConditionalFormatInverseOps(ops: readonly Operation[]): Operation[] {
	const seen = new Set<string>()
	const deduped: Operation[] = []
	for (const op of ops) {
		const key =
			op.op === 'deleteConditionalFormat'
				? `${op.op}:${op.sheet}:${op.range ?? ''}:${op.priority ?? ''}:${op.ruleIndex ?? ''}`
				: JSON.stringify(op)
		if (seen.has(key)) continue
		seen.add(key)
		deduped.push(op)
	}
	return deduped
}

function dataValidationRuleFromSheet(validation: MutationJournalDataValidationPreimage): {
	readonly rule: DataValidationRule | null
	readonly issues: readonly MutationJournalIssue[]
} {
	const source = validation.validation
	if (!source?.type || !isDataValidationType(source.type)) {
		return {
			rule: null,
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore data validation at ${validation.sheet}!${validation.range} with unsupported type ${source?.type ?? '<missing>'}`,
					surface: 'data-validations',
					reason: 'value-unsupported',
					refs: [`${validation.sheet}!${validation.range}`],
				},
			],
		}
	}
	const issues: MutationJournalIssue[] = []
	if (source.source === 'x14' || source.uid !== undefined) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Data validation extension metadata at ${validation.sheet}!${validation.range} cannot be restored with public operations`,
			surface: 'x14-metadata',
			reason: 'x14-metadata',
			refs: [`${validation.sheet}!${validation.range}`],
		})
	}
	const materializedDefaultFields = [
		source.allowBlank === undefined ? 'allowBlank' : null,
		source.showErrorMessage === undefined ? 'showErrorMessage' : null,
	].filter((field): field is string => field !== null)
	if (materializedDefaultFields.length > 0 && source.source !== 'x14' && source.uid === undefined) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Data validation default attributes ${materializedDefaultFields.join(', ')} at ${validation.sheet}!${validation.range} cannot be restored exactly with public operations`,
			surface: 'data-validations',
			reason: 'data-validation-default-attributes',
			refs: [`${validation.sheet}!${validation.range}`],
		})
	}
	return {
		rule: {
			type: source.type,
			...(isDataValidationOperator(source.operator) ? { operator: source.operator } : {}),
			...(source.formula1 !== undefined ? { formula1: source.formula1 } : {}),
			...(source.formula2 !== undefined ? { formula2: source.formula2 } : {}),
			...(source.allowBlank !== undefined ? { allowBlank: source.allowBlank } : {}),
			...(source.showDropDown !== undefined ? { showDropDown: source.showDropDown } : {}),
			...(source.showErrorMessage !== undefined
				? { showErrorMessage: source.showErrorMessage }
				: {}),
			...(source.errorTitle !== undefined ? { errorTitle: source.errorTitle } : {}),
			...(source.error !== undefined ? { errorMessage: source.error } : {}),
			...(source.errorStyle !== undefined ? { errorStyle: source.errorStyle } : {}),
			...(source.imeMode !== undefined ? { imeMode: source.imeMode } : {}),
			...(source.showInputMessage !== undefined
				? { showInputMessage: source.showInputMessage }
				: {}),
			...(source.promptTitle !== undefined ? { promptTitle: source.promptTitle } : {}),
			...(source.prompt !== undefined ? { prompt: source.prompt } : {}),
		},
		issues,
	}
}

function autoFilterPreimage(
	workbook: Workbook,
	sheetName: string,
): MutationJournalAutoFilterPreimage {
	const autoFilter = workbook.getSheet(sheetName)?.autoFilter
	return {
		sheet: sheetName,
		autoFilter: autoFilter ? cloneAutoFilter(autoFilter) : null,
	}
}

function restoreAutoFilterOps(preimage: MutationJournalAutoFilterPreimage): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const autoFilter = preimage.autoFilter
	if (!autoFilter) {
		return {
			inverseOps: [{ op: 'clearAutoFilter', sheet: preimage.sheet }],
			issues: [],
		}
	}
	const inverseOps: Operation[] = [
		{ op: 'setAutoFilter', sheet: preimage.sheet, range: autoFilter.ref },
	]
	const issues: MutationJournalIssue[] = []
	for (const column of autoFilter.columns) {
		if (
			column.kind === 'filters' &&
			column.values !== undefined &&
			column.blank === undefined &&
			column.calendarType === undefined &&
			column.dateGroupItems === undefined &&
			column.hiddenButton === undefined &&
			column.showButton === undefined
		) {
			inverseOps.push({
				op: 'setAutoFilter',
				sheet: preimage.sheet,
				range: autoFilter.ref,
				column: column.colId,
				values: [...column.values],
			})
			continue
		}
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `AutoFilter column ${column.colId} on ${preimage.sheet}!${autoFilter.ref} cannot be fully restored with public operations`,
			surface: 'auto-filters',
			reason: 'auto-filter-column-metadata',
			refs: [`${preimage.sheet}!${autoFilter.ref}`],
		})
	}
	if (autoFilter.uid !== undefined) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `AutoFilter extension metadata on ${preimage.sheet}!${autoFilter.ref} cannot be restored with public operations`,
			surface: 'auto-filters',
			reason: 'auto-filter-extension-metadata',
			refs: [`${preimage.sheet}!${autoFilter.ref}`],
		})
	}
	if (autoFilter.sortState) {
		const [firstCondition, ...extraConditions] = autoFilter.sortState.conditions
		if (
			autoFilter.sortState.caseSensitive !== undefined ||
			autoFilter.sortState.columnSort !== undefined ||
			autoFilter.sortState.sortMethod !== undefined ||
			autoFilter.sortState.preservedAttributes !== undefined ||
			extraConditions.length > 0 ||
			firstCondition?.customList !== undefined ||
			firstCondition?.dxfId !== undefined ||
			firstCondition?.iconSet !== undefined ||
			firstCondition?.iconId !== undefined
		) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `AutoFilter sort metadata on ${preimage.sheet}!${autoFilter.ref} cannot be fully restored with public operations`,
				surface: 'auto-filters',
				reason: 'auto-filter-sort-metadata',
				refs: [`${preimage.sheet}!${autoFilter.ref}`],
			})
		}
		if (firstCondition) {
			inverseOps.push({
				op: 'setAutoFilter',
				sheet: preimage.sheet,
				range: autoFilter.ref,
				sortRef: autoFilter.sortState.ref,
				sortBy: firstCondition.ref,
				...(firstCondition.descending !== undefined
					? { descending: firstCondition.descending }
					: {}),
			})
		}
	}
	return { inverseOps, issues }
}

function conditionalFormatPreimage(
	workbook: Workbook,
	sheetName: string,
	range: string | undefined,
): MutationJournalConditionalFormatPreimage {
	const formats = workbook
		.getSheet(sheetName)
		?.conditionalFormats.filter((cf) => range === undefined || cf.sqref === range)
	return {
		sheet: sheetName,
		...(range !== undefined ? { range } : {}),
		formats: (formats ?? []).map(cloneConditionalFormat),
	}
}

function restoreConditionalFormatOps(
	sheet: string,
	formats: readonly SheetConditionalFormat[],
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const inverseOps: Operation[] = []
	const issues: MutationJournalIssue[] = []
	for (const format of formats) {
		inverseOps.push({ op: 'deleteConditionalFormat', sheet, range: format.sqref })
		format.rules.forEach((rule, index) => {
			const converted = conditionalFormatRuleFromSheet(sheet, format.sqref, rule)
			issues.push(...converted.issues)
			if (!converted.rule) return
			inverseOps.push({
				op: 'setConditionalFormat',
				sheet,
				range: format.sqref,
				rule: converted.rule,
				mode: index === 0 ? 'replace' : 'append',
			})
		})
	}
	return { inverseOps, issues }
}

function conditionalFormatRestoreOrderIssues(
	workbook: Workbook,
	sheetName: string,
	formats: readonly SheetConditionalFormat[],
): readonly MutationJournalIssue[] {
	if (formats.length === 0) return []
	const sheet = workbook.getSheet(sheetName)
	if (!sheet) return []
	const duplicateRanges = duplicateStrings(formats.map((format) => format.sqref))
	if (duplicateRanges.length > 0) {
		return [
			{
				code: 'LOSSY_INVERSE',
				message: `Duplicate conditional format metadata on ${sheetName} cannot be restored exactly with public operations`,
				surface: 'conditional-formats',
				reason: 'metadata-duplicate',
				refs: duplicateRanges.map((range) => `${sheetName}!${range}`),
			},
		]
	}
	const ranges = new Set(formats.map((format) => format.sqref))
	const indexes = sheet.conditionalFormats
		.map((format, index) => (ranges.has(format.sqref) ? index : -1))
		.filter((index) => index >= 0)
	if (indexes.length !== formats.length) return []
	const suffixStart = sheet.conditionalFormats.length - indexes.length
	const restoresSuffix = indexes.every((index, offset) => index === suffixStart + offset)
	if (restoresSuffix) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Conditional-format order on ${sheetName} cannot be restored exactly with public operations`,
			surface: 'conditional-formats',
			reason: 'metadata-order',
			refs: formats.map((format) => `${sheetName}!${format.sqref}`),
		},
	]
}

function conditionalFormatRuleFromSheet(
	sheet: string,
	range: string,
	rule: SheetConditionalFormatRule,
): {
	readonly rule: ConditionalFormatRule | null
	readonly issues: readonly MutationJournalIssue[]
} {
	const issues: MutationJournalIssue[] = []
	if (!isConditionalFormatType(rule.type)) {
		return {
			rule: null,
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore conditional format at ${sheet}!${range} with unsupported type ${rule.type}`,
					surface: 'conditional-formats',
					reason: 'value-unsupported',
					refs: [`${sheet}!${range}`],
				},
			],
		}
	}
	if (
		rule.dxfId !== undefined ||
		rule.rank !== undefined ||
		rule.percent !== undefined ||
		rule.bottom !== undefined ||
		rule.aboveAverage !== undefined ||
		rule.equalAverage !== undefined ||
		rule.stdDev !== undefined ||
		rule.text !== undefined ||
		rule.timePeriod !== undefined ||
		rule.preservedRuleAttributes !== undefined ||
		rule.preservedRuleChildXml !== undefined
	) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Conditional format metadata at ${sheet}!${range} cannot be fully restored with public operations`,
			surface: 'conditional-formats',
			reason: 'package-part-preservation',
			refs: [`${sheet}!${range}`],
		})
	}
	return {
		rule: {
			type: rule.type,
			...(isConditionalFormatOperator(rule.operator) ? { operator: rule.operator } : {}),
			...(rule.formulas[0] !== undefined ? { formula: rule.formulas[0] } : {}),
			...(rule.formulas[1] !== undefined ? { formula2: rule.formulas[1] } : {}),
			...(rule.priority !== undefined ? { priority: rule.priority } : {}),
			...(rule.stopIfTrue !== undefined ? { stopIfTrue: rule.stopIfTrue } : {}),
			...(rule.style ? { style: cloneCellStyle(rule.style) as StyleInput } : {}),
			...(rule.colorScale ? { colorScale: clonePlain(rule.colorScale) } : {}),
			...(rule.dataBar ? { dataBar: clonePlain(rule.dataBar) } : {}),
			...(rule.iconSet ? { iconSet: clonePlain(rule.iconSet) } : {}),
		},
		issues,
	}
}

function definedNamePreimage(
	workbook: Workbook,
	name: string,
	scopeSheetName: string | undefined,
): MutationJournalDefinedNamePreimage {
	const scope = scopeSheetName
		? sheetScopeForName(workbook, scopeSheetName)
		: ({ kind: 'workbook' } as const)
	const definedName = scope ? workbook.definedNames.getEntry(name, scope) : undefined
	return {
		name,
		...(scopeSheetName !== undefined ? { scope: scopeSheetName } : {}),
		definedName: definedName ? cloneDefinedName(definedName) : null,
	}
}

function pageSetupPreimage(
	workbook: Workbook,
	sheetName: string,
): MutationJournalPageSetupPreimage {
	const sheet = workbook.getSheet(sheetName)
	return {
		sheet: sheetName,
		pageSetup: sheet?.pageSetup ? clonePageSetup(sheet.pageSetup) : null,
		pageMargins: sheet?.pageMargins ? { ...sheet.pageMargins } : null,
	}
}

function restorePageSetupOps(
	preimage: MutationJournalPageSetupPreimage,
	touched: Extract<Operation, { op: 'setPageSetup' }>['setup'],
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const issues: MutationJournalIssue[] = []
	if (!preimage.pageSetup) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Page setup for ${preimage.sheet} cannot be removed with public operations`,
			surface: 'page-setup',
			reason: 'page-setup-unsettable',
			refs: [preimage.sheet],
		})
	}
	const setup = preimage.pageSetup ? touchedPageSetupToInput(preimage, touched, issues) : null
	const margins = touched.margins
		? touchedPageMarginsToInput(preimage, touched.margins, issues)
		: undefined
	if (touched.margins && !preimage.pageMargins) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Page margins for ${preimage.sheet} cannot be removed with public operations`,
			surface: 'page-setup',
			reason: 'page-margins-unsettable',
			refs: [preimage.sheet],
		})
	}
	if (!setup && !margins) return { inverseOps: [], issues }
	return {
		inverseOps: [
			{
				op: 'setPageSetup',
				sheet: preimage.sheet,
				setup: {
					...(setup ?? {}),
					...(margins ? { margins } : {}),
				},
			},
		],
		issues,
	}
}

function touchedPageSetupToInput(
	preimage: MutationJournalPageSetupPreimage,
	touched: Extract<Operation, { op: 'setPageSetup' }>['setup'],
	issues: MutationJournalIssue[],
): Omit<Extract<Operation, { op: 'setPageSetup' }>['setup'], 'margins'> | null {
	const setup = preimage.pageSetup
	if (!setup) return null
	const input: {
		orientation?: 'portrait' | 'landscape'
		paperSize?: number
		scale?: number
		fitToWidth?: number
		fitToHeight?: number
	} = {}
	if (touched.orientation !== undefined) {
		if (setup.orientation === 'portrait' || setup.orientation === 'landscape') {
			input.orientation = setup.orientation
		} else {
			addPageSetupLossyIssue(preimage.sheet, issues)
		}
	}
	if (touched.paperSize !== undefined) {
		if (setup.paperSize !== undefined) input.paperSize = setup.paperSize
		else addPageSetupLossyIssue(preimage.sheet, issues)
	}
	if (touched.scale !== undefined) {
		if (setup.scale !== undefined) input.scale = setup.scale
		else addPageSetupLossyIssue(preimage.sheet, issues)
	}
	if (touched.fitToWidth !== undefined) {
		if (setup.fitToWidth !== undefined) input.fitToWidth = setup.fitToWidth
		else addPageSetupLossyIssue(preimage.sheet, issues)
	}
	if (touched.fitToHeight !== undefined) {
		if (setup.fitToHeight !== undefined) input.fitToHeight = setup.fitToHeight
		else addPageSetupLossyIssue(preimage.sheet, issues)
	}
	return Object.keys(input).length > 0 ? input : null
}

function touchedPageMarginsToInput(
	preimage: MutationJournalPageSetupPreimage,
	touched: NonNullable<Extract<Operation, { op: 'setPageSetup' }>['setup']['margins']>,
	issues: MutationJournalIssue[],
): NonNullable<Extract<Operation, { op: 'setPageSetup' }>['setup']['margins']> | null {
	const margins = preimage.pageMargins
	if (!margins) return null
	const input: {
		left?: number
		right?: number
		top?: number
		bottom?: number
		header?: number
		footer?: number
	} = {}
	if (touched.left !== undefined) {
		if (margins.left !== undefined) input.left = margins.left
		else addPageMarginsLossyIssue(preimage.sheet, issues)
	}
	if (touched.right !== undefined) {
		if (margins.right !== undefined) input.right = margins.right
		else addPageMarginsLossyIssue(preimage.sheet, issues)
	}
	if (touched.top !== undefined) {
		if (margins.top !== undefined) input.top = margins.top
		else addPageMarginsLossyIssue(preimage.sheet, issues)
	}
	if (touched.bottom !== undefined) {
		if (margins.bottom !== undefined) input.bottom = margins.bottom
		else addPageMarginsLossyIssue(preimage.sheet, issues)
	}
	if (touched.header !== undefined) {
		if (margins.header !== undefined) input.header = margins.header
		else addPageMarginsLossyIssue(preimage.sheet, issues)
	}
	if (touched.footer !== undefined) {
		if (margins.footer !== undefined) input.footer = margins.footer
		else addPageMarginsLossyIssue(preimage.sheet, issues)
	}
	return Object.keys(input).length > 0 ? input : null
}

function addPageSetupLossyIssue(sheet: string, issues: MutationJournalIssue[]): void {
	if (issues.some((issue) => issue.message.startsWith(`Page setup for ${sheet} contains`))) return
	issues.push({
		code: 'LOSSY_INVERSE',
		message: `Page setup for ${sheet} contains metadata that cannot be restored with public operations`,
		surface: 'page-setup',
		reason: 'page-setup-unsettable',
		refs: [sheet],
	})
}

function addPageMarginsLossyIssue(sheet: string, issues: MutationJournalIssue[]): void {
	if (issues.some((issue) => issue.message.startsWith(`Page margins for ${sheet} cannot`))) return
	issues.push({
		code: 'LOSSY_INVERSE',
		message: `Page margins for ${sheet} cannot be removed with public operations`,
		surface: 'page-setup',
		reason: 'page-margins-unsettable',
		refs: [sheet],
	})
}

function restoreDefinedNameOps(
	workbook: Workbook,
	preimage: MutationJournalDefinedNamePreimage,
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const definedName = preimage.definedName
	if (!definedName) {
		return {
			inverseOps: [
				{
					op: 'deleteDefinedName',
					name: preimage.name,
					...(preimage.scope !== undefined ? { scope: preimage.scope } : {}),
				},
			],
			issues: [],
		}
	}
	const scope =
		definedName.scope.kind === 'sheet'
			? sheetNameForId(workbook, definedName.scope.sheetId)
			: undefined
	const refs = [`${scope ? `${scope}!` : ''}${definedName.name}`]
	const issues: MutationJournalIssue[] = []
	if (definedName.hidden !== undefined || definedName.extraAttributes !== undefined) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Defined name ${refs[0]} has metadata that cannot be restored with public operations`,
			surface: 'defined-names',
			reason: 'defined-name-metadata',
			refs,
		})
	}
	if (definedName.scope.kind === 'sheet' && scope === undefined) {
		return {
			inverseOps: [],
			issues: [
				...issues,
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore sheet-scoped defined name ${definedName.name} because its sheet scope is missing`,
					surface: 'defined-names',
					reason: 'operation-unsupported',
				},
			],
		}
	}
	return {
		inverseOps: [
			{
				op: 'setDefinedName',
				name: definedName.name,
				ref: definedName.formula,
				...(scope !== undefined ? { scope } : {}),
			},
		],
		issues,
	}
}

function tableColumnPreimage(
	workbook: Workbook,
	tableName: string,
	column: string | number,
): MutationJournalTableColumnPreimage {
	const located = findTableColumn(workbook, tableName, column)
	return {
		table: tableName,
		column,
		columnIndex: located?.columnIndex ?? -1,
		columnState: located?.column ? cloneTableColumn(located.column) : null,
	}
}

function restoreTableColumnOps(
	op: Extract<Operation, { op: 'setTableColumn' }>,
	preimage: MutationJournalTableColumnPreimage,
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const column = preimage.columnState
	if (!column) {
		return {
			inverseOps: [],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore table column ${op.table}[${String(op.column)}] because it was not found before the edit`,
					surface: 'tables',
					reason: 'operation-unsupported',
				},
			],
		}
	}
	const inverse: Extract<Operation, { op: 'setTableColumn' }> = {
		op: 'setTableColumn',
		table: op.table,
		column: op.newName ?? op.column,
		...(op.newName !== undefined ? { newName: column.name } : {}),
		...(op.formula !== undefined ? { formula: column.formula ?? null } : {}),
		...(op.totalsRowFunction !== undefined
			? { totalsRowFunction: column.totalsRowFunction ?? null }
			: {}),
		...(op.totalsRowFormula !== undefined
			? { totalsRowFormula: column.totalsRowFormula ?? null }
			: {}),
		...(op.totalsRowLabel !== undefined ? { totalsRowLabel: column.totalsRowLabel ?? null } : {}),
	}
	return { inverseOps: [inverse], issues: [] }
}

function tablePreimage(workbook: Workbook, tableName: string): MutationJournalTablePreimage {
	const located = findTable(workbook, tableName)
	return {
		table: located ? cloneTable(located.table) : null,
		sheet: located?.sheet.name ?? null,
		ref: located ? toRangeString(located.table.ref) : null,
	}
}

function restoreTableOps(preimage: MutationJournalTablePreimage): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	if (!preimage.table) {
		return {
			inverseOps: [],
			issues: [
				{
					code: 'UNSUPPORTED_VALUE',
					message: 'Cannot restore deleted table because it was not found before the edit',
					surface: 'tables',
					reason: 'operation-unsupported',
				},
			],
		}
	}
	return restoreExistingTableOps(preimage.table, preimage.sheet ?? undefined, {
		includeCreate: true,
	})
}

function restoreExistingTableOps(
	table: Table,
	sheet: string | undefined,
	options: { readonly includeCreate: boolean; readonly tableName?: string },
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const issues = tableRestoreIssues(table, sheet, options.includeCreate)
	if (!sheet) {
		return {
			inverseOps: [],
			issues: [
				...issues,
				{
					code: 'UNSUPPORTED_VALUE',
					message: `Cannot restore table ${table.name} because its sheet is missing`,
					surface: 'tables',
					reason: 'operation-unsupported',
				},
			],
		}
	}
	const inverseOps: Operation[] = options.includeCreate
		? [
				{
					op: 'createTable',
					sheet,
					ref: toRangeString(table.ref),
					name: table.name,
					hasHeaders: table.hasHeaders,
				},
			]
		: [
				{
					op: 'resizeTable',
					table: options.tableName ?? table.name,
					ref: toRangeString(table.ref),
				},
			]
	if (table.tableStyleInfo) {
		inverseOps.push(tableStyleSetOperation(table.name, table.tableStyleInfo))
	}
	table.columns.forEach((column, index) => {
		const op = tableColumnRestoreOperation(table.name, index, column, {
			restoreName: options.includeCreate,
		})
		if (op) inverseOps.push(op)
	})
	return { inverseOps, issues }
}

function tableStylePreimage(
	workbook: Workbook,
	tableName: string,
): MutationJournalTableStylePreimage {
	const style = findTable(workbook, tableName)?.table.tableStyleInfo
	return {
		table: tableName,
		style: style ? { ...style } : null,
	}
}

function tableStyleSetOperation(
	table: string,
	style: TableStyleInfo | null,
): Extract<Operation, { op: 'setTableStyle' }> {
	return {
		op: 'setTableStyle',
		table,
		styleName: style?.name ?? null,
		...(style?.showFirstColumn !== undefined ? { showFirstColumn: style.showFirstColumn } : {}),
		...(style?.showLastColumn !== undefined ? { showLastColumn: style.showLastColumn } : {}),
		...(style?.showRowStripes !== undefined ? { showRowStripes: style.showRowStripes } : {}),
		...(style?.showColumnStripes !== undefined
			? { showColumnStripes: style.showColumnStripes }
			: {}),
	}
}

function tableStyleRestoreIssues(
	op: Extract<Operation, { op: 'setTableStyle' }>,
	style: TableStyleInfo | null,
): MutationJournalIssue[] {
	const touchedBooleanWithNoPreimage =
		(op.showFirstColumn !== undefined && style?.showFirstColumn === undefined) ||
		(op.showLastColumn !== undefined && style?.showLastColumn === undefined) ||
		(op.showRowStripes !== undefined && style?.showRowStripes === undefined) ||
		(op.showColumnStripes !== undefined && style?.showColumnStripes === undefined)
	return touchedBooleanWithNoPreimage
		? [
				{
					code: 'LOSSY_INVERSE',
					message: `Table style flags for ${op.table} cannot be cleared with public operations`,
					surface: 'tables',
					reason: 'table-metadata',
				},
			]
		: []
}

function tableColumnRestoreOperation(
	table: string,
	columnIndex: number,
	column: TableColumn,
	options: { readonly restoreName: boolean },
): Extract<Operation, { op: 'setTableColumn' }> | null {
	const op: Extract<Operation, { op: 'setTableColumn' }> = {
		op: 'setTableColumn',
		table,
		column: columnIndex,
		...(options.restoreName ? { newName: column.name } : {}),
		...(column.formula !== undefined ? { formula: column.formula } : {}),
		...(column.totalsRowFunction !== undefined
			? { totalsRowFunction: column.totalsRowFunction }
			: {}),
		...(column.totalsRowFormula !== undefined ? { totalsRowFormula: column.totalsRowFormula } : {}),
		...(column.totalsRowLabel !== undefined ? { totalsRowLabel: column.totalsRowLabel } : {}),
	}
	return Object.keys(op).length > 3 ? op : null
}

function tableRestoreIssues(
	table: Table,
	sheet: string | undefined,
	includeCreate: boolean,
): MutationJournalIssue[] {
	const refs = [sheet ? `${sheet}!${toRangeString(table.ref)}` : table.name]
	const issues: MutationJournalIssue[] = []
	if (
		table.partPath !== undefined ||
		table.contentType !== undefined ||
		table.contentTypeSource !== undefined ||
		table.sourcePartPath !== undefined ||
		table.sourceRelationshipPart !== undefined ||
		table.sourceRelationshipId !== undefined ||
		table.sourceRelationshipType !== undefined ||
		table.sourceRelationshipRawType !== undefined ||
		table.sourceRelationshipRawTarget !== undefined ||
		table.sourceRelationshipResolvedTarget !== undefined ||
		table.sourceRelationshipTargetMode !== undefined ||
		table.uid !== undefined ||
		table.tableType !== undefined ||
		table.insertRow !== undefined ||
		table.insertRowShift !== undefined ||
		table.altText !== undefined ||
		table.altTextSummary !== undefined ||
		table.autoFilter !== undefined ||
		table.sortState !== undefined ||
		table.dxfId !== undefined ||
		table.dataCellStyle !== undefined ||
		table.headerRowDxfId !== undefined ||
		table.headerRowCellStyle !== undefined ||
		table.dataDxfId !== undefined ||
		table.totalsRowDxfId !== undefined ||
		table.headerRowBorderDxfId !== undefined ||
		table.tableBorderDxfId !== undefined ||
		table.queryTable !== undefined
	) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Table ${table.name} has metadata that cannot be restored with public operations`,
			surface: 'tables',
			reason: 'table-metadata',
			refs,
		})
	}
	if (includeCreate && table.hasTotals) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Table ${table.name} totals-row state cannot be recreated with public operations`,
			surface: 'tables',
			reason: 'table-metadata',
			refs,
		})
	}
	for (const column of table.columns) {
		if (
			column.id !== undefined ||
			column.uid !== undefined ||
			column.uniqueName !== undefined ||
			column.formulaIsArray !== undefined ||
			column.xmlColumnPr !== undefined ||
			column.queryTableFieldId !== undefined ||
			column.dataCellStyle !== undefined ||
			column.dataDxfId !== undefined ||
			column.headerRowDxfId !== undefined ||
			column.totalsRowDxfId !== undefined
		) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Table column ${table.name}[${column.name}] has metadata that cannot be restored with public operations`,
				surface: 'tables',
				reason: 'table-metadata',
				refs,
			})
		}
	}
	return issues
}

function missingTableIssues(table: string): readonly MutationJournalIssue[] {
	return [
		{
			code: 'UNSUPPORTED_VALUE',
			message: `Cannot restore table ${table} because it was not found before the edit`,
			surface: 'tables',
			reason: 'operation-unsupported',
		},
	]
}

function structuralDeletePreimage(
	workbook: Workbook,
	sheetName: string,
	axis: 'row' | 'col',
	at: number,
	count: number,
): MutationJournalStructuralPreimage {
	const refs = structuralDeletedCellRefs(workbook.getSheet(sheetName), axis, at, count)
	return {
		sheet: sheetName,
		axis,
		at,
		count,
		deletedCells: cellPreimages(workbook, sheetName, refs),
	}
}

function structuralDeletedCellRefs(
	sheet: Sheet | undefined,
	axis: 'row' | 'col',
	at: number,
	count: number,
): string[] {
	const usedRange = sheet?.cells.usedRange()
	if (!sheet || !usedRange) return []
	const deleted =
		axis === 'row'
			? {
					start: { row: at, col: usedRange.start.col },
					end: { row: at + count - 1, col: usedRange.end.col },
				}
			: {
					start: { row: usedRange.start.row, col: at },
					end: { row: usedRange.end.row, col: at + count - 1 },
				}
	const refs: string[] = []
	for (const [row, col] of sheet.cells.getRange(deleted)) {
		refs.push(toA1({ row, col }))
	}
	return refs
}

function structuralDeleteIssues(
	workbook: Workbook,
	preimage: MutationJournalStructuralPreimage,
): readonly MutationJournalIssue[] {
	const sheet = workbook.getSheet(preimage.sheet)
	if (!sheet) return []
	const refs = [
		`${preimage.sheet}!${preimage.axis === 'row' ? preimage.at + 1 : toA1({ row: 0, col: preimage.at }).replace(/\d+$/, '')}`,
	]
	const issues: MutationJournalIssue[] = []
	const affected = structuralAffectedRange(preimage.axis, preimage.at, preimage.count)
	if (
		hasDeletedAxisDimensions(sheet, preimage.axis, preimage.at, preimage.count) ||
		hasMapRefInRange(sheet.comments, affected) ||
		hasMapRefInRange(sheet.hyperlinks, affected) ||
		sheet.threadedComments.some((comment) => refInRange(comment.ref, affected)) ||
		sheet.merges.some((merge) => rangesOverlap(merge, affected)) ||
		sheet.dataValidations.some((validation) => sqrefOverlaps(validation.sqref, affected)) ||
		sheet.conditionalFormats.some((format) => sqrefOverlaps(format.sqref, affected)) ||
		sheet.tables.some((table) => rangesOverlap(table.ref, affected)) ||
		(sheet.autoFilter?.ref ? sqrefOverlaps(sheet.autoFilter.ref, affected) : false)
	) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Deleted ${preimage.axis === 'row' ? 'row' : 'column'} metadata on ${preimage.sheet} cannot be fully restored with public operations`,
			surface: preimage.axis === 'row' ? 'row-layout' : 'column-layout',
			reason: preimage.axis === 'row' ? 'row-layout-created' : 'column-layout-created',
			refs,
		})
	}
	issues.push(...structuralFormulaReferenceIssues(workbook, preimage))
	issues.push(...structuralX14MetadataIssues(sheet, preimage, affected))
	issues.push(...structuralRepresentedMetadataIssues(workbook, sheet, preimage, affected))
	return issues
}

function structuralX14MetadataIssues(
	sheet: Sheet,
	preimage: MutationJournalStructuralPreimage,
	affected: RangeRef,
): readonly MutationJournalIssue[] {
	const refs = [
		...sheet.x14DataValidations
			.filter((validation) => !validation.deleted && sqrefOverlaps(validation.sqref, affected))
			.map((validation) => `${sheet.name}!x14Validation:${validation.sqref}:${validation.index}`),
		...sheet.x14ConditionalFormats
			.filter((format) => !format.deleted && sqrefOverlaps(format.sqref, affected))
			.map((format) => `${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}`),
	]
	if (refs.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Deleted ${preimage.axis === 'row' ? 'row' : 'column'} x14 metadata on ${preimage.sheet} cannot be fully restored with public operations`,
			surface: 'x14-metadata',
			reason: 'x14-metadata',
			refs,
		},
	]
}

function structuralFormulaReferenceIssues(
	workbook: Workbook,
	preimage: MutationJournalStructuralPreimage,
): readonly MutationJournalIssue[] {
	const refs = structuralFormulaReferenceLocations(workbook, preimage)
	if (refs.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Deleted ${preimage.axis === 'row' ? 'row' : 'column'} formula references on ${preimage.sheet} cannot be restored with public operations`,
			surface: 'formulas',
			reason: 'formula-reference-rewrite',
			refs,
		},
	]
}

function structuralRepresentedMetadataIssues(
	workbook: Workbook,
	sheet: Sheet,
	preimage: MutationJournalStructuralPreimage,
	affected: RangeRef,
): readonly MutationJournalIssue[] {
	const refs = [
		...sheet.ignoredErrors
			.filter((entry) => sqrefOverlaps(entry.sqref, affected))
			.map((entry) => `${sheet.name}!ignoredError:${entry.sqref}`),
		...sortStateStructuralRefs(sheet.name, sheet.sortState, affected, 'sortState'),
		...sheet.advancedFilters.flatMap((filter, index) => [
			...(filter.ref && sqrefOverlaps(filter.ref, affected)
				? [`${sheet.name}!advancedFilter:${index}:${filter.ref}`]
				: []),
			...autoFilterStructuralRefs(
				sheet.name,
				filter.autoFilter,
				affected,
				`advancedFilter:${index}:autoFilter`,
			),
		]),
		...sheet.imageRefs
			.filter((image) => anchorOverlapsAffected(image.anchor, affected))
			.map((image, index) => `${sheet.name}!image:${image.drawingPartPath}:${index}`),
		...sheet.drawingObjectRefs
			.filter((object) => anchorOverlapsAffected(object.anchor, affected))
			.map((object, index) => `${sheet.name}!drawing:${object.drawingPartPath}:${index}`),
		...chartStructuralFormulaRefs(workbook, preimage),
		...pivotStructuralRefs(workbook, preimage, affected),
	]
	if (refs.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Deleted ${preimage.axis === 'row' ? 'row' : 'column'} represented metadata on ${preimage.sheet} cannot be fully restored with public operations`,
			surface: 'package-parts',
			reason: 'package-part-preservation',
			refs,
		},
	]
}

function autoFilterStructuralRefs(
	sheetName: string,
	autoFilter: AutoFilter | null | undefined,
	affected: RangeRef,
	prefix: string,
): string[] {
	if (!autoFilter) return []
	return [
		...(sqrefOverlaps(autoFilter.ref, affected)
			? [`${sheetName}!${prefix}:${autoFilter.ref}`]
			: []),
		...sortStateStructuralRefs(sheetName, autoFilter.sortState, affected, `${prefix}:sortState`),
	]
}

function sortStateStructuralRefs(
	sheetName: string,
	sortState: AutoFilter['sortState'] | null,
	affected: RangeRef,
	prefix: string,
): string[] {
	if (!sortState) return []
	return [
		...(sqrefOverlaps(sortState.ref, affected) ? [`${sheetName}!${prefix}:${sortState.ref}`] : []),
		...sortState.conditions
			.filter((condition) => sqrefOverlaps(condition.ref, affected))
			.map((condition, index) => `${sheetName}!${prefix}:condition:${index}:${condition.ref}`),
	]
}

function structuralFormulaReferenceLocations(
	workbook: Workbook,
	preimage: MutationJournalStructuralPreimage,
): string[] {
	const refs: string[] = []
	for (const sheet of workbook.sheets) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formula) continue
			if (formulaReferencesDeletedAxis(workbook, cell.formula, sheet.name, preimage)) {
				refs.push(`${sheet.name}!${toA1({ row, col })}`)
			}
		}
		for (const validation of sheet.dataValidations) {
			pushMetadataFormulaReferenceLocation(
				refs,
				workbook,
				validation.formula1,
				sheet.name,
				preimage,
				`${sheet.name}!validation:${validation.sqref}:formula1`,
			)
			pushMetadataFormulaReferenceLocation(
				refs,
				workbook,
				validation.formula2,
				sheet.name,
				preimage,
				`${sheet.name}!validation:${validation.sqref}:formula2`,
			)
		}
		for (const validation of sheet.x14DataValidations) {
			if (validation.deleted) continue
			pushMetadataFormulaReferenceLocation(
				refs,
				workbook,
				validation.formula1,
				sheet.name,
				preimage,
				`${sheet.name}!x14Validation:${validation.sqref}:formula1`,
			)
			pushMetadataFormulaReferenceLocation(
				refs,
				workbook,
				validation.formula2,
				sheet.name,
				preimage,
				`${sheet.name}!x14Validation:${validation.sqref}:formula2`,
			)
		}
		sheet.conditionalFormats.forEach((format, formatIndex) => {
			format.rules.forEach((rule, ruleIndex) => {
				rule.formulas.forEach((formula, formulaIndex) => {
					pushMetadataFormulaReferenceLocation(
						refs,
						workbook,
						formula,
						sheet.name,
						preimage,
						`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:${formulaIndex}`,
					)
				})
				pushConditionalFormatValueObjectReferenceLocations(
					refs,
					workbook,
					rule.colorScale?.cfvo,
					sheet.name,
					preimage,
					`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:colorScale.cfvo`,
				)
				pushConditionalFormatValueObjectReferenceLocations(
					refs,
					workbook,
					rule.dataBar?.cfvo,
					sheet.name,
					preimage,
					`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:dataBar.cfvo`,
				)
				pushConditionalFormatValueObjectReferenceLocations(
					refs,
					workbook,
					rule.iconSet?.cfvo,
					sheet.name,
					preimage,
					`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:iconSet.cfvo`,
				)
			})
		})
		sheet.x14ConditionalFormats.forEach((format) => {
			if (format.deleted) return
			format.formulas.forEach((formula, formulaIndex) => {
				pushMetadataFormulaReferenceLocation(
					refs,
					workbook,
					formula,
					sheet.name,
					preimage,
					`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:${formulaIndex}`,
				)
			})
			pushConditionalFormatValueObjectReferenceLocations(
				refs,
				workbook,
				format.colorScale?.cfvo,
				sheet.name,
				preimage,
				`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:colorScale.cfvo`,
			)
			pushConditionalFormatValueObjectReferenceLocations(
				refs,
				workbook,
				format.dataBar?.cfvo,
				sheet.name,
				preimage,
				`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:dataBar.cfvo`,
			)
			pushConditionalFormatValueObjectReferenceLocations(
				refs,
				workbook,
				format.iconSet?.cfvo,
				sheet.name,
				preimage,
				`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:iconSet.cfvo`,
			)
		})
		for (const table of sheet.tables) {
			for (const column of table.columns) {
				pushMetadataFormulaReferenceLocation(
					refs,
					workbook,
					column.formula,
					sheet.name,
					preimage,
					`${sheet.name}!table:${table.name}:${column.name}:formula`,
				)
				pushMetadataFormulaReferenceLocation(
					refs,
					workbook,
					column.totalsRowFormula,
					sheet.name,
					preimage,
					`${sheet.name}!table:${table.name}:${column.name}:totalsRowFormula`,
				)
			}
		}
	}
	for (const name of workbook.definedNames.list()) {
		const scopeSheet =
			name.scope.kind === 'sheet' ? sheetNameForId(workbook, name.scope.sheetId) : undefined
		if (
			formulaReferencesDeletedAxis(workbook, name.formula, scopeSheet ?? preimage.sheet, preimage)
		) {
			refs.push(definedNameJournalKey(workbook, name))
		}
	}
	return refs
}

function pushConditionalFormatValueObjectReferenceLocations(
	refs: string[],
	workbook: Workbook,
	entries: readonly SheetConditionalFormatValueObject[] | undefined,
	ownerSheet: string,
	preimage: MutationJournalStructuralPreimage,
	location: string,
): void {
	entries?.forEach((entry, index) => {
		pushMetadataFormulaReferenceLocation(
			refs,
			workbook,
			entry.value,
			ownerSheet,
			preimage,
			`${location}:${index}`,
		)
	})
}

function chartStructuralFormulaRefs(
	workbook: Workbook,
	preimage: MutationJournalStructuralPreimage,
): string[] {
	const refs: string[] = []
	for (const chart of workbook.chartParts) {
		chart.series.forEach((series, seriesIndex) => {
			for (const field of ['nameRef', 'categoryRef', 'valueRef'] as const) {
				const formula = series[field]
				if (
					formula !== undefined &&
					formulaReferencesDeletedAxis(
						workbook,
						formula,
						chart.sheetName ?? preimage.sheet,
						preimage,
					)
				) {
					refs.push(`chart:${chart.partPath}:series:${seriesIndex}:${field}`)
				}
			}
		})
	}
	return refs
}

function pivotStructuralRefs(
	workbook: Workbook,
	preimage: MutationJournalStructuralPreimage,
	affected: RangeRef,
): string[] {
	const refs: string[] = []
	for (const cache of workbook.pivotCaches) {
		if (
			cache.sourceRef !== undefined &&
			refTextOverlapsAffected(cache.sourceRef, cache.sourceSheet, preimage, affected)
		) {
			refs.push(`pivotCache:${cache.partPath}:sourceRef`)
		}
	}
	for (const pivot of workbook.pivotTables) {
		if (
			pivot.locationRef &&
			refTextOverlapsAffected(pivot.locationRef, pivot.sheetName, preimage, affected)
		) {
			refs.push(`pivotTable:${pivot.partPath}:locationRef`)
		}
		if (
			pivot.location?.ref &&
			refTextOverlapsAffected(pivot.location.ref, pivot.sheetName, preimage, affected)
		) {
			refs.push(`pivotTable:${pivot.partPath}:location.ref`)
		}
	}
	return refs
}

function refTextOverlapsAffected(
	refText: string,
	ownerSheet: string | undefined,
	preimage: MutationJournalStructuralPreimage,
	affected: RangeRef,
): boolean {
	const split = splitSheetQualifiedRefText(refText)
	const sheetName = split?.sheet ?? ownerSheet
	if (sheetName === undefined || !sameSheetName(sheetName, preimage.sheet)) return false
	const ref = refTextToRange(split?.ref ?? refText)
	return ref ? rangesOverlap(ref, affected) : false
}

function splitSheetQualifiedRefText(
	input: string,
): { readonly sheet: string; readonly ref: string } | null {
	const bang = input.lastIndexOf('!')
	if (bang < 0) return null
	const sheet = input.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'")
	const ref = input.slice(bang + 1)
	return sheet && ref ? { sheet, ref } : null
}

function refTextToRange(input: string): RangeRef | null {
	const normalized = input.replace(/\$/g, '')
	try {
		return parseRange(normalized)
	} catch {
		try {
			const ref = parseA1(normalized)
			return { start: ref, end: ref }
		} catch {
			return null
		}
	}
}

function anchorOverlapsAffected(anchor: SheetImageAnchor | undefined, affected: RangeRef): boolean {
	if (!anchor || anchor.kind === 'absolute') return false
	const range =
		anchor.kind === 'oneCell'
			? {
					start: { row: anchor.from.row, col: anchor.from.col },
					end: { row: anchor.from.row, col: anchor.from.col },
				}
			: {
					start: {
						row: Math.min(anchor.from.row, anchor.to.row),
						col: Math.min(anchor.from.col, anchor.to.col),
					},
					end: {
						row: Math.max(anchor.from.row, anchor.to.row),
						col: Math.max(anchor.from.col, anchor.to.col),
					},
				}
	return rangesOverlap(range, affected)
}

function pushMetadataFormulaReferenceLocation(
	refs: string[],
	workbook: Workbook,
	formula: string | undefined,
	ownerSheet: string,
	preimage: MutationJournalStructuralPreimage,
	location: string,
): void {
	if (formula === undefined) return
	if (formulaReferencesDeletedAxis(workbook, formula, ownerSheet, preimage)) refs.push(location)
}

function formulaReferencesDeletedAxis(
	workbook: Workbook,
	formula: string,
	ownerSheet: string,
	preimage: MutationJournalStructuralPreimage,
): boolean {
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return false
	return extractRefs(parsed.value).some((ref) =>
		formulaRefOverlapsDeletedAxis(workbook, ref, ownerSheet, preimage),
	)
}

function formulaRefOverlapsDeletedAxis(
	workbook: Workbook,
	ref: FormulaRef,
	ownerSheet: string,
	preimage: MutationJournalStructuralPreimage,
): boolean {
	if (ref.kind === 'sheetSpan') {
		return (
			sheetSpanIncludes(workbook, ref.startSheet, ref.endSheet, preimage.sheet) &&
			formulaRefOverlapsDeletedAxis(workbook, ref.target, preimage.sheet, preimage)
		)
	}
	const sheet = 'sheet' in ref && ref.sheet !== undefined ? ref.sheet : ownerSheet
	if (!sameSheetName(sheet, preimage.sheet)) return false
	const end = preimage.at + preimage.count - 1
	switch (ref.kind) {
		case 'cell':
			return preimage.axis === 'row'
				? ref.ref.row >= preimage.at && ref.ref.row <= end
				: ref.ref.col >= preimage.at && ref.ref.col <= end
		case 'range':
			return preimage.axis === 'row'
				? ref.start.row <= end && ref.end.row >= preimage.at
				: ref.start.col <= end && ref.end.col >= preimage.at
		case 'wholeRowRange':
			return preimage.axis === 'row' && ref.startRow <= end && ref.endRow >= preimage.at
		case 'wholeColumnRange':
			return preimage.axis === 'col' && ref.startCol <= end && ref.endCol >= preimage.at
		default:
			return false
	}
}

function sheetSpanIncludes(
	workbook: Workbook,
	startSheet: string,
	endSheet: string,
	targetSheet: string,
): boolean {
	const start = workbook.sheets.findIndex((sheet) => sheet.name === startSheet)
	const end = workbook.sheets.findIndex((sheet) => sheet.name === endSheet)
	const target = workbook.sheets.findIndex((sheet) => sheet.name === targetSheet)
	if (start < 0 || end < 0 || target < 0) return false
	return target >= Math.min(start, end) && target <= Math.max(start, end)
}

function definedNameJournalKey(workbook: Workbook, entry: DefinedName): string {
	if (entry.scope.kind === 'workbook') return `name:${entry.name}`
	const sheetName = sheetNameForId(workbook, entry.scope.sheetId) ?? 'Sheet'
	return `name:${sheetName}!${entry.name}`
}

function structuralAffectedRange(axis: 'row' | 'col', at: number, count: number): RangeRef {
	const maxSpreadsheetIndex = Number.MAX_SAFE_INTEGER
	return axis === 'row'
		? {
				start: { row: at, col: 0 },
				end: { row: at + count - 1, col: maxSpreadsheetIndex },
			}
		: {
				start: { row: 0, col: at },
				end: { row: maxSpreadsheetIndex, col: at + count - 1 },
			}
}

function hasDeletedAxisDimensions(
	sheet: Sheet,
	axis: 'row' | 'col',
	at: number,
	count: number,
): boolean {
	const end = at + count
	if (axis === 'row') {
		for (const row of sheet.rowHeights.keys()) if (row >= at && row < end) return true
		for (const row of sheet.rowDefs.keys()) if (row >= at && row < end) return true
		return false
	}
	for (const col of sheet.colWidths.keys()) if (col >= at && col < end) return true
	return sheet.colDefs.some((def) => def.max >= at && def.min < end)
}

function hasMapRefInRange<T>(map: ReadonlyMap<string, T>, range: RangeRef): boolean {
	for (const ref of map.keys()) {
		if (refInRange(ref, range)) return true
	}
	return false
}

function refInRange(ref: string, range: RangeRef): boolean {
	try {
		const parsed = parseA1(ref)
		return (
			parsed.row >= range.start.row &&
			parsed.row <= range.end.row &&
			parsed.col >= range.start.col &&
			parsed.col <= range.end.col
		)
	} catch {
		return false
	}
}

function sqrefOverlaps(sqref: string, range: RangeRef): boolean {
	return sqref
		.split(/\s+/)
		.filter(Boolean)
		.some((part) => {
			try {
				return rangesOverlap(parseRange(part), range)
			} catch {
				return false
			}
		})
}

function parseSqrefRanges(sqref: string): RangeRef[] {
	const ranges: RangeRef[] = []
	for (const token of sqref.trim().split(/\s+/)) {
		if (!token) continue
		try {
			ranges.push(parseRange(token))
		} catch {
			return []
		}
	}
	return ranges
}

function rangeContainsRange(outer: RangeRef, inner: RangeRef): boolean {
	return (
		inner.start.row >= outer.start.row &&
		inner.end.row <= outer.end.row &&
		inner.start.col >= outer.start.col &&
		inner.end.col <= outer.end.col
	)
}

function shiftRange(range: RangeRef, rowDelta: number, colDelta: number): RangeRef {
	return {
		start: { row: range.start.row + rowDelta, col: range.start.col + colDelta },
		end: { row: range.end.row + rowDelta, col: range.end.col + colDelta },
	}
}

function rangesToSqref(ranges: readonly RangeRef[]): string {
	return ranges.map(rangeToA1).join(' ')
}

function rangesOverlap(a: RangeRef, b: RangeRef): boolean {
	return (
		a.start.row <= b.end.row &&
		a.end.row >= b.start.row &&
		a.start.col <= b.end.col &&
		a.end.col >= b.start.col
	)
}

interface SortRangeDataValidationMove {
	readonly sourceIndex: number
	readonly source: MutationJournalDataValidationPreimage
	readonly targetRange: string
}

interface SortRangeConditionalFormatMove {
	readonly sourceIndex: number
	readonly source: SheetConditionalFormat
	readonly targetRange: string
}

function sortRangeDataValidationRestoration(
	sheet: Sheet,
	op: Extract<Operation, { op: 'sortRange' }>,
	dataRange: RangeRef,
): {
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const moves = sortRangeDataValidationMoves(sheet, op, dataRange)
	if (moves.length === 0) return EMPTY_METADATA_RESTORATION
	const inverseOps: Operation[] = []
	for (const range of uniqueStrings(moves.map((move) => move.targetRange))) {
		inverseOps.push({ op: 'deleteDataValidation', sheet: sheet.name, range })
	}
	const issues: MutationJournalIssue[] = [
		...sortRangeDataValidationOrderIssues(sheet, op.range, moves),
		...sortRangeDataValidationDuplicateIssues(sheet.name, op.range, moves),
	]
	for (const move of moves) {
		const restored = restoreDataValidationOps(move.source)
		inverseOps.push(...restored.inverseOps)
		issues.push(...restored.issues)
	}
	return {
		inverseOps,
		preimages: [{ kind: 'data-validations', validations: moves.map((move) => move.source) }],
		issues,
	}
}

function sortRangeDataValidationMoves(
	sheet: Sheet,
	op: Extract<Operation, { op: 'sortRange' }>,
	dataRange: RangeRef,
): SortRangeDataValidationMove[] {
	const rowTargets = sortRangeRowTargets(sheet, op, dataRange)
	const moves: SortRangeDataValidationMove[] = []
	for (let index = 0; index < sheet.dataValidations.length; index++) {
		const validation = sheet.dataValidations[index]
		if (!validation) continue
		const parsed = parseSortRowScopedSqref(validation.sqref, dataRange)
		if (!parsed) continue
		const targetRow = rowTargets.get(parsed.row)
		if (targetRow === undefined) continue
		const source: MutationJournalDataValidationPreimage = {
			sheet: sheet.name,
			range: validation.sqref,
			validation: { ...validation },
		}
		moves.push({
			sourceIndex: index,
			source,
			targetRange: rowScopedSqref(parsed, targetRow),
		})
	}
	return moves
}

function sortRangeRowTargets(
	sheet: Sheet,
	op: Extract<Operation, { op: 'sortRange' }>,
	dataRange: RangeRef,
): Map<number, number> {
	const range = parseRange(op.range)
	const columns = sortRangeResolvedColumns(sheet, range, op.by)
	const rows = Array.from({ length: dataRange.end.row - dataRange.start.row + 1 }, (_, offset) => ({
		originalRow: dataRange.start.row + offset,
		originalIndex: offset,
	}))
	rows.sort((left, right) => {
		for (const column of columns) {
			const leftValue = sheet.cells.readValue(left.originalRow, column.col) ?? EMPTY
			const rightValue = sheet.cells.readValue(right.originalRow, column.col) ?? EMPTY
			const result = compareValues(leftValue, rightValue)
			if (result !== 0) return column.descending ? -result : result
		}
		return left.originalIndex - right.originalIndex
	})
	return new Map(rows.map((row, index) => [row.originalRow, dataRange.start.row + index]))
}

function sortRangeResolvedColumns(
	sheet: Sheet,
	range: RangeRef,
	specs: Extract<Operation, { op: 'sortRange' }>['by'],
): Array<{ readonly col: number; readonly descending: boolean }> {
	const headerMap = new Map<string, number>()
	for (let col = range.start.col; col <= range.end.col; col++) {
		const value = sheet.cells.readValue(range.start.row, col)
		if (value?.kind === 'string' && value.value.trim() !== '') {
			headerMap.set(value.value.trim().toLowerCase(), col)
		}
	}
	return specs.map((spec) => {
		if (typeof spec.column === 'number') {
			return {
				col: range.start.col + Math.trunc(spec.column) - 1,
				descending: spec.descending ?? false,
			}
		}
		const trimmed = spec.column.trim()
		const headerCol = headerMap.get(trimmed.toLowerCase())
		if (headerCol !== undefined) return { col: headerCol, descending: spec.descending ?? false }
		return {
			col: parseA1(`${trimmed.toUpperCase()}1`).col,
			descending: spec.descending ?? false,
		}
	})
}

function parseSortRowScopedSqref(
	sqref: string,
	range: RangeRef,
): { readonly row: number; readonly startCol: number; readonly endCol: number } | null {
	if (sqref.includes(' ')) return null
	try {
		const parsed = parseRange(sqref)
		if (parsed.start.row !== parsed.end.row) return null
		if (parsed.start.row < range.start.row || parsed.start.row > range.end.row) return null
		if (parsed.start.col < range.start.col || parsed.end.col > range.end.col) return null
		return { row: parsed.start.row, startCol: parsed.start.col, endCol: parsed.end.col }
	} catch {
		return null
	}
}

function rowScopedSqref(
	parsed: { readonly startCol: number; readonly endCol: number },
	targetRow: number,
): string {
	const start = toA1({ row: targetRow, col: parsed.startCol })
	const end = toA1({ row: targetRow, col: parsed.endCol })
	return start === end ? start : `${start}:${end}`
}

function sortRangeDataValidationOrderIssues(
	sheet: Sheet,
	range: string,
	moves: readonly SortRangeDataValidationMove[],
): readonly MutationJournalIssue[] {
	if (moves.length === 0) return []
	const moved = new Set(moves.map((move) => move.sourceIndex))
	const firstMoved = Math.min(...moved)
	for (let index = firstMoved; index < sheet.dataValidations.length; index++) {
		if (moved.has(index)) continue
		return [
			{
				code: 'LOSSY_INVERSE',
				message: `Sorted data validation order on ${sheet.name}!${range} cannot be restored exactly with public operations`,
				surface: 'data-validations',
				reason: 'metadata-order',
				refs: [`${sheet.name}!${range}`],
			},
		]
	}
	return []
}

function sortRangeDataValidationDuplicateIssues(
	sheetName: string,
	range: string,
	moves: readonly SortRangeDataValidationMove[],
): readonly MutationJournalIssue[] {
	const sourceDuplicates = duplicateStrings(moves.map((move) => move.source.range))
	const targetDuplicates = duplicateStrings(moves.map((move) => move.targetRange))
	if (sourceDuplicates.length === 0 && targetDuplicates.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Sorted duplicate data validations on ${sheetName}!${range} cannot be restored exactly with public operations`,
			surface: 'data-validations',
			reason: 'metadata-duplicate',
			refs: [`${sheetName}!${range}`],
		},
	]
}

function sortRangeConditionalFormatRestoration(
	sheet: Sheet,
	op: Extract<Operation, { op: 'sortRange' }>,
	dataRange: RangeRef,
): {
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const moves = sortRangeConditionalFormatMoves(sheet, op, dataRange)
	if (moves.length === 0) return EMPTY_METADATA_RESTORATION
	const inverseOps: Operation[] = []
	for (const range of uniqueStrings(moves.map((move) => move.targetRange))) {
		inverseOps.push({ op: 'deleteConditionalFormat', sheet: sheet.name, range })
	}
	const restored = restoreConditionalFormatOps(
		sheet.name,
		moves.map((move) => move.source),
	)
	return {
		inverseOps: [...inverseOps, ...restored.inverseOps],
		preimages: [
			{
				kind: 'conditional-formats',
				conditionalFormats: {
					sheet: sheet.name,
					formats: moves.map((move) => cloneConditionalFormat(move.source)),
				},
			},
		],
		issues: [
			...sortRangeConditionalFormatOrderIssues(sheet, op.range, moves),
			...sortRangeConditionalFormatDuplicateIssues(sheet.name, op.range, moves),
			...restored.issues,
		],
	}
}

function sortRangeConditionalFormatMoves(
	sheet: Sheet,
	op: Extract<Operation, { op: 'sortRange' }>,
	dataRange: RangeRef,
): SortRangeConditionalFormatMove[] {
	const rowTargets = sortRangeRowTargets(sheet, op, dataRange)
	const moves: SortRangeConditionalFormatMove[] = []
	for (let index = 0; index < sheet.conditionalFormats.length; index++) {
		const format = sheet.conditionalFormats[index]
		if (!format) continue
		const parsed = parseSortRowScopedSqref(format.sqref, dataRange)
		if (!parsed) continue
		const targetRow = rowTargets.get(parsed.row)
		if (targetRow === undefined) continue
		moves.push({
			sourceIndex: index,
			source: cloneConditionalFormat(format),
			targetRange: rowScopedSqref(parsed, targetRow),
		})
	}
	return moves
}

function sortRangeConditionalFormatOrderIssues(
	sheet: Sheet,
	range: string,
	moves: readonly SortRangeConditionalFormatMove[],
): readonly MutationJournalIssue[] {
	if (moves.length === 0) return []
	const moved = new Set(moves.map((move) => move.sourceIndex))
	const firstMoved = Math.min(...moved)
	for (let index = firstMoved; index < sheet.conditionalFormats.length; index++) {
		if (moved.has(index)) continue
		return [
			{
				code: 'LOSSY_INVERSE',
				message: `Sorted conditional format order on ${sheet.name}!${range} cannot be restored exactly with public operations`,
				surface: 'conditional-formats',
				reason: 'metadata-order',
				refs: [`${sheet.name}!${range}`],
			},
		]
	}
	return []
}

function sortRangeConditionalFormatDuplicateIssues(
	sheetName: string,
	range: string,
	moves: readonly SortRangeConditionalFormatMove[],
): readonly MutationJournalIssue[] {
	const sourceDuplicates = duplicateStrings(moves.map((move) => move.source.sqref))
	const targetDuplicates = duplicateStrings(moves.map((move) => move.targetRange))
	if (sourceDuplicates.length === 0 && targetDuplicates.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Sorted duplicate conditional formats on ${sheetName}!${range} cannot be restored exactly with public operations`,
			surface: 'conditional-formats',
			reason: 'metadata-duplicate',
			refs: [`${sheetName}!${range}`],
		},
	]
}

function uniqueStrings(values: readonly string[]): string[] {
	return [...new Set(values)]
}

function duplicateStrings(values: readonly string[]): string[] {
	const seen = new Set<string>()
	const duplicates = new Set<string>()
	for (const value of values) {
		if (seen.has(value)) duplicates.add(value)
		else seen.add(value)
	}
	return [...duplicates]
}

function sortRangeMetadataIssues(
	workbook: Workbook,
	op: Extract<Operation, { op: 'sortRange' }>,
): readonly MutationJournalIssue[] {
	const sheet = workbook.getSheet(op.sheet)
	if (!sheet) return []
	const dataRange = sortRangeDataRange(sheet, op.range, op.by)
	if (!dataRange) return []
	const hasRowHeights = [...sheet.rowHeights.keys()].some(
		(row) => row >= dataRange.start.row && row <= dataRange.end.row,
	)
	const hasRowDefs = [...sheet.rowDefs.keys()].some(
		(row) => row >= dataRange.start.row && row <= dataRange.end.row,
	)
	const issues: MutationJournalIssue[] = []
	if (hasRowHeights || hasRowDefs) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Sorted row metadata on ${op.sheet}!${op.range} cannot be fully restored with public operations`,
			surface: 'row-layout',
			reason: 'row-layout-created',
			refs: [`${op.sheet}!${op.range}`],
		})
	}
	if (sheet.threadedComments.some((comment) => refInRange(comment.ref, dataRange))) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Sorted threaded comment metadata on ${op.sheet}!${op.range} cannot be fully restored with public operations`,
			surface: 'comments',
			reason: 'threaded-comment-selector',
			refs: [`${op.sheet}!${op.range}`],
		})
	}
	if (
		sheet.x14DataValidations.some(
			(validation) => !validation.deleted && sqrefOverlaps(validation.sqref, dataRange),
		) ||
		sheet.x14ConditionalFormats.some(
			(format) => !format.deleted && sqrefOverlaps(format.sqref, dataRange),
		) ||
		sheet.ignoredErrors.some((entry) => sqrefOverlaps(entry.sqref, dataRange))
	) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: `Sorted x14 row metadata on ${op.sheet}!${op.range} cannot be fully restored with public operations`,
			surface: 'x14-metadata',
			reason: 'x14-metadata',
			refs: [`${op.sheet}!${op.range}`],
		})
	}
	return issues
}

function sortRangeDataRange(
	sheet: Sheet,
	rangeText: string,
	specs: Extract<Operation, { op: 'sortRange' }>['by'],
): RangeRef | null {
	const range = parseRange(rangeText)
	const startRow = sortRangeHasHeaderRow(sheet, range, specs)
		? range.start.row + 1
		: range.start.row
	if (startRow > range.end.row) return null
	return {
		start: { row: startRow, col: range.start.col },
		end: { ...range.end },
	}
}

function sortRangeHasHeaderRow(
	sheet: Sheet,
	range: RangeRef,
	specs: Extract<Operation, { op: 'sortRange' }>['by'],
): boolean {
	const headerMap = new Set<string>()
	for (let col = range.start.col; col <= range.end.col; col++) {
		const value = sheet.cells.readValue(range.start.row, col)
		if (value?.kind === 'string' && value.value.trim() !== '') {
			headerMap.add(value.value.trim().toLowerCase())
		}
	}
	return specs.some(
		(spec) => typeof spec.column === 'string' && headerMap.has(spec.column.trim().toLowerCase()),
	)
}

function copyRangeCellModeSupported(mode: string): boolean {
	return (
		mode === 'all' ||
		mode === 'values' ||
		mode === 'formulas' ||
		mode === 'formats' ||
		mode === 'styles' ||
		mode === 'validations' ||
		mode === 'comments' ||
		mode === 'hyperlinks'
	)
}

function copyRangeOverwritesFormulas(mode: string): boolean {
	return mode === 'all' || mode === 'values' || mode === 'formulas'
}

function copyRangeTransfersConditionalFormats(mode: string): boolean {
	return mode === 'all' || mode === 'formats' || mode === 'styles'
}

function copyRangeTransfersMerges(mode: string): boolean {
	return mode === 'all' || mode === 'formats' || mode === 'styles'
}

function copyRangeRestoration(
	cells: readonly MutationJournalCellPreimage[],
	mode: string,
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	if (mode === 'formats' || mode === 'styles') {
		return { inverseOps: styleInverseOps(cells), issues: [] }
	}
	const { inverseOps, issues } = inverseCellOps(cells)
	return {
		inverseOps: mode === 'all' ? [...inverseOps, ...styleInverseOps(cells)] : inverseOps,
		issues,
	}
}

function moveRangeSourceRestoration(
	cells: readonly MutationJournalCellPreimage[],
	mode: string,
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	if (mode === 'formats' || mode === 'styles') {
		return { inverseOps: styleInverseOps(cells), issues: [] }
	}
	const { inverseOps, issues } = inverseCellOps(cells)
	return {
		inverseOps: [...inverseOps, ...styleInverseOps(cells)],
		issues,
	}
}

const EMPTY_METADATA_RESTORATION: {
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} = { inverseOps: [], preimages: [], issues: [] }

function mergeTransferRestoration(
	workbook: Workbook,
	sourceSheetName: string,
	sourceRange: RangeRef,
	targetSheetName: string,
	targetRange: RangeRef,
	move: boolean,
): {
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const sourceSheet = workbook.getSheet(sourceSheetName)
	const targetSheet = workbook.getSheet(targetSheetName)
	if (!sourceSheet || !targetSheet) return EMPTY_METADATA_RESTORATION
	const rowDelta = targetRange.start.row - sourceRange.start.row
	const colDelta = targetRange.start.col - sourceRange.start.col
	const sourceMerges = sourceSheet.merges.filter((merge) => rangesOverlap(merge, sourceRange))
	const targetMerges = sourceMerges.map((merge) => shiftRange(merge, rowDelta, colDelta))
	const targetReplacedMerges = targetSheet.merges.filter((merge) =>
		rangesOverlap(merge, targetRange),
	)
	const validationIssues = mergeTransferValidationIssues(
		sourceSheetName,
		sourceRange,
		targetSheetName,
		targetRange,
		sourceMerges,
		targetReplacedMerges,
		move,
		sourceSheet === targetSheet,
	)
	const hasMergeEffect = sourceMerges.length > 0 || targetReplacedMerges.length > 0
	if (!hasMergeEffect && validationIssues.length === 0) return EMPTY_METADATA_RESTORATION
	if (validationIssues.length > 0) {
		return {
			inverseOps: [],
			preimages: mergeTransferPreimages(
				workbook,
				sourceSheetName,
				sourceMerges,
				targetSheetName,
				targetReplacedMerges,
				targetMerges,
			),
			issues: validationIssues,
		}
	}

	const inverseOps: Operation[] = []
	for (const merge of uniqueRanges(targetMerges)) {
		inverseOps.push({ op: 'unmergeCells', sheet: targetSheetName, range: rangeToA1(merge) })
	}
	const issues: MutationJournalIssue[] = [
		...mergeTransferDuplicateIssues(sourceSheetName, sourceRange, sourceMerges),
		...mergeTransferDuplicateIssues(targetSheetName, targetRange, targetReplacedMerges),
		...mergeTransferDuplicateIssues(targetSheetName, targetRange, targetMerges),
	]
	if (move && sourceSheet === targetSheet) {
		const removed = mergeRangesInOriginalOrder(sourceSheet.merges, [
			...sourceMerges,
			...targetReplacedMerges,
		])
		issues.push(
			...mergeTransferOrderIssues(
				sourceSheetName,
				sourceRange,
				sourceSheet.merges,
				removed,
				'Moved merge order',
			),
		)
		for (const merge of uniqueRanges(removed)) {
			inverseOps.push({ op: 'mergeCells', sheet: sourceSheetName, range: rangeToA1(merge) })
		}
	} else {
		issues.push(
			...mergeTransferOrderIssues(
				targetSheetName,
				targetRange,
				targetSheet.merges,
				targetReplacedMerges,
				'Replaced target merge order',
			),
		)
		for (const merge of uniqueRanges(targetReplacedMerges)) {
			inverseOps.push({ op: 'mergeCells', sheet: targetSheetName, range: rangeToA1(merge) })
		}
		if (move) {
			issues.push(
				...mergeTransferOrderIssues(
					sourceSheetName,
					sourceRange,
					sourceSheet.merges,
					sourceMerges,
					'Moved merge order',
				),
			)
			for (const merge of uniqueRanges(sourceMerges)) {
				inverseOps.push({ op: 'mergeCells', sheet: sourceSheetName, range: rangeToA1(merge) })
			}
		}
	}

	return {
		inverseOps,
		preimages: mergeTransferPreimages(
			workbook,
			sourceSheetName,
			sourceMerges,
			targetSheetName,
			targetReplacedMerges,
			targetMerges,
		),
		issues,
	}
}

function mergeTransferValidationIssues(
	sourceSheetName: string,
	sourceRange: RangeRef,
	targetSheetName: string,
	targetRange: RangeRef,
	sourceMerges: readonly RangeRef[],
	targetReplacedMerges: readonly RangeRef[],
	move: boolean,
	sameSheet: boolean,
): readonly MutationJournalIssue[] {
	const issues: MutationJournalIssue[] = []
	for (const merge of sourceMerges) {
		if (!rangeContainsRange(sourceRange, merge)) {
			issues.push({
				code: 'UNSUPPORTED_OPERATION',
				message: `Cannot journal ${move ? 'moveRange' : 'copyRange'} merge transfer because ${sourceSheetName}!${rangeToA1(sourceRange)} partially overlaps merged range ${sourceSheetName}!${rangeToA1(merge)}`,
				surface: 'merged-cells',
				reason: 'merge-overlap',
				refs: [
					`${sourceSheetName}!${rangeToA1(sourceRange)}`,
					`${sourceSheetName}!${rangeToA1(merge)}`,
				],
			})
		}
	}
	if (
		sourceMerges.length > 0 &&
		sameSheet &&
		!sameRange(sourceRange, targetRange) &&
		rangesOverlap(sourceRange, targetRange)
	) {
		issues.push({
			code: 'UNSUPPORTED_OPERATION',
			message: `Cannot journal ${move ? 'moveRange' : 'copyRange'} merge transfer onto overlapping target ${targetSheetName}!${rangeToA1(targetRange)}`,
			surface: 'merged-cells',
			reason: 'merge-overlap',
			refs: [
				`${sourceSheetName}!${rangeToA1(sourceRange)}`,
				`${targetSheetName}!${rangeToA1(targetRange)}`,
			],
		})
	}
	for (const merge of targetReplacedMerges) {
		if (!rangeContainsRange(targetRange, merge)) {
			issues.push({
				code: 'UNSUPPORTED_OPERATION',
				message: `Cannot journal ${move ? 'moveRange' : 'copyRange'} merge transfer because ${targetSheetName}!${rangeToA1(targetRange)} partially overlaps merged range ${targetSheetName}!${rangeToA1(merge)}`,
				surface: 'merged-cells',
				reason: 'merge-overlap',
				refs: [
					`${targetSheetName}!${rangeToA1(targetRange)}`,
					`${targetSheetName}!${rangeToA1(merge)}`,
				],
			})
		}
	}
	return issues
}

function mergeTransferPreimages(
	workbook: Workbook,
	sourceSheetName: string,
	sourceMerges: readonly RangeRef[],
	targetSheetName: string,
	targetReplacedMerges: readonly RangeRef[],
	targetMerges: readonly RangeRef[],
): MutationJournalPreimage[] {
	const preimages = new Map<string, MutationJournalPreimage>()
	const add = (sheet: string, merge: RangeRef) => {
		const range = rangeToA1(merge)
		preimages.set(`${sheet}!${range}`, {
			kind: 'merge',
			merge: mergePreimage(workbook, sheet, range),
		})
	}
	for (const merge of sourceMerges) add(sourceSheetName, merge)
	for (const merge of targetReplacedMerges) add(targetSheetName, merge)
	for (const merge of targetMerges) add(targetSheetName, merge)
	return [...preimages.values()]
}

function mergeTransferOrderIssues(
	sheetName: string,
	contextRange: RangeRef,
	originalMerges: readonly RangeRef[],
	removedMerges: readonly RangeRef[],
	label: string,
): readonly MutationJournalIssue[] {
	if (removedMerges.length === 0) return []
	const removed = new Set(removedMerges.map(rangeCoordinateKey))
	const firstRemoved = originalMerges.findIndex((merge) => removed.has(rangeCoordinateKey(merge)))
	if (firstRemoved < 0) return []
	for (let index = firstRemoved; index < originalMerges.length; index++) {
		const merge = originalMerges[index]
		if (merge && removed.has(rangeCoordinateKey(merge))) continue
		return [
			{
				code: 'LOSSY_INVERSE',
				message: `${label} on ${sheetName}!${rangeToA1(contextRange)} cannot be restored exactly with public operations`,
				surface: 'merged-cells',
				reason: 'metadata-order',
				refs: [`${sheetName}!${rangeToA1(contextRange)}`],
			},
		]
	}
	return []
}

function mergeTransferDuplicateIssues(
	sheetName: string,
	contextRange: RangeRef,
	merges: readonly RangeRef[],
): readonly MutationJournalIssue[] {
	if (duplicateStrings(merges.map(rangeCoordinateKey)).length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Duplicate merge metadata on ${sheetName}!${rangeToA1(contextRange)} cannot be restored exactly with public operations`,
			surface: 'merged-cells',
			reason: 'metadata-duplicate',
			refs: [`${sheetName}!${rangeToA1(contextRange)}`],
		},
	]
}

function mergeRangesInOriginalOrder(
	originalMerges: readonly RangeRef[],
	removedMerges: readonly RangeRef[],
): RangeRef[] {
	const removed = new Set(removedMerges.map(rangeCoordinateKey))
	return originalMerges.filter((merge) => removed.has(rangeCoordinateKey(merge))).map(cloneRange)
}

function uniqueRanges(ranges: readonly RangeRef[]): RangeRef[] {
	const seen = new Set<string>()
	const unique: RangeRef[] = []
	for (const range of ranges) {
		const key = rangeCoordinateKey(range)
		if (seen.has(key)) continue
		seen.add(key)
		unique.push(cloneRange(range))
	}
	return unique
}

function cloneRange(range: RangeRef): RangeRef {
	return { start: { ...range.start }, end: { ...range.end } }
}

function rangeCoordinateKey(range: RangeRef): string {
	return `${range.start.row}:${range.start.col}:${range.end.row}:${range.end.col}`
}

function copyRangeAllMetadataRestoration(
	workbook: Workbook,
	sourceSheetName: string,
	sourceRange: RangeRef,
	targetSheetName: string,
	targetRange: RangeRef,
): {
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const targetRefs = refsInParsedRange(targetRange)
	const targetComments = commentPreimages(workbook, targetSheetName, targetRefs)
	const targetHyperlinks = hyperlinkPreimages(workbook, targetSheetName, targetRefs)
	const validations = validationTransferRestoration(
		workbook,
		sourceSheetName,
		sourceRange,
		targetSheetName,
		targetRange,
		false,
	)
	return {
		inverseOps: [
			...restoreCommentOps(targetComments),
			...restoreHyperlinkOps(targetHyperlinks),
			...validations.inverseOps,
		],
		preimages: [
			...targetComments.map((comment) => ({ kind: 'comment' as const, comment })),
			...targetHyperlinks.map((hyperlink) => ({ kind: 'hyperlink' as const, hyperlink })),
			...validations.preimages,
		],
		issues: [
			...commentTransferIssues(
				workbook,
				sourceSheetName,
				rangeToA1(sourceRange),
				targetSheetName,
				targetRange,
			),
			...validations.issues,
		],
	}
}

function moveRangeAllMetadataRestoration(
	workbook: Workbook,
	sourceSheetName: string,
	sourceRange: RangeRef,
	targetSheetName: string,
	targetRange: RangeRef,
): {
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const sourceRefs = refsInParsedRange(sourceRange)
	const targetRefs = refsInParsedRange(targetRange)
	const sourceComments = commentPreimages(workbook, sourceSheetName, sourceRefs)
	const targetComments = commentPreimages(workbook, targetSheetName, targetRefs)
	const sourceHyperlinks = hyperlinkPreimages(workbook, sourceSheetName, sourceRefs)
	const targetHyperlinks = hyperlinkPreimages(workbook, targetSheetName, targetRefs)
	const validations = validationTransferRestoration(
		workbook,
		sourceSheetName,
		sourceRange,
		targetSheetName,
		targetRange,
		true,
	)
	return {
		inverseOps: [
			...moveRangeCommentRestoreOps(sourceComments, targetComments),
			...moveRangeHyperlinkRestoreOps(sourceHyperlinks, targetHyperlinks),
			...validations.inverseOps,
		],
		preimages: [
			...targetComments.map((comment) => ({ kind: 'comment' as const, comment })),
			...sourceComments.map((comment) => ({ kind: 'comment' as const, comment })),
			...targetHyperlinks.map((hyperlink) => ({ kind: 'hyperlink' as const, hyperlink })),
			...sourceHyperlinks.map((hyperlink) => ({ kind: 'hyperlink' as const, hyperlink })),
			...validations.preimages,
		],
		issues: [
			...commentTransferIssues(
				workbook,
				sourceSheetName,
				rangeToA1(sourceRange),
				targetSheetName,
				targetRange,
				{ move: true },
			),
			...validations.issues,
		],
	}
}

function moveRangeOverlapIssues(
	op: Extract<Operation, { op: 'moveRange' }>,
	sourceRange: RangeRef,
	targetRange: RangeRef,
): readonly MutationJournalIssue[] {
	if (op.targetSheet !== undefined && op.targetSheet !== op.sheet) return []
	if (!rangesOverlap(sourceRange, targetRange)) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Overlapping moveRange ${op.sheet}!${op.source} to ${op.target} cannot be fully restored with public operations`,
			surface: 'cells',
			reason: 'operation-unsupported',
			refs: [`${op.sheet}!${op.source}`, `${op.sheet}!${rangeToA1(targetRange)}`],
		},
	]
}

function moveRangeFormulaSurfaceRestoration(
	workbook: Workbook,
	op: Extract<Operation, { op: 'moveRange' }>,
	sourceRange: RangeRef,
	targetRange: RangeRef,
): {
	readonly inverseOps: readonly Operation[]
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const sourceSheetName = op.sheet
	const targetSheetName = op.targetSheet ?? op.sheet
	const cellRefsBySheet = new Map<string, string[]>()
	const dataValidationPreimages = new Map<string, MutationJournalDataValidationPreimage>()
	const conditionalFormatPreimages = new Map<string, MutationJournalConditionalFormatPreimage>()
	const hyperlinkLocationPreimages = new Map<string, MutationJournalHyperlinkPreimage>()
	const metadataRefs: string[] = []
	for (const sheet of workbook.sheets) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cell.formula) continue
			if (sheet.name === sourceSheetName && rangeContainsCell(sourceRange, { row, col })) continue
			if (sheet.name === targetSheetName && rangeContainsCell(targetRange, { row, col })) continue
			if (
				formulaReferencesMovedRange(
					workbook,
					cell.formula,
					sheet.name,
					sourceSheetName,
					sourceRange,
				)
			) {
				pushSheetRef(cellRefsBySheet, sheet.name, toA1({ row, col }))
			}
		}
		collectMovedRangeDataValidationFormulaRestorations(
			dataValidationPreimages,
			metadataRefs,
			workbook,
			sheet,
			sourceSheetName,
			sourceRange,
		)
		collectMovedRangeConditionalFormatFormulaRestorations(
			conditionalFormatPreimages,
			metadataRefs,
			workbook,
			sheet,
			sourceSheetName,
			sourceRange,
		)
		for (const [ref, hyperlink] of sheet.hyperlinks) {
			if (
				hyperlink.location &&
				formulaReferencesMovedRange(
					workbook,
					hyperlink.location,
					sheet.name,
					sourceSheetName,
					sourceRange,
				)
			) {
				hyperlinkLocationPreimages.set(
					`${sheet.name}!${ref}`,
					hyperlinkPreimage(workbook, sheet.name, ref),
				)
			}
		}
		pushMovedRangeMetadataFormulaRefs(metadataRefs, workbook, sheet, sourceSheetName, sourceRange)
	}
	const definedNames: MutationJournalDefinedNamePreimage[] = []
	for (const name of workbook.definedNames.list()) {
		const scopeSheet =
			name.scope.kind === 'sheet' ? sheetNameForId(workbook, name.scope.sheetId) : undefined
		if (
			formulaReferencesMovedRange(
				workbook,
				name.formula,
				scopeSheet ?? sourceSheetName,
				sourceSheetName,
				sourceRange,
			)
		) {
			definedNames.push(definedNamePreimageFromEntry(workbook, name))
		}
	}
	const cells = uniqueCellPreimages(
		[...cellRefsBySheet].flatMap(([sheet, refs]) => cellEditPreimages(workbook, sheet, refs)),
	)
	const { inverseOps: cellInverseOps, issues: cellIssues } = inverseCellOps(cells)
	const definedNameRestorations = definedNames.map((preimage) =>
		restoreDefinedNameOps(workbook, preimage),
	)
	const dataValidationRestorations = [...dataValidationPreimages.values()].map((preimage) =>
		restoreDataValidationOps(preimage),
	)
	const conditionalFormatRestorations = [...conditionalFormatPreimages.values()].map((preimage) =>
		restoreConditionalFormatOps(preimage.sheet, preimage.formats),
	)
	const metadataIssues = moveRangeFormulaRewriteIssues(op, metadataRefs)
	const preimages: MutationJournalPreimage[] = [
		...(cells.length > 0 ? [{ kind: 'cells' as const, cells }] : []),
		...definedNames.map((definedName) => ({ kind: 'defined-name' as const, definedName })),
		...(dataValidationPreimages.size > 0
			? [
					{
						kind: 'data-validations' as const,
						validations: [...dataValidationPreimages.values()],
					},
				]
			: []),
		...conditionalFormatPreimages.values().map((conditionalFormats) => ({
			kind: 'conditional-formats' as const,
			conditionalFormats,
		})),
		...hyperlinkLocationPreimages.values().map((hyperlink) => ({
			kind: 'hyperlink' as const,
			hyperlink,
		})),
	]
	if (preimages.length === 0 && metadataIssues.length === 0) return EMPTY_METADATA_RESTORATION
	return {
		inverseOps: [
			...cellInverseOps,
			...definedNameRestorations.flatMap((restoration) => restoration.inverseOps),
			...dataValidationRestorations.flatMap((restoration) => restoration.inverseOps),
			...conditionalFormatRestorations.flatMap((restoration) => restoration.inverseOps),
			...restoreHyperlinkOps([...hyperlinkLocationPreimages.values()]),
		],
		preimages,
		issues: [
			...cellIssues,
			...definedNameRestorations.flatMap((restoration) => restoration.issues),
			...dataValidationRestorations.flatMap((restoration) => restoration.issues),
			...conditionalFormatRestorations.flatMap((restoration) => restoration.issues),
			...metadataIssues,
		],
	}
}

function moveRangeFormulaRewriteIssues(
	op: Extract<Operation, { op: 'moveRange' }>,
	metadataRefs: readonly string[],
): readonly MutationJournalIssue[] {
	if (metadataRefs.length === 0) return []
	const x14Refs = metadataRefs.filter((ref) => ref.includes('!x14'))
	const formulaRefs = metadataRefs.filter((ref) => !ref.includes('!x14'))
	const message = `moveRange formula reference rewrites for ${op.sheet}!${op.source} cannot be fully restored with public operations`
	const issues: MutationJournalIssue[] = []
	if (formulaRefs.length > 0) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message,
			surface: 'formulas',
			reason: 'formula-reference-rewrite',
			refs: formulaRefs,
		})
	}
	if (x14Refs.length > 0) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message,
			surface: 'x14-metadata',
			reason: 'x14-metadata',
			refs: x14Refs,
		})
	}
	return issues
}

function pushSheetRef(refsBySheet: Map<string, string[]>, sheet: string, ref: string): void {
	const refs = refsBySheet.get(sheet)
	if (refs) {
		refs.push(ref)
		return
	}
	refsBySheet.set(sheet, [ref])
}

function definedNamePreimageFromEntry(
	workbook: Workbook,
	entry: DefinedName,
): MutationJournalDefinedNamePreimage {
	const scope =
		entry.scope.kind === 'sheet' ? sheetNameForId(workbook, entry.scope.sheetId) : undefined
	return {
		name: entry.name,
		...(scope !== undefined ? { scope } : {}),
		definedName: cloneDefinedName(entry),
	}
}

function collectMovedRangeDataValidationFormulaRestorations(
	preimages: Map<string, MutationJournalDataValidationPreimage>,
	lossyRefs: string[],
	workbook: Workbook,
	sheet: Sheet,
	sourceSheetName: string,
	sourceRange: RangeRef,
): void {
	for (const validation of sheet.dataValidations) {
		const refs: string[] = []
		if (
			validation.formula1 !== undefined &&
			formulaReferencesMovedRange(
				workbook,
				validation.formula1,
				sheet.name,
				sourceSheetName,
				sourceRange,
			)
		) {
			refs.push(`${sheet.name}!validation:${validation.sqref}:formula1`)
		}
		if (
			validation.formula2 !== undefined &&
			formulaReferencesMovedRange(
				workbook,
				validation.formula2,
				sheet.name,
				sourceSheetName,
				sourceRange,
			)
		) {
			refs.push(`${sheet.name}!validation:${validation.sqref}:formula2`)
		}
		if (refs.length === 0) continue
		if (dataValidationSqrefCount(sheet, validation.sqref) !== 1) {
			lossyRefs.push(...refs)
			continue
		}
		preimages.set(`${sheet.name}!${validation.sqref}`, {
			sheet: sheet.name,
			range: validation.sqref,
			validation: { ...validation },
		})
	}
}

function collectMovedRangeConditionalFormatFormulaRestorations(
	preimages: Map<string, MutationJournalConditionalFormatPreimage>,
	lossyRefs: string[],
	workbook: Workbook,
	sheet: Sheet,
	sourceSheetName: string,
	sourceRange: RangeRef,
): void {
	sheet.conditionalFormats.forEach((format, formatIndex) => {
		const refs: string[] = []
		format.rules.forEach((rule, ruleIndex) => {
			rule.formulas.forEach((formula, formulaIndex) => {
				pushMovedRangeFormulaRef(
					refs,
					workbook,
					formula,
					sheet.name,
					sourceSheetName,
					sourceRange,
					`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:${formulaIndex}`,
				)
			})
			pushMovedRangeValueObjectFormulaRefs(
				refs,
				workbook,
				rule.colorScale?.cfvo,
				sheet.name,
				sourceSheetName,
				sourceRange,
				`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:colorScale.cfvo`,
			)
			pushMovedRangeValueObjectFormulaRefs(
				refs,
				workbook,
				rule.dataBar?.cfvo,
				sheet.name,
				sourceSheetName,
				sourceRange,
				`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:dataBar.cfvo`,
			)
			pushMovedRangeValueObjectFormulaRefs(
				refs,
				workbook,
				rule.iconSet?.cfvo,
				sheet.name,
				sourceSheetName,
				sourceRange,
				`${sheet.name}!conditionalFormat:${format.sqref}:${formatIndex}:${ruleIndex}:iconSet.cfvo`,
			)
		})
		if (refs.length === 0) return
		if (conditionalFormatSqrefCount(sheet, format.sqref) !== 1) {
			lossyRefs.push(...refs)
			return
		}
		preimages.set(
			`${sheet.name}!${format.sqref}`,
			conditionalFormatPreimage(workbook, sheet.name, format.sqref),
		)
	})
}

function pushMovedRangeMetadataFormulaRefs(
	refs: string[],
	workbook: Workbook,
	sheet: Sheet,
	sourceSheetName: string,
	sourceRange: RangeRef,
): void {
	for (const validation of sheet.x14DataValidations) {
		if (validation.deleted) continue
		pushMovedRangeFormulaRef(
			refs,
			workbook,
			validation.formula1,
			sheet.name,
			sourceSheetName,
			sourceRange,
			`${sheet.name}!x14Validation:${validation.sqref}:formula1`,
		)
		pushMovedRangeFormulaRef(
			refs,
			workbook,
			validation.formula2,
			sheet.name,
			sourceSheetName,
			sourceRange,
			`${sheet.name}!x14Validation:${validation.sqref}:formula2`,
		)
	}
	sheet.x14ConditionalFormats.forEach((format) => {
		if (format.deleted) return
		format.formulas.forEach((formula, formulaIndex) => {
			pushMovedRangeFormulaRef(
				refs,
				workbook,
				formula,
				sheet.name,
				sourceSheetName,
				sourceRange,
				`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:${formulaIndex}`,
			)
		})
		pushMovedRangeValueObjectFormulaRefs(
			refs,
			workbook,
			format.colorScale?.cfvo,
			sheet.name,
			sourceSheetName,
			sourceRange,
			`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:colorScale.cfvo`,
		)
		pushMovedRangeValueObjectFormulaRefs(
			refs,
			workbook,
			format.dataBar?.cfvo,
			sheet.name,
			sourceSheetName,
			sourceRange,
			`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:dataBar.cfvo`,
		)
		pushMovedRangeValueObjectFormulaRefs(
			refs,
			workbook,
			format.iconSet?.cfvo,
			sheet.name,
			sourceSheetName,
			sourceRange,
			`${sheet.name}!x14ConditionalFormat:${format.sqref}:${format.index}:iconSet.cfvo`,
		)
	})
	for (const table of sheet.tables) {
		for (const column of table.columns) {
			pushMovedRangeFormulaRef(
				refs,
				workbook,
				column.formula,
				sheet.name,
				sourceSheetName,
				sourceRange,
				`${sheet.name}!table:${table.name}:${column.name}:formula`,
			)
			pushMovedRangeFormulaRef(
				refs,
				workbook,
				column.totalsRowFormula,
				sheet.name,
				sourceSheetName,
				sourceRange,
				`${sheet.name}!table:${table.name}:${column.name}:totalsRowFormula`,
			)
		}
	}
}

function pushMovedRangeValueObjectFormulaRefs(
	refs: string[],
	workbook: Workbook,
	entries: readonly SheetConditionalFormatValueObject[] | undefined,
	ownerSheet: string,
	sourceSheetName: string,
	sourceRange: RangeRef,
	location: string,
): void {
	entries?.forEach((entry, index) => {
		pushMovedRangeFormulaRef(
			refs,
			workbook,
			entry.value,
			ownerSheet,
			sourceSheetName,
			sourceRange,
			`${location}:${index}`,
		)
	})
}

function pushMovedRangeFormulaRef(
	refs: string[],
	workbook: Workbook,
	formula: string | undefined,
	ownerSheet: string,
	sourceSheetName: string,
	sourceRange: RangeRef,
	location: string,
): void {
	if (formula === undefined) return
	if (formulaReferencesMovedRange(workbook, formula, ownerSheet, sourceSheetName, sourceRange)) {
		refs.push(location)
	}
}

function formulaReferencesMovedRange(
	workbook: Workbook,
	formula: string,
	ownerSheet: string,
	sourceSheetName: string,
	sourceRange: RangeRef,
): boolean {
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return false
	return extractRefs(parsed.value).some((ref) =>
		formulaRefReferencesMovedRange(workbook, ref, ownerSheet, sourceSheetName, sourceRange),
	)
}

function formulaRefReferencesMovedRange(
	workbook: Workbook,
	ref: FormulaRef,
	ownerSheet: string,
	sourceSheetName: string,
	sourceRange: RangeRef,
): boolean {
	if (ref.kind === 'sheetSpan') {
		return (
			sheetSpanIncludes(workbook, ref.startSheet, ref.endSheet, sourceSheetName) &&
			formulaRefReferencesMovedRange(
				workbook,
				ref.target,
				sourceSheetName,
				sourceSheetName,
				sourceRange,
			)
		)
	}
	const sheet = 'sheet' in ref && ref.sheet !== undefined ? ref.sheet : ownerSheet
	if (!sameSheetName(sheet, sourceSheetName)) return false
	switch (ref.kind) {
		case 'cell':
			return rangeContainsCell(sourceRange, ref.ref)
		case 'range':
			return rangeContainsRange(sourceRange, { start: ref.start, end: ref.end })
		case 'wholeRowRange':
			return ref.startRow >= sourceRange.start.row && ref.endRow <= sourceRange.end.row
		case 'wholeColumnRange':
			return ref.startCol >= sourceRange.start.col && ref.endCol <= sourceRange.end.col
		default:
			return false
	}
}

function transferTargetRange(sourceRangeText: string, targetRef: string): RangeRef {
	const source = parseRange(sourceRangeText)
	const target = parseA1(targetRef)
	return {
		start: target,
		end: {
			row: target.row + source.end.row - source.start.row,
			col: target.col + source.end.col - source.start.col,
		},
	}
}

function refsInParsedRange(range: RangeRef): string[] {
	const refs: string[] = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			refs.push(toA1({ row, col }))
		}
	}
	return refs
}

function rangeToA1(range: RangeRef): string {
	const start = toA1(range.start)
	const end = toA1(range.end)
	return start === end ? start : `${start}:${end}`
}

function cellEditPreimages(
	workbook: Workbook,
	sheetName: string,
	refs: readonly string[],
	options: { readonly blockedSpillBlockers?: boolean } = {},
): MutationJournalCellPreimage[] {
	return cellPreimages(
		workbook,
		sheetName,
		formulaBindingEditRefs(workbook, sheetName, refs, options),
	)
}

function formulaBindingJournalAddendum(cells: readonly MutationJournalCellPreimage[]): {
	readonly preimages: readonly MutationJournalPreimage[]
	readonly issues: readonly MutationJournalIssue[]
} {
	if (cells.length === 0) return { preimages: [], issues: [] }
	const { issues } = inverseCellOps(cells)
	return { preimages: [{ kind: 'cells', cells }], issues }
}

function workbookFormulaBindingCellPreimages(workbook: Workbook): MutationJournalCellPreimage[] {
	const cells: MutationJournalCellPreimage[] = []
	for (const sheet of workbook.sheets) {
		if (sheet.cells.formulaInfoCellCount() === 0) continue
		const refs: string[] = []
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (isMaterializedFormulaBindingInfo(cell.formulaInfo)) refs.push(toA1({ row, col }))
		}
		cells.push(...cellPreimages(workbook, sheet.name, refs))
	}
	return cells
}

function uniqueCellPreimages(
	cells: readonly MutationJournalCellPreimage[],
): MutationJournalCellPreimage[] {
	const unique = new Map<string, MutationJournalCellPreimage>()
	for (const cell of cells) unique.set(`${cell.sheet}!${cell.ref}`, cell)
	return [...unique.values()]
}

function tableColumnFormulaBindingCellPreimages(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setTableColumn' }>,
): MutationJournalCellPreimage[] {
	if (op.newName !== undefined) return workbookFormulaBindingCellPreimages(workbook)
	const located = findTable(workbook, op.table)
	if (!located) return []
	const columnIndex = tableColumnIndex(located.table, op.column)
	if (columnIndex < 0) return []
	const col = located.table.ref.start.col + columnIndex
	const refs = new Set<string>()
	if (op.formula !== undefined) {
		const bodyStart = located.table.ref.start.row + (located.table.hasHeaders ? 1 : 0)
		const bodyEnd = located.table.ref.end.row - (located.table.hasTotals ? 1 : 0)
		for (let row = bodyStart; row <= bodyEnd; row++) refs.add(toA1({ row, col }))
	}
	if (
		located.table.hasTotals &&
		(op.totalsRowFunction !== undefined ||
			op.totalsRowFormula !== undefined ||
			op.totalsRowLabel !== undefined)
	) {
		refs.add(toA1({ row: located.table.ref.end.row, col }))
	}
	if (refs.size === 0) return []
	return formulaBindingOnlyPreimages(workbook, located.sheet.name, [...refs])
}

function tableColumnIndex(table: Table, columnSelector: string | number): number {
	return typeof columnSelector === 'number'
		? columnSelector
		: table.columns.findIndex(
				(column) => column.name.toLowerCase() === columnSelector.toLowerCase(),
			)
}

function formulaBindingOnlyPreimages(
	workbook: Workbook,
	sheetName: string,
	refs: readonly string[],
): MutationJournalCellPreimage[] {
	return cellEditPreimages(workbook, sheetName, refs).filter((cell) =>
		isMaterializedFormulaBindingInfo(cell.formulaInfo),
	)
}

function isMaterializedFormulaBindingInfo(
	formulaInfo: Cell['formulaInfo'],
): formulaInfo is Exclude<NonNullable<Cell['formulaInfo']>, { kind: 'array' }> {
	return formulaInfo !== undefined && formulaInfo.kind !== 'array'
}

function formulaBindingEditRefs(
	workbook: Workbook,
	sheetName: string,
	refs: readonly string[],
	options: { readonly blockedSpillBlockers?: boolean } = {},
): string[] {
	const sheet = workbook.getSheet(sheetName)
	const parsedRefs = refs.map((ref) => parseA1(ref))
	const expanded = new Map<string, { readonly row: number; readonly col: number }>()
	const push = (row: number, col: number) => expanded.set(toA1({ row, col }), { row, col })
	for (const ref of parsedRefs) push(ref.row, ref.col)
	if (!sheet || sheet.cells.formulaInfoCellCount() === 0) return [...expanded.keys()]

	for (const ref of parsedRefs) {
		const binding = sheet.cells.get(ref.row, ref.col)?.formulaInfo
		if (!binding) continue
		if (binding.kind === 'shared') {
			for (const [row, col, cell] of sheet.cells.iterate()) {
				if (sameSharedFormulaBinding(binding, cell.formulaInfo)) push(row, col)
			}
			continue
		}
		if (isSpillFormulaBinding(binding)) {
			const dynamicAnchorRef =
				binding.kind === 'dynamicArray'
					? `${formatSheetNameForFormula(sheet.name)}!${toA1(ref)}`
					: undefined
			for (const [row, col, cell] of sheet.cells.iterate()) {
				const candidateRef = `${formatSheetNameForFormula(sheet.name)}!${toA1({ row, col })}`
				if (sameSpillFormulaBinding(binding, cell.formulaInfo, dynamicAnchorRef, candidateRef)) {
					push(row, col)
				}
			}
		}
	}

	for (const [row, col, cell] of sheet.cells.iterate()) {
		const binding = cell.formulaInfo
		if (binding?.kind !== 'dataTable') continue
		const tableRange = dataTableFormulaRange(binding, row, col)
		if (parsedRefs.some((ref) => rangeContainsCell(tableRange, ref))) push(row, col)
	}
	if (options.blockedSpillBlockers !== false) {
		for (const [row, col, cell] of sheet.cells.iterate()) {
			const binding = cell.formulaInfo
			if (binding?.kind !== 'blockedSpill') continue
			if (
				parsedRefs.some(
					(ref) =>
						formulaBindingRangeContainsCell(binding.ref, sheet.name, ref) ||
						binding.blockingRefs.some((blockingRef) =>
							formulaBindingRangeContainsCell(blockingRef, sheet.name, ref),
						),
				)
			) {
				push(row, col)
			}
		}
	}
	return [...expanded.keys()]
}

function formatSheetNameForFormula(sheet: string): string {
	if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet)) return sheet
	return `'${sheet.replace(/'/g, "''")}'`
}

function sameSharedFormulaBinding(
	binding: Extract<NonNullable<Cell['formulaInfo']>, { kind: 'shared' }>,
	candidate: Cell['formulaInfo'],
): boolean {
	if (candidate?.kind !== 'shared') return false
	if (binding.sharedIndex !== candidate.sharedIndex) return false
	if (binding.masterRef !== undefined && candidate.masterRef !== undefined) {
		return sameFormulaCellRef(binding.masterRef, candidate.masterRef)
	}
	return true
}

function isSpillFormulaBinding(
	binding: NonNullable<Cell['formulaInfo']>,
): binding is Extract<
	NonNullable<Cell['formulaInfo']>,
	{ kind: 'dynamicArray' | 'spill' | 'blockedSpill' }
> {
	return (
		binding.kind === 'dynamicArray' || binding.kind === 'spill' || binding.kind === 'blockedSpill'
	)
}

function sameSpillFormulaBinding(
	binding: Extract<
		NonNullable<Cell['formulaInfo']>,
		{ kind: 'dynamicArray' | 'spill' | 'blockedSpill' }
	>,
	candidate: Cell['formulaInfo'],
	dynamicAnchorRef?: string,
	candidateRef?: string,
): boolean {
	if (!candidate) return false
	if (binding.kind === 'dynamicArray') {
		if (candidate.kind === 'dynamicArray') {
			if (candidateRef !== undefined && dynamicAnchorRef !== undefined) {
				return sameFormulaCellRef(candidateRef, dynamicAnchorRef)
			}
			return candidate.metadataIndex === binding.metadataIndex
		}
		return (
			dynamicAnchorRef !== undefined &&
			(candidate.kind === 'spill' || candidate.kind === 'blockedSpill') &&
			sameFormulaCellRef(candidate.anchorRef, dynamicAnchorRef)
		)
	}
	if (candidate.kind === 'dynamicArray') {
		return candidateRef !== undefined && sameFormulaCellRef(candidateRef, binding.anchorRef)
	}
	if (candidate.kind !== 'spill' && candidate.kind !== 'blockedSpill') return false
	return sameFormulaCellRef(candidate.anchorRef, binding.anchorRef)
}

function sameFormulaCellRef(left: string, right: string): boolean {
	if (left === right) return true
	try {
		const leftRange = parseRange(left)
		const rightRange = parseRange(right)
		return (
			sameOptionalSheetName(leftRange.sheet, rightRange.sheet) &&
			leftRange.start.row === rightRange.start.row &&
			leftRange.start.col === rightRange.start.col &&
			leftRange.end.row === rightRange.end.row &&
			leftRange.end.col === rightRange.end.col
		)
	} catch {
		return false
	}
}

function sameOptionalSheetName(left: string | undefined, right: string | undefined): boolean {
	if (left === undefined || right === undefined) return left === right
	return left.toLowerCase() === right.toLowerCase()
}

function sameSheetName(left: string, right: string): boolean {
	return left.toLowerCase() === right.toLowerCase()
}

function dataTableFormulaRange(
	binding: Extract<NonNullable<Cell['formulaInfo']>, { kind: 'dataTable' }>,
	row: number,
	col: number,
): RangeRef {
	if (binding.ref) {
		try {
			return parseRange(binding.ref)
		} catch {
			// Fall back to the formula cell when imported metadata carries a malformed range.
		}
	}
	return { start: { row, col }, end: { row, col } }
}

function formulaBindingRangeContainsCell(
	refText: string,
	sheetName: string,
	ref: { readonly row: number; readonly col: number },
): boolean {
	try {
		const range = parseRange(refText)
		if (range.sheet !== undefined && !sameOptionalSheetName(range.sheet, sheetName)) return false
		return rangeContainsCell(range, ref)
	} catch {
		return false
	}
}

function rangeContainsCell(
	range: RangeRef,
	ref: { readonly row: number; readonly col: number },
): boolean {
	return (
		ref.row >= range.start.row &&
		ref.row <= range.end.row &&
		ref.col >= range.start.col &&
		ref.col <= range.end.col
	)
}

function cellPreimages(
	workbook: Workbook,
	sheetName: string,
	refs: readonly string[],
): MutationJournalCellPreimage[] {
	const sheet = workbook.getSheet(sheetName)
	return refs.map((ref) => {
		const parsed = parseA1(ref)
		const existing = sheet?.cells.get(parsed.row, parsed.col)
		const styleId = existing?.styleId ?? DEFAULT_STYLE_ID
		const style = cloneCellStyle(workbook.styles.get(styleId) ?? {})
		return {
			sheet: sheetName,
			ref: toA1(parsed),
			dateSystem: workbook.calcSettings.dateSystem,
			existed: existing !== undefined,
			value: cloneCellValue(existing?.value ?? EMPTY),
			formula: existing?.formula ?? null,
			...(existing?.formulaInfo ? { formulaInfo: cloneFormulaInfo(existing.formulaInfo) } : {}),
			styleId,
			style,
		}
	})
}

function formulaBindingIssueSurface(
	formulaInfo: NonNullable<Cell['formulaInfo']>,
): MutationJournalSurface {
	switch (formulaInfo.kind) {
		case 'shared':
			return 'shared-formulas'
		case 'array':
			return 'legacy-arrays'
		case 'dynamicArray':
			return 'dynamic-arrays'
		case 'dataTable':
			return 'data-tables'
		case 'spill':
		case 'blockedSpill':
			return 'spills'
	}
}

function inverseCellOps(cells: readonly MutationJournalCellPreimage[]): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const inverseOps: Operation[] = []
	const issues: MutationJournalIssue[] = []
	const scalarUpdatesBySheet = new Map<string, Array<{ ref: string; value: InputValue }>>()
	for (const cell of cells) {
		if (cell.formulaInfo) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Formula binding metadata for ${cell.sheet}!${cell.ref} cannot be restored with public operations`,
				surface: formulaBindingIssueSurface(cell.formulaInfo),
				reason: 'formula-binding-metadata',
				refs: [`${cell.sheet}!${cell.ref}`],
			})
		}
		if (!cell.existed) {
			inverseOps.push({ op: 'clearRange', sheet: cell.sheet, range: cell.ref, what: 'all' })
			continue
		}
		if (cell.formula) {
			const cacheRestore = formulaCacheRestoreOps(cell)
			inverseOps.push(...cacheRestore.inverseOps)
			issues.push(...cacheRestore.issues)
			inverseOps.push({
				op: 'setFormula',
				sheet: cell.sheet,
				ref: cell.ref,
				formula: cell.formula,
			})
			continue
		}
		if (cell.value.kind === 'richText') {
			const runs = richTextRunsToOperationRuns(cell.value.runs)
			if (runs.supported) {
				inverseOps.push({ op: 'setRichText', sheet: cell.sheet, ref: cell.ref, runs: runs.runs })
				continue
			}
			issues.push({
				code: 'UNSUPPORTED_VALUE',
				message: `Cannot restore richText at ${cell.sheet}!${cell.ref} with setRichText`,
				surface: 'cells',
				reason: 'rich-text-unsupported-runs',
				refs: [`${cell.sheet}!${cell.ref}`],
			})
			continue
		}
		const input = cellValueToInput(cell.value, cell.dateSystem ?? '1900')
		if (input.supported) {
			const updates = scalarUpdatesBySheet.get(cell.sheet) ?? []
			updates.push({ ref: cell.ref, value: input.value })
			scalarUpdatesBySheet.set(cell.sheet, updates)
			continue
		}
		issues.push({
			code: 'UNSUPPORTED_VALUE',
			message: `Cannot restore ${cell.value.kind} at ${cell.sheet}!${cell.ref} with setCells`,
			surface: 'cells',
			reason: 'value-unsupported',
			refs: [`${cell.sheet}!${cell.ref}`],
		})
	}
	for (const [sheet, updates] of scalarUpdatesBySheet) {
		inverseOps.push({ op: 'setCells', sheet, updates })
	}
	return { inverseOps, issues }
}

function formulaCacheRestoreOps(cell: MutationJournalCellPreimage): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	const input = cellValueToInput(cell.value, cell.dateSystem ?? '1900')
	if (input.supported) {
		return {
			inverseOps: [
				{
					op: 'setCells',
					sheet: cell.sheet,
					updates: [{ ref: cell.ref, value: input.value }],
				},
			],
			issues: [],
		}
	}
	if (cell.value.kind === 'richText') {
		const runs = richTextRunsToOperationRuns(cell.value.runs)
		if (runs.supported) {
			return {
				inverseOps: [{ op: 'setRichText', sheet: cell.sheet, ref: cell.ref, runs: runs.runs }],
				issues: [],
			}
		}
	}
	return {
		inverseOps: [],
		issues: [
			{
				code: 'LOSSY_INVERSE',
				message: `Formula cache for ${cell.sheet}!${cell.ref} cannot be restored with public operations`,
				surface: 'formulas',
				reason: 'formula-cache-unsupported-value',
				refs: [`${cell.sheet}!${cell.ref}`],
			},
		],
	}
}

function richTextRunsToOperationRuns(runs: Extract<CellValue, { kind: 'richText' }>['runs']):
	| {
			readonly supported: true
			readonly runs: Extract<Operation, { op: 'setRichText' }>['runs']
	  }
	| { readonly supported: false } {
	const operationRuns: Array<Extract<Operation, { op: 'setRichText' }>['runs'][number]> = []
	for (const run of runs) {
		if (
			run.strikethrough !== undefined ||
			run.fontName !== undefined ||
			run.fontSize !== undefined ||
			(run.color !== undefined && typeof run.color !== 'string')
		) {
			return { supported: false }
		}
		operationRuns.push({
			text: run.text,
			...(run.bold !== undefined ? { bold: run.bold } : {}),
			...(run.italic !== undefined ? { italic: run.italic } : {}),
			...(run.underline !== undefined ? { underline: run.underline } : {}),
			...(run.color !== undefined ? { color: run.color } : {}),
		})
	}
	return { supported: true, runs: operationRuns }
}

function styleInverseOps(cells: readonly MutationJournalCellPreimage[]): Operation[] {
	const inverseOps: Operation[] = []
	for (const cell of cells) {
		if (!cell.existed) {
			inverseOps.push({ op: 'clearRange', sheet: cell.sheet, range: cell.ref, what: 'all' })
			continue
		}
		inverseOps.push({ op: 'clearRange', sheet: cell.sheet, range: cell.ref, what: 'styles' })
		if (Object.keys(cell.style).length > 0) {
			inverseOps.push({
				op: 'setStyle',
				sheet: cell.sheet,
				range: cell.ref,
				style: cell.style,
			})
		}
	}
	return inverseOps
}

function cellValueToInput(
	value: CellValue,
	dateSystem: '1900' | '1904' = '1900',
): { readonly supported: true; readonly value: InputValue } | { readonly supported: false } {
	switch (value.kind) {
		case 'empty':
			return { supported: true, value: null }
		case 'number':
		case 'string':
		case 'boolean':
			return { supported: true, value: value.value }
		case 'date':
			return dateCellValueToInput(value.serial, dateSystem)
		default:
			return { supported: false }
	}
}

function dateCellValueToInput(
	serial: number,
	dateSystem: '1900' | '1904',
): { readonly supported: true; readonly value: Date } | { readonly supported: false } {
	if (!Number.isInteger(serial)) return { supported: false }
	const parts = serialToDate(serial, dateSystem)
	if (!parts) return { supported: false }
	const date = new Date(0)
	date.setFullYear(parts.year, parts.month - 1, parts.day)
	date.setHours(0, 0, 0, 0)
	if (
		dateToSerial(date.getFullYear(), date.getMonth() + 1, date.getDate(), dateSystem) !== serial
	) {
		return { supported: false }
	}
	return { supported: true, value: date }
}

function refsInRange(rangeText: string): string[] {
	const range = parseRange(rangeText)
	const refs: string[] = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			refs.push(toA1({ row, col }))
		}
	}
	return refs
}

function sameRange(a: RangeRef, b: RangeRef): boolean {
	return (
		a.start.row === b.start.row &&
		a.start.col === b.start.col &&
		a.end.row === b.end.row &&
		a.end.col === b.end.col
	)
}

function sheetScopeForName(
	workbook: Workbook,
	sheetName: string,
): Extract<DefinedNameScope, { kind: 'sheet' }> | null {
	const sheet = workbook.getSheet(sheetName)
	return sheet ? { kind: 'sheet', sheetId: sheet.id } : null
}

function sheetNameForId(workbook: Workbook, sheetId: string): string | undefined {
	return workbook.sheets.find((sheet) => sheet.id === sheetId)?.name
}

function findTable(
	workbook: Workbook,
	tableName: string,
): { readonly sheet: Sheet; readonly table: Table } | null {
	return findTableMatches(workbook, tableName)[0] ?? null
}

function findTableMatches(
	workbook: Workbook,
	tableName: string,
): readonly { readonly sheet: Sheet; readonly table: Table }[] {
	const matches: { sheet: Sheet; table: Table }[] = []
	for (const sheet of workbook.sheets) {
		for (const table of sheet.tables) {
			if (table.name.toLowerCase() === tableName.toLowerCase()) matches.push({ sheet, table })
		}
	}
	return matches
}

function findTableColumn(
	workbook: Workbook,
	tableName: string,
	columnSelector: string | number,
): { readonly columnIndex: number; readonly column: TableColumn } | null {
	const located = findTable(workbook, tableName)
	if (!located) return null
	const columnIndex =
		typeof columnSelector === 'number'
			? columnSelector
			: located.table.columns.findIndex(
					(column) => column.name.toLowerCase() === columnSelector.toLowerCase(),
				)
	const column = located.table.columns[columnIndex]
	return column ? { columnIndex, column } : null
}

function commentPreimage(
	workbook: Workbook,
	sheetName: string,
	refText: string,
): MutationJournalCommentPreimage {
	const ref = refText.toUpperCase()
	const comment = findComment(workbook.getSheet(sheetName), ref)
	return {
		sheet: sheetName,
		ref,
		comment: comment ? { ...comment } : null,
	}
}

function commentPreimages(
	workbook: Workbook,
	sheetName: string,
	refs: readonly string[],
): MutationJournalCommentPreimage[] {
	return refs.map((ref) => commentPreimage(workbook, sheetName, ref))
}

function restoreCommentOps(comments: readonly MutationJournalCommentPreimage[]): Operation[] {
	return comments.map((comment) =>
		comment.comment
			? {
					op: 'setComment',
					sheet: comment.sheet,
					ref: comment.ref,
					text: comment.comment.text,
					...(comment.comment.author !== undefined ? { author: comment.comment.author } : {}),
				}
			: { op: 'deleteComment', sheet: comment.sheet, ref: comment.ref },
	)
}

function moveRangeCommentRestoreOps(
	sourceComments: readonly MutationJournalCommentPreimage[],
	targetComments: readonly MutationJournalCommentPreimage[],
): Operation[] {
	const targetCleanupOps = targetComments
		.filter((_, index) => sourceComments[index]?.comment)
		.map((comment) => ({ op: 'deleteComment' as const, sheet: comment.sheet, ref: comment.ref }))
	return [
		...targetCleanupOps,
		...restoreCommentOps(sourceComments),
		...restoreCommentOps(targetComments),
	]
}

function restoreSetCommentOps(
	op: Extract<Operation, { op: 'setComment' }>,
	comment: MutationJournalCommentPreimage,
	threadedComments: readonly MutationJournalThreadedCommentPreimage[],
): {
	readonly inverseOps: readonly Operation[]
	readonly issues: readonly MutationJournalIssue[]
} {
	if (!comment.comment) {
		return {
			inverseOps: [{ op: 'deleteComment', sheet: op.sheet, ref: comment.ref }],
			issues:
				threadedComments.length > 0
					? [
							{
								code: 'LOSSY_INVERSE',
								message: `Created legacy comment at ${op.sheet}!${comment.ref} cannot be removed without deleting threaded comments with public operations`,
								surface: 'comments',
								reason: 'threaded-comment-selector',
								refs: threadedCommentIssueRefs(op.sheet, threadedComments),
							},
						]
					: [],
		}
	}
	if (comment.comment.author === undefined && op.author !== undefined) {
		const baseInverse: Operation = {
			op: 'setComment',
			sheet: op.sheet,
			ref: comment.ref,
			text: comment.comment.text,
		}
		if (!comment.comment.legacyDrawing && threadedComments.length === 0) {
			return {
				inverseOps: [{ op: 'deleteComment', sheet: op.sheet, ref: comment.ref }, baseInverse],
				issues: [],
			}
		}
		return {
			inverseOps: [baseInverse],
			issues: [
				{
					code: 'LOSSY_INVERSE',
					message: `Comment author at ${op.sheet}!${comment.ref} cannot be removed exactly with public operations`,
					surface: 'comments',
					reason: 'comment-author-removal',
					refs: [
						`${op.sheet}!${comment.ref}`,
						...threadedCommentIssueRefs(op.sheet, threadedComments),
					],
				},
			],
		}
	}
	return {
		inverseOps: [
			{
				op: 'setComment',
				sheet: op.sheet,
				ref: comment.ref,
				text: comment.comment.text,
				...(comment.comment.author !== undefined ? { author: comment.comment.author } : {}),
			},
		],
		issues: [],
	}
}

function commentRestoreIssues(
	comments: readonly MutationJournalCommentPreimage[],
): readonly MutationJournalIssue[] {
	return comments
		.filter((preimage) => preimage.comment?.legacyDrawing)
		.map((preimage) => ({
			code: 'LOSSY_INVERSE' as const,
			message: `Legacy comment drawing metadata for ${preimage.sheet}!${preimage.ref} cannot be restored with public operations`,
			surface: 'comments' as const,
			reason: 'legacy-comment-drawing' as const,
			refs: [`${preimage.sheet}!${preimage.ref}`],
		}))
}

function commentTransferIssues(
	workbook: Workbook,
	sourceSheetName: string,
	sourceRangeText: string,
	targetSheetName: string,
	targetRange: RangeRef,
	options: { readonly move?: boolean } = {},
): MutationJournalIssue[] {
	const sourceSheet = workbook.getSheet(sourceSheetName)
	const targetSheet = workbook.getSheet(targetSheetName)
	const sourceRange = parseRange(sourceRangeText)
	const refs = new Set<string>()
	const issues: MutationJournalIssue[] = []

	if (threadedCommentsOverlap(sourceSheet, sourceRange))
		refs.add(`${sourceSheetName}!${sourceRangeText}`)
	if (threadedCommentsOverlap(targetSheet, targetRange)) {
		refs.add(`${targetSheetName}!${rangeToA1(targetRange)}`)
	}
	if (refs.size > 0) {
		issues.push({
			code: 'LOSSY_INVERSE',
			message: 'Threaded comment range transfers cannot be fully restored with public operations',
			surface: 'comments',
			reason: 'threaded-comment-selector',
			refs: [...refs],
		})
	}

	for (const preimage of commentPreimages(
		workbook,
		targetSheetName,
		refsInParsedRange(targetRange),
	)) {
		if (preimage.comment?.legacyDrawing) {
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Legacy comment drawing metadata for ${preimage.sheet}!${preimage.ref} cannot be restored with public operations`,
				surface: 'comments',
				reason: 'legacy-comment-drawing',
				refs: [`${preimage.sheet}!${preimage.ref}`],
			})
		}
	}
	if (options.move) {
		for (const preimage of commentPreimages(
			workbook,
			sourceSheetName,
			refsInParsedRange(sourceRange),
		)) {
			if (preimage.comment?.legacyDrawing) {
				issues.push({
					code: 'LOSSY_INVERSE',
					message: `Legacy comment drawing metadata for ${preimage.sheet}!${preimage.ref} cannot be restored with public operations`,
					surface: 'comments',
					reason: 'legacy-comment-drawing',
					refs: [`${preimage.sheet}!${preimage.ref}`],
				})
			}
		}
	}

	if (sourceSheet && targetSheet) {
		const rowDelta = targetRange.start.row - sourceRange.start.row
		const colDelta = targetRange.start.col - sourceRange.start.col
		for (const [ref, comment] of sourceSheet.comments) {
			const sourceRef = parseA1(ref)
			if (!rangeContainsCell(sourceRange, sourceRef) || !comment.legacyDrawing) continue
			const targetRef = toA1({ row: sourceRef.row + rowDelta, col: sourceRef.col + colDelta })
			const targetComment = findComment(targetSheet, targetRef)
			if (!targetComment) continue
			issues.push({
				code: 'LOSSY_INVERSE',
				message: `Copied legacy comment drawing metadata for ${targetSheetName}!${targetRef} cannot be removed with public operations`,
				surface: 'comments',
				reason: 'legacy-comment-drawing',
				refs: [`${sourceSheetName}!${toA1(sourceRef)}`, `${targetSheetName}!${targetRef}`],
			})
		}
	}
	return issues
}

function threadedCommentsOverlap(sheet: Sheet | undefined, range: RangeRef): boolean {
	if (!sheet) return false
	return sheet.threadedComments.some((comment) => rangeContainsCell(range, parseA1(comment.ref)))
}

function findComment(sheet: Sheet | undefined, ref: string): SheetComment | null {
	if (!sheet) return null
	for (const [commentRef, comment] of sheet.comments) {
		if (commentRef.toUpperCase() === ref) return comment
	}
	return null
}

function hyperlinkPreimage(
	workbook: Workbook,
	sheetName: string,
	refText: string,
): MutationJournalHyperlinkPreimage {
	const ref = refText.toUpperCase()
	const hyperlink = findHyperlink(workbook.getSheet(sheetName), ref)
	return {
		sheet: sheetName,
		ref,
		hyperlink: hyperlink ? { ...hyperlink } : null,
	}
}

function hyperlinkPreimages(
	workbook: Workbook,
	sheetName: string,
	refs: readonly string[],
): MutationJournalHyperlinkPreimage[] {
	return refs.map((ref) => hyperlinkPreimage(workbook, sheetName, ref))
}

function restoreHyperlinkOps(hyperlinks: readonly MutationJournalHyperlinkPreimage[]): Operation[] {
	return hyperlinks.map((hyperlink) =>
		hyperlink.hyperlink
			? setHyperlinkInverse(hyperlink.sheet, hyperlink.ref, hyperlink.hyperlink)
			: { op: 'deleteHyperlink', sheet: hyperlink.sheet, ref: hyperlink.ref },
	)
}

function moveRangeHyperlinkRestoreOps(
	sourceHyperlinks: readonly MutationJournalHyperlinkPreimage[],
	targetHyperlinks: readonly MutationJournalHyperlinkPreimage[],
): Operation[] {
	const targetCleanupOps = targetHyperlinks
		.filter((_, index) => sourceHyperlinks[index]?.hyperlink)
		.map((hyperlink) => ({
			op: 'deleteHyperlink' as const,
			sheet: hyperlink.sheet,
			ref: hyperlink.ref,
		}))
	return [
		...targetCleanupOps,
		...restoreHyperlinkOps(sourceHyperlinks),
		...restoreHyperlinkOps(targetHyperlinks),
	]
}

function findHyperlink(sheet: Sheet | undefined, ref: string): SheetHyperlink | null {
	if (!sheet) return null
	for (const [linkRef, hyperlink] of sheet.hyperlinks) {
		if (linkRef.toUpperCase() === ref) return hyperlink
	}
	return null
}

function threadedCommentPreimage(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setThreadedComment' }>,
): MutationJournalThreadedCommentPreimage {
	const sheet = workbook.getSheet(op.sheet)
	if (
		!sheet ||
		(op.partPath === undefined &&
			op.threadedCommentId === undefined &&
			op.ref === undefined &&
			op.commentIndex === undefined) ||
		(op.commentIndex !== undefined && (op.commentIndex < 0 || !Number.isInteger(op.commentIndex)))
	) {
		return { sheet: op.sheet, commentIndex: null, threadedComment: null }
	}
	const matches = sheet.threadedComments
		.map((comment, index) => ({ comment, index }))
		.filter(({ comment, index }) => {
			if (op.commentIndex !== undefined && index !== op.commentIndex) return false
			if (op.partPath !== undefined && comment.partPath !== op.partPath) return false
			if (op.threadedCommentId !== undefined && comment.id !== op.threadedCommentId) return false
			if (op.ref !== undefined && comment.ref.toUpperCase() !== op.ref.toUpperCase()) return false
			return true
		})
	const match = matches.length === 1 ? matches[0] : undefined
	return {
		sheet: op.sheet,
		commentIndex: match?.index ?? null,
		threadedComment: match ? cloneThreadedComment(match.comment) : null,
	}
}

function threadedCommentPreimagesAtRef(
	workbook: Workbook,
	sheetName: string,
	refText: string,
): MutationJournalThreadedCommentPreimage[] {
	const sheet = workbook.getSheet(sheetName)
	if (!sheet) return []
	const ref = refText.toUpperCase()
	return sheet.threadedComments
		.map((comment, index) => ({ comment, index }))
		.filter(({ comment }) => comment.ref.toUpperCase() === ref)
		.map(({ comment, index }) => ({
			sheet: sheetName,
			commentIndex: index,
			threadedComment: cloneThreadedComment(comment),
		}))
}

function deleteCommentThreadedCommentIssues(
	sheetName: string,
	refText: string,
	threadedComments: readonly MutationJournalThreadedCommentPreimage[],
): readonly MutationJournalIssue[] {
	if (threadedComments.length === 0) return []
	return [
		{
			code: 'LOSSY_INVERSE',
			message: `Threaded comments deleted at ${sheetName}!${refText.toUpperCase()} cannot be recreated with public operations`,
			surface: 'comments',
			reason: 'threaded-comment-selector',
			refs: threadedCommentIssueRefs(sheetName, threadedComments),
		},
	]
}

function threadedCommentIssueRefs(
	sheetName: string,
	threadedComments: readonly MutationJournalThreadedCommentPreimage[],
): string[] {
	return threadedComments.map((preimage) => {
		const id = preimage.threadedComment?.id
		return id
			? `${sheetName}!threadedComment:${id}`
			: `${sheetName}!threadedComment:${preimage.commentIndex ?? 'unknown'}`
	})
}

function threadedCommentStableSelector(
	preimage: MutationJournalThreadedCommentPreimage,
): Omit<Extract<Operation, { op: 'setThreadedComment' }>, 'op' | 'sheet' | 'text'> {
	const comment = preimage.threadedComment
	if (!comment) return {}
	if (comment.partPath !== undefined && comment.id !== undefined) {
		return { partPath: comment.partPath, threadedCommentId: comment.id }
	}
	if (comment.id !== undefined) return { threadedCommentId: comment.id }
	if (comment.partPath !== undefined) return { partPath: comment.partPath }
	if (preimage.commentIndex !== null) return { commentIndex: preimage.commentIndex }
	return { ref: comment.ref }
}

function drawingTextPreimage(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setDrawingText' }>,
): MutationJournalDrawingTextPreimage {
	const sheet = workbook.getSheet(op.sheet)
	if (
		!sheet ||
		(op.drawingPartPath === undefined &&
			op.id === undefined &&
			op.name === undefined &&
			op.drawingObjectIndex === undefined) ||
		(op.drawingObjectIndex !== undefined &&
			(op.drawingObjectIndex < 0 || !Number.isInteger(op.drawingObjectIndex)))
	) {
		return { sheet: op.sheet, drawingObjectIndex: null, drawingObject: null }
	}
	const matches = sheet.drawingObjectRefs
		.map((object, index) => ({ object, index }))
		.filter(({ object, index }) => {
			if (op.drawingObjectIndex !== undefined && index !== op.drawingObjectIndex) return false
			if (op.drawingPartPath !== undefined && object.drawingPartPath !== op.drawingPartPath) {
				return false
			}
			if (op.id !== undefined && object.id !== op.id) return false
			if (op.name !== undefined && object.name !== op.name) return false
			return true
		})
	const match = matches.length === 1 ? matches[0] : undefined
	return {
		sheet: op.sheet,
		drawingObjectIndex: match?.index ?? null,
		drawingObject: match ? cloneDrawingObjectRef(match.object) : null,
	}
}

function drawingObjectStableSelector(
	preimage: MutationJournalDrawingTextPreimage,
): Omit<Extract<Operation, { op: 'setDrawingText' }>, 'op' | 'sheet' | 'text'> {
	const object = preimage.drawingObject
	if (!object) return {}
	if (object.drawingPartPath !== undefined && object.id !== undefined) {
		return { drawingPartPath: object.drawingPartPath, id: object.id }
	}
	if (object.id !== undefined) return { id: object.id }
	if (object.drawingPartPath !== undefined && object.name !== undefined) {
		return { drawingPartPath: object.drawingPartPath, name: object.name }
	}
	if (object.name !== undefined) return { name: object.name }
	if (preimage.drawingObjectIndex !== null) {
		return { drawingObjectIndex: preimage.drawingObjectIndex }
	}
	return { drawingPartPath: object.drawingPartPath }
}

function chartSeriesPreimage(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setChartSeriesSource' }>,
): MutationJournalChartSeriesPreimage {
	if (
		op.seriesIndex < 0 ||
		!Number.isInteger(op.seriesIndex) ||
		(op.chartIndex !== undefined && (op.chartIndex < 0 || !Number.isInteger(op.chartIndex)))
	) {
		return { chartIndex: null, seriesIndex: op.seriesIndex, chart: null, series: null }
	}
	let candidates = workbook.chartParts.map((chart, index) => ({ chart, index }))
	if (op.partPath !== undefined) {
		candidates = candidates.filter(({ chart }) => chart.partPath === op.partPath)
	}
	if (op.sheet !== undefined) {
		const sheetName = op.sheet
		candidates = candidates.filter(
			({ chart }) => chart.sheetName !== undefined && sameSheetName(chart.sheetName, sheetName),
		)
	}
	if (op.chartIndex !== undefined) {
		const chart = candidates[op.chartIndex]
		candidates = chart ? [chart] : []
	}
	const match = candidates.length === 1 ? candidates[0] : undefined
	const series = match?.chart.series[op.seriesIndex]
	return {
		chartIndex: match?.index ?? null,
		seriesIndex: op.seriesIndex,
		chart: match ? cloneChartPart(match.chart) : null,
		series: series ? cloneChartSeries(series) : null,
	}
}

function chartSeriesInverseOperation(
	op: Extract<Operation, { op: 'setChartSeriesSource' }>,
	preimage: MutationJournalChartSeriesPreimage,
): { readonly exact: boolean; readonly op?: Operation } {
	if (!preimage.chart || !preimage.series) return { exact: false }
	const refs: { nameRef?: string; categoryRef?: string; valueRef?: string } = {}
	let exact = true
	for (const field of ['nameRef', 'categoryRef', 'valueRef'] as const) {
		if (op[field] === undefined) continue
		const value = preimage.series[field]
		if (value === undefined) {
			exact = false
			continue
		}
		refs[field] = value
	}
	const inverse: Extract<Operation, { op: 'setChartSeriesSource' }> = {
		op: 'setChartSeriesSource',
		seriesIndex: preimage.seriesIndex,
		...chartSeriesStableSelector(preimage),
		...refs,
	}
	return hasChartSeriesRefUpdate(inverse) ? { exact, op: inverse } : { exact: false }
}

function chartSeriesStableSelector(
	preimage: MutationJournalChartSeriesPreimage,
): Omit<
	Extract<Operation, { op: 'setChartSeriesSource' }>,
	'op' | 'seriesIndex' | 'nameRef' | 'categoryRef' | 'valueRef'
> {
	const chart = preimage.chart
	if (!chart) return {}
	if (chart.partPath !== undefined) return { partPath: chart.partPath }
	if (chart.sheetName !== undefined && preimage.chartIndex !== null) {
		return { sheet: chart.sheetName, chartIndex: preimage.chartIndex }
	}
	if (chart.sheetName !== undefined) return { sheet: chart.sheetName }
	if (preimage.chartIndex !== null) return { chartIndex: preimage.chartIndex }
	return {}
}

function hasChartSeriesRefUpdate(op: Extract<Operation, { op: 'setChartSeriesSource' }>): boolean {
	return op.nameRef !== undefined || op.categoryRef !== undefined || op.valueRef !== undefined
}

function pivotCachePreimage(
	workbook: Workbook,
	op: Extract<Operation, { op: 'setPivotCache' }>,
): MutationJournalPivotCachePreimage {
	let expectedCacheId = op.cacheId
	if (op.pivotTable !== undefined) {
		const pivots = workbook.pivotTables.filter((pivot) => pivot.name === op.pivotTable)
		if (pivots.length !== 1 || pivots[0]?.cacheId === undefined) {
			return { cacheIndex: null, cache: null }
		}
		if (expectedCacheId !== undefined && expectedCacheId !== pivots[0].cacheId) {
			return { cacheIndex: null, cache: null }
		}
		expectedCacheId = pivots[0].cacheId
	}
	const matches = workbook.pivotCaches
		.map((cache, index) => ({ cache, index }))
		.filter(({ cache }) => {
			if (op.partPath !== undefined && cache.partPath !== op.partPath) return false
			if (expectedCacheId !== undefined && cache.cacheId !== expectedCacheId) return false
			return true
		})
	const match = matches.length === 1 ? matches[0] : undefined
	return {
		cacheIndex: match?.index ?? null,
		cache: match ? clonePivotCacheInfo(match.cache) : null,
	}
}

function pivotCacheInverseOperation(
	op: Extract<Operation, { op: 'setPivotCache' }>,
	preimage: MutationJournalPivotCachePreimage,
): { readonly exact: boolean; readonly op?: Operation } {
	if (!preimage.cache) return { exact: false }
	const fields: {
		sourceSheet?: string
		sourceRef?: string
		refreshOnLoad?: boolean
		enableRefresh?: boolean
		invalid?: boolean
		saveData?: boolean
	} = {}
	let exact = true
	for (const field of [
		'sourceSheet',
		'sourceRef',
		'refreshOnLoad',
		'enableRefresh',
		'invalid',
		'saveData',
	] as const) {
		if (op[field] === undefined) continue
		switch (field) {
			case 'sourceSheet': {
				const value = preimage.cache.sourceSheet
				if (value === undefined) {
					exact = false
					continue
				}
				fields.sourceSheet = value
				break
			}
			case 'sourceRef': {
				const value = preimage.cache.sourceRef
				if (value === undefined) {
					exact = false
					continue
				}
				fields.sourceRef = value
				break
			}
			case 'refreshOnLoad': {
				const value = preimage.cache.refreshOnLoad
				if (value === undefined) {
					exact = false
					continue
				}
				fields.refreshOnLoad = value
				break
			}
			case 'enableRefresh': {
				const value = preimage.cache.enableRefresh
				if (value === undefined) {
					exact = false
					continue
				}
				fields.enableRefresh = value
				break
			}
			case 'invalid': {
				const value = preimage.cache.invalid
				if (value === undefined) {
					exact = false
					continue
				}
				fields.invalid = value
				break
			}
			case 'saveData': {
				const value = preimage.cache.saveData
				if (value === undefined) {
					exact = false
					continue
				}
				fields.saveData = value
				break
			}
		}
	}
	const inverse: Extract<Operation, { op: 'setPivotCache' }> = {
		op: 'setPivotCache',
		partPath: preimage.cache.partPath,
		...fields,
	}
	return hasPivotCacheUpdate(inverse) ? { exact, op: inverse } : { exact: false }
}

function hasPivotCacheUpdate(op: Extract<Operation, { op: 'setPivotCache' }>): boolean {
	return (
		op.sourceSheet !== undefined ||
		op.sourceRef !== undefined ||
		op.refreshOnLoad !== undefined ||
		op.enableRefresh !== undefined ||
		op.invalid !== undefined ||
		op.saveData !== undefined
	)
}

function setHyperlinkInverse(
	sheet: string,
	ref: string,
	hyperlink: SheetHyperlink,
): Extract<Operation, { op: 'setHyperlink' }> {
	return {
		op: 'setHyperlink',
		sheet,
		ref,
		...(hyperlink.target !== undefined ? { url: hyperlink.target } : {}),
		...(hyperlink.location !== undefined ? { location: hyperlink.location } : {}),
		...(hyperlink.display !== undefined ? { display: hyperlink.display } : {}),
		...(hyperlink.tooltip !== undefined ? { tooltip: hyperlink.tooltip } : {}),
	}
}

function cloneFormulaInfo(formulaInfo: NonNullable<Cell['formulaInfo']>): Cell['formulaInfo'] {
	if (formulaInfo.kind === 'blockedSpill') {
		return { ...formulaInfo, blockingRefs: [...formulaInfo.blockingRefs] }
	}
	return { ...formulaInfo }
}

function cloneCellValue(value: CellValue): CellValue {
	switch (value.kind) {
		case 'richText':
			return { kind: 'richText', runs: value.runs.map((run) => ({ ...run })) }
		case 'array':
			return {
				kind: 'array',
				rows: value.rows.map((row) => row.map(cloneScalarCellValue)),
			}
		default:
			return { ...value }
	}
}

function cloneScalarCellValue(value: ScalarCellValue): ScalarCellValue {
	return cloneCellValue(value) as ScalarCellValue
}

function cloneDefinedName(definedName: DefinedName): DefinedName {
	return {
		...definedName,
		scope: { ...definedName.scope },
		...(definedName.extraAttributes
			? { extraAttributes: definedName.extraAttributes.map((attribute) => ({ ...attribute })) }
			: {}),
	}
}

function clonePageSetup(setup: SheetPageSetup): SheetPageSetup {
	return { ...setup }
}

function cloneTableColumn(column: TableColumn): TableColumn {
	return {
		...column,
		...(column.xmlColumnPr ? { xmlColumnPr: { ...column.xmlColumnPr } } : {}),
	}
}

function cloneTable(table: Table): Table {
	return {
		...table,
		ref: { start: { ...table.ref.start }, end: { ...table.ref.end } },
		columns: table.columns.map(cloneTableColumn),
		...(table.autoFilter ? { autoFilter: cloneAutoFilter(table.autoFilter) } : {}),
		...(table.sortState
			? {
					sortState: {
						...table.sortState,
						...(table.sortState.preservedAttributes
							? { preservedAttributes: { ...table.sortState.preservedAttributes } }
							: {}),
						conditions: table.sortState.conditions.map((condition) => ({ ...condition })),
					},
				}
			: {}),
		...(table.tableStyleInfo ? { tableStyleInfo: { ...table.tableStyleInfo } } : {}),
		...(table.queryTable ? { queryTable: { ...table.queryTable } } : {}),
	}
}

function cloneThreadedComment(comment: SheetThreadedComment): SheetThreadedComment {
	return { ...comment }
}

function cloneDrawingObjectRef(object: SheetDrawingObjectRef): SheetDrawingObjectRef {
	return {
		...object,
		...(object.anchor ? { anchor: clonePlain(object.anchor) } : {}),
		...(object.relIds ? { relIds: [...object.relIds] } : {}),
		...(object.relationshipRefs
			? { relationshipRefs: object.relationshipRefs.map((ref) => ({ ...ref })) }
			: {}),
	}
}

function cloneChartPart(chart: ChartPartInfo): ChartPartInfo {
	return {
		...chart,
		series: chart.series.map(cloneChartSeries),
	}
}

function cloneChartSeries(series: ChartSeriesInfo): ChartSeriesInfo {
	return { ...series }
}

function cloneConditionalFormat(format: SheetConditionalFormat): SheetConditionalFormat {
	return {
		...format,
		rules: format.rules.map((rule) => ({
			...rule,
			formulas: [...rule.formulas],
			...(rule.style ? { style: cloneCellStyle(rule.style) } : {}),
			...(rule.preservedRuleAttributes
				? { preservedRuleAttributes: { ...rule.preservedRuleAttributes } }
				: {}),
			...(rule.preservedRuleChildXml
				? { preservedRuleChildXml: [...rule.preservedRuleChildXml] }
				: {}),
			...(rule.colorScale ? { colorScale: clonePlain(rule.colorScale) } : {}),
			...(rule.dataBar ? { dataBar: clonePlain(rule.dataBar) } : {}),
			...(rule.iconSet ? { iconSet: clonePlain(rule.iconSet) } : {}),
		})),
	}
}

function cloneAutoFilter(autoFilter: AutoFilter): AutoFilter {
	return {
		...autoFilter,
		columns: autoFilter.columns.map((column) => ({
			...column,
			...(column.values ? { values: [...column.values] } : {}),
			...(column.dateGroupItems
				? { dateGroupItems: column.dateGroupItems.map((item) => ({ ...item })) }
				: {}),
			...(column.customFilters
				? { customFilters: column.customFilters.map((filter) => ({ ...filter })) }
				: {}),
		})),
		...(autoFilter.sortState
			? {
					sortState: {
						...autoFilter.sortState,
						...(autoFilter.sortState.preservedAttributes
							? { preservedAttributes: { ...autoFilter.sortState.preservedAttributes } }
							: {}),
						conditions: autoFilter.sortState.conditions.map((condition) => ({ ...condition })),
					},
				}
			: {}),
	}
}

function clonePlain<T>(value: T): T {
	return JSON.parse(JSON.stringify(value)) as T
}

function isDataValidationType(value: string): value is DataValidationRule['type'] {
	return DATA_VALIDATION_TYPES.has(value as DataValidationRule['type'])
}

function isDataValidationOperator(
	value: string | undefined,
): value is NonNullable<DataValidationRule['operator']> {
	return (
		value !== undefined &&
		DATA_VALIDATION_OPERATORS.has(value as NonNullable<DataValidationRule['operator']>)
	)
}

function isConditionalFormatType(value: string): value is ConditionalFormatRule['type'] {
	return CONDITIONAL_FORMAT_TYPES.has(value as ConditionalFormatRule['type'])
}

function isConditionalFormatOperator(
	value: string | undefined,
): value is NonNullable<ConditionalFormatRule['operator']> {
	return (
		value !== undefined &&
		CONDITIONAL_FORMAT_OPERATORS.has(value as NonNullable<ConditionalFormatRule['operator']>)
	)
}

const DATA_VALIDATION_TYPES = new Set<DataValidationRule['type']>([
	'list',
	'whole',
	'decimal',
	'date',
	'time',
	'textLength',
	'custom',
])

const DATA_VALIDATION_OPERATORS = new Set<NonNullable<DataValidationRule['operator']>>([
	'between',
	'notBetween',
	'equal',
	'notEqual',
	'greaterThan',
	'lessThan',
	'greaterThanOrEqual',
	'lessThanOrEqual',
])

const CONDITIONAL_FORMAT_TYPES = new Set<ConditionalFormatRule['type']>([
	'cellIs',
	'expression',
	'colorScale',
	'dataBar',
	'iconSet',
	'top10',
	'aboveAverage',
	'duplicateValues',
	'containsText',
])

const CONDITIONAL_FORMAT_OPERATORS = new Set<NonNullable<ConditionalFormatRule['operator']>>([
	'greaterThan',
	'lessThan',
	'equal',
	'between',
	'greaterThanOrEqual',
	'lessThanOrEqual',
	'notEqual',
	'notBetween',
])
