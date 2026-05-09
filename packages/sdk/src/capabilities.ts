export type CapabilityStatus =
	| 'excel-equivalent'
	| 'editable'
	| 'inspectable'
	| 'preserved'
	| 'unsafe-blocked'
	| 'unsupported'

export type CapabilityPriority = 'P0' | 'P1' | 'P2' | 'P3'
export type CapabilitySurface = 'cli' | 'sdk' | 'mcp' | 'api'

export interface ExcelCapability {
	readonly id: string
	readonly family: string
	readonly label: string
	readonly status: CapabilityStatus
	readonly priority: CapabilityPriority
	readonly surfaceCoverage: Readonly<Record<CapabilitySurface, boolean>>
	readonly ossBaseline: Readonly<Record<string, string>>
	readonly tests: readonly string[]
	readonly gapReason: string
	readonly nextMilestone: string
}

export interface CapabilityFilters {
	readonly feature?: string
	readonly family?: string
	readonly priority?: CapabilityPriority
	readonly status?: CapabilityStatus
	readonly gapsOnly?: boolean
}

export interface CapabilitySummary {
	readonly total: number
	readonly byStatus: Readonly<Record<CapabilityStatus, number>>
	readonly byPriority: Readonly<Record<CapabilityPriority, number>>
	readonly gaps: number
}

const ALL_SURFACES: Readonly<Record<CapabilitySurface, boolean>> = {
	cli: true,
	sdk: true,
	mcp: true,
	api: true,
}

const OSS_BASELINE = {
	SheetJS: 'Broad format IO and formula preservation; no formula execution or agent workflow.',
	XlsxWriter:
		'Excellent write-only XLSX authoring; cannot read or safely modify existing workbooks.',
	openpyxl: 'Broad Python read/write; charts and pivots can be lossy or preserve-oriented.',
	ApachePOI: 'Broad JVM XLSX support; charts, pivots, and active content remain limited.',
	ClosedXML: 'Ergonomic .NET workbook API; less agent-native and incomplete full-file fidelity.',
	Excelize: 'Strong Go XLSX API; pivots and rich features exist but not full Excel parity.',
	HyperFormula: 'Strong headless formula engine; not a full XLSX fidelity platform.',
	IronCalc: 'Promising formula/workbook engine; feature coverage is still incomplete.',
	Univer: 'Modern spreadsheet platform; not optimized as a compact headless file CLI.',
} as const

function cap(
	id: string,
	family: string,
	label: string,
	status: CapabilityStatus,
	priority: CapabilityPriority,
	nextMilestone: string,
	gapReason = status === 'excel-equivalent'
		? 'No known functional gap at the current parity bar.'
		: 'Capability is classified for staged Excel parity work.',
	tests: readonly string[] = [],
	surfaceCoverage: Readonly<Record<CapabilitySurface, boolean>> = ALL_SURFACES,
): ExcelCapability {
	return {
		id,
		family,
		label,
		status,
		priority,
		surfaceCoverage,
		ossBaseline: OSS_BASELINE,
		tests,
		gapReason,
		nextMilestone,
	}
}

export const EXCEL_CAPABILITIES: readonly ExcelCapability[] = [
	cap(
		'workbook.xlsx-package',
		'workbook/package',
		'XLSX package read/write',
		'editable',
		'P0',
		'Keep expanding no-op preservation and semantic roundtrip corpus.',
		'Core package IO exists with preservation capsules; exact fidelity must keep broadening.',
		['packages/sdk/src/sdk.test.ts', 'packages/io-xlsx/src/writer/writer.test.ts'],
	),
	cap(
		'workbook.xlsm-package',
		'workbook/package',
		'XLSM package and macro container',
		'inspectable',
		'P0',
		'Add macro-sheet inventory and signature invalidation policy.',
		'Macro containers are inventoried and preserved; macro semantics are not executable or editable.',
		['packages/io-xlsx/src/reader/active-content.test.ts'],
	),
	cap(
		'workbook.properties',
		'workbook/package',
		'Workbook properties',
		'inspectable',
		'P1',
		'Add editable document/core/custom property operations.',
		'Properties are visible through workbook metadata but lack first-class edit operations.',
	),
	cap(
		'workbook.themes',
		'workbook/package',
		'Themes and theme colors',
		'preserved',
		'P1',
		'Expose theme inventory and safe theme replacement.',
		'Theme parts are preserved but not yet a structured edit surface.',
	),
	cap(
		'workbook.views',
		'workbook/package',
		'Workbook views and active tab',
		'inspectable',
		'P1',
		'Add view edit operations and parity fixtures.',
		'Views are inspectable metadata today.',
	),
	cap(
		'workbook.protection',
		'workbook/package',
		'Workbook protection',
		'editable',
		'P1',
		'Add cross-version hash fixtures and protected workbook roundtrips.',
		'Workbook protection operations exist but need broader compatibility fixtures.',
	),
	cap(
		'workbook.calc-settings',
		'workbook/package',
		'Calculation settings and chain metadata',
		'inspectable',
		'P1',
		'Expose calc mode, iteration, precision, and chain preservation audits.',
		'Calc metadata is reported but not fully editable.',
	),
	cap(
		'workbook.external-links',
		'workbook/package',
		'External workbook links',
		'editable',
		'P0',
		'Add formula reference rewrite audits and cross-workbook fixture checks.',
		'External link targets can be rewritten safely; formula text auditing remains preserve-first.',
		[
			'packages/io-xlsx/src/reader/reader.test.ts',
			'packages/io-xlsx/src/writer/writer.test.ts',
			'packages/engine/src/operations.test.ts',
		],
	),

	cap(
		'sheets.values',
		'sheets/ranges',
		'Cell values',
		'excel-equivalent',
		'P0',
		'Maintain speed and roundtrip benchmarks.',
		'No known gap for core scalar cell values.',
		['packages/sdk/src/sdk.test.ts'],
	),
	cap(
		'sheets.formulas',
		'sheets/ranges',
		'Formula text edit/preserve',
		'editable',
		'P0',
		'Finish shared/array/dynamic formula rewrite guarantees.',
		'Formula text can be edited, but advanced binding semantics still need full parity.',
	),
	cap(
		'sheets.styles',
		'sheets/ranges',
		'Styles and number formats',
		'editable',
		'P0',
		'Expand border/fill/theme style fixtures and loss audits.',
		'Common style operations exist; deeper style fidelity remains fixture-driven.',
	),
	cap(
		'sheets.conditional-formatting',
		'sheets/ranges',
		'Conditional formatting',
		'editable',
		'P0',
		'Add priority reorder operations, overlap diagnostics, and extLst/x14 rule preservation.',
		'CellIs/expression rules and visual colorScale/dataBar/iconSet rules round-trip with core metadata.',
		['packages/io-xlsx/src/writer/conditional-format-visual.test.ts'],
	),
	cap(
		'sheets.rich-text',
		'sheets/ranges',
		'Rich text runs',
		'editable',
		'P1',
		'Add read/modify fixtures for shared strings and inline rich text.',
		'Set operations exist, but mixed rich text roundtrip coverage is still growing.',
	),
	cap(
		'sheets.comments',
		'sheets/ranges',
		'Cell comments',
		'editable',
		'P1',
		'Add threaded comment inventory plus edit blockers.',
		'Legacy comments are editable; threaded comments are preserve-first.',
	),
	cap(
		'sheets.threaded-comments',
		'sheets/ranges',
		'Threaded comments',
		'preserved',
		'P1',
		'Expose author/thread metadata and block lossy writes.',
		'Threaded comments are preserved but not semantically editable.',
	),
	cap(
		'sheets.hyperlinks',
		'sheets/ranges',
		'Hyperlinks',
		'editable',
		'P1',
		'Add internal/external hyperlink rewrite fixtures.',
		'Basic hyperlink operations exist.',
	),
	cap(
		'sheets.merged-cells',
		'sheets/ranges',
		'Merged cells',
		'editable',
		'P1',
		'Add overlap and copy/paste merge behavior tests.',
		'Merge operations exist; transform semantics need more coverage.',
	),
	cap(
		'sheets.copy-paste',
		'sheets/ranges',
		'Copy, move, and paste modes',
		'editable',
		'P0',
		'Add merge-aware overlap handling and cross-sheet paste fixtures.',
		'Copy/move supports Excel-like modes for all, values, formulas, formats/styles, validations, comments, and hyperlinks.',
		['packages/engine/src/operations.test.ts'],
	),
	cap(
		'sheets.panes',
		'sheets/ranges',
		'Freeze panes',
		'editable',
		'P1',
		'Add split pane and workbook view fixtures.',
		'Freeze panes are editable; split panes are not fully modeled.',
	),
	cap(
		'sheets.dimensions',
		'sheets/ranges',
		'Row heights and column widths',
		'editable',
		'P1',
		'Add default dimension and hidden outline fixtures.',
		'Explicit dimensions are editable.',
	),
	cap(
		'sheets.hidden-grouped-rows-cols',
		'sheets/ranges',
		'Hidden and grouped rows/columns',
		'editable',
		'P1',
		'Add outline summary and collapsed-state roundtrips.',
		'Group/hide operations exist; Excel outline edge cases need coverage.',
	),

	cap(
		'tables.tables',
		'tables/data',
		'Excel tables',
		'editable',
		'P0',
		'Add table style edit operations, column rename rewrites, and totals-row materialization fixtures.',
		'Create/delete/rename/resize plus calculated-column formula and totals metadata edits exist.',
		['packages/engine/src/operations.test.ts'],
	),
	cap(
		'tables.structured-refs',
		'tables/data',
		'Structured references',
		'editable',
		'P0',
		'Broaden Excel-ground-truth fixtures for totals, headers, current-row refs, and rewrites.',
		'Structured references parse, evaluate, and participate in table/formula rewrite flows.',
	),
	cap(
		'tables.filters',
		'tables/data',
		'Auto filters and criteria',
		'editable',
		'P0',
		'Represent criteria, dynamic filters, top10, color/icon filters, and state.',
		'Filter ranges can be set/cleared; full criteria semantics are missing.',
	),
	cap(
		'tables.sorts',
		'tables/data',
		'Sort state and range sorting',
		'editable',
		'P1',
		'Preserve and expose sort state separately from executing sort operations.',
		'Range sort operations exist; persisted sort-state fidelity needs coverage.',
	),
	cap(
		'tables.data-validation',
		'tables/data',
		'Data validation',
		'editable',
		'P1',
		'Add named-list validation fixtures and table-resize rewrite tests.',
		'Validation operations and XLSX roundtrip cover formulas, prompts, errors, dropdown policy, and IME mode.',
		['packages/io-xlsx/src/writer/data-validation.test.ts'],
	),
	cap(
		'tables.advanced-filters',
		'tables/data',
		'Advanced filters',
		'unsupported',
		'P2',
		'Inventory advanced filter definitions and preserve parts before edit support.',
		'Advanced filter definitions are not modeled yet.',
	),

	cap(
		'formulas.functions',
		'formula engine',
		'Excel function coverage',
		'editable',
		'P0',
		'Keep expanding Excel-generated edge-case fixtures beyond the common-function audit.',
		'The common-function audit reports 413/413 functions present, including ISFORMULA.',
		[
			'fixtures/formulas/missing-formula-audit.ts',
			'packages/formulas/src/functions/functions.test.ts',
		],
	),
	cap(
		'formulas.dynamic-arrays',
		'formula engine',
		'Dynamic arrays and spills',
		'editable',
		'P0',
		'Expand Excel-ground-truth coverage for nested spill and resize edge cases.',
		'Dynamic array execution, spill ranges, blocked spills, and resize recalculation are implemented for core modern functions.',
	),
	cap(
		'formulas.lambda-let',
		'formula engine',
		'LAMBDA and LET',
		'editable',
		'P0',
		'Add recursion limits, optional-argument helpers, and broader ground-truth fixtures.',
		'LAMBDA, LET, MAP, REDUCE, SCAN, BYROW, BYCOL, and MAKEARRAY execute in the headless engine.',
	),
	cap(
		'formulas.names',
		'formula engine',
		'Defined names',
		'editable',
		'P0',
		'Add scoped name dependency tracing and external name refs.',
		'Defined names are editable; advanced binding semantics need more tests.',
	),
	cap(
		'formulas.volatile-refs',
		'formula engine',
		'Volatile functions and refs',
		'editable',
		'P1',
		'Add more dependency invalidation tests for volatile references across sheets and names.',
		'NOW, TODAY, RAND, INDIRECT, and OFFSET participate in recalculation paths with deterministic test controls.',
	),
	cap(
		'formulas.external-refs',
		'formula engine',
		'External references',
		'inspectable',
		'P1',
		'Preserve external refs symbolically and add configurable resolution hooks.',
		'External references are visible but not resolved.',
	),
	cap(
		'formulas.iterative-calc',
		'formula engine',
		'Iterative calculation',
		'inspectable',
		'P1',
		'Implement max-iteration/max-change settings and circular convergence.',
		'Circular refs are detected; iterative calc is not Excel-equivalent.',
	),
	cap(
		'formulas.spill-diagnostics',
		'formula engine',
		'Spill diagnostics',
		'inspectable',
		'P1',
		'Expose machine-readable blocked-range causes in plan/check output.',
		'Blocked spills produce #SPILL! and spill ownership is tracked for recalculation cleanup.',
	),

	cap(
		'visuals.charts',
		'visuals',
		'Charts',
		'editable',
		'P0',
		'Add axes, anchors, and styling inventory after safe series source edits.',
		'Chart type, title, and series source refs are inspectable; series source refs are editable while unsupported styling is preserved.',
		[
			'packages/engine/src/operations.test.ts',
			'packages/io-xlsx/src/reader/reader.test.ts',
			'packages/io-xlsx/src/writer/writer.test.ts',
			'packages/sdk/src/visual-inventory.test.ts',
		],
	),
	cap(
		'visuals.chartsheets',
		'visuals',
		'Chartsheets',
		'unsupported',
		'P1',
		'Inventory chartsheets and block lossy writes.',
		'Chartsheets are not modeled as first-class sheets.',
	),
	cap(
		'visuals.images',
		'visuals',
		'Images',
		'editable',
		'P0',
		'Extend insert/delete to complex preserved drawings without losing charts or shapes.',
		'Image inventory, replacement, and generated-image insert/delete exist; complex preserved drawing edits remain guarded.',
		[
			'packages/engine/src/operations.test.ts',
			'packages/io-xlsx/src/writer/writer.test.ts',
			'packages/sdk/src/visual-inventory.test.ts',
		],
	),
	cap(
		'visuals.drawings-shapes',
		'visuals',
		'Drawings and shapes',
		'preserved',
		'P0',
		'Inspect drawing relationships and preserve unsupported styling exactly.',
		'Shape semantics are preserve-first.',
	),
	cap(
		'visuals.text-boxes',
		'visuals',
		'Text boxes',
		'preserved',
		'P1',
		'Expose text box inventory and safe text replacement.',
		'Text boxes are not semantically editable.',
	),
	cap(
		'visuals.sparklines',
		'visuals',
		'Sparklines',
		'preserved',
		'P2',
		'Inventory sparkline groups and source ranges.',
		'Sparkline XML is preserve-first.',
	),

	cap(
		'analytics.pivots',
		'analytics',
		'Pivot tables',
		'inspectable',
		'P0',
		'Add filter state edits and calculated output refresh after field-level inventory.',
		'Pivot cache fields and row/column/page/data field layout are inspectable; full management and recalculation are not complete.',
		['packages/io-xlsx/src/reader/reader.test.ts'],
	),
	cap(
		'analytics.pivot-caches',
		'analytics',
		'Pivot caches',
		'editable',
		'P0',
		'Add cache-record recalculation and broader Excel fixture coverage.',
		'Pivot cache source and refresh metadata are editable; output recalculation is not implemented.',
		['packages/io-xlsx/src/reader/reader.test.ts', 'packages/engine/src/operations.test.ts'],
	),
	cap(
		'analytics.slicers',
		'analytics',
		'Slicers',
		'inspectable',
		'P0',
		'Expose slicer links and safely preserve/filter state.',
		'Slicers are inventory-only today.',
	),
	cap(
		'analytics.timelines',
		'analytics',
		'Timelines',
		'inspectable',
		'P1',
		'Expose safe timeline filter/edit operations and pivot refresh impact warnings.',
		'Timeline caches and timeline UI parts are inventoried; semantic edits are still preserve-first.',
		['packages/io-xlsx/src/reader/timeline.test.ts'],
	),
	cap(
		'analytics.data-model',
		'analytics',
		'Data model',
		'preserved',
		'P1',
		'Inventory model parts and block lossy writes.',
		'Data model semantics are not modeled.',
	),
	cap(
		'analytics.getpivotdata',
		'analytics',
		'GETPIVOTDATA',
		'unsupported',
		'P1',
		'Implement GETPIVOTDATA over inspectable pivot metadata.',
		'Pivot output querying is not implemented.',
	),

	cap(
		'active.vba-macros',
		'active content',
		'VBA macros',
		'unsafe-blocked',
		'P0',
		'Add macro module summaries without exposing or executing code.',
		'VBA project parts are inventoried, preserved, and blocked by the loss audit before writes.',
		['packages/io-xlsx/src/reader/active-content.test.ts'],
	),
	cap(
		'active.activex-controls',
		'active content',
		'ActiveX controls',
		'inspectable',
		'P1',
		'Expose control relationship metadata and linked worksheet anchors.',
		'ActiveX parts are inventoried, preserved, and blocked by the loss audit before writes.',
		['packages/io-xlsx/src/reader/active-content.test.ts'],
	),
	cap(
		'active.form-controls',
		'active content',
		'Form controls',
		'inspectable',
		'P1',
		'Expose form control metadata, linked cells, and macro bindings.',
		'Control property parts are inventoried and preserved; linked behavior is not semantically editable.',
		['packages/io-xlsx/src/reader/active-content.test.ts'],
	),
	cap(
		'active.signatures',
		'active content',
		'Digital signatures',
		'preserved',
		'P1',
		'Detect invalidation risk and require explicit policy before writes.',
		'Signatures are not re-signed after edits.',
	),

	cap(
		'connections.external-workbooks',
		'connections',
		'External workbook connections',
		'inspectable',
		'P0',
		'List connection targets and formula refs with refresh/loss policy.',
		'External references are visible but not refreshable.',
	),
	cap(
		'connections.query-tables',
		'connections',
		'Query tables',
		'preserved',
		'P1',
		'Inventory query tables and refresh metadata.',
		'Query table semantics are preserve-first.',
	),
	cap(
		'connections.power-query',
		'connections',
		'Power Query',
		'preserved',
		'P1',
		'Inventory mashup parts and block lossy writes.',
		'Power Query execution is out of scope for current engine.',
	),
	cap(
		'connections.refresh-metadata',
		'connections',
		'Refresh metadata',
		'inspectable',
		'P1',
		'Expose refresh-on-open and stale cache indicators.',
		'Refresh metadata is not fully editable.',
	),

	cap(
		'agent.inspect',
		'agent UX',
		'Inspect workbook state',
		'editable',
		'P0',
		'Add capability warnings directly into inspect payloads.',
		'Inspect exists; warnings need to be tied to this registry.',
	),
	cap(
		'agent.search',
		'agent UX',
		'Search workbook content',
		'editable',
		'P0',
		'Add formula/metadata scoped search facets.',
		'Cell search exists; metadata search can expand.',
	),
	cap(
		'agent.read',
		'agent UX',
		'Token-efficient reads',
		'editable',
		'P0',
		'Expand compact/object/TSV modes and row-window parity across surfaces.',
		'Agent read helpers exist; every surface should expose the same options.',
	),
	cap(
		'agent.plan',
		'agent UX',
		'Plan safe edits',
		'editable',
		'P0',
		'Make plan digests and preservation audits mandatory in docs and examples.',
		'Plan command/API/MCP exists as the recommended edit preview path.',
	),
	cap(
		'agent.preview',
		'agent UX',
		'Preview operations',
		'editable',
		'P0',
		'Keep preview as a compatibility alias for plan.',
		'Preview exists and powers plan.',
	),
	cap(
		'agent.commit',
		'agent UX',
		'Commit safe edits',
		'editable',
		'P0',
		'Add allow-loss policy and backup defaults for unsafe features.',
		'Commit writes atomically with hash guards; deeper loss policy is next.',
	),
	cap(
		'agent.recalc',
		'agent UX',
		'Recalculate formulas',
		'editable',
		'P0',
		'Broaden Excel ground-truth formula gates.',
		'Recalc exists but Excel formula parity is ongoing.',
	),
	cap(
		'agent.verify',
		'agent UX',
		'Verify and lint',
		'editable',
		'P0',
		'Make suggested fixes machine-readable for every checker result.',
		'Check/lint exist; recovery actions need richer taxonomy.',
	),
	cap(
		'agent.diff',
		'agent UX',
		'Semantic diff',
		'editable',
		'P0',
		'Add visual/metadata diff classes.',
		'Core semantic diff exists.',
	),
	cap(
		'agent.export',
		'agent UX',
		'Export workbook data',
		'editable',
		'P1',
		'Expose stable export metadata and unsupported feature warnings.',
		'CSV/TSV/JSON/XLSX exports exist.',
	),
]

export function isCapabilityGap(status: CapabilityStatus): boolean {
	return status !== 'excel-equivalent' && status !== 'editable'
}

export function listCapabilities(filters: CapabilityFilters = {}): readonly ExcelCapability[] {
	const feature = filters.feature?.toLowerCase()
	const family = filters.family?.toLowerCase()
	return EXCEL_CAPABILITIES.filter((capability) => {
		if (feature) {
			const haystack = `${capability.id} ${capability.label} ${capability.family}`.toLowerCase()
			if (!haystack.includes(feature)) return false
		}
		if (family && capability.family.toLowerCase() !== family) return false
		if (filters.priority && capability.priority !== filters.priority) return false
		if (filters.status && capability.status !== filters.status) return false
		if (filters.gapsOnly && !isCapabilityGap(capability.status)) return false
		return true
	})
}

export function getCapability(feature: string): ExcelCapability | undefined {
	const needle = feature.toLowerCase()
	return EXCEL_CAPABILITIES.find(
		(capability) =>
			capability.id.toLowerCase() === needle || capability.label.toLowerCase() === needle,
	)
}

export function summarizeCapabilities(
	capabilities: readonly ExcelCapability[] = EXCEL_CAPABILITIES,
): CapabilitySummary {
	const byStatus: Record<CapabilityStatus, number> = {
		'excel-equivalent': 0,
		editable: 0,
		inspectable: 0,
		preserved: 0,
		'unsafe-blocked': 0,
		unsupported: 0,
	}
	const byPriority: Record<CapabilityPriority, number> = { P0: 0, P1: 0, P2: 0, P3: 0 }
	let gaps = 0
	for (const capability of capabilities) {
		byStatus[capability.status] += 1
		byPriority[capability.priority] += 1
		if (isCapabilityGap(capability.status)) gaps += 1
	}
	return { total: capabilities.length, byStatus, byPriority, gaps }
}
