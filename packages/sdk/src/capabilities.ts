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
		'editable',
		'P1',
		'Add docProps core/custom property editing after workbookPr operation coverage.',
		'WorkbookPr properties are inspectable and editable through setWorkbookProperties with merge/replace and null-clearing semantics; document property parts remain a future extension.',
		['packages/engine/src/workbook-ops.test.ts', 'packages/sdk/src/ops-schema.test.ts'],
	),
	cap(
		'workbook.themes',
		'workbook/package',
		'Themes and theme colors',
		'inspectable',
		'P1',
		'Add safe theme replacement with previewed style impact.',
		'Theme name, color scheme, theme color slots, system color fallbacks, and major/minor font names are inspectable while the source theme part remains preserved.',
		[
			'packages/io-xlsx/src/reader/theme-inventory.test.ts',
			'packages/sdk/src/theme-inventory.test.ts',
		],
	),
	cap(
		'workbook.views',
		'workbook/package',
		'Workbook views and active tab',
		'editable',
		'P1',
		'Add additional parity fixtures for multi-window and hidden workbook views.',
		'Workbook views are inspectable and editable through setWorkbookView with merge/replace, append, delete, and null-clearing semantics.',
		['packages/engine/src/workbook-ops.test.ts', 'packages/sdk/src/ops-schema.test.ts'],
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
		'editable',
		'P1',
		'Expose calc-chain preservation audits and broader plan warnings for date-system changes.',
		'Calc mode, full-calc flags, date system, calc IDs, and iterative calculation settings are editable through setCalcSettings with validation and date1904 synchronization.',
		['packages/engine/src/workbook-ops.test.ts', 'packages/sdk/src/ops-schema.test.ts'],
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
		'Add overlap diagnostics and extLst/x14 rule preservation.',
		'CellIs/expression and visual rules round-trip; operations can append overlapping rules, delete individual priorities, and reassign rule priority order.',
		[
			'packages/engine/src/operations.test.ts',
			'packages/io-xlsx/src/writer/conditional-format-visual.test.ts',
		],
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
		'inspectable',
		'P1',
		'Add explicit edit blockers and author/person roundtrip fixtures.',
		'Threaded comments expose ref, text, thread IDs, parent IDs, person IDs, author names, timestamps, and preservation/loss-audit features; semantic edits remain blocked.',
		[
			'packages/io-xlsx/src/reader/threaded-comments.test.ts',
			'packages/sdk/src/threaded-comments.test.ts',
		],
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
		'Add icon filter evaluation plus richer filter-state edit operations.',
		'Filter ranges plus explicit value, date group, custom comparison, top/bottom, dynamic, and color criteria are parsed, preserved, and evaluated for formula-visible row semantics.',
		['packages/engine/src/calc.test.ts'],
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
		'inspectable',
		'P2',
		'Add safe advanced-filter edit operations after full fixture coverage.',
		'Advanced filter/custom sheet view auto-filters are inventoried with filter and sort counts while raw worksheet XML remains preserved for roundtrip.',
		[
			'packages/io-xlsx/src/reader/advanced-filter-sparkline.test.ts',
			'packages/sdk/src/advanced-filter-sparkline.test.ts',
		],
	),

	cap(
		'formulas.functions',
		'formula engine',
		'Excel function coverage',
		'editable',
		'P0',
		'Keep deepening Excel-generated edge-case fixtures beyond tracked common-function smoke coverage.',
		'The common-function audit enforces full registry presence separately from full tracked JSON semantic corpus coverage.',
		[
			'fixtures/formulas/missing-formula-audit.ts',
			'fixtures/formulas/missing-formula-audit.test.ts',
			'fixtures/formulas/conformance.test.ts',
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
		'Chart type, title, and series source refs are inspectable; series source refs are editable while opaque styling is preserved.',
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
		'inspectable',
		'P1',
		'Add chartsheet editing only after lossless chart-sheet package preservation is guaranteed.',
		'Chartsheets are inventoried separately from worksheet grids and blocked by the loss audit before writes.',
		[
			'packages/io-xlsx/src/reader/chartsheet.test.ts',
			'packages/sdk/src/chartsheet-inventory.test.ts',
		],
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
		'inspectable',
		'P0',
		'Add safe shape/text replacement operations while preserving opaque styling exactly.',
		'Drawing object inventory exposes shape, connector, group shape, graphic frame, text box anchors, names, text, and relationship IDs.',
		['packages/io-xlsx/src/reader/drawing.test.ts', 'packages/sdk/src/visual-inventory.test.ts'],
	),
	cap(
		'visuals.text-boxes',
		'visuals',
		'Text boxes',
		'inspectable',
		'P1',
		'Add safe text replacement with plan-time loss audit for opaque rich text styling.',
		'Text box inventory exposes anchor, name, description, and plain text content.',
		['packages/io-xlsx/src/reader/drawing.test.ts', 'packages/sdk/src/visual-inventory.test.ts'],
	),
	cap(
		'visuals.sparklines',
		'visuals',
		'Sparklines',
		'inspectable',
		'P2',
		'Add safe sparkline source/style edit operations after preservation fixtures.',
		'Sparkline groups are inventoried with type, source range, location range, count, marker flags, and series color while extension XML remains preserved.',
		[
			'packages/io-xlsx/src/reader/advanced-filter-sparkline.test.ts',
			'packages/sdk/src/advanced-filter-sparkline.test.ts',
		],
	),

	cap(
		'analytics.pivots',
		'analytics',
		'Pivot tables',
		'inspectable',
		'P0',
		'Add filter state edits and calculated output refresh after field-level inventory.',
		'Pivot layout is inspectable and inspect payloads expose refresh/stale-output plans; full management and recalculation are not complete.',
		['packages/io-xlsx/src/reader/reader.test.ts', 'packages/sdk/src/pivot-refresh-plan.test.ts'],
	),
	cap(
		'analytics.pivot-caches',
		'analytics',
		'Pivot caches',
		'editable',
		'P0',
		'Add cache-record recalculation and broader Excel fixture coverage.',
		'Pivot cache source and refresh metadata are editable; SDK refresh plans tell agents when Excel refresh is required because output recalculation is not implemented.',
		[
			'packages/io-xlsx/src/reader/reader.test.ts',
			'packages/engine/src/operations.test.ts',
			'packages/sdk/src/pivot-refresh-plan.test.ts',
		],
	),
	cap(
		'analytics.slicers',
		'analytics',
		'Slicers',
		'inspectable',
		'P0',
		'Add safe slicer filter-state edit operations with pivot refresh impact warnings.',
		'Slicer UI parts, cache links, pivot table links, and tabular selected/no-data item states are inspectable while package bytes are preserved.',
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
		'inspectable',
		'P1',
		'Add relationship-level summaries and explicit lossy-write policy gates.',
		'Data model package parts are inventoried with type, content type, relationship type, and relationship counts while Power Pivot/data-model execution remains out of scope.',
		['packages/io-xlsx/src/reader/data-model.test.ts', 'packages/sdk/src/data-model.test.ts'],
	),
	cap(
		'analytics.getpivotdata',
		'analytics',
		'GETPIVOTDATA',
		'inspectable',
		'P1',
		'Resolve output values after headless pivot-cache recalculation exists.',
		'GETPIVOTDATA-style SDK queries resolve matching pivot table/data-field metadata and return explicit warnings that output values are not recalculated headlessly.',
		['packages/sdk/src/get-pivot-data.test.ts'],
	),

	cap(
		'active.vba-macros',
		'active content',
		'VBA macros',
		'inspectable',
		'P0',
		'Add CFB module-name summaries without exposing source code.',
		'VBA project parts are inventoried with opaque byte-size summaries, preserved, and marked with executionPolicy=blocked; macro execution and semantic edits remain blocked by policy.',
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
		'inspectable',
		'P1',
		'Require explicit re-sign or allow-loss policy before writing signed packages.',
		'Digital signature origin/signature parts are inventoried and preserved with an explicit invalidation-risk compatibility note; signatures are not re-signed after edits.',
		['packages/io-xlsx/src/reader/signature.test.ts'],
	),

	cap(
		'connections.external-workbooks',
		'connections',
		'External workbook connections',
		'editable',
		'P0',
		'Add optional resolution hooks for controlled headless external workbook reads.',
		'External link targets are inspectable/editable, and SDK inspect maps formula/name usages before rewrites.',
		[
			'packages/engine/src/operations.test.ts',
			'packages/sdk/src/external-reference-usages.test.ts',
		],
	),
	cap(
		'connections.query-tables',
		'connections',
		'Query tables',
		'inspectable',
		'P1',
		'Add safe refresh metadata edits and table-output impact planning.',
		'Query table parts are inventoried with sheet anchors, connection IDs, refresh flags, and preservation/loss-audit features; query execution remains preserve-first.',
		[
			'packages/io-xlsx/src/reader/connections.test.ts',
			'packages/sdk/src/connection-inventory.test.ts',
		],
	),
	cap(
		'connections.power-query',
		'connections',
		'Power Query',
		'inspectable',
		'P1',
		'Add richer mashup metadata summaries while continuing to block lossy writes.',
		'Power Query mashup/customData parts are inventoried and preserved with explicit compatibility features; Power Query execution is out of scope for the current engine.',
		[
			'packages/io-xlsx/src/reader/connections.test.ts',
			'packages/sdk/src/connection-inventory.test.ts',
		],
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
		'Add larger-session cache telemetry and table-aware changedSince polling.',
		'MCP and SDK expose row-windowed cells, objects, TSV, compact sparse reads, MCP column pruning/header selection, and compact changedSince/changeToken polling.',
		['apps/mcp/src/index.test.ts', 'packages/sdk/src/sdk.test.ts'],
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
		'Expose stable export metadata and preserved-feature warnings.',
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
