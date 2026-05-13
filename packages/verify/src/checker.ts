import type { CellFormulaBinding, RangeRef, Workbook } from '@ascend/core'
import { indexToColumn, parseA1, parseRange, toA1 } from '@ascend/core'
import {
	analyzeWorkbookDependencies,
	analyzeWorkbookFormulas,
	cellHasFormula,
	defaultCalcContext,
	findTableRangeOverlaps,
	parseCellKey,
	recalculate,
	type WorkbookDependencyAnalysis,
	type WorkbookFormulaAnalysis,
} from '@ascend/engine'
import {
	extractRefs,
	type FormulaNode,
	type FormulaRef,
	parseFormula,
	type StructuredRefNode,
} from '@ascend/formulas'
import { isError, levenshtein } from '@ascend/schema'

export interface CheckResult {
	readonly passed: boolean
	readonly issues: readonly CheckIssue[]
}

export interface CheckIssue {
	readonly rule: string
	readonly severity: 'error' | 'warning' | 'info'
	readonly message: string
	readonly refs?: readonly string[]
	readonly suggestedFix?: string
	readonly details?: Readonly<Record<string, unknown>>
}

export interface CheckAnalysis {
	readonly formulas?: WorkbookFormulaAnalysis
	readonly dependencies?: WorkbookDependencyAnalysis
	readonly packageGraph?: VerifyPackageGraph
}

export interface VerifyPackageGraph {
	readonly parts: readonly VerifyPackageGraphPart[]
	readonly relationships: readonly VerifyPackageGraphRelationship[]
	readonly contentTypeDefaults?: readonly VerifyPackageContentTypeDefault[]
	readonly contentTypeOverrides?: readonly VerifyPackageContentTypeOverride[]
}

export interface VerifyPackageGraphPart {
	readonly path: string
	readonly featureFamily?: string
	readonly ownerScope?: string
	readonly contentType?: string
	readonly preservationPolicy?: string
	readonly threadedCommentPersons?: readonly VerifyThreadedCommentPersonEntry[]
}

export interface VerifyThreadedCommentPersonEntry {
	readonly id: string
	readonly displayName?: string
	readonly index: number
}

export interface VerifyPackageGraphRelationship {
	readonly sourcePartPath: string
	readonly relationshipPartPath: string
	readonly id: string
	readonly type: string
	readonly rawType?: string
	readonly rawTarget: string
	readonly resolvedTarget?: string
	readonly targetMode?: string
	readonly featureFamily?: string
}

export interface VerifyPackageContentTypeDefault {
	readonly extension: string
	readonly contentType: string
}

export interface VerifyPackageContentTypeOverride {
	readonly partPath: string
	readonly contentType: string
}

function findClosestSheetName(target: string, sheetNames: readonly string[]): string | null {
	if (sheetNames.length === 0) return null
	let best: string | null = null
	let bestDist = Number.POSITIVE_INFINITY
	const targetLower = target.toLowerCase()
	for (const name of sheetNames) {
		const dist = levenshtein(targetLower, name.toLowerCase())
		if (dist < bestDist) {
			bestDist = dist
			best = name
		}
	}
	if (best !== null && bestDist <= Math.max(target.length, best.length) * 0.5) {
		return best
	}
	return null
}

function checkBrokenRefs(
	_wb: Workbook,
	analysis: WorkbookFormulaAnalysis,
	sheetNames: readonly string[],
): CheckIssue[] {
	const issues: CheckIssue[] = []

	for (const formula of analysis.formulas.values()) {
		if (!formula.ast) continue
		const cellAddr = `${formula.sheetName}!${toA1({ row: formula.row, col: formula.col })}`
		for (const ref of formula.refs) {
			if (ref.kind === 'sheetSpan') {
				const start = analysis.sheetNameIndex.get(ref.startSheet.toLowerCase())
				const end = analysis.sheetNameIndex.get(ref.endSheet.toLowerCase())
				if (start === undefined) {
					const closest = findClosestSheetName(ref.startSheet, sheetNames)
					issues.push({
						rule: 'broken-refs',
						severity: 'error',
						message: `Reference to non-existent sheet "${ref.startSheet}"`,
						refs: [cellAddr],
						...(closest ? { suggestedFix: `Did you mean sheet "${closest}"?` } : {}),
					})
				}
				if (end === undefined) {
					const closest = findClosestSheetName(ref.endSheet, sheetNames)
					issues.push({
						rule: 'broken-refs',
						severity: 'error',
						message: `Reference to non-existent sheet "${ref.endSheet}"`,
						refs: [cellAddr],
						...(closest ? { suggestedFix: `Did you mean sheet "${closest}"?` } : {}),
					})
				}
				if (start !== undefined && end !== undefined && start > end) {
					issues.push({
						rule: 'broken-refs',
						severity: 'error',
						message: `Invalid 3D sheet span "${ref.startSheet}:${ref.endSheet}"`,
						refs: [cellAddr],
					})
				}
				continue
			}
			if (ref.sheet?.startsWith('[')) continue
			if (ref.sheet && !analysis.sheetNameIndex.has(ref.sheet.toLowerCase())) {
				const closest = findClosestSheetName(ref.sheet, sheetNames)
				issues.push({
					rule: 'broken-refs',
					severity: 'error',
					message: `Reference to non-existent sheet "${ref.sheet}"`,
					refs: [cellAddr],
					...(closest ? { suggestedFix: `Did you mean sheet "${closest}"?` } : {}),
				})
			}
		}
	}

	return issues
}

function checkExternalRefs(wb: Workbook, analysis: WorkbookFormulaAnalysis): CheckIssue[] {
	const issues: CheckIssue[] = []

	for (const formula of analysis.formulas.values()) {
		if (!formula.ast) continue
		const cellAddr = `${formula.sheetName}!${toA1({ row: formula.row, col: formula.col })}`
		for (const ref of formula.refs) {
			if (ref.kind === 'sheetSpan') continue
			if (ref.sheet?.startsWith('[')) {
				issues.push({
					rule: 'external-refs',
					severity: 'warning',
					message: `External workbook reference: ${ref.sheet}`,
					refs: [cellAddr],
					suggestedFix:
						'Replace external reference with a local copy of the data or a defined name',
				})
			}
		}
	}

	for (const entry of wb.definedNames.list()) {
		for (const reference of externalReferencesInFormula({
			field: 'formula',
			formula: entry.formula,
		})) {
			issues.push({
				rule: 'external-refs',
				severity: 'warning',
				message: `Defined name "${entry.name}" references external workbook "${reference.externalTarget}"`,
				refs: [entry.name],
				suggestedFix:
					'Replace the external defined-name formula with local workbook data or verify external link metadata before writing.',
				details: {
					kind: 'defined-name-external-reference',
					name: entry.name,
					scope: entry.scope,
					formula: entry.formula,
					externalSheet: reference.sheetName,
					externalTarget: reference.externalTarget,
				},
			})
		}
	}

	return issues
}

function checkCalcFreshness(
	wb: Workbook,
	analysis: WorkbookFormulaAnalysis,
	packageGraph?: VerifyPackageGraph,
): CheckIssue[] {
	const calcChainParts =
		packageGraph?.parts.filter((part) => part.featureFamily === 'preservedCalcChain') ?? []
	if (analysis.formulas.size === 0 && calcChainParts.length === 0) return []

	const reasons: string[] = []
	if (wb.calcSettings.calcMode === 'manual') reasons.push('manual calculation mode')
	if (wb.calcSettings.fullCalcOnLoad) reasons.push('full recalculation requested on load')
	if (wb.calcSettings.calcCompleted === false) reasons.push('calculation not completed')
	if (wb.calcSettings.forceFullCalc === true) reasons.push('forced full recalculation')
	if (reasons.length === 0) return []

	const refs = [...analysis.formulas.values()]
		.slice(0, 5)
		.map((formula) => `${formula.sheetName}!${toA1({ row: formula.row, col: formula.col })}`)
	return [
		{
			rule: 'calc-freshness',
			severity: 'warning',
			message: `Workbook calculation metadata indicates stale formula caches: ${reasons.join(', ')}`,
			...(refs.length > 0 ? { refs } : {}),
			suggestedFix:
				'Recalculate the workbook with Ascend or Excel before trusting cached formula values or preserved calcChain ordering.',
			details: {
				kind: 'stale-calculation-metadata',
				reasons,
				formulaCount: analysis.formulas.size,
				calcChainParts: calcChainParts.map((part) => part.path),
				calcMode: wb.calcSettings.calcMode,
				fullCalcOnLoad: wb.calcSettings.fullCalcOnLoad,
				calcCompleted: wb.calcSettings.calcCompleted,
				calcOnSave: wb.calcSettings.calcOnSave,
				forceFullCalc: wb.calcSettings.forceFullCalc,
			},
		},
	]
}

interface ChartSeriesReferenceEntry {
	readonly field: 'nameRef' | 'categoryRef' | 'valueRef'
	readonly reference: string
}

function chartSeriesReferenceEntries(series: {
	readonly nameRef?: string
	readonly categoryRef?: string
	readonly valueRef?: string
}): ChartSeriesReferenceEntry[] {
	const entries: ChartSeriesReferenceEntry[] = []
	if (series.nameRef) entries.push({ field: 'nameRef', reference: series.nameRef })
	if (series.categoryRef) entries.push({ field: 'categoryRef', reference: series.categoryRef })
	if (series.valueRef) entries.push({ field: 'valueRef', reference: series.valueRef })
	return entries
}

function chartReferenceSheetName(reference: string): { sheetName?: string; external: boolean } {
	const trimmed = reference.trim()
	if (trimmed.startsWith('[')) return { external: true }
	if (trimmed.startsWith("'")) {
		let sheetName = ''
		for (let i = 1; i < trimmed.length; i++) {
			const char = trimmed[i]
			if (char === "'" && trimmed[i + 1] === "'") {
				sheetName += "'"
				i++
				continue
			}
			if (char === "'" && trimmed[i + 1] === '!') {
				return { sheetName, external: sheetName.startsWith('[') }
			}
			sheetName += char
		}
		return { external: false }
	}
	const bang = trimmed.indexOf('!')
	if (bang === -1) return { external: false }
	const sheetName = trimmed.slice(0, bang)
	return { sheetName, external: sheetName.startsWith('[') }
}

function externalTargetFromSheetName(sheetName: string): string {
	const match = /^\[[^\]]+\]/.exec(sheetName)
	return match?.[0] ?? sheetName
}

function checkChartSeriesReferences(wb: Workbook, sheetNames: readonly string[]): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNameSet = new Set(sheetNames.map((name) => name.toLowerCase()))
	const tablesByName = new Map<
		string,
		readonly { readonly name: string; readonly columns: readonly { readonly name: string }[] }[]
	>()
	for (const table of wb.sheets.flatMap((sheet) => sheet.tables)) {
		const normalized = table.name.toLowerCase()
		tablesByName.set(normalized, [...(tablesByName.get(normalized) ?? []), table])
	}
	for (const chart of wb.chartParts) {
		for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
			const series = chart.series[seriesIndex]
			if (!series) continue
			for (const entry of chartSeriesReferenceEntries(series)) {
				const parsed = chartReferenceSheetName(entry.reference)
				if (parsed.external) {
					issues.push({
						rule: 'chart-series-integrity',
						severity: 'warning',
						message: `Chart series ${entry.field} in "${chart.partPath}" references external workbook ${parsed.sheetName ? `"${externalTargetFromSheetName(parsed.sheetName)}"` : 'data'}`,
						refs: [`${chart.partPath}#series${seriesIndex}`],
						suggestedFix:
							'Replace the chart source with local workbook data or verify the external link metadata before editing chart ranges.',
						details: {
							kind: 'chart-series-external-reference',
							partPath: chart.partPath,
							seriesIndex,
							field: entry.field,
							reference: entry.reference,
							...(parsed.sheetName ? { externalSheet: parsed.sheetName } : {}),
							...(parsed.sheetName
								? { externalTarget: externalTargetFromSheetName(parsed.sheetName) }
								: {}),
							...(chart.sheetName ? { ownerSheet: chart.sheetName } : {}),
							...(chart.chartType ? { chartType: chart.chartType } : {}),
						},
					})
					continue
				}
				if (parsed.sheetName && !sheetNameSet.has(parsed.sheetName.toLowerCase())) {
					const closest = findClosestSheetName(parsed.sheetName, sheetNames)
					issues.push({
						rule: 'chart-series-integrity',
						severity: 'warning',
						message: `Chart series ${entry.field} in "${chart.partPath}" references non-existent sheet "${parsed.sheetName}"`,
						refs: [`${chart.partPath}#series${seriesIndex}`],
						suggestedFix: closest
							? `Did you mean sheet "${closest}"?`
							: 'Repair the chart series source reference before editing chart data ranges.',
						details: {
							partPath: chart.partPath,
							seriesIndex,
							field: entry.field,
							reference: entry.reference,
							sheetName: parsed.sheetName,
							...(chart.sheetName ? { ownerSheet: chart.sheetName } : {}),
							...(chart.chartType ? { chartType: chart.chartType } : {}),
						},
					})
					continue
				}
				for (const structuredRef of chartStructuredReferences(entry.reference)) {
					const tableMatches = tablesByName.get(structuredRef.tableName.toLowerCase()) ?? []
					if (tableMatches.length === 0) {
						issues.push({
							rule: 'chart-series-integrity',
							severity: 'warning',
							message: `Chart series ${entry.field} in "${chart.partPath}" references non-existent table "${structuredRef.tableName}"`,
							refs: [`${chart.partPath}#series${seriesIndex}`],
							suggestedFix:
								'Repair the chart source table reference before editing chart data ranges.',
							details: {
								kind: 'chart-series-missing-table-reference',
								partPath: chart.partPath,
								seriesIndex,
								field: entry.field,
								reference: entry.reference,
								tableName: structuredRef.tableName,
								...(chart.sheetName ? { ownerSheet: chart.sheetName } : {}),
								...(chart.chartType ? { chartType: chart.chartType } : {}),
							},
						})
						continue
					}
					const missingColumns = missingChartStructuredReferenceColumns(structuredRef, tableMatches)
					if (missingColumns.length === 0) continue
					issues.push({
						rule: 'chart-series-integrity',
						severity: 'warning',
						message: `Chart series ${entry.field} in "${chart.partPath}" references missing table column(s) ${missingColumns.map((column) => `"${column}"`).join(', ')} in "${structuredRef.tableName}"`,
						refs: [`${chart.partPath}#series${seriesIndex}`],
						suggestedFix:
							'Repair the chart source structured reference before editing chart data ranges.',
						details: {
							kind: 'chart-series-missing-table-column-reference',
							partPath: chart.partPath,
							seriesIndex,
							field: entry.field,
							reference: entry.reference,
							tableName: structuredRef.tableName,
							missingColumns,
							...(structuredRef.column ? { column: structuredRef.column } : {}),
							...(structuredRef.endColumn ? { endColumn: structuredRef.endColumn } : {}),
							...(chart.sheetName ? { ownerSheet: chart.sheetName } : {}),
							...(chart.chartType ? { chartType: chart.chartType } : {}),
						},
					})
				}
			}
		}
	}
	return issues
}

interface ChartStructuredReference {
	readonly tableName: string
	readonly column?: string
	readonly endColumn?: string
}

function chartStructuredReferences(reference: string): ChartStructuredReference[] {
	const parsed = parseFormulaText(reference)
	if (parsed.ok) {
		return structuredReferencesForFormula(parsed.value).flatMap((structuredRef) =>
			structuredRef.table
				? [
						{
							tableName: structuredRef.table,
							...(structuredRef.column ? { column: structuredRef.column } : {}),
							...(structuredRef.endColumn ? { endColumn: structuredRef.endColumn } : {}),
						},
					]
				: [],
		)
	}
	const local = reference.trim().replace(/^=/, '').split('!').pop()?.trim() ?? ''
	const match = /^'?([A-Za-z_\\][\w.]*)'?\s*\[/.exec(local)
	return match?.[1] ? [{ tableName: match[1] }] : []
}

function missingChartStructuredReferenceColumns(
	reference: ChartStructuredReference,
	tableMatches: readonly { readonly columns: readonly { readonly name: string }[] }[],
): string[] {
	const referencedColumns = [reference.column, reference.endColumn].filter(
		(column): column is string => column !== undefined && column.length > 0,
	)
	if (referencedColumns.length === 0) return []
	const availableColumns = new Set(
		tableMatches.flatMap((table) => table.columns.map((column) => column.name.toLowerCase())),
	)
	return referencedColumns.filter((column) => !availableColumns.has(column.toLowerCase()))
}

function checkChartPartOwnership(
	wb: Workbook,
	sheetNames: readonly string[],
	packageGraph?: VerifyPackageGraph,
): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNameSet = new Set(sheetNames.map((name) => name.toLowerCase()))
	const chartSheetOwnerByPartPath = new Map<string, string>()
	for (const chartSheet of wb.chartSheets) {
		for (const partPath of chartSheet.chartPartPaths) {
			if (!chartSheetOwnerByPartPath.has(partPath)) {
				chartSheetOwnerByPartPath.set(partPath, chartSheet.name)
			}
		}
	}
	for (const chart of wb.chartParts) {
		const chartSheetOwner = chartSheetOwnerByPartPath.get(chart.partPath)
		if (chart.sheetName && sheetNameSet.has(chart.sheetName.toLowerCase()) && chartSheetOwner) {
			issues.push({
				rule: 'chart-part-ownership',
				severity: 'warning',
				message: `Chart part "${chart.partPath}" is attributed to both worksheet "${chart.sheetName}" and chartsheet "${chartSheetOwner}"`,
				refs: [chart.partPath],
				suggestedFix:
					'Choose either worksheet drawing ownership or chartsheet ownership before editing this chart.',
				details: {
					kind: 'chart-worksheet-chartsheet-owner-ambiguity',
					partPath: chart.partPath,
					ownerSheet: chart.sheetName,
					ownerChartSheet: chartSheetOwner,
					...(chart.chartType ? { chartType: chart.chartType } : {}),
				},
			})
			continue
		}
		if (chart.sheetName && sheetNameSet.has(chart.sheetName.toLowerCase())) continue
		if (!chart.sheetName && chartSheetOwner) continue
		if (chart.sheetName && !sheetNameSet.has(chart.sheetName.toLowerCase())) {
			const closest = findClosestSheetName(chart.sheetName, sheetNames)
			issues.push({
				rule: 'chart-part-ownership',
				severity: 'warning',
				message: `Chart part "${chart.partPath}" is attributed to non-existent sheet "${chart.sheetName}"`,
				refs: [chart.partPath],
				suggestedFix: closest
					? `Did you mean sheet "${closest}"?`
					: 'Inspect drawing and chartsheet relationships before editing this chart.',
				details: {
					partPath: chart.partPath,
					ownerSheet: chart.sheetName,
					...(chart.chartType ? { chartType: chart.chartType } : {}),
				},
			})
			continue
		}
		issues.push({
			rule: 'chart-part-ownership',
			severity: 'warning',
			message: `Chart part "${chart.partPath}" is not attributed to a worksheet or chartsheet`,
			refs: [chart.partPath],
			suggestedFix: 'Inspect drawing and chartsheet relationships before editing this chart.',
			details: {
				partPath: chart.partPath,
				...(chart.chartType ? { chartType: chart.chartType } : {}),
			},
		})
	}
	if (packageGraph) {
		const graphPartByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
		const chartModelPartPaths = new Set(wb.chartParts.map((chart) => chart.partPath))
		const graphRelationshipsByTarget = relationshipsByTarget(packageGraph.relationships)
		for (const chart of wb.chartParts) {
			if (graphPartByPath.has(chart.partPath)) continue
			issues.push({
				rule: 'chart-part-ownership',
				severity: 'error',
				message: `Chart model references missing package part "${chart.partPath}"`,
				refs: [chart.partPath],
				suggestedFix:
					'Restore the chart package part or remove the stale chart metadata before writing.',
				details: {
					kind: 'missing-chart-package-part',
					partPath: chart.partPath,
					...(chart.sheetName ? { ownerSheet: chart.sheetName } : {}),
					...(chart.chartType ? { chartType: chart.chartType } : {}),
				},
			})
		}
		for (const part of packageGraph.parts) {
			if (!isChartPackagePart(part)) continue
			if (chartModelPartPaths.has(part.path)) continue
			issues.push({
				rule: 'chart-part-ownership',
				severity: 'warning',
				message: `Chart package part "${part.path}" is not claimed by the workbook chart model`,
				refs: [part.path],
				suggestedFix:
					'Reconnect the chart part through its drawing or chartsheet owner, or remove the orphan chart sidecar before writing.',
				details: {
					kind: 'orphan-chart-package-part',
					partPath: part.path,
					ownerScope: part.ownerScope,
					contentType: part.contentType,
					incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
						packageRelationshipDetails,
					),
				},
			})
		}
	}
	return issues
}

function checkCircularRefs(wb: Workbook, analysis: WorkbookDependencyAnalysis): CheckIssue[] {
	return analysis.cycles.map((cycle) => {
		const refs = cycle.map((key) => {
			const [si, row, col] = parseCellKey(key)
			const s = wb.sheets[si]
			const sheetName = s ? s.name : `Sheet${si}`
			return `${sheetName}!${toA1({ row, col })}`
		})
		return {
			rule: 'circular-refs',
			severity: 'error' as const,
			message: `Circular reference detected involving ${refs.length} cell(s)`,
			refs,
			suggestedFix: `Break the cycle by removing one of the references: ${refs.join(' → ')} → ${refs[0]}`,
		}
	})
}

function suggestedFixForError(errorType: string): string | null {
	switch (errorType) {
		case '#REF!':
			return 'Check that all referenced cells and ranges still exist; a row, column, or sheet may have been deleted'
		case '#NAME?':
			return 'Check for misspelled function names or undefined named ranges'
		case '#DIV/0!':
			return 'Add a check for zero before dividing (e.g. IF(B1=0, 0, A1/B1))'
		default:
			return null
	}
}

function checkFormulaErrors(wb: Workbook, analysis: WorkbookFormulaAnalysis): CheckIssue[] {
	const issues: CheckIssue[] = []

	for (const formula of analysis.formulas.values()) {
		const sheet = wb.sheets[formula.sheetIndex]
		const cell = sheet?.cells.get(formula.row, formula.col)
		if (!sheet || !cell || !cellHasFormula(cell)) continue
		if (isError(cell.value)) {
			const errorType = cell.value.value
			const fix = suggestedFixForError(errorType)
			issues.push({
				rule: 'formula-errors',
				severity: 'warning',
				message: `Formula evaluates to ${errorType}`,
				refs: [`${sheet.name}!${toA1({ row: formula.row, col: formula.col })}`],
				...(fix ? { suggestedFix: fix } : {}),
			})
		}
	}

	return issues
}

function blockedSpillIssue(
	sheetName: string,
	row: number,
	col: number,
	binding: Extract<CellFormulaBinding, { kind: 'blockedSpill' }>,
): CheckIssue {
	const anchorRef = `${sheetName}!${toA1({ row, col })}`
	const blockingRefs = binding.blockingRefs.map((ref) => `${sheetName}!${ref}`)
	return {
		rule: 'spill-diagnostics',
		severity: 'warning',
		message: `Formula spill is blocked by ${blockingRefs.length} occupied cell(s)`,
		refs: [anchorRef, ...blockingRefs],
		suggestedFix: `Clear or move the blocking cell(s): ${blockingRefs.join(', ')}`,
		details: {
			error: '#SPILL!',
			cause: 'occupied-cell',
			spillRange: `${sheetName}!${binding.ref}`,
			blockingRefs,
		},
	}
}

function unknownSpillIssue(sheetName: string, row: number, col: number): CheckIssue {
	return {
		rule: 'spill-diagnostics',
		severity: 'warning',
		message: 'Formula spill is blocked, but the blocking range was not captured',
		refs: [`${sheetName}!${toA1({ row, col })}`],
		suggestedFix: 'Recalculate the workbook to refresh spill-block diagnostics.',
		details: { error: '#SPILL!', cause: 'unknown' },
	}
}

function checkBlockedSpills(wb: Workbook, formulas: WorkbookFormulaAnalysis): CheckIssue[] {
	const issues: CheckIssue[] = []
	const unknownSpills: Array<{ sheetIndex: number; row: number; col: number }> = []
	for (const formula of formulas.formulas.values()) {
		const sheetIndex = formula.sheetIndex
		const sheet = wb.sheets[sheetIndex]
		if (!sheet) continue
		const cell = sheet.cells.get(formula.row, formula.col)
		if (!cell || !cellHasFormula(cell) || !isError(cell.value) || cell.value.value !== '#SPILL!')
			continue
		const binding = cell.formulaInfo
		if (binding?.kind === 'blockedSpill') {
			issues.push(blockedSpillIssue(sheet.name, formula.row, formula.col, binding))
			continue
		}
		unknownSpills.push({ sheetIndex, row: formula.row, col: formula.col })
	}
	if (unknownSpills.length > 0) {
		const recalculated = wb.clone()
		recalculate(
			recalculated,
			defaultCalcContext({
				dateSystem: recalculated.calcSettings.dateSystem,
				iterativeCalc: recalculated.calcSettings.iterativeCalc,
			}),
		)
		for (const spill of unknownSpills) {
			const sheet = wb.sheets[spill.sheetIndex]
			const refreshedSheet = recalculated.sheets[spill.sheetIndex]
			if (!sheet || !refreshedSheet) continue
			const refreshed = refreshedSheet.cells.readFormulaInfo(spill.row, spill.col)
			if (refreshed?.kind === 'blockedSpill') {
				issues.push(blockedSpillIssue(sheet.name, spill.row, spill.col, refreshed))
				continue
			}
			issues.push(unknownSpillIssue(sheet.name, spill.row, spill.col))
		}
	}
	return issues
}

function checkOrphanedNames(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNames = new Set(wb.sheets.map((s) => s.name.toLowerCase()))
	const sheetNameList = wb.sheets.map((s) => s.name)

	for (const entry of wb.definedNames.list()) {
		const name = entry.name
		const ref = entry.formula
		for (const missing of missingSheetReferencesInFormula(
			{ field: 'formula', formula: ref },
			sheetNames,
		)) {
			const closest = findClosestSheetName(missing.sheetName, sheetNameList)
			issues.push({
				rule: 'orphaned-names',
				severity: 'warning',
				message: `Defined name "${name}" references non-existent sheet "${missing.sheetName}"`,
				refs: [ref],
				suggestedFix: closest
					? `Did you mean sheet "${closest}"?`
					: 'Repair or remove the defined name before using it in formulas or write workflows.',
				details: {
					kind: 'defined-name-missing-sheet-reference',
					name,
					scope: entry.scope,
					field: missing.field,
					formula: ref,
					missingSheet: missing.sheetName,
				},
			})
		}
	}

	return issues
}

function rangesOverlap2D(a: RangeRef, b: RangeRef): boolean {
	const r1 = Math.min(a.start.row, a.end.row)
	const r2 = Math.max(a.start.row, a.end.row)
	const c1 = Math.min(a.start.col, a.end.col)
	const c2 = Math.max(a.start.col, a.end.col)
	const q1 = Math.min(b.start.row, b.end.row)
	const q2 = Math.max(b.start.row, b.end.row)
	const d1 = Math.min(b.start.col, b.end.col)
	const d2 = Math.max(b.start.col, b.end.col)
	return r1 <= q2 && q1 <= r2 && c1 <= d2 && d1 <= c2
}

function checkMergeOverlaps(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	for (const sheet of wb.sheets) {
		const merges = sheet.merges
		for (let i = 0; i < merges.length; i++) {
			const a = merges[i]
			if (!a) continue
			for (let j = i + 1; j < merges.length; j++) {
				const b = merges[j]
				if (!b) continue
				if (rangesOverlap2D(a, b)) {
					const ra = `${indexToColumn(a.start.col)}${a.start.row + 1}:${indexToColumn(a.end.col)}${a.end.row + 1}`
					const rb = `${indexToColumn(b.start.col)}${b.start.row + 1}:${indexToColumn(b.end.col)}${b.end.row + 1}`
					issues.push({
						rule: 'merge-overlap',
						severity: 'warning',
						message: `Overlapping merged ranges on sheet "${sheet.name}"`,
						refs: [`${sheet.name}!${ra}`, `${sheet.name}!${rb}`],
						suggestedFix: 'Unmerge or resize one of the ranges so they do not intersect',
					})
				}
			}
		}
	}
	return issues
}

function checkTableIntegrity(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	const tablesByName = new Map<string, TableIntegrityEntry>()
	const tablesById = new Map<string, TableIntegrityEntry>()

	for (const sheet of wb.sheets) {
		for (const table of sheet.tables) {
			const rangeWidth = table.ref.end.col - table.ref.start.col + 1
			const ref = rangeToA1(table.ref)
			if (table.columns.length !== rangeWidth) {
				issues.push({
					rule: 'table-integrity',
					severity: 'error',
					message: `Table "${table.name}" has ${table.columns.length} columns but range spans ${rangeWidth}`,
					refs: [`${sheet.name}!${ref}`],
					details: {
						kind: 'table-column-count-mismatch',
						tableName: table.name,
						rangeWidth,
						columnCount: table.columns.length,
						partPath: table.partPath,
					},
				})
			}

			const columnsByName = new Map<string, TableColumnIntegrityEntry>()
			for (let columnIndex = 0; columnIndex < table.columns.length; columnIndex++) {
				const column = table.columns[columnIndex]
				if (!column) continue
				const normalizedColumnName = column.name.toLowerCase()
				const columnEntry: TableColumnIntegrityEntry = {
					columnName: column.name,
					columnIndex,
					ref: table.hasHeaders
						? toA1({ row: table.ref.start.row, col: table.ref.start.col + columnIndex })
						: ref,
				}
				const existingColumn = columnsByName.get(normalizedColumnName)
				if (existingColumn) {
					issues.push({
						rule: 'table-integrity',
						severity: 'error',
						message: `Table "${table.name}" has duplicate column name "${column.name}"`,
						refs: table.hasHeaders
							? [
									`${sheet.name}!${ref}`,
									`${sheet.name}!${existingColumn.ref}`,
									`${sheet.name}!${columnEntry.ref}`,
								]
							: [`${sheet.name}!${ref}`],
						suggestedFix:
							'Rename one of the duplicate table columns before using structured references or table column edit operations.',
						details: {
							kind: 'duplicate-table-column-name',
							tableName: table.name,
							sheetName: sheet.name,
							ref,
							normalizedName: normalizedColumnName,
							first: existingColumn,
							duplicate: columnEntry,
							partPath: table.partPath,
						},
					})
				} else {
					columnsByName.set(normalizedColumnName, columnEntry)
				}

				const formulaEntries: X14FormulaReferenceEntry[] = []
				if (column.formula) {
					formulaEntries.push({ field: `columns[${columnIndex}].formula`, formula: column.formula })
				}
				if (column.totalsRowFormula) {
					formulaEntries.push({
						field: `columns[${columnIndex}].totalsRowFormula`,
						formula: column.totalsRowFormula,
					})
				}
				pushExternalMetadataReferenceIssues(issues, {
					rule: 'table-integrity',
					source: 'tableColumn',
					sheetName: sheet.name,
					index: columnIndex,
					references: formulaEntries.flatMap(externalReferencesInFormula),
					refs: [`${sheet.name}!${ref}`],
					suggestedFix:
						'Replace table column formulas with local workbook data or verify the external link metadata before editing table formulas.',
					details: {
						tableName: table.name,
						ref,
						columnName: column.name,
						...(table.partPath ? { partPath: table.partPath } : {}),
					},
				})
			}

			const entry: TableIntegrityEntry = {
				tableName: table.name,
				tableId: String(table.id),
				sheetName: sheet.name,
				ref,
				...(table.partPath ? { partPath: table.partPath } : {}),
			}
			const normalizedName = table.name.toLowerCase()
			const existingName = tablesByName.get(normalizedName)
			if (existingName) {
				issues.push({
					rule: 'table-integrity',
					severity: 'error',
					message: `Duplicate table name "${table.name}" conflicts with "${existingName.tableName}"`,
					refs: [`${existingName.sheetName}!${existingName.ref}`, `${sheet.name}!${ref}`],
					suggestedFix:
						'Rename one table before using structured references or table edit operations.',
					details: {
						kind: 'duplicate-table-name',
						normalizedName,
						first: existingName,
						duplicate: entry,
					},
				})
			} else {
				tablesByName.set(normalizedName, entry)
			}

			const existingId = tablesById.get(entry.tableId)
			if (existingId) {
				issues.push({
					rule: 'table-integrity',
					severity: 'error',
					message: `Duplicate table id "${entry.tableId}" on table "${table.name}"`,
					refs: [`${existingId.sheetName}!${existingId.ref}`, `${sheet.name}!${ref}`],
					suggestedFix:
						'Regenerate table ids so every table part has a workbook-unique identifier.',
					details: {
						kind: 'duplicate-table-id',
						tableId: entry.tableId,
						first: existingId,
						duplicate: entry,
					},
				})
			} else {
				tablesById.set(entry.tableId, entry)
			}
		}

		for (const overlap of findTableRangeOverlaps(sheet)) {
			const leftRef = rangeToA1(overlap.leftRef)
			const rightRef = rangeToA1(overlap.rightRef)
			issues.push({
				rule: 'table-integrity',
				severity: 'error',
				message: `Table "${overlap.left.name}" overlaps table "${overlap.right.name}" on sheet "${sheet.name}"`,
				refs: [`${sheet.name}!${leftRef}`, `${sheet.name}!${rightRef}`],
				suggestedFix:
					'Resize, move, or delete one table so worksheet cells have unambiguous table ownership.',
				details: {
					kind: 'overlapping-table-ranges',
					left: {
						tableName: overlap.left.name,
						ref: leftRef,
						partPath: overlap.left.partPath,
					},
					right: {
						tableName: overlap.right.name,
						ref: rightRef,
						partPath: overlap.right.partPath,
					},
				},
			})
		}
	}

	return issues
}

function checkTablePackageGraphIntegrity(
	wb: Workbook,
	packageGraph?: VerifyPackageGraph,
): CheckIssue[] {
	if (!packageGraph) return []
	const issues: CheckIssue[] = []
	const graphPartByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
	const graphRelationshipsByTarget = relationshipsByTarget(packageGraph.relationships)
	const claimedTableParts = new Map<string, TableIntegrityEntry>()

	for (const sheet of wb.sheets) {
		for (const table of sheet.tables) {
			if (!table.partPath) continue
			const tableRef = rangeToA1(table.ref)
			const entry: TableIntegrityEntry = {
				tableName: table.name,
				tableId: String(table.id),
				sheetName: sheet.name,
				ref: tableRef,
				partPath: table.partPath,
			}
			const existingClaim = claimedTableParts.get(table.partPath)
			if (existingClaim) {
				issues.push({
					rule: 'table-package-integrity',
					severity: 'error',
					message: `Table package part "${table.partPath}" is claimed by multiple workbook tables`,
					refs: [
						`${existingClaim.sheetName}!${existingClaim.ref}`,
						`${sheet.name}!${tableRef}`,
						table.partPath,
					],
					suggestedFix:
						'Give each table a distinct worksheet table relationship before editing table topology.',
					details: {
						kind: 'duplicate-table-part-owner',
						first: existingClaim,
						duplicate: entry,
					},
				})
			} else {
				claimedTableParts.set(table.partPath, entry)
			}

			const graphPart = graphPartByPath.get(table.partPath)
			if (!graphPart) {
				issues.push({
					rule: 'table-package-integrity',
					severity: 'error',
					message: `Table "${table.name}" references missing package part "${table.partPath}"`,
					refs: [`${sheet.name}!${tableRef}`, table.partPath],
					suggestedFix:
						'Restore the table package part or remove the stale table relationship before writing.',
					details: {
						kind: 'missing-table-part',
						table: entry,
					},
				})
				continue
			}

			if (!table.sourcePartPath || !table.sourceRelationshipId) continue
			const incomingRelationships = graphRelationshipsByTarget.get(table.partPath) ?? []
			const matchingRelationship = incomingRelationships.find(
				(relationship) =>
					relationship.sourcePartPath === table.sourcePartPath &&
					relationship.id === table.sourceRelationshipId &&
					isTableRelationshipType(relationship.type),
			)
			if (!matchingRelationship) {
				issues.push({
					rule: 'table-package-integrity',
					severity: 'error',
					message: `Table "${table.name}" relationship "${table.sourceRelationshipId}" does not bind to "${table.partPath}" in the package graph`,
					refs: [`${sheet.name}!${tableRef}`, table.partPath],
					suggestedFix:
						'Repair the worksheet table relationship id/target binding before editing table structure.',
					details: {
						kind: 'table-relationship-binding-mismatch',
						table: {
							...entry,
							sourcePartPath: table.sourcePartPath,
							sourceRelationshipPart: table.sourceRelationshipPart,
							sourceRelationshipId: table.sourceRelationshipId,
							sourceRelationshipType: table.sourceRelationshipType,
							sourceRelationshipRawTarget: table.sourceRelationshipRawTarget,
							sourceRelationshipResolvedTarget: table.sourceRelationshipResolvedTarget,
						},
						graphPart,
						incomingRelationships: incomingRelationships.map(packageRelationshipDetails),
					},
				})
			}
		}
	}

	for (const part of packageGraph.parts) {
		if (!isTablePart(part)) continue
		if (claimedTableParts.has(part.path)) continue
		issues.push({
			rule: 'table-package-integrity',
			severity: 'warning',
			message: `Table package part "${part.path}" is not claimed by any workbook table model`,
			refs: [part.path],
			suggestedFix:
				'Inspect worksheet table relationships and reconnect or intentionally remove the orphan table sidecar before table edits.',
			details: {
				kind: 'orphan-table-part',
				partPath: part.path,
				ownerScope: part.ownerScope,
				contentType: part.contentType,
				incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
					packageRelationshipDetails,
				),
			},
		})
	}

	return issues
}

function checkTableQueryTableIntegrity(
	wb: Workbook,
	packageGraph?: VerifyPackageGraph,
): CheckIssue[] {
	const issues: CheckIssue[] = []
	const graphPartByPath = new Map(packageGraph?.parts.map((part) => [part.path, part]) ?? [])
	const graphRelationshipsByTarget = new Map<string, VerifyPackageGraphRelationship[]>()
	for (const relationship of packageGraph?.relationships ?? []) {
		if (!relationship.resolvedTarget) continue
		const relationships = graphRelationshipsByTarget.get(relationship.resolvedTarget)
		if (relationships) relationships.push(relationship)
		else graphRelationshipsByTarget.set(relationship.resolvedTarget, [relationship])
	}
	const workbookConnectionsById = workbookConnectionPartsById(wb)
	const claimedQueryParts = new Map<string, TableQueryTableIntegrityEntry>()

	for (const sheet of wb.sheets) {
		for (const table of sheet.tables) {
			const queryTable = table.queryTable
			if (!queryTable) continue
			const tableRef = rangeToA1(table.ref)
			const entry: TableQueryTableIntegrityEntry = {
				tableName: table.name,
				sheetName: sheet.name,
				ref: tableRef,
				queryTablePartPath: queryTable.partPath,
				relationshipId: queryTable.relationshipId,
				...(table.partPath ? { tablePartPath: table.partPath } : {}),
			}

			if (!table.partPath) {
				issues.push({
					rule: 'table-query-integrity',
					severity: 'error',
					message: `Table "${table.name}" has a queryTable sidecar but no table part path`,
					refs: [`${sheet.name}!${tableRef}`, queryTable.partPath],
					suggestedFix:
						'Preserve or repair the table part path before editing queryTable-backed table data.',
					details: {
						kind: 'query-table-missing-table-part',
						table: entry,
					},
				})
			}
			if (table.tableType && table.tableType !== 'queryTable') {
				issues.push({
					rule: 'table-query-integrity',
					severity: 'warning',
					message: `Table "${table.name}" owns queryTable part "${queryTable.partPath}" but tableType is "${table.tableType}"`,
					refs: [`${sheet.name}!${tableRef}`, queryTable.partPath],
					suggestedFix:
						'Confirm whether this table is query-backed before resizing or refreshing it.',
					details: {
						kind: 'query-table-type-mismatch',
						table: entry,
						tableType: table.tableType,
					},
				})
			}
			if (!isQueryTableRelationshipType(queryTable.relationshipType)) {
				issues.push({
					rule: 'table-query-integrity',
					severity: 'warning',
					message: `Table "${table.name}" queryTable relationship "${queryTable.relationshipId}" has unexpected type "${queryTable.relationshipType}"`,
					refs: [`${sheet.name}!${tableRef}`, queryTable.partPath],
					suggestedFix:
						'Inspect the table relationship type dialect before preserving queryTable sidecars.',
					details: {
						kind: 'query-table-relationship-type-mismatch',
						table: entry,
						relationshipType: queryTable.relationshipType,
						...(queryTable.relationshipRawType
							? { relationshipRawType: queryTable.relationshipRawType }
							: {}),
					},
				})
			}
			const fieldIds = new Map<number, TableColumnQueryFieldEntry>()
			const missingFieldIdColumns: TableColumnQueryFieldEntry[] = []
			for (let columnIndex = 0; columnIndex < table.columns.length; columnIndex++) {
				const column = table.columns[columnIndex]
				if (!column) continue
				const columnEntry: TableColumnQueryFieldEntry = {
					columnName: column.name,
					columnIndex,
					ref: table.hasHeaders
						? `${sheet.name}!${toA1({
								row: table.ref.start.row,
								col: table.ref.start.col + columnIndex,
							})}`
						: `${sheet.name}!${tableRef}`,
				}
				if (column.queryTableFieldId === undefined) {
					missingFieldIdColumns.push(columnEntry)
					continue
				}
				const existingColumn = fieldIds.get(column.queryTableFieldId)
				if (existingColumn) {
					issues.push({
						rule: 'table-query-integrity',
						severity: 'error',
						message: `Table "${table.name}" has duplicate queryTableFieldId "${column.queryTableFieldId}"`,
						refs: [
							`${sheet.name}!${tableRef}`,
							existingColumn.ref,
							columnEntry.ref,
							queryTable.partPath,
						],
						suggestedFix:
							'Repair the tableColumn queryTableFieldId bindings before editing query-backed table columns.',
						details: {
							kind: 'duplicate-query-table-field-id',
							table: entry,
							queryTableFieldId: column.queryTableFieldId,
							first: existingColumn,
							duplicate: columnEntry,
						},
					})
				} else {
					fieldIds.set(column.queryTableFieldId, columnEntry)
				}
			}
			if (missingFieldIdColumns.length > 0) {
				issues.push({
					rule: 'table-query-integrity',
					severity: 'error',
					message: `Table "${table.name}" is queryTable-backed but ${missingFieldIdColumns.length} table column(s) lack queryTableFieldId bindings`,
					refs: [`${sheet.name}!${tableRef}`, queryTable.partPath],
					suggestedFix:
						'Restore tableColumn queryTableFieldId bindings or rebuild the queryTable sidecar before editing query-backed table columns.',
					details: {
						kind: 'missing-query-table-field-id',
						table: entry,
						columns: missingFieldIdColumns,
					},
				})
			}

			const existingClaim = claimedQueryParts.get(queryTable.partPath)
			if (existingClaim) {
				issues.push({
					rule: 'table-query-integrity',
					severity: 'error',
					message: `QueryTable part "${queryTable.partPath}" is claimed by multiple tables`,
					refs: [
						`${existingClaim.sheetName}!${existingClaim.ref}`,
						`${sheet.name}!${tableRef}`,
						queryTable.partPath,
					],
					suggestedFix:
						'Give each query-backed table a distinct queryTable part relationship before editing table topology.',
					details: {
						kind: 'duplicate-query-table-part-owner',
						first: existingClaim,
						duplicate: entry,
					},
				})
			} else {
				claimedQueryParts.set(queryTable.partPath, entry)
			}

			if (!packageGraph) continue
			const graphPart = graphPartByPath.get(queryTable.partPath)
			if (!graphPart) {
				issues.push({
					rule: 'table-query-integrity',
					severity: 'error',
					message: `Table "${table.name}" references missing queryTable package part "${queryTable.partPath}"`,
					refs: [`${sheet.name}!${tableRef}`, queryTable.partPath],
					suggestedFix:
						'Restore the queryTable package part or remove the stale queryTable relationship before writing.',
					details: {
						kind: 'missing-query-table-part',
						table: entry,
					},
				})
				continue
			}
			const incomingRelationships = graphRelationshipsByTarget.get(queryTable.partPath) ?? []
			const matchingRelationship = incomingRelationships.find(
				(relationship) =>
					relationship.sourcePartPath === table.partPath &&
					relationship.id === queryTable.relationshipId,
			)
			if (!matchingRelationship) {
				issues.push({
					rule: 'table-query-integrity',
					severity: 'error',
					message: `Table "${table.name}" queryTable relationship "${queryTable.relationshipId}" does not bind to "${queryTable.partPath}" in the package graph`,
					refs: [`${sheet.name}!${tableRef}`, queryTable.partPath],
					suggestedFix:
						'Repair the table relationship id/target binding before editing query-backed table structure.',
					details: {
						kind: 'query-table-relationship-binding-mismatch',
						table: entry,
						graphPart,
						incomingRelationships: incomingRelationships.map(queryTableRelationshipDetails),
					},
				})
			}
		}
	}
	for (const part of wb.connectionParts) {
		if (part.kind !== 'queryTable') continue
		issues.push(...queryTableConnectionIdIssues(part, workbookConnectionsById))
		if (claimedQueryParts.has(part.partPath)) continue
		const sheet = part.sheetName
			? wb.sheets.find((entry) => entry.name === part.sheetName)
			: undefined
		const entry = worksheetQueryTableIntegrityEntry(part)
		if (part.relType && !isQueryTableRelationshipType(part.relType)) {
			issues.push({
				rule: 'table-query-integrity',
				severity: 'warning',
				message: `Worksheet queryTable part "${part.partPath}" has unexpected relationship type "${part.relType}"`,
				refs: [part.sheetName ?? part.partPath, part.partPath],
				suggestedFix:
					'Inspect the worksheet queryTable relationship type dialect before preserving queryTable sidecars.',
				details: {
					kind: 'worksheet-query-table-relationship-type-mismatch',
					connectionPart: entry,
					relationshipType: part.relType,
				},
			})
		}
		if (packageGraph) {
			const graphPart = graphPartByPath.get(part.partPath)
			if (!graphPart) {
				issues.push({
					rule: 'table-query-integrity',
					severity: 'error',
					message: `Worksheet queryTable inventory references missing package part "${part.partPath}"`,
					refs: [part.sheetName ?? part.partPath, part.partPath],
					suggestedFix:
						'Restore the worksheet queryTable package part or remove the stale queryTable inventory before writing.',
					details: {
						kind: 'missing-worksheet-query-table-part',
						connectionPart: entry,
					},
				})
			} else {
				const incomingRelationships = graphRelationshipsByTarget.get(part.partPath) ?? []
				const matchingRelationship = incomingRelationships.find((relationship) => {
					if (!isQueryTableRelationshipType(relationship.type)) return false
					if (sheet?.preservedXml?.partPath) {
						return relationship.sourcePartPath === sheet.preservedXml.partPath
					}
					return relationship.sourcePartPath.includes('/worksheets/')
				})
				if (!matchingRelationship) {
					issues.push({
						rule: 'table-query-integrity',
						severity: 'error',
						message: `Worksheet queryTable inventory does not bind "${part.partPath}" to an owning worksheet relationship`,
						refs: [part.sheetName ?? part.partPath, part.partPath],
						suggestedFix:
							'Repair the worksheet relationship target for this queryTable part before editing query-backed data.',
						details: {
							kind: 'worksheet-query-table-relationship-binding-mismatch',
							connectionPart: entry,
							graphPart,
							incomingRelationships: incomingRelationships.map(queryTableRelationshipDetails),
						},
					})
				}
			}
		}
		claimedQueryParts.set(part.partPath, {
			tableName: part.name ?? '(worksheet queryTable)',
			sheetName: part.sheetName ?? '(unknown sheet)',
			ref: part.partPath,
			queryTablePartPath: part.partPath,
			relationshipId: '(worksheet relationship)',
		})
	}

	if (packageGraph) {
		for (const part of packageGraph.parts) {
			if (part.featureFamily !== 'preservedQueryTable') continue
			if (claimedQueryParts.has(part.path)) continue
			issues.push({
				rule: 'table-query-integrity',
				severity: 'warning',
				message: `QueryTable package part "${part.path}" is not claimed by any workbook table model`,
				refs: [part.path],
				suggestedFix:
					'Inspect table relationships and reconnect or intentionally remove the orphan queryTable sidecar before table edits.',
				details: {
					kind: 'orphan-query-table-part',
					partPath: part.path,
					ownerScope: part.ownerScope,
					contentType: part.contentType,
					incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
						queryTableRelationshipDetails,
					),
				},
			})
		}
	}

	return issues
}

interface TableIntegrityEntry {
	readonly tableName: string
	readonly tableId: string
	readonly sheetName: string
	readonly ref: string
	readonly partPath?: string
}

interface TableColumnIntegrityEntry {
	readonly columnName: string
	readonly columnIndex: number
	readonly ref: string
}

interface TableQueryTableIntegrityEntry {
	readonly tableName: string
	readonly sheetName: string
	readonly ref: string
	readonly tablePartPath?: string
	readonly queryTablePartPath: string
	readonly relationshipId: string
}

interface TableColumnQueryFieldEntry {
	readonly columnName: string
	readonly columnIndex: number
	readonly ref: string
}

interface WorkbookConnectionIntegrityEntry {
	readonly partPath: string
	readonly connectionId: number
	readonly name?: string
}

function workbookConnectionPartsById(
	wb: Workbook,
): Map<number, WorkbookConnectionIntegrityEntry[]> {
	const byId = new Map<number, WorkbookConnectionIntegrityEntry[]>()
	for (const part of wb.connectionParts) {
		if (part.kind !== 'connection' || part.connectionId === undefined) continue
		const entry = workbookConnectionIntegrityEntry(part)
		const entries = byId.get(part.connectionId)
		if (entries) entries.push(entry)
		else byId.set(part.connectionId, [entry])
	}
	return byId
}

function workbookConnectionIntegrityEntry(
	part: Workbook['connectionParts'][number],
): WorkbookConnectionIntegrityEntry {
	return {
		partPath: part.partPath,
		connectionId: part.connectionId ?? -1,
		...(part.name ? { name: part.name } : {}),
	}
}

function queryTableConnectionIdIssues(
	part: Workbook['connectionParts'][number],
	workbookConnectionsById: ReadonlyMap<number, readonly WorkbookConnectionIntegrityEntry[]>,
): CheckIssue[] {
	const entry = worksheetQueryTableIntegrityEntry(part)
	const refs = [part.sheetName ?? part.partPath, part.partPath]
	if (part.connectionId === undefined) {
		return [
			{
				rule: 'table-query-integrity',
				severity: 'error',
				message: `QueryTable part "${part.partPath}" is missing connectionId binding`,
				refs,
				suggestedFix:
					'Restore the queryTable connectionId so the queryTable can bind to one workbook connection before refresh or write.',
				details: {
					kind: 'query-table-connection-id-missing',
					connectionPart: entry,
				},
			},
		]
	}
	const workbookConnections = workbookConnectionsById.get(part.connectionId) ?? []
	if (workbookConnections.length === 0) {
		return [
			{
				rule: 'table-query-integrity',
				severity: 'error',
				message: `QueryTable part "${part.partPath}" connectionId "${part.connectionId}" does not match a workbook connection`,
				refs,
				suggestedFix:
					'Restore a matching workbook connection entry or update the queryTable connectionId before editing query-backed data.',
				details: {
					kind: 'query-table-connection-id-missing-workbook-connection',
					connectionPart: entry,
					connectionId: part.connectionId,
				},
			},
		]
	}
	if (workbookConnections.length > 1) {
		return [
			{
				rule: 'table-query-integrity',
				severity: 'error',
				message: `QueryTable part "${part.partPath}" connectionId "${part.connectionId}" matches multiple workbook connections`,
				refs,
				suggestedFix:
					'Repair duplicate workbook connection ids before editing query-backed data or refresh metadata.',
				details: {
					kind: 'query-table-connection-id-ambiguous',
					connectionPart: entry,
					connectionId: part.connectionId,
					workbookConnections,
				},
			},
		]
	}
	return []
}

function worksheetQueryTableIntegrityEntry(
	part: Workbook['connectionParts'][number],
): Readonly<Record<string, unknown>> {
	return {
		partPath: part.partPath,
		...(part.sheetName ? { sheetName: part.sheetName } : {}),
		...(part.name ? { name: part.name } : {}),
		...(part.connectionId !== undefined ? { connectionId: part.connectionId } : {}),
		relationshipCount: part.relationshipCount,
	}
}

function isQueryTableRelationshipType(relationshipType: string): boolean {
	return relationshipType.toLowerCase().endsWith('/relationships/querytable')
}

function queryTableRelationshipDetails(
	relationship: VerifyPackageGraphRelationship,
): Readonly<Record<string, unknown>> {
	return packageRelationshipDetails(relationship)
}

function checkPivotSlicerTimelineIntegrity(
	wb: Workbook,
	packageGraph?: VerifyPackageGraph,
): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNames = new Set(wb.sheets.map((sheet) => sheet.name))
	const pivotCachesById = new Map<number, Workbook['pivotCaches'][number]>()
	const duplicateCacheIds = new Map<number, Workbook['pivotCaches'][number][]>()
	for (const cache of wb.pivotCaches) {
		if (cache.cacheId === undefined) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot cache "${cache.partPath}" is missing a workbook cacheId`,
				refs: [cache.partPath],
				suggestedFix:
					'Restore the workbook pivotCache id before editing pivot tables or cache definitions.',
				details: { kind: 'pivot-cache-id-missing', cache: pivotCacheSummary(cache) },
			})
			continue
		}
		const existing = pivotCachesById.get(cache.cacheId)
		if (existing)
			duplicateCacheIds.set(cache.cacheId, [
				...(duplicateCacheIds.get(cache.cacheId) ?? [existing]),
				cache,
			])
		else pivotCachesById.set(cache.cacheId, cache)
	}
	for (const [cacheId, caches] of duplicateCacheIds) {
		issues.push({
			rule: 'pivot-integrity',
			severity: 'error',
			message: `Pivot cache id "${cacheId}" is used by multiple pivot cache definitions`,
			refs: caches.map((cache) => cache.partPath),
			suggestedFix:
				'Repair workbook pivotCache ids before editing pivot tables or cache definitions.',
			details: {
				kind: 'duplicate-pivot-cache-id',
				cacheId,
				partPaths: caches.map((cache) => cache.partPath),
			},
		})
	}

	for (const pivot of wb.pivotTables) {
		if (!sheetNames.has(pivot.sheetName)) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot table "${pivot.name ?? pivot.partPath}" belongs to missing sheet "${pivot.sheetName}"`,
				refs: [pivot.partPath, pivot.sheetName],
				suggestedFix:
					'Restore the owning worksheet or remove the stale pivot table metadata before writing.',
				details: { kind: 'pivot-table-sheet-missing', pivotTable: pivotSummary(pivot) },
			})
		}
		if (pivot.cacheId === undefined) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot table "${pivot.name ?? pivot.partPath}" is missing cacheId binding`,
				refs: [pivot.partPath],
				suggestedFix:
					'Reconnect the pivot table to a workbook pivot cache definition before editing pivot output.',
				details: { kind: 'pivot-table-cache-id-absent', pivotTable: pivotSummary(pivot) },
			})
		} else if (duplicateCacheIds.has(pivot.cacheId)) {
			const caches = duplicateCacheIds.get(pivot.cacheId) ?? []
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot table "${pivot.name ?? pivot.partPath}" cacheId "${pivot.cacheId}" is ambiguous`,
				refs: [pivot.partPath, ...caches.map((cache) => cache.partPath)],
				suggestedFix:
					'Repair duplicate workbook pivotCache ids before resolving pivot table cache bindings.',
				details: {
					kind: 'pivot-table-cache-id-ambiguous',
					pivotTable: pivotSummary(pivot),
					caches: caches.map(pivotCacheSummary),
				},
			})
		} else if (!pivotCachesById.has(pivot.cacheId)) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot table "${pivot.name ?? pivot.partPath}" references missing cacheId "${pivot.cacheId}"`,
				refs: [pivot.partPath],
				suggestedFix:
					'Reconnect the pivot table cacheId to a workbook pivot cache definition before editing pivot output.',
				details: { kind: 'pivot-table-cache-id-missing', pivotTable: pivotSummary(pivot) },
			})
		}
	}

	for (const cache of wb.pivotCaches) {
		issues.push(...checkPivotCacheSourceIntegrity(wb, cache, sheetNames))
		issues.push(...checkPivotCacheRecordIntegrity(cache))
		issues.push(...checkPivotRefreshIntegrity(wb, cache))
	}
	issues.push(...checkSlicerTimelineModelIntegrity(wb))
	if (packageGraph) {
		issues.push(...checkPivotPackageGraphIntegrity(wb, packageGraph))
		issues.push(...checkSlicerTimelinePackageGraphIntegrity(wb, packageGraph))
	}
	return issues
}

function checkPivotCacheSourceIntegrity(
	wb: Workbook,
	cache: Workbook['pivotCaches'][number],
	sheetNames: ReadonlySet<string>,
): CheckIssue[] {
	const issues: CheckIssue[] = []
	if (cache.sourceType === 'worksheet' && cache.sourceSheet && !sheetNames.has(cache.sourceSheet)) {
		issues.push({
			rule: 'pivot-source-integrity',
			severity: 'error',
			message: `Pivot cache "${cache.partPath}" references missing source sheet "${cache.sourceSheet}"`,
			refs: [cache.partPath, cache.sourceSheet],
			suggestedFix:
				'Restore the source worksheet or update the pivot cache worksheetSource before writing.',
			details: { kind: 'pivot-cache-source-sheet-missing', cache: pivotCacheSummary(cache) },
		})
	}
	if (cache.sourceRef) {
		try {
			parseRange(cache.sourceRef)
		} catch {
			issues.push({
				rule: 'pivot-source-integrity',
				severity: 'error',
				message: `Pivot cache "${cache.partPath}" has invalid worksheetSource ref "${cache.sourceRef}"`,
				refs: [cache.partPath, cache.sourceRef],
				suggestedFix:
					'Repair worksheetSource ref to a valid A1 range before relying on pivot refresh metadata.',
				details: { kind: 'pivot-cache-source-ref-invalid', cache: pivotCacheSummary(cache) },
			})
		}
	}
	if (!cache.sourceName) return issues
	const tableMatches = wb.sheets.flatMap((sheet) =>
		sheet.tables
			.filter((table) => table.name === cache.sourceName)
			.map((table) => ({
				sheetName: sheet.name,
				tableName: table.name,
				ref: rangeToA1(table.ref),
			})),
	)
	const hasDefinedName = wb.definedNames.has(cache.sourceName)
	if (tableMatches.length === 0 && !hasDefinedName) {
		issues.push({
			rule: 'pivot-source-integrity',
			severity: 'error',
			message: `Pivot cache "${cache.partPath}" references missing source name "${cache.sourceName}"`,
			refs: [cache.partPath, cache.sourceName],
			suggestedFix:
				'Restore the source table/defined name or update the pivot cache sourceName before writing.',
			details: { kind: 'pivot-cache-source-table-missing', cache: pivotCacheSummary(cache) },
		})
	}
	if (cache.sourceSheet && tableMatches.some((table) => table.sheetName !== cache.sourceSheet)) {
		issues.push({
			rule: 'pivot-source-integrity',
			severity: 'warning',
			message: `Pivot cache "${cache.partPath}" sourceName "${cache.sourceName}" is not on source sheet "${cache.sourceSheet}"`,
			refs: [cache.partPath, ...tableMatches.map((table) => `${table.sheetName}!${table.ref}`)],
			suggestedFix:
				'Align worksheetSource sheet/name metadata before refreshing or rewriting pivot sources.',
			details: {
				kind: 'pivot-cache-source-table-sheet-mismatch',
				cache: pivotCacheSummary(cache),
				tables: tableMatches,
			},
		})
	}
	return issues
}

function checkPivotCacheRecordIntegrity(cache: Workbook['pivotCaches'][number]): CheckIssue[] {
	const issues: CheckIssue[] = []
	const records = cache.records
	if (!records) {
		if (cache.recordsPartPath) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'warning',
				message: `Pivot cache "${cache.partPath}" references records part "${cache.recordsPartPath}" but no records were parsed`,
				refs: [cache.partPath, cache.recordsPartPath],
				suggestedFix:
					'Load or restore the pivotCacheRecords payload before relying on saved cache rows.',
				details: {
					kind: 'pivot-cache-records-unparsed',
					cache: pivotCacheSummary(cache),
				},
			})
		}
		return issues
	}
	if (!cache.recordsPartPath) {
		issues.push({
			rule: 'pivot-integrity',
			severity: 'error',
			message: `Pivot cache "${cache.partPath}" has parsed records but no cache definition records relationship`,
			refs: [cache.partPath, records.partPath],
			suggestedFix:
				'Restore the pivotCacheDefinition relationship to its records part before cache edits.',
			details: {
				kind: 'pivot-cache-records-relationship-missing',
				cache: pivotCacheSummary(cache),
				recordsPartPath: records.partPath,
			},
		})
	} else if (records.partPath !== cache.recordsPartPath) {
		issues.push({
			rule: 'pivot-integrity',
			severity: 'error',
			message: `Pivot cache "${cache.partPath}" records payload path does not match definition relationship`,
			refs: [cache.partPath, cache.recordsPartPath, records.partPath],
			suggestedFix:
				'Make the parsed records payload and pivotCacheDefinition records relationship point to the same part.',
			details: {
				kind: 'pivot-cache-records-part-path-mismatch',
				cache: pivotCacheSummary(cache),
				expectedRecordsPartPath: cache.recordsPartPath,
				actualRecordsPartPath: records.partPath,
			},
		})
	}
	const counts = [
		['cacheRecordCount', cache.recordCount],
		['recordsDeclaredCount', records.declaredCount],
		['recordsParsedCount', records.parsedCount],
	] as const
	const definedCounts = counts.flatMap(([key, value]) =>
		value === undefined ? [] : ([[key, value]] as const),
	)
	const distinctCounts = new Set(definedCounts.map(([, count]) => count))
	if (distinctCounts.size > 1) {
		issues.push({
			rule: 'pivot-integrity',
			severity: 'error',
			message: `Pivot cache "${cache.partPath}" record counts disagree between cache definition and records part`,
			refs: [cache.partPath, records.partPath],
			suggestedFix:
				'Regenerate or refresh pivot cache records so recordCount and pivotCacheRecords count agree.',
			details: {
				kind: 'pivot-cache-record-count-mismatch',
				cache: pivotCacheSummary(cache),
				counts: Object.fromEntries(definedCounts),
			},
		})
	}
	const expectedWidth = cache.fields.length
	if (expectedWidth > 0) {
		const mismatched = [...records.preview, ...(records.materializedRecords ?? [])]
			.filter((record) => record.values.length !== expectedWidth)
			.slice(0, 5)
		if (mismatched.length > 0) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot cache "${cache.partPath}" has records whose value count does not match ${expectedWidth} cache field(s)`,
				refs: [cache.partPath, records.partPath],
				suggestedFix:
					'Repair pivot cache records before using saved cache rows for output auditing.',
				details: {
					kind: 'pivot-cache-record-width-mismatch',
					cache: pivotCacheSummary(cache),
					expectedWidth,
					mismatchedRecords: mismatched.map((record) => ({
						index: record.index,
						actualWidth: record.values.length,
					})),
				},
			})
		}
	}
	return issues
}

function checkPivotRefreshIntegrity(
	wb: Workbook,
	cache: Workbook['pivotCaches'][number],
): CheckIssue[] {
	const issues: CheckIssue[] = []
	const needsRefresh =
		cache.invalid === true ||
		cache.refreshOnLoad === true ||
		cache.upgradeOnRefresh === true ||
		cache.saveData === false
	if (!needsRefresh) {
		if (cache.records?.materializedComplete === false) {
			issues.push({
				rule: 'pivot-refresh-integrity',
				severity: 'info',
				message: `Pivot cache "${cache.partPath}" records were only partially materialized for verification`,
				refs: [cache.partPath, cache.records.partPath],
				suggestedFix:
					'Reopen with pivotCacheRecordMaterializeLimit: "all" before claiming complete cache-row audit coverage.',
				details: {
					kind: 'pivot-records-not-materialized',
					cache: pivotCacheSummary(cache),
					materializedCount: cache.records.materializedCount,
					parsedCount: cache.records.parsedCount,
				},
			})
		}
		return issues
	}
	issues.push({
		rule: 'pivot-refresh-integrity',
		severity: 'warning',
		message: `Pivot cache "${cache.partPath}" is marked for pivot-aware refresh`,
		refs: [cache.partPath],
		suggestedFix:
			'Preserve the cache and require Excel or a pivot-aware refresh engine before trusting saved pivot output.',
		details: {
			kind: 'pivot-cache-refresh-required',
			cache: pivotCacheSummary(cache),
			signals: {
				invalid: cache.invalid,
				refreshOnLoad: cache.refreshOnLoad,
				upgradeOnRefresh: cache.upgradeOnRefresh,
				saveData: cache.saveData,
			},
		},
	})
	issues.push({
		rule: 'pivot-refresh-integrity',
		severity: 'warning',
		message: `Pivot cache "${cache.partPath}" requires a pivot-aware refresh; headless refresh is not supported`,
		refs: [cache.partPath],
		suggestedFix:
			'Open the workbook in Excel or another pivot-aware engine before treating saved pivot cache or output cells as current.',
		details: {
			kind: 'pivot-headless-refresh-unsupported',
			cache: pivotCacheSummary(cache),
			signals: {
				invalid: cache.invalid,
				refreshOnLoad: cache.refreshOnLoad,
				upgradeOnRefresh: cache.upgradeOnRefresh,
				saveData: cache.saveData,
			},
		},
	})
	if (cache.enableRefresh === false) {
		issues.push({
			rule: 'pivot-refresh-integrity',
			severity: 'warning',
			message: `Pivot cache "${cache.partPath}" needs refresh but enableRefresh is false`,
			refs: [cache.partPath],
			suggestedFix:
				'Inspect workbook policy before attempting headless pivot refresh or cache rewrites.',
			details: { kind: 'pivot-cache-refresh-disabled', cache: pivotCacheSummary(cache) },
		})
	}
	for (const pivot of wb.pivotTables.filter((pivot) => pivot.cacheId === cache.cacheId)) {
		if (cache.invalid === true) {
			issues.push({
				rule: 'pivot-refresh-integrity',
				severity: 'warning',
				message: `Pivot table "${pivot.name ?? pivot.partPath}" saved output is stale after pivot source or filter edits`,
				refs: [pivot.partPath, cache.partPath],
				suggestedFix:
					'Refresh the pivot cache in Excel or another pivot-aware engine before trusting saved pivot output cells.',
				details: {
					kind: 'pivot-saved-output-stale-after-source-edit',
					pivotTable: pivotSummary(pivot),
					cache: pivotCacheSummary(cache),
				},
			})
		}
		issues.push({
			rule: 'pivot-refresh-integrity',
			severity: 'warning',
			message: `Pivot table "${pivot.name ?? pivot.partPath}" may show stale saved output from cacheId "${cache.cacheId}"`,
			refs: [pivot.partPath, cache.partPath],
			suggestedFix:
				'Refresh the pivot cache before treating worksheet pivot output cells as authoritative.',
			details: {
				kind: 'pivot-output-may-be-stale',
				pivotTable: pivotSummary(pivot),
				cache: pivotCacheSummary(cache),
			},
		})
	}
	return issues
}

function checkPivotPackageGraphIntegrity(
	wb: Workbook,
	packageGraph: VerifyPackageGraph,
): CheckIssue[] {
	const issues: CheckIssue[] = []
	const graphPartByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
	const graphRelationshipsByTarget = relationshipsByTarget(packageGraph.relationships)
	const claimedPivotTables = new Set(wb.pivotTables.map((pivot) => pivot.partPath))
	const claimedCaches = new Set(wb.pivotCaches.map((cache) => cache.partPath))
	const claimedRecords = new Set(
		wb.pivotCaches.flatMap((cache) => (cache.recordsPartPath ? [cache.recordsPartPath] : [])),
	)

	for (const cache of wb.pivotCaches) {
		const workbookRelationships = packageGraph.relationships.filter(
			(rel) =>
				rel.sourcePartPath === 'xl/workbook.xml' &&
				rel.resolvedTarget === cache.partPath &&
				isPivotCacheDefinitionRelationshipType(rel.type),
		)
		if (!graphPartByPath.has(cache.partPath)) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot cache metadata references missing package part "${cache.partPath}"`,
				refs: [cache.partPath],
				suggestedFix: 'Restore the pivot cache definition part before writing.',
				details: { kind: 'missing-pivot-cache-part', cache: pivotCacheSummary(cache) },
			})
		}
		if (workbookRelationships.length > 1) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot cache "${cache.partPath}" is targeted by multiple workbook pivotCache relationships`,
				refs: [
					cache.partPath,
					...workbookRelationships.map((rel) => `${rel.relationshipPartPath}#${rel.id}`),
				],
				suggestedFix:
					'Repair workbook pivotCache relationship ids so each cache definition has a single workbook binding.',
				details: {
					kind: 'duplicate-pivot-cache-workbook-relationships',
					cache: pivotCacheSummary(cache),
					workbookRelationships: workbookRelationships.map(packageRelationshipDetails),
				},
			})
		}
		if (
			graphPartByPath.has(cache.partPath) &&
			cache.relId === undefined &&
			workbookRelationships.length === 0
		) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot cache "${cache.partPath}" is missing its workbook pivotCache relationship`,
				refs: [cache.partPath],
				suggestedFix:
					'Restore the workbook pivotCache relationship before editing or writing pivot caches.',
				details: {
					kind: 'pivot-cache-workbook-relationship-missing',
					cache: pivotCacheSummary(cache),
				},
			})
		}
		if (cache.recordsPartPath && !graphPartByPath.has(cache.recordsPartPath)) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot cache "${cache.partPath}" references missing records part "${cache.recordsPartPath}"`,
				refs: [cache.partPath, cache.recordsPartPath],
				suggestedFix:
					'Restore the pivotCacheRecords part or remove the stale cache records relationship.',
				details: { kind: 'missing-pivot-cache-records-part', cache: pivotCacheSummary(cache) },
			})
		}
		if (cache.relId) {
			const workbookRel = packageGraph.relationships.find(
				(rel) =>
					rel.sourcePartPath === 'xl/workbook.xml' &&
					rel.id === cache.relId &&
					rel.resolvedTarget === cache.partPath &&
					isPivotCacheDefinitionRelationshipType(rel.type),
			)
			if (!workbookRel) {
				issues.push({
					rule: 'pivot-integrity',
					severity: 'error',
					message: `Workbook pivotCache relationship "${cache.relId}" does not bind to "${cache.partPath}" in the package graph`,
					refs: [`xl/_rels/workbook.xml.rels#${cache.relId}`, cache.partPath],
					suggestedFix:
						'Repair workbook pivotCache relationship id/target binding before editing pivot caches.',
					details: {
						kind: 'pivot-cache-workbook-relationship-binding-mismatch',
						cache: pivotCacheSummary(cache),
						incomingRelationships: (graphRelationshipsByTarget.get(cache.partPath) ?? []).map(
							packageRelationshipDetails,
						),
					},
				})
			}
		}
		if (cache.recordsPartPath) {
			const recordsRel = packageGraph.relationships.find(
				(rel) =>
					rel.sourcePartPath === cache.partPath &&
					rel.resolvedTarget === cache.recordsPartPath &&
					isPivotCacheRecordsRelationshipType(rel.type),
			)
			if (!recordsRel) {
				issues.push({
					rule: 'pivot-integrity',
					severity: 'error',
					message: `Pivot cache records relationship from "${cache.partPath}" does not bind to "${cache.recordsPartPath}"`,
					refs: [cache.partPath, cache.recordsPartPath],
					suggestedFix:
						'Repair the pivotCacheDefinition relationship to its records part before cache edits.',
					details: {
						kind: 'pivot-cache-records-relationship-binding-mismatch',
						cache: pivotCacheSummary(cache),
						incomingRelationships: (
							graphRelationshipsByTarget.get(cache.recordsPartPath) ?? []
						).map(packageRelationshipDetails),
					},
				})
			}
		}
	}

	for (const pivot of wb.pivotTables) {
		if (!graphPartByPath.has(pivot.partPath)) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot table metadata references missing package part "${pivot.partPath}"`,
				refs: [pivot.partPath],
				suggestedFix: 'Restore the pivot table part before writing.',
				details: { kind: 'missing-pivot-table-part', pivotTable: pivotSummary(pivot) },
			})
		}
		const worksheetOwner = (graphRelationshipsByTarget.get(pivot.partPath) ?? []).find(
			(rel) => isPivotTableRelationshipType(rel.type) && isWorksheetPartPath(rel.sourcePartPath),
		)
		if (!worksheetOwner) {
			issues.push({
				rule: 'pivot-integrity',
				severity: 'error',
				message: `Pivot table part "${pivot.partPath}" is not bound from an owning worksheet relationship`,
				refs: [pivot.partPath],
				suggestedFix:
					'Reconnect the pivot table part from its worksheet relationship before writing.',
				details: {
					kind: 'pivot-table-worksheet-relationship-binding-mismatch',
					pivotTable: pivotSummary(pivot),
					incomingRelationships: (graphRelationshipsByTarget.get(pivot.partPath) ?? []).map(
						packageRelationshipDetails,
					),
				},
			})
		}
	}

	for (const part of packageGraph.parts) {
		if (isPivotCacheDefinitionPart(part) && !claimedCaches.has(part.path)) {
			issues.push(orphanPivotPartIssue('orphan-pivot-cache-part', part, graphRelationshipsByTarget))
		}
		if (isPivotCacheRecordsPart(part) && !claimedRecords.has(part.path)) {
			issues.push(
				orphanPivotPartIssue('orphan-pivot-cache-records-part', part, graphRelationshipsByTarget),
			)
		}
		if (isPivotTablePart(part) && !claimedPivotTables.has(part.path)) {
			issues.push(orphanPivotPartIssue('orphan-pivot-table-part', part, graphRelationshipsByTarget))
		}
	}
	return issues
}

function checkSlicerTimelineModelIntegrity(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	const pivotTablesByName = new Map(
		wb.pivotTables.filter((pivot) => pivot.name).map((pivot) => [pivot.name as string, pivot]),
	)
	const pivotCacheIdsByTable = new Map(
		wb.pivotTables.map((pivot) => [pivot.partPath, linkedPivotCacheIds(wb, pivot)]),
	)
	const slicerCachesByName = new Map(
		wb.slicerCaches.filter((cache) => cache.name).map((cache) => [cache.name as string, cache]),
	)
	for (const slicer of wb.slicers) {
		if (slicer.cacheName && !slicerCachesByName.has(slicer.cacheName)) {
			issues.push({
				rule: 'slicer-integrity',
				severity: 'error',
				message: `Slicer "${slicer.name ?? slicer.partPath}" references missing slicer cache "${slicer.cacheName}"`,
				refs: [slicer.partPath, slicer.cacheName],
				suggestedFix: 'Reconnect the slicer to an existing slicer cache before writing.',
				details: { kind: 'slicer-cache-name-missing', slicer },
			})
		}
	}
	for (const cache of wb.slicerCaches) {
		for (const pivotTableName of cache.pivotTableNames) {
			const pivot = pivotTablesByName.get(pivotTableName)
			if (!pivot) {
				issues.push(
					linkedPivotMissingIssue(
						'slicer-integrity',
						'slicer-cache-pivot-table-missing',
						cache.partPath,
						pivotTableName,
					),
				)
			} else if (
				cache.pivotCacheId !== undefined &&
				pivot.cacheId !== undefined &&
				cache.pivotCacheId !== pivot.cacheId &&
				!(pivotCacheIdsByTable.get(pivot.partPath) ?? new Set()).has(cache.pivotCacheId)
			) {
				issues.push(
					cachePivotIdMismatchIssue(
						'slicer-integrity',
						'slicer-cache-pivot-cache-id-mismatch',
						cache.partPath,
						cache.pivotCacheId,
						pivot,
					),
				)
			}
		}
	}

	const timelineCachesByName = new Map(
		wb.timelineCaches.filter((cache) => cache.name).map((cache) => [cache.name as string, cache]),
	)
	for (const timeline of wb.timelines) {
		if (timeline.cacheName && !timelineCachesByName.has(timeline.cacheName)) {
			issues.push({
				rule: 'timeline-integrity',
				severity: 'error',
				message: `Timeline "${timeline.name ?? timeline.partPath}" references missing timeline cache "${timeline.cacheName}"`,
				refs: [timeline.partPath, timeline.cacheName],
				suggestedFix: 'Reconnect the timeline to an existing timeline cache before writing.',
				details: { kind: 'timeline-cache-name-missing', timeline },
			})
		}
	}
	for (const cache of wb.timelineCaches) {
		if (
			cache.state?.pivotCacheId !== undefined &&
			cache.pivotCacheId !== undefined &&
			cache.state.pivotCacheId !== cache.pivotCacheId
		) {
			issues.push({
				rule: 'timeline-integrity',
				severity: 'error',
				message: `Timeline cache "${cache.partPath}" state pivotCacheId does not match cache pivotCacheId`,
				refs: [cache.partPath],
				suggestedFix: 'Repair timeline state pivot cache ids before editing timeline selections.',
				details: {
					kind: 'timeline-state-pivot-cache-id-mismatch',
					partPath: cache.partPath,
					pivotCacheId: cache.pivotCacheId,
					statePivotCacheId: cache.state.pivotCacheId,
				},
			})
		}
		if (cache.state?.filterPivotName && !pivotTablesByName.has(cache.state.filterPivotName)) {
			issues.push(
				linkedPivotMissingIssue(
					'timeline-integrity',
					'timeline-state-filter-pivot-missing',
					cache.partPath,
					cache.state.filterPivotName,
				),
			)
		}
		for (const pivotTableName of cache.pivotTableNames) {
			const pivot = pivotTablesByName.get(pivotTableName)
			if (!pivot) {
				issues.push(
					linkedPivotMissingIssue(
						'timeline-integrity',
						'timeline-cache-pivot-table-missing',
						cache.partPath,
						pivotTableName,
					),
				)
			} else if (
				cache.pivotCacheId !== undefined &&
				pivot.cacheId !== undefined &&
				cache.pivotCacheId !== pivot.cacheId &&
				!(pivotCacheIdsByTable.get(pivot.partPath) ?? new Set()).has(cache.pivotCacheId)
			) {
				issues.push(
					cachePivotIdMismatchIssue(
						'timeline-integrity',
						'timeline-cache-pivot-cache-id-mismatch',
						cache.partPath,
						cache.pivotCacheId,
						pivot,
					),
				)
			}
		}
	}
	return issues
}

function linkedPivotCacheIds(
	wb: Workbook,
	pivot: Workbook['pivotTables'][number],
): ReadonlySet<number> {
	const ids = new Set<number>()
	if (pivot.cacheId !== undefined) ids.add(pivot.cacheId)
	const cache = wb.pivotCaches.find((entry) => entry.cacheId === pivot.cacheId)
	if (cache?.extensionCacheId !== undefined) ids.add(cache.extensionCacheId)
	return ids
}

function checkSlicerTimelinePackageGraphIntegrity(
	wb: Workbook,
	packageGraph: VerifyPackageGraph,
): CheckIssue[] {
	const issues: CheckIssue[] = []
	const graphPartByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
	const graphRelationshipsByTarget = relationshipsByTarget(packageGraph.relationships)
	const claimedSlicerCaches = new Set(wb.slicerCaches.map((cache) => cache.partPath))
	const claimedSlicers = new Set(wb.slicers.map((slicer) => slicer.partPath))
	const claimedTimelineCaches = new Set(wb.timelineCaches.map((cache) => cache.partPath))
	const claimedTimelines = new Set(wb.timelines.map((timeline) => timeline.partPath))

	for (const cache of wb.slicerCaches) {
		if (!graphPartByPath.has(cache.partPath)) {
			issues.push(
				missingAnalyticalPartIssue(
					'slicer-integrity',
					'slicer-cache-missing-package-part',
					cache.partPath,
				),
			)
		}
		if (
			!packageGraph.relationships.some(
				(rel) =>
					rel.sourcePartPath === 'xl/workbook.xml' &&
					rel.resolvedTarget === cache.partPath &&
					isSlicerCacheRelationshipType(rel.type),
			)
		) {
			issues.push(
				workbookAnalyticalBindingIssue(
					'slicer-integrity',
					'slicer-cache-workbook-relationship-binding-mismatch',
					cache.partPath,
					graphRelationshipsByTarget,
				),
			)
		}
		for (const slicer of wb.slicers.filter((entry) => entry.cacheName === cache.name)) {
			if (
				!packageGraph.relationships.some(
					(rel) =>
						rel.sourcePartPath === cache.partPath &&
						rel.resolvedTarget === slicer.partPath &&
						isSlicerRelationshipType(rel.type),
				)
			) {
				issues.push(
					analyticalUiBindingIssue(
						'slicer-integrity',
						'slicer-cache-ui-relationship-binding-mismatch',
						cache.partPath,
						slicer.partPath,
						graphRelationshipsByTarget,
					),
				)
			}
			if (
				!packageGraph.relationships.some(
					(rel) =>
						rel.sourcePartPath.startsWith('xl/worksheets/') &&
						rel.resolvedTarget === slicer.partPath &&
						isSlicerRelationshipType(rel.type),
				)
			) {
				issues.push(
					worksheetAnalyticalUiBindingIssue(
						'slicer-integrity',
						'slicer-worksheet-relationship-binding-mismatch',
						slicer.partPath,
						graphRelationshipsByTarget,
					),
				)
			}
		}
	}
	for (const cache of wb.timelineCaches) {
		if (!graphPartByPath.has(cache.partPath)) {
			issues.push(
				missingAnalyticalPartIssue(
					'timeline-integrity',
					'timeline-cache-missing-package-part',
					cache.partPath,
				),
			)
		}
		if (
			!packageGraph.relationships.some(
				(rel) =>
					rel.sourcePartPath === 'xl/workbook.xml' &&
					rel.resolvedTarget === cache.partPath &&
					isTimelineCacheRelationshipType(rel.type),
			)
		) {
			issues.push(
				workbookAnalyticalBindingIssue(
					'timeline-integrity',
					'timeline-cache-workbook-relationship-binding-mismatch',
					cache.partPath,
					graphRelationshipsByTarget,
				),
			)
		}
		for (const timeline of wb.timelines.filter((entry) => entry.cacheName === cache.name)) {
			if (
				!packageGraph.relationships.some(
					(rel) =>
						rel.sourcePartPath === cache.partPath &&
						rel.resolvedTarget === timeline.partPath &&
						isTimelineRelationshipType(rel.type),
				)
			) {
				issues.push(
					analyticalUiBindingIssue(
						'timeline-integrity',
						'timeline-cache-ui-relationship-binding-mismatch',
						cache.partPath,
						timeline.partPath,
						graphRelationshipsByTarget,
					),
				)
			}
			if (
				!packageGraph.relationships.some(
					(rel) =>
						rel.sourcePartPath.startsWith('xl/worksheets/') &&
						rel.resolvedTarget === timeline.partPath &&
						isTimelineRelationshipType(rel.type),
				)
			) {
				issues.push(
					worksheetAnalyticalUiBindingIssue(
						'timeline-integrity',
						'timeline-worksheet-relationship-binding-mismatch',
						timeline.partPath,
						graphRelationshipsByTarget,
					),
				)
			}
		}
	}
	for (const part of packageGraph.parts) {
		if (isSlicerCachePart(part) && !claimedSlicerCaches.has(part.path)) {
			issues.push(
				orphanAnalyticalPartIssue(
					'slicer-integrity',
					'orphan-slicer-cache-part',
					part,
					graphRelationshipsByTarget,
				),
			)
		}
		if (isSlicerPart(part) && !claimedSlicers.has(part.path)) {
			issues.push(
				orphanAnalyticalPartIssue(
					'slicer-integrity',
					'orphan-slicer-part',
					part,
					graphRelationshipsByTarget,
				),
			)
		}
		if (isTimelineCachePart(part) && !claimedTimelineCaches.has(part.path)) {
			issues.push(
				orphanAnalyticalPartIssue(
					'timeline-integrity',
					'orphan-timeline-cache-part',
					part,
					graphRelationshipsByTarget,
				),
			)
		}
		if (isTimelinePart(part) && !claimedTimelines.has(part.path)) {
			issues.push(
				orphanAnalyticalPartIssue(
					'timeline-integrity',
					'orphan-timeline-part',
					part,
					graphRelationshipsByTarget,
				),
			)
		}
	}
	return issues
}

function pivotSummary(pivot: Workbook['pivotTables'][number]): Readonly<Record<string, unknown>> {
	return {
		partPath: pivot.partPath,
		sheetName: pivot.sheetName,
		...(pivot.name ? { name: pivot.name } : {}),
		...(pivot.cacheId !== undefined ? { cacheId: pivot.cacheId } : {}),
		...(pivot.locationRef ? { locationRef: pivot.locationRef } : {}),
	}
}

function pivotCacheSummary(
	cache: Workbook['pivotCaches'][number],
): Readonly<Record<string, unknown>> {
	return {
		partPath: cache.partPath,
		...(cache.cacheId !== undefined ? { cacheId: cache.cacheId } : {}),
		...(cache.relId ? { relId: cache.relId } : {}),
		...(cache.recordsPartPath ? { recordsPartPath: cache.recordsPartPath } : {}),
		...(cache.sourceType ? { sourceType: cache.sourceType } : {}),
		...(cache.sourceSheet ? { sourceSheet: cache.sourceSheet } : {}),
		...(cache.sourceRef ? { sourceRef: cache.sourceRef } : {}),
		...(cache.sourceName ? { sourceName: cache.sourceName } : {}),
	}
}

function linkedPivotMissingIssue(
	rule: 'slicer-integrity' | 'timeline-integrity',
	kind: string,
	partPath: string,
	pivotTableName: string,
): CheckIssue {
	return {
		rule,
		severity: 'error',
		message: `Analytical cache "${partPath}" references missing pivot table "${pivotTableName}"`,
		refs: [partPath, pivotTableName],
		suggestedFix:
			'Reconnect the cache to an existing pivot table or remove the stale pivot-table binding before writing.',
		details: { kind, partPath, pivotTableName },
	}
}

function cachePivotIdMismatchIssue(
	rule: 'slicer-integrity' | 'timeline-integrity',
	kind: string,
	partPath: string,
	pivotCacheId: number,
	pivot: Workbook['pivotTables'][number],
): CheckIssue {
	return {
		rule,
		severity: 'error',
		message: `Analytical cache "${partPath}" pivotCacheId "${pivotCacheId}" does not match linked pivot table "${pivot.name ?? pivot.partPath}"`,
		refs: [partPath, pivot.partPath],
		suggestedFix:
			'Align slicer/timeline cache pivotCacheId with the linked pivot table cache before editing filters.',
		details: { kind, partPath, pivotCacheId, pivotTable: pivotSummary(pivot) },
	}
}

function missingAnalyticalPartIssue(
	rule: 'slicer-integrity' | 'timeline-integrity',
	kind: string,
	partPath: string,
): CheckIssue {
	return {
		rule,
		severity: 'error',
		message: `Analytical cache metadata references missing package part "${partPath}"`,
		refs: [partPath],
		suggestedFix: 'Restore the analytical cache package part before writing.',
		details: { kind, partPath },
	}
}

function workbookAnalyticalBindingIssue(
	rule: 'slicer-integrity' | 'timeline-integrity',
	kind: string,
	partPath: string,
	graphRelationshipsByTarget: ReadonlyMap<string, readonly VerifyPackageGraphRelationship[]>,
): CheckIssue {
	return {
		rule,
		severity: 'error',
		message: `Workbook analytical cache relationship does not bind to "${partPath}" in the package graph`,
		refs: [partPath],
		suggestedFix:
			'Repair the workbook relationship id/target binding before editing slicer or timeline caches.',
		details: {
			kind,
			partPath,
			incomingRelationships: (graphRelationshipsByTarget.get(partPath) ?? []).map(
				packageRelationshipDetails,
			),
		},
	}
}

function analyticalUiBindingIssue(
	rule: 'slicer-integrity' | 'timeline-integrity',
	kind: string,
	cachePartPath: string,
	uiPartPath: string,
	graphRelationshipsByTarget: ReadonlyMap<string, readonly VerifyPackageGraphRelationship[]>,
): CheckIssue {
	return {
		rule,
		severity: 'error',
		message: `Analytical cache "${cachePartPath}" does not bind to UI part "${uiPartPath}" in the package graph`,
		refs: [cachePartPath, uiPartPath],
		suggestedFix:
			'Repair the slicer/timeline cache-to-UI package relationship before editing analytical filters.',
		details: {
			kind,
			cachePartPath,
			uiPartPath,
			incomingRelationships: (graphRelationshipsByTarget.get(uiPartPath) ?? []).map(
				packageRelationshipDetails,
			),
		},
	}
}

function worksheetAnalyticalUiBindingIssue(
	rule: 'slicer-integrity' | 'timeline-integrity',
	kind: string,
	uiPartPath: string,
	graphRelationshipsByTarget: ReadonlyMap<string, readonly VerifyPackageGraphRelationship[]>,
): CheckIssue {
	return {
		rule,
		severity: 'error',
		message: `Analytical UI part "${uiPartPath}" is not bound from an owning worksheet relationship`,
		refs: [uiPartPath],
		suggestedFix:
			'Repair the worksheet slicer/timeline package relationship before editing analytical filter UI state.',
		details: {
			kind,
			uiPartPath,
			incomingRelationships: (graphRelationshipsByTarget.get(uiPartPath) ?? []).map(
				packageRelationshipDetails,
			),
		},
	}
}

function orphanPivotPartIssue(
	kind: string,
	part: VerifyPackageGraphPart,
	graphRelationshipsByTarget: ReadonlyMap<string, readonly VerifyPackageGraphRelationship[]>,
): CheckIssue {
	return {
		rule: 'pivot-integrity',
		severity: 'warning',
		message: `Pivot package part "${part.path}" is not claimed by workbook pivot metadata`,
		refs: [part.path],
		suggestedFix:
			'Reconnect the pivot sidecar to workbook pivot metadata or remove the orphan part before writing.',
		details: {
			kind,
			partPath: part.path,
			ownerScope: part.ownerScope,
			contentType: part.contentType,
			incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
				packageRelationshipDetails,
			),
		},
	}
}

function orphanAnalyticalPartIssue(
	rule: 'slicer-integrity' | 'timeline-integrity',
	kind: string,
	part: VerifyPackageGraphPart,
	graphRelationshipsByTarget: ReadonlyMap<string, readonly VerifyPackageGraphRelationship[]>,
): CheckIssue {
	return {
		rule,
		severity: 'warning',
		message: `Analytical package part "${part.path}" is not claimed by workbook ${rule.replace('-integrity', '')} metadata`,
		refs: [part.path],
		suggestedFix:
			'Reconnect the analytical sidecar to workbook metadata or remove the orphan part before writing.',
		details: {
			kind,
			partPath: part.path,
			ownerScope: part.ownerScope,
			contentType: part.contentType,
			incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
				packageRelationshipDetails,
			),
		},
	}
}

function isPivotCacheDefinitionRelationshipType(relationshipType: string | undefined): boolean {
	return /\/relationships\/pivotcachedefinition$/i.test(relationshipType ?? '')
}

function isPivotCacheRecordsRelationshipType(relationshipType: string | undefined): boolean {
	return /\/relationships\/pivotcacherecords$/i.test(relationshipType ?? '')
}

function isPivotTableRelationshipType(relationshipType: string | undefined): boolean {
	return /\/relationships\/pivottable$/i.test(relationshipType ?? '')
}

function isTableRelationshipType(relationshipType: string | undefined): boolean {
	return /\/relationships\/table$/i.test(relationshipType ?? '')
}

function isSlicerCacheRelationshipType(relationshipType: string | undefined): boolean {
	return /\/relationships\/slicercache$/i.test(relationshipType ?? '')
}

function isSlicerRelationshipType(relationshipType: string | undefined): boolean {
	return /\/relationships\/slicer$/i.test(relationshipType ?? '')
}

function isTimelineCacheRelationshipType(relationshipType: string | undefined): boolean {
	return /\/relationships\/timelinecache$/i.test(relationshipType ?? '')
}

function isTimelineRelationshipType(relationshipType: string | undefined): boolean {
	return /\/relationships\/timeline$/i.test(relationshipType ?? '')
}

function isPivotCacheDefinitionPart(part: VerifyPackageGraphPart): boolean {
	return (
		part.featureFamily === 'preservedPivot' &&
		/(^|\/)pivotCache\/pivotCacheDefinition\d+\.xml$/i.test(part.path)
	)
}

function isPivotCacheRecordsPart(part: VerifyPackageGraphPart): boolean {
	return (
		part.featureFamily === 'preservedPivot' &&
		/(^|\/)pivotCache\/pivotCacheRecords\d+\.xml$/i.test(part.path)
	)
}

function isPivotTablePart(part: VerifyPackageGraphPart): boolean {
	return (
		part.featureFamily === 'preservedPivot' &&
		/(^|\/)pivotTables\/pivotTable\d+\.xml$/i.test(part.path)
	)
}

function isTablePart(part: VerifyPackageGraphPart): boolean {
	return (
		part.featureFamily === 'preservedTable' &&
		/(^|\/)tables\/(?!_rels\/)[^/]+\.xml$/i.test(part.path)
	)
}

function isSlicerCachePart(part: VerifyPackageGraphPart): boolean {
	return (
		part.featureFamily === 'preservedSlicer' &&
		/(^|\/)slicerCaches\/(?!_rels\/)[^/]+\.xml$/i.test(part.path)
	)
}

function isSlicerPart(part: VerifyPackageGraphPart): boolean {
	return (
		part.featureFamily === 'preservedSlicer' &&
		/(^|\/)slicers\/(?!_rels\/)[^/]+\.xml$/i.test(part.path)
	)
}

function isTimelineCachePart(part: VerifyPackageGraphPart): boolean {
	return (
		part.featureFamily === 'preservedTimeline' &&
		/(^|\/)timelineCaches\/(?!_rels\/)[^/]+\.xml$/i.test(part.path)
	)
}

function isTimelinePart(part: VerifyPackageGraphPart): boolean {
	return (
		part.featureFamily === 'preservedTimeline' &&
		/(^|\/)timelines\/(?!_rels\/)[^/]+\.xml$/i.test(part.path)
	)
}

function packageRelationshipDetails(
	relationship: VerifyPackageGraphRelationship,
): Readonly<Record<string, unknown>> {
	return {
		sourcePartPath: relationship.sourcePartPath,
		relationshipPartPath: relationship.relationshipPartPath,
		id: relationship.id,
		type: relationship.type,
		rawTarget: relationship.rawTarget,
		...(relationship.rawType ? { rawType: relationship.rawType } : {}),
		...(relationship.resolvedTarget ? { resolvedTarget: relationship.resolvedTarget } : {}),
	}
}

interface ConditionalFormatPriorityEntry {
	readonly source: 'conditionalFormat' | 'x14ConditionalFormat'
	readonly sheetName: string
	readonly priority?: number
	readonly sqref: string
	readonly formatIndex: number
	readonly ruleIndex?: number
	readonly ruleType?: string
	readonly ranges: readonly RangeRef[]
}

function rangeToA1(range: RangeRef): string {
	const start = `${indexToColumn(range.start.col)}${range.start.row + 1}`
	const end = `${indexToColumn(range.end.col)}${range.end.row + 1}`
	return start === end ? start : `${start}:${end}`
}

function parseSqrefRanges(sqref: string): RangeRef[] {
	const ranges: RangeRef[] = []
	for (const token of sqref.trim().split(/\s+/)) {
		if (!token) continue
		try {
			ranges.push(parseRange(token))
		} catch {
			// Malformed sqref values are preserved by the reader; other checks can flag them later.
		}
	}
	return ranges
}

interface X14FormulaReferenceEntry {
	readonly field: string
	readonly formula: string
}

interface MissingSheetReference {
	readonly field: string
	readonly reference: string
	readonly sheetName: string
	readonly token?: string
}

interface MissingTableReference {
	readonly field: string
	readonly reference: string
	readonly tableName: string
	readonly column?: string
	readonly endColumn?: string
}

interface ExternalMetadataReference {
	readonly field: string
	readonly reference: string
	readonly sheetName: string
	readonly externalTarget: string
	readonly token?: string
}

interface InvalidSqrefReference {
	readonly field: 'sqref'
	readonly reference: string
	readonly token: string
	readonly reason?: string
}

interface DeletedFormulaReference {
	readonly field: string
	readonly reference: string
	readonly error: '#REF!'
}

type MetadataReferenceSource =
	| 'conditionalFormat'
	| 'x14ConditionalFormat'
	| 'dataValidation'
	| 'x14DataValidation'

function sqrefTokens(sqref: string): string[] {
	return sqref.trim().split(/\s+/).filter(Boolean)
}

function x14EntryRefs(sheetName: string, sqref: string, fallback: string): string[] {
	const refs = sqrefTokens(sqref).map((token) =>
		chartReferenceSheetName(token).sheetName ? token : `${sheetName}!${token}`,
	)
	return refs.length > 0 ? refs : [`${sheetName}#${fallback}`]
}

function invalidSqrefReferences(sqref: string): InvalidSqrefReference[] {
	const invalid: InvalidSqrefReference[] = []
	for (const token of sqrefTokens(sqref)) {
		try {
			parseRange(token)
		} catch (error) {
			invalid.push({
				field: 'sqref',
				reference: sqref,
				token,
				...(error instanceof Error ? { reason: error.message } : {}),
			})
		}
	}
	return invalid
}

function missingSheetReferencesInSqref(
	sqref: string,
	sheetNameSet: ReadonlySet<string>,
): MissingSheetReference[] {
	const missing: MissingSheetReference[] = []
	for (const token of sqrefTokens(sqref)) {
		try {
			parseRange(token)
		} catch {
			continue
		}
		const parsed = chartReferenceSheetName(token)
		if (parsed.external || !parsed.sheetName) continue
		if (sheetNameSet.has(parsed.sheetName.toLowerCase())) continue
		missing.push({
			field: 'sqref',
			reference: sqref,
			sheetName: parsed.sheetName,
			token,
		})
	}
	return missing
}

function externalReferencesInSqref(sqref: string): ExternalMetadataReference[] {
	const external: ExternalMetadataReference[] = []
	for (const token of sqrefTokens(sqref)) {
		const parsed = chartReferenceSheetName(token)
		if (!parsed.external) continue
		external.push({
			field: 'sqref',
			reference: sqref,
			sheetName: parsed.sheetName ?? token,
			externalTarget: externalTargetFromSheetName(parsed.sheetName ?? token),
			token,
		})
	}
	return external
}

function parseFormulaText(formula: string) {
	const trimmed = formula.trim()
	return parseFormula(trimmed.startsWith('=') ? trimmed.slice(1) : trimmed)
}

function missingSheetReferencesInFormula(
	entry: X14FormulaReferenceEntry,
	sheetNameSet: ReadonlySet<string>,
): MissingSheetReference[] {
	const parsed = parseFormulaText(entry.formula)
	if (!parsed.ok) return []
	const missing: MissingSheetReference[] = []
	for (const sheetName of sheetNamesForFormula(parsed.value)) {
		if (sheetName.startsWith('[')) continue
		if (sheetNameSet.has(sheetName.toLowerCase())) continue
		missing.push({
			field: entry.field,
			reference: entry.formula,
			sheetName,
		})
	}
	return missing
}

function missingTableReferencesInFormula(
	entry: X14FormulaReferenceEntry,
	tableNameSet: ReadonlySet<string>,
): MissingTableReference[] {
	const parsed = parseFormulaText(entry.formula)
	if (!parsed.ok) return []
	const missing: MissingTableReference[] = []
	for (const reference of structuredReferencesForFormula(parsed.value)) {
		if (!reference.table || tableNameSet.has(reference.table.toLowerCase())) continue
		missing.push({
			field: entry.field,
			reference: entry.formula,
			tableName: reference.table,
			...(reference.column ? { column: reference.column } : {}),
			...(reference.endColumn ? { endColumn: reference.endColumn } : {}),
		})
	}
	return missing
}

function externalReferencesInFormula(entry: X14FormulaReferenceEntry): ExternalMetadataReference[] {
	const parsed = parseFormulaText(entry.formula)
	if (!parsed.ok) return []
	const external: ExternalMetadataReference[] = []
	for (const sheetName of sheetNamesForFormula(parsed.value)) {
		if (!sheetName.startsWith('[')) continue
		external.push({
			field: entry.field,
			reference: entry.formula,
			sheetName,
			externalTarget: externalTargetFromSheetName(sheetName),
		})
	}
	return external
}

function deletedReferencesInFormula(entry: X14FormulaReferenceEntry): DeletedFormulaReference[] {
	const parsed = parseFormulaText(entry.formula)
	if (!parsed.ok || !formulaHasDeletedReference(parsed.value)) return []
	return [{ field: entry.field, reference: entry.formula, error: '#REF!' }]
}

function sheetNamesForFormula(node: FormulaNode): string[] {
	const sheetNames = new Map<string, string>()
	for (const ref of extractRefs(node)) {
		for (const sheetName of sheetNamesForFormulaRef(ref)) addFormulaSheetName(sheetNames, sheetName)
	}
	collectSheetQualifiedNameReferences(node, sheetNames)
	return [...sheetNames.values()]
}

function formulaHasDeletedReference(node: FormulaNode): boolean {
	switch (node.type) {
		case 'error':
			return node.value === '#REF!'
		case 'binary':
			return formulaHasDeletedReference(node.left) || formulaHasDeletedReference(node.right)
		case 'dynamicRangeRef':
			return formulaHasDeletedReference(node.start) || formulaHasDeletedReference(node.end)
		case 'unary':
			return formulaHasDeletedReference(node.operand)
		case 'function':
			return node.args.some(formulaHasDeletedReference)
		case 'array':
			return node.rows.some((row) => row.some(formulaHasDeletedReference))
		case 'spillRef':
			return formulaHasDeletedReference(node.target)
		case 'sheetSpanRef':
			return formulaHasDeletedReference(node.target)
		default:
			return false
	}
}

function addFormulaSheetName(sheetNames: Map<string, string>, sheetName: string): void {
	const normalized = sheetName.toLowerCase()
	if (!sheetNames.has(normalized)) sheetNames.set(normalized, sheetName)
}

function sheetNamesForFormulaRef(ref: FormulaRef): string[] {
	if (ref.kind === 'sheetSpan') return [ref.startSheet, ref.endSheet]
	return ref.sheet ? [ref.sheet] : []
}

function structuredReferencesForFormula(node: FormulaNode): StructuredRefNode[] {
	const references: StructuredRefNode[] = []
	collectStructuredReferences(node, references)
	return references
}

function collectStructuredReferences(node: FormulaNode, references: StructuredRefNode[]): void {
	switch (node.type) {
		case 'structuredRef':
			references.push(node)
			break
		case 'binary':
			collectStructuredReferences(node.left, references)
			collectStructuredReferences(node.right, references)
			break
		case 'dynamicRangeRef':
			collectStructuredReferences(node.start, references)
			collectStructuredReferences(node.end, references)
			break
		case 'unary':
			collectStructuredReferences(node.operand, references)
			break
		case 'function':
			for (const arg of node.args) collectStructuredReferences(arg, references)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) collectStructuredReferences(cell, references)
			}
			break
		case 'spillRef':
			collectStructuredReferences(node.target, references)
			break
		case 'sheetSpanRef':
			collectStructuredReferences(node.target, references)
			break
		default:
			break
	}
}

function collectSheetQualifiedNameReferences(
	node: FormulaNode,
	sheetNames: Map<string, string>,
): void {
	switch (node.type) {
		case 'name':
			if (node.sheet) addFormulaSheetName(sheetNames, node.sheet)
			break
		case 'binary':
			collectSheetQualifiedNameReferences(node.left, sheetNames)
			collectSheetQualifiedNameReferences(node.right, sheetNames)
			break
		case 'dynamicRangeRef':
			collectSheetQualifiedNameReferences(node.start, sheetNames)
			collectSheetQualifiedNameReferences(node.end, sheetNames)
			break
		case 'unary':
			collectSheetQualifiedNameReferences(node.operand, sheetNames)
			break
		case 'function':
			for (const arg of node.args) collectSheetQualifiedNameReferences(arg, sheetNames)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) collectSheetQualifiedNameReferences(cell, sheetNames)
			}
			break
		case 'spillRef':
			collectSheetQualifiedNameReferences(node.target, sheetNames)
			break
		case 'sheetSpanRef':
			if (node.startSheet) addFormulaSheetName(sheetNames, node.startSheet)
			if (node.endSheet) addFormulaSheetName(sheetNames, node.endSheet)
			collectSheetQualifiedNameReferences(node.target, sheetNames)
			break
		default:
			break
	}
}

function pushX14MissingTableIssues(
	issues: CheckIssue[],
	params: {
		readonly rule: 'conditional-format-integrity' | 'data-validation-integrity'
		readonly source: MetadataReferenceSource
		readonly sheetName: string
		readonly index: number
		readonly sqref: string
		readonly missing: readonly MissingTableReference[]
		readonly fallbackRef: string
		readonly details?: Readonly<Record<string, unknown>>
	},
): void {
	for (const missing of params.missing) {
		issues.push({
			rule: params.rule,
			severity: 'error',
			message: `${params.source} ${missing.field} on sheet "${params.sheetName}" references missing table "${missing.tableName}"`,
			refs: x14EntryRefs(params.sheetName, params.sqref, params.fallbackRef),
			suggestedFix: 'Repair the structured reference table name before writing workbook metadata.',
			details: {
				kind: metadataMissingTableKind(params.source),
				source: params.source,
				sheetName: params.sheetName,
				index: params.index,
				field: missing.field,
				reference: missing.reference,
				tableName: missing.tableName,
				...(missing.column ? { column: missing.column } : {}),
				...(missing.endColumn ? { endColumn: missing.endColumn } : {}),
				...params.details,
			},
		})
	}
}

function metadataMissingTableKind(source: MetadataReferenceSource): string {
	switch (source) {
		case 'dataValidation':
			return 'data-validation-formula-missing-table'
		case 'conditionalFormat':
			return 'conditional-format-formula-missing-table'
		default:
			return 'x14-formula-missing-table'
	}
}

function formulaHasDetectableReferences(formula: string | undefined): boolean {
	if (!formula) return false
	const parsed = parseFormulaText(formula)
	return (
		parsed.ok && (extractRefs(parsed.value).length > 0 || formulaHasDeletedReference(parsed.value))
	)
}

function x14FormulaLiveFields(entries: readonly X14FormulaReferenceEntry[]): string[] {
	const fields: string[] = []
	for (const entry of entries) {
		if (formulaHasDetectableReferences(entry.formula)) fields.push(entry.field)
	}
	return fields
}

function workbookTableNameSet(wb: Workbook): ReadonlySet<string> {
	const names = new Set<string>()
	for (const sheet of wb.sheets) {
		for (const table of sheet.tables) names.add(table.name.toLowerCase())
	}
	return names
}

function pushDuplicateX14IndexIssues(
	issues: CheckIssue[],
	params: {
		readonly rule: 'conditional-format-integrity' | 'data-validation-integrity'
		readonly source: 'x14ConditionalFormat' | 'x14DataValidation'
		readonly sheetName: string
		readonly entries: readonly { readonly index: number; readonly sqref: string }[]
		readonly kind: 'duplicate-x14-conditional-format-index' | 'duplicate-x14-data-validation-index'
		readonly fallbackPrefix: string
	},
): void {
	const byIndex = new Map<number, readonly { readonly index: number; readonly sqref: string }[]>()
	for (const entry of params.entries) {
		byIndex.set(entry.index, [...(byIndex.get(entry.index) ?? []), entry])
	}
	for (const [index, entries] of byIndex) {
		if (entries.length < 2) continue
		issues.push({
			rule: params.rule,
			severity: 'error',
			message: `${params.source} index ${index} appears ${entries.length} times on sheet "${params.sheetName}"`,
			refs: entries.flatMap((entry) =>
				x14EntryRefs(params.sheetName, entry.sqref, `${params.fallbackPrefix}[${index}]`),
			),
			suggestedFix:
				'Repair duplicate x14 document-order indexes before writing preserved extension metadata.',
			details: {
				kind: params.kind,
				source: params.source,
				sheetName: params.sheetName,
				index,
				count: entries.length,
				sqrefs: entries.map((entry) => entry.sqref),
			},
		})
	}
}

function pushX14MissingSheetIssues(
	issues: CheckIssue[],
	params: {
		readonly rule: 'conditional-format-integrity' | 'data-validation-integrity'
		readonly source: MetadataReferenceSource
		readonly sheetName: string
		readonly index: number
		readonly sqref: string
		readonly missing: readonly MissingSheetReference[]
		readonly sheetNames: readonly string[]
		readonly fallbackRef: string
		readonly details?: Readonly<Record<string, unknown>>
	},
): void {
	for (const missing of params.missing) {
		const closest = findClosestSheetName(missing.sheetName, params.sheetNames)
		issues.push({
			rule: params.rule,
			severity: 'error',
			message: `${params.source} ${missing.field} on sheet "${params.sheetName}" references non-existent sheet "${missing.sheetName}"`,
			refs: x14EntryRefs(params.sheetName, params.sqref, params.fallbackRef),
			suggestedFix: closest
				? `Did you mean sheet "${closest}"?`
				: 'Repair the metadata reference before writing workbook metadata.',
			details: {
				kind: metadataMissingSheetKind(params.source, missing.field),
				source: params.source,
				sheetName: params.sheetName,
				index: params.index,
				field: missing.field,
				reference: missing.reference,
				missingSheet: missing.sheetName,
				...(missing.token ? { token: missing.token } : {}),
				...params.details,
			},
		})
	}
}

function metadataMissingSheetKind(source: MetadataReferenceSource, field: string): string {
	if (source === 'dataValidation') {
		return field === 'sqref'
			? 'data-validation-sqref-missing-sheet'
			: 'data-validation-formula-missing-sheet'
	}
	if (source === 'conditionalFormat') {
		return field === 'sqref'
			? 'conditional-format-sqref-missing-sheet'
			: 'conditional-format-formula-missing-sheet'
	}
	return field === 'sqref' ? 'x14-sqref-missing-sheet' : 'x14-formula-missing-sheet'
}

function pushInvalidX14SqrefIssues(
	issues: CheckIssue[],
	params: {
		readonly rule: 'conditional-format-integrity' | 'data-validation-integrity'
		readonly source: MetadataReferenceSource
		readonly sheetName: string
		readonly index: number
		readonly sqref: string
		readonly invalid: readonly InvalidSqrefReference[]
		readonly fallbackRef: string
		readonly details?: Readonly<Record<string, unknown>>
	},
): void {
	for (const invalid of params.invalid) {
		issues.push({
			rule: params.rule,
			severity: 'error',
			message: `${params.source} sqref on sheet "${params.sheetName}" contains invalid range token "${invalid.token}"`,
			refs: x14EntryRefs(params.sheetName, params.sqref, params.fallbackRef),
			suggestedFix: 'Repair the metadata sqref range before writing workbook metadata.',
			details: {
				kind: metadataInvalidSqrefKind(params.source),
				source: params.source,
				sheetName: params.sheetName,
				index: params.index,
				field: invalid.field,
				reference: invalid.reference,
				token: invalid.token,
				...(invalid.reason ? { reason: invalid.reason } : {}),
				...params.details,
			},
		})
	}
}

function metadataInvalidSqrefKind(source: MetadataReferenceSource): string {
	switch (source) {
		case 'dataValidation':
			return 'data-validation-sqref-invalid'
		case 'conditionalFormat':
			return 'conditional-format-sqref-invalid'
		default:
			return 'x14-sqref-invalid'
	}
}

function pushX14DeletedFormulaReferenceIssues(
	issues: CheckIssue[],
	params: {
		readonly rule: 'conditional-format-integrity' | 'data-validation-integrity'
		readonly source: MetadataReferenceSource
		readonly sheetName: string
		readonly index: number
		readonly sqref: string
		readonly deleted: readonly DeletedFormulaReference[]
		readonly fallbackRef: string
		readonly details?: Readonly<Record<string, unknown>>
	},
): void {
	for (const deleted of params.deleted) {
		issues.push({
			rule: params.rule,
			severity: 'error',
			message: `${params.source} ${deleted.field} on sheet "${params.sheetName}" contains deleted reference ${deleted.error}`,
			refs: x14EntryRefs(params.sheetName, params.sqref, params.fallbackRef),
			suggestedFix: 'Repair the deleted formula reference before writing workbook metadata.',
			details: {
				kind: metadataDeletedFormulaKind(params.source),
				source: params.source,
				sheetName: params.sheetName,
				index: params.index,
				field: deleted.field,
				reference: deleted.reference,
				error: deleted.error,
				...params.details,
			},
		})
	}
}

function metadataDeletedFormulaKind(source: MetadataReferenceSource): string {
	switch (source) {
		case 'dataValidation':
			return 'data-validation-formula-deleted-reference'
		case 'conditionalFormat':
			return 'conditional-format-formula-deleted-reference'
		default:
			return 'x14-formula-deleted-reference'
	}
}

function pushExternalMetadataReferenceIssues(
	issues: CheckIssue[],
	params: {
		readonly rule: 'conditional-format-integrity' | 'data-validation-integrity' | 'table-integrity'
		readonly source:
			| 'conditionalFormat'
			| 'x14ConditionalFormat'
			| 'dataValidation'
			| 'x14DataValidation'
			| 'tableColumn'
		readonly sheetName: string
		readonly index: number
		readonly references: readonly ExternalMetadataReference[]
		readonly refs: readonly string[]
		readonly suggestedFix: string
		readonly details?: Readonly<Record<string, unknown>>
	},
): void {
	for (const reference of params.references) {
		issues.push({
			rule: params.rule,
			severity: 'warning',
			message: `${params.source} ${reference.field} on sheet "${params.sheetName}" references external workbook "${reference.externalTarget}"`,
			refs: params.refs,
			suggestedFix: params.suggestedFix,
			details: {
				kind: `${params.source}-external-reference`,
				source: params.source,
				sheetName: params.sheetName,
				index: params.index,
				field: reference.field,
				reference: reference.reference,
				externalSheet: reference.sheetName,
				externalTarget: reference.externalTarget,
				...(reference.token ? { token: reference.token } : {}),
				...params.details,
			},
		})
	}
}

function x14ConditionalFormatFormulaEntries(
	format: Workbook['sheets'][number]['x14ConditionalFormats'][number],
): X14FormulaReferenceEntry[] {
	const entries: X14FormulaReferenceEntry[] = format.formulas.map((formula, index) => ({
		field: `formulas[${index}]`,
		formula,
	}))
	for (let index = 0; index < (format.colorScale?.cfvo.length ?? 0); index++) {
		const value = format.colorScale?.cfvo[index]?.value
		if (value !== undefined)
			entries.push({ field: `colorScale.cfvo[${index}].value`, formula: value })
	}
	for (let index = 0; index < (format.dataBar?.cfvo.length ?? 0); index++) {
		const value = format.dataBar?.cfvo[index]?.value
		if (value !== undefined) entries.push({ field: `dataBar.cfvo[${index}].value`, formula: value })
	}
	for (let index = 0; index < (format.iconSet?.cfvo.length ?? 0); index++) {
		const value = format.iconSet?.cfvo[index]?.value
		if (value !== undefined) entries.push({ field: `iconSet.cfvo[${index}].value`, formula: value })
	}
	return entries
}

function conditionalFormatRuleFormulaEntries(
	rule: Workbook['sheets'][number]['conditionalFormats'][number]['rules'][number],
	ruleIndex: number,
): X14FormulaReferenceEntry[] {
	const entries: X14FormulaReferenceEntry[] = rule.formulas.map((formula, index) => ({
		field: `rules[${ruleIndex}].formulas[${index}]`,
		formula,
	}))
	for (let index = 0; index < (rule.colorScale?.cfvo.length ?? 0); index++) {
		const value = rule.colorScale?.cfvo[index]?.value
		if (value !== undefined) {
			entries.push({ field: `rules[${ruleIndex}].colorScale.cfvo[${index}].value`, formula: value })
		}
	}
	for (let index = 0; index < (rule.dataBar?.cfvo.length ?? 0); index++) {
		const value = rule.dataBar?.cfvo[index]?.value
		if (value !== undefined) {
			entries.push({ field: `rules[${ruleIndex}].dataBar.cfvo[${index}].value`, formula: value })
		}
	}
	for (let index = 0; index < (rule.iconSet?.cfvo.length ?? 0); index++) {
		const value = rule.iconSet?.cfvo[index]?.value
		if (value !== undefined) {
			entries.push({ field: `rules[${ruleIndex}].iconSet.cfvo[${index}].value`, formula: value })
		}
	}
	return entries
}

function checkX14DataValidationIntegrity(
	wb: Workbook,
	sheetNames: readonly string[],
): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNameSet = new Set(sheetNames.map((name) => name.toLowerCase()))
	const tableNameSet = workbookTableNameSet(wb)

	for (const sheet of wb.sheets) {
		pushDuplicateX14IndexIssues(issues, {
			rule: 'data-validation-integrity',
			source: 'x14DataValidation',
			sheetName: sheet.name,
			entries: sheet.x14DataValidations.map((entry) => ({
				index: entry.index,
				sqref: entry.sqref,
			})),
			kind: 'duplicate-x14-data-validation-index',
			fallbackPrefix: 'x14DataValidation',
		})
		for (let index = 0; index < sheet.dataValidations.length; index++) {
			const validation = sheet.dataValidations[index]
			if (!validation) continue
			const formulaEntries: X14FormulaReferenceEntry[] = []
			if (validation.formula1) {
				formulaEntries.push({ field: 'formula1', formula: validation.formula1 })
			}
			if (validation.formula2) {
				formulaEntries.push({ field: 'formula2', formula: validation.formula2 })
			}
			const fallbackRef = `dataValidation[${index}]`
			pushInvalidX14SqrefIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'dataValidation',
				sheetName: sheet.name,
				index,
				sqref: validation.sqref,
				invalid: invalidSqrefReferences(validation.sqref),
				fallbackRef,
				details: {
					sqref: validation.sqref,
					...(validation.type ? { validationType: validation.type } : {}),
				},
			})
			pushX14MissingSheetIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'dataValidation',
				sheetName: sheet.name,
				index,
				sqref: validation.sqref,
				missing: [
					...missingSheetReferencesInSqref(validation.sqref, sheetNameSet),
					...formulaEntries.flatMap((entry) =>
						missingSheetReferencesInFormula(entry, sheetNameSet),
					),
				],
				sheetNames,
				fallbackRef,
				details: {
					sqref: validation.sqref,
					...(validation.type ? { validationType: validation.type } : {}),
				},
			})
			pushX14MissingTableIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'dataValidation',
				sheetName: sheet.name,
				index,
				sqref: validation.sqref,
				missing: formulaEntries.flatMap((entry) =>
					missingTableReferencesInFormula(entry, tableNameSet),
				),
				fallbackRef,
				details: {
					sqref: validation.sqref,
					...(validation.type ? { validationType: validation.type } : {}),
				},
			})
			pushX14DeletedFormulaReferenceIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'dataValidation',
				sheetName: sheet.name,
				index,
				sqref: validation.sqref,
				deleted: formulaEntries.flatMap(deletedReferencesInFormula),
				fallbackRef,
				details: {
					sqref: validation.sqref,
					...(validation.type ? { validationType: validation.type } : {}),
				},
			})
			pushExternalMetadataReferenceIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'dataValidation',
				sheetName: sheet.name,
				index,
				references: [
					...externalReferencesInSqref(validation.sqref),
					...formulaEntries.flatMap(externalReferencesInFormula),
				],
				refs: x14EntryRefs(sheet.name, validation.sqref, fallbackRef),
				suggestedFix:
					'Replace validation references with local workbook ranges or verify the external link metadata before editing validations.',
				details: {
					sqref: validation.sqref,
					...(validation.type ? { validationType: validation.type } : {}),
				},
			})
		}
		for (const validation of sheet.x14DataValidations) {
			const formulaEntries: X14FormulaReferenceEntry[] = []
			if (validation.formula1) {
				formulaEntries.push({ field: 'formula1', formula: validation.formula1 })
			}
			if (validation.formula2) {
				formulaEntries.push({ field: 'formula2', formula: validation.formula2 })
			}
			const fallbackRef = `x14DataValidation[${validation.index}]`
			pushInvalidX14SqrefIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'x14DataValidation',
				sheetName: sheet.name,
				index: validation.index,
				sqref: validation.sqref,
				invalid: invalidSqrefReferences(validation.sqref),
				fallbackRef,
			})
			pushX14MissingSheetIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'x14DataValidation',
				sheetName: sheet.name,
				index: validation.index,
				sqref: validation.sqref,
				missing: [
					...missingSheetReferencesInSqref(validation.sqref, sheetNameSet),
					...formulaEntries.flatMap((entry) =>
						missingSheetReferencesInFormula(entry, sheetNameSet),
					),
				],
				sheetNames,
				fallbackRef,
			})
			pushX14MissingTableIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'x14DataValidation',
				sheetName: sheet.name,
				index: validation.index,
				sqref: validation.sqref,
				missing: formulaEntries.flatMap((entry) =>
					missingTableReferencesInFormula(entry, tableNameSet),
				),
				fallbackRef,
			})
			pushX14DeletedFormulaReferenceIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'x14DataValidation',
				sheetName: sheet.name,
				index: validation.index,
				sqref: validation.sqref,
				deleted: formulaEntries.flatMap(deletedReferencesInFormula),
				fallbackRef,
			})
			pushExternalMetadataReferenceIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'x14DataValidation',
				sheetName: sheet.name,
				index: validation.index,
				references: [
					...externalReferencesInSqref(validation.sqref),
					...formulaEntries.flatMap(externalReferencesInFormula),
				],
				refs: x14EntryRefs(sheet.name, validation.sqref, fallbackRef),
				suggestedFix:
					'Replace x14 validation references with local workbook ranges or verify the external link metadata before writing preserved extension metadata.',
				details: {
					sqref: validation.sqref,
					...(validation.type ? { validationType: validation.type } : {}),
				},
			})

			if (!validation.deleted) continue
			const liveFields = [
				...(validation.sqref.trim() ? ['sqref'] : []),
				...x14FormulaLiveFields(formulaEntries),
			]
			if (liveFields.length === 0) continue
			issues.push({
				rule: 'data-validation-integrity',
				severity: 'warning',
				message: `Deleted x14DataValidation on sheet "${sheet.name}" still carries live references`,
				refs: x14EntryRefs(sheet.name, validation.sqref, fallbackRef),
				suggestedFix:
					'Clear the deleted x14 data-validation entry or restore it before writing preserved extension metadata.',
				details: {
					kind: 'deleted-x14-data-validation-live-refs',
					source: 'x14DataValidation',
					sheetName: sheet.name,
					index: validation.index,
					liveFields,
					sqref: validation.sqref,
				},
			})
		}
	}

	return issues
}

function firstOverlappingRange(
	left: ConditionalFormatPriorityEntry,
	right: ConditionalFormatPriorityEntry,
): [RangeRef, RangeRef] | null {
	for (const leftRange of left.ranges) {
		for (const rightRange of right.ranges) {
			if (rangesOverlap2D(leftRange, rightRange)) return [leftRange, rightRange]
		}
	}
	return null
}

function checkConditionalFormatIntegrity(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNames = wb.sheets.map((sheet) => sheet.name)
	const sheetNameSet = new Set(sheetNames.map((name) => name.toLowerCase()))
	const tableNameSet = workbookTableNameSet(wb)
	for (const sheet of wb.sheets) {
		const entries: ConditionalFormatPriorityEntry[] = []
		for (let formatIndex = 0; formatIndex < sheet.conditionalFormats.length; formatIndex++) {
			const format = sheet.conditionalFormats[formatIndex]
			if (!format) continue
			const ranges = parseSqrefRanges(format.sqref)
			const fallbackRef = `conditionalFormat[${formatIndex}]`
			pushInvalidX14SqrefIssues(issues, {
				rule: 'conditional-format-integrity',
				source: 'conditionalFormat',
				sheetName: sheet.name,
				index: formatIndex,
				sqref: format.sqref,
				invalid: invalidSqrefReferences(format.sqref),
				fallbackRef,
				details: { sqref: format.sqref },
			})
			pushX14MissingSheetIssues(issues, {
				rule: 'conditional-format-integrity',
				source: 'conditionalFormat',
				sheetName: sheet.name,
				index: formatIndex,
				sqref: format.sqref,
				missing: missingSheetReferencesInSqref(format.sqref, sheetNameSet),
				sheetNames,
				fallbackRef,
				details: { sqref: format.sqref },
			})
			pushExternalMetadataReferenceIssues(issues, {
				rule: 'conditional-format-integrity',
				source: 'conditionalFormat',
				sheetName: sheet.name,
				index: formatIndex,
				references: externalReferencesInSqref(format.sqref),
				refs: x14EntryRefs(sheet.name, format.sqref, fallbackRef),
				suggestedFix:
					'Replace conditional-format ranges with local workbook ranges or verify the external link metadata before editing conditional formats.',
				details: {
					sqref: format.sqref,
				},
			})
			for (let ruleIndex = 0; ruleIndex < format.rules.length; ruleIndex++) {
				const rule = format.rules[ruleIndex]
				if (!rule) continue
				const formulaEntries = conditionalFormatRuleFormulaEntries(rule, ruleIndex)
				const ruleDetails = {
					sqref: format.sqref,
					ruleIndex,
					ruleType: rule.type,
				}
				pushX14MissingSheetIssues(issues, {
					rule: 'conditional-format-integrity',
					source: 'conditionalFormat',
					sheetName: sheet.name,
					index: formatIndex,
					sqref: format.sqref,
					missing: formulaEntries.flatMap((entry) =>
						missingSheetReferencesInFormula(entry, sheetNameSet),
					),
					sheetNames,
					fallbackRef,
					details: ruleDetails,
				})
				pushX14MissingTableIssues(issues, {
					rule: 'conditional-format-integrity',
					source: 'conditionalFormat',
					sheetName: sheet.name,
					index: formatIndex,
					sqref: format.sqref,
					missing: formulaEntries.flatMap((entry) =>
						missingTableReferencesInFormula(entry, tableNameSet),
					),
					fallbackRef,
					details: ruleDetails,
				})
				pushX14DeletedFormulaReferenceIssues(issues, {
					rule: 'conditional-format-integrity',
					source: 'conditionalFormat',
					sheetName: sheet.name,
					index: formatIndex,
					sqref: format.sqref,
					deleted: formulaEntries.flatMap(deletedReferencesInFormula),
					fallbackRef,
					details: ruleDetails,
				})
				pushExternalMetadataReferenceIssues(issues, {
					rule: 'conditional-format-integrity',
					source: 'conditionalFormat',
					sheetName: sheet.name,
					index: formatIndex,
					references: formulaEntries.flatMap(externalReferencesInFormula),
					refs: x14EntryRefs(sheet.name, format.sqref, fallbackRef),
					suggestedFix:
						'Replace conditional-format formulas with local workbook ranges or verify the external link metadata before editing conditional formats.',
					details: ruleDetails,
				})
				entries.push({
					source: 'conditionalFormat',
					sheetName: sheet.name,
					sqref: format.sqref,
					formatIndex,
					ruleIndex,
					ruleType: rule.type,
					ranges,
					...(rule.priority !== undefined ? { priority: rule.priority } : {}),
				})
			}
		}
		pushDuplicateX14IndexIssues(issues, {
			rule: 'conditional-format-integrity',
			source: 'x14ConditionalFormat',
			sheetName: sheet.name,
			entries: sheet.x14ConditionalFormats.map((entry) => ({
				index: entry.index,
				sqref: entry.sqref,
			})),
			kind: 'duplicate-x14-conditional-format-index',
			fallbackPrefix: 'x14ConditionalFormat',
		})
		for (const format of sheet.x14ConditionalFormats) {
			const fallbackRef = `x14ConditionalFormat[${format.index}]`
			const formulaEntries = x14ConditionalFormatFormulaEntries(format)
			pushInvalidX14SqrefIssues(issues, {
				rule: 'conditional-format-integrity',
				source: 'x14ConditionalFormat',
				sheetName: sheet.name,
				index: format.index,
				sqref: format.sqref,
				invalid: invalidSqrefReferences(format.sqref),
				fallbackRef,
			})
			pushX14MissingSheetIssues(issues, {
				rule: 'conditional-format-integrity',
				source: 'x14ConditionalFormat',
				sheetName: sheet.name,
				index: format.index,
				sqref: format.sqref,
				missing: [
					...missingSheetReferencesInSqref(format.sqref, sheetNameSet),
					...formulaEntries.flatMap((entry) =>
						missingSheetReferencesInFormula(entry, sheetNameSet),
					),
				],
				sheetNames,
				fallbackRef,
			})
			pushX14MissingTableIssues(issues, {
				rule: 'conditional-format-integrity',
				source: 'x14ConditionalFormat',
				sheetName: sheet.name,
				index: format.index,
				sqref: format.sqref,
				missing: formulaEntries.flatMap((entry) =>
					missingTableReferencesInFormula(entry, tableNameSet),
				),
				fallbackRef,
			})
			pushX14DeletedFormulaReferenceIssues(issues, {
				rule: 'conditional-format-integrity',
				source: 'x14ConditionalFormat',
				sheetName: sheet.name,
				index: format.index,
				sqref: format.sqref,
				deleted: formulaEntries.flatMap(deletedReferencesInFormula),
				fallbackRef,
			})
			pushExternalMetadataReferenceIssues(issues, {
				rule: 'conditional-format-integrity',
				source: 'x14ConditionalFormat',
				sheetName: sheet.name,
				index: format.index,
				references: [
					...externalReferencesInSqref(format.sqref),
					...formulaEntries.flatMap(externalReferencesInFormula),
				],
				refs: x14EntryRefs(sheet.name, format.sqref, fallbackRef),
				suggestedFix:
					'Replace x14 conditional-format references with local workbook ranges or verify the external link metadata before writing preserved extension metadata.',
				details: {
					sqref: format.sqref,
					...(format.type ? { ruleType: format.type } : {}),
				},
			})

			if (format.deleted) {
				const liveFields = [
					...(format.sqref.trim() ? ['sqref'] : []),
					...x14FormulaLiveFields(formulaEntries),
				]
				if (liveFields.length > 0) {
					issues.push({
						rule: 'conditional-format-integrity',
						severity: 'warning',
						message: `Deleted x14ConditionalFormat on sheet "${sheet.name}" still carries live references`,
						refs: x14EntryRefs(sheet.name, format.sqref, fallbackRef),
						suggestedFix:
							'Clear the deleted x14 conditional-format entry or restore it before writing preserved extension metadata.',
						details: {
							kind: 'deleted-x14-conditional-format-live-refs',
							source: 'x14ConditionalFormat',
							sheetName: sheet.name,
							index: format.index,
							liveFields,
							sqref: format.sqref,
						},
					})
				}
				continue
			}
			entries.push({
				source: 'x14ConditionalFormat',
				sheetName: sheet.name,
				sqref: format.sqref,
				formatIndex: format.index,
				ranges: parseSqrefRanges(format.sqref),
				...(format.priority !== undefined ? { priority: format.priority } : {}),
				...(format.type ? { ruleType: format.type } : {}),
			})
		}
		for (const entry of entries) {
			if (entry.priority !== undefined && entry.priority <= 0) {
				issues.push({
					rule: 'conditional-format-integrity',
					severity: 'warning',
					message: `Conditional format on sheet "${sheet.name}" has non-positive priority ${entry.priority}`,
					refs: [`${sheet.name}!${entry.sqref}`],
					suggestedFix:
						'Assign a positive, unique priority before relying on conditional-format order.',
					details: {
						kind: 'conditional-format-nonpositive-priority',
						source: entry.source,
						priority: entry.priority,
						formatIndex: entry.formatIndex,
						...(entry.ruleIndex !== undefined ? { ruleIndex: entry.ruleIndex } : {}),
						...(entry.ruleType ? { ruleType: entry.ruleType } : {}),
					},
				})
			}
		}
		for (let i = 0; i < entries.length; i++) {
			const left = entries[i]
			if (!left) continue
			for (let j = i + 1; j < entries.length; j++) {
				const right = entries[j]
				if (!right || left.priority === undefined || left.priority !== right.priority) continue
				const overlap = firstOverlappingRange(left, right)
				if (!overlap) continue
				issues.push({
					rule: 'conditional-format-integrity',
					severity: 'warning',
					message: `Overlapping conditional formats on sheet "${sheet.name}" share priority ${left.priority}`,
					refs: [
						`${sheet.name}!${rangeToA1(overlap[0])}`,
						`${sheet.name}!${rangeToA1(overlap[1])}`,
					],
					suggestedFix:
						'Give overlapping conditional-format rules distinct priorities before editing or reordering them.',
					details: {
						kind: 'conditional-format-priority-collision',
						priority: left.priority,
						left: {
							source: left.source,
							sqref: left.sqref,
							formatIndex: left.formatIndex,
							...(left.ruleIndex !== undefined ? { ruleIndex: left.ruleIndex } : {}),
							...(left.ruleType ? { ruleType: left.ruleType } : {}),
						},
						right: {
							source: right.source,
							sqref: right.sqref,
							formatIndex: right.formatIndex,
							...(right.ruleIndex !== undefined ? { ruleIndex: right.ruleIndex } : {}),
							...(right.ruleType ? { ruleType: right.ruleType } : {}),
						},
					},
				})
			}
		}
		for (let i = 0; i < entries.length; i++) {
			const left = entries[i]
			if (!left) continue
			for (let j = i + 1; j < entries.length; j++) {
				const right = entries[j]
				if (!right || left.source === right.source) continue
				if (left.priority !== undefined && right.priority !== undefined) continue
				const overlap = firstOverlappingRange(left, right)
				if (!overlap) continue
				issues.push({
					rule: 'conditional-format-integrity',
					severity: 'warning',
					message: `Overlapping legacy and x14 conditional formats on sheet "${sheet.name}" have ambiguous priority metadata`,
					refs: [
						`${sheet.name}!${rangeToA1(overlap[0])}`,
						`${sheet.name}!${rangeToA1(overlap[1])}`,
					],
					suggestedFix:
						'Assign explicit priorities or separate the ranges before editing conditional-format order.',
					details: {
						kind: 'ambiguous-x14-legacy-overlap',
						left: {
							source: left.source,
							sqref: left.sqref,
							formatIndex: left.formatIndex,
							...(left.priority !== undefined ? { priority: left.priority } : {}),
							...(left.ruleIndex !== undefined ? { ruleIndex: left.ruleIndex } : {}),
							...(left.ruleType ? { ruleType: left.ruleType } : {}),
						},
						right: {
							source: right.source,
							sqref: right.sqref,
							formatIndex: right.formatIndex,
							...(right.priority !== undefined ? { priority: right.priority } : {}),
							...(right.ruleIndex !== undefined ? { ruleIndex: right.ruleIndex } : {}),
							...(right.ruleType ? { ruleType: right.ruleType } : {}),
						},
					},
				})
			}
		}
	}
	return issues
}

interface ThreadedCommentIntegrityEntry {
	readonly sheetName: string
	readonly ref: string
	readonly index: number
	readonly partPath: string
	readonly id?: string
	readonly parentId?: string
	readonly personId?: string
	readonly author?: string
}

function isThreadedCommentRelationshipType(relationshipType: string | undefined): boolean {
	return relationshipType?.toLowerCase().endsWith('/relationships/threadedcomment') ?? false
}

function isCommentsRelationshipType(relationshipType: string | undefined): boolean {
	return relationshipType?.toLowerCase().endsWith('/relationships/comments') ?? false
}

function isVmlDrawingRelationshipType(relationshipType: string | undefined): boolean {
	return relationshipType?.toLowerCase().endsWith('/relationships/vmldrawing') ?? false
}

function isWorksheetPartPath(partPath: string): boolean {
	return /(^|\/)worksheets\/sheet\d+\.xml$/i.test(partPath)
}

const MAX_WORKSHEET_ROW_COUNT = 1_048_576
const MAX_WORKSHEET_COLUMN_COUNT = 16_384

function parseWorksheetCellRef(ref: string): { readonly row: number; readonly col: number } | null {
	try {
		const cellRef = parseA1(ref)
		if (
			cellRef.row < 0 ||
			cellRef.col < 0 ||
			cellRef.row >= MAX_WORKSHEET_ROW_COUNT ||
			cellRef.col >= MAX_WORKSHEET_COLUMN_COUNT
		) {
			return null
		}
		return cellRef
	} catch {
		return null
	}
}

function isThreadedCommentPart(part: VerifyPackageGraphPart): boolean {
	return (
		part.featureFamily === 'preservedThreadedComments' && /\/threadedComments\//i.test(part.path)
	)
}

function isThreadedPersonPart(part: VerifyPackageGraphPart): boolean {
	return part.featureFamily === 'preservedThreadedComments' && /\/persons\//i.test(part.path)
}

function checkThreadedCommentIntegrity(
	wb: Workbook,
	packageGraph?: VerifyPackageGraph,
): CheckIssue[] {
	const issues: CheckIssue[] = []
	const graphPartByPath = new Map(packageGraph?.parts.map((part) => [part.path, part]) ?? [])
	const graphRelationshipsByTarget = new Map<string, VerifyPackageGraphRelationship[]>()
	for (const relationship of packageGraph?.relationships ?? []) {
		if (!relationship.resolvedTarget) continue
		const relationships = graphRelationshipsByTarget.get(relationship.resolvedTarget)
		if (relationships) relationships.push(relationship)
		else graphRelationshipsByTarget.set(relationship.resolvedTarget, [relationship])
	}
	const idsByPart = new Map<string, Map<string, ThreadedCommentIntegrityEntry>>()
	const idsByWorkbook = new Map<string, ThreadedCommentIntegrityEntry>()
	const rootRefsByPart = new Map<string, Map<string, ThreadedCommentIntegrityEntry>>()
	const authorsByPersonId = new Map<string, ThreadedCommentIntegrityEntry>()
	const claimedPartsBySheet = new Map<string, Set<string>>()
	const sheetsByName = new Map(wb.sheets.map((sheet) => [sheet.name, sheet]))
	let threadedCommentCount = 0
	let threadedCommentsWithPersonIds = 0

	for (const sheet of wb.sheets) {
		for (let index = 0; index < sheet.threadedComments.length; index++) {
			const comment = sheet.threadedComments[index]
			if (!comment) continue
			threadedCommentCount++
			const partPath = comment.partPath ?? '(unknown threaded comment part)'
			if (comment.partPath) {
				let sheetNames = claimedPartsBySheet.get(comment.partPath)
				if (!sheetNames) {
					sheetNames = new Set()
					claimedPartsBySheet.set(comment.partPath, sheetNames)
				}
				sheetNames.add(sheet.name)
			}
			if (!comment.ref) {
				issues.push({
					rule: 'threaded-comment-integrity',
					severity: 'error',
					message: `Threaded comment at index ${index} on sheet "${sheet.name}" is missing a cell reference`,
					refs: [`${sheet.name}!(missing ref)`],
					suggestedFix:
						'Restore the threadedComment ref before writing; comments without cell refs cannot be bound to worksheet cells.',
					details: {
						kind: 'missing-threaded-comment-ref',
						partPath,
						commentIndex: index,
						...(comment.id ? { id: comment.id } : {}),
					},
				})
			} else {
				if (!parseWorksheetCellRef(comment.ref)) {
					issues.push({
						rule: 'threaded-comment-integrity',
						severity: 'error',
						message: `Threaded comment on sheet "${sheet.name}" has invalid cell reference "${comment.ref}"`,
						refs: [`${sheet.name}!${comment.ref}`],
						suggestedFix:
							'Repair the threadedComment ref before writing; invalid refs cannot be round-tripped safely.',
						details: {
							kind: 'invalid-threaded-comment-ref',
							partPath,
							commentIndex: index,
							ref: comment.ref,
							...(comment.id ? { id: comment.id } : {}),
						},
					})
				}
			}
			if (!comment.partPath) {
				issues.push({
					rule: 'threaded-comment-integrity',
					severity: 'warning',
					message: `Threaded comment at ${sheet.name}!${comment.ref || '(missing ref)'} has no source part path`,
					refs: [`${sheet.name}!${comment.ref || '(missing ref)'}`],
					suggestedFix:
						'Preserve the threadedComments part path before editing or writing threaded comment metadata.',
					details: {
						kind: 'missing-threaded-comment-part-path',
						commentIndex: index,
						...(comment.id ? { id: comment.id } : {}),
					},
				})
			}
			if (comment.id) {
				let ids = idsByPart.get(partPath)
				if (!ids) {
					ids = new Map()
					idsByPart.set(partPath, ids)
				}
				const existing = ids.get(comment.id)
				if (existing) {
					issues.push({
						rule: 'threaded-comment-integrity',
						severity: 'warning',
						message: `Duplicate threaded comment id "${comment.id}" on sheet "${sheet.name}"`,
						refs: [`${existing.sheetName}!${existing.ref}`, `${sheet.name}!${comment.ref}`],
						suggestedFix:
							'Inspect the threadedComments part before editing this thread; duplicate ids make replies ambiguous.',
						details: {
							kind: 'duplicate-threaded-comment-id',
							partPath,
							id: comment.id,
							firstCommentIndex: existing.index,
							duplicateCommentIndex: index,
							...(existing.sheetName !== sheet.name
								? { firstSheetName: existing.sheetName, duplicateSheetName: sheet.name }
								: {}),
						},
					})
				} else {
					ids.set(comment.id, {
						sheetName: sheet.name,
						ref: comment.ref,
						index,
						partPath,
						id: comment.id,
						...(comment.parentId ? { parentId: comment.parentId } : {}),
					})
				}
				const existingWorkbookId = idsByWorkbook.get(comment.id)
				if (existingWorkbookId && existingWorkbookId.partPath !== partPath) {
					issues.push({
						rule: 'threaded-comment-integrity',
						severity: 'warning',
						message: `Duplicate threaded comment id "${comment.id}" across threaded comment parts`,
						refs: [
							`${existingWorkbookId.sheetName}!${existingWorkbookId.ref}`,
							`${sheet.name}!${comment.ref}`,
						],
						suggestedFix:
							'Assign workbook-unique threadedComment ids before editing replies or merging threaded comment parts.',
						details: {
							kind: 'duplicate-threaded-comment-id-across-parts',
							id: comment.id,
							firstPartPath: existingWorkbookId.partPath,
							duplicatePartPath: partPath,
							firstSheetName: existingWorkbookId.sheetName,
							duplicateSheetName: sheet.name,
							firstCommentIndex: existingWorkbookId.index,
							duplicateCommentIndex: index,
						},
					})
				} else if (!existingWorkbookId) {
					idsByWorkbook.set(comment.id, {
						sheetName: sheet.name,
						ref: comment.ref,
						index,
						partPath,
						id: comment.id,
					})
				}
			} else {
				issues.push({
					rule: 'threaded-comment-integrity',
					severity: 'warning',
					message: `Threaded comment at ${sheet.name}!${comment.ref || '(missing ref)'} is missing an id`,
					refs: [`${sheet.name}!${comment.ref || '(missing ref)'}`],
					suggestedFix:
						'Restore the threadedComment id before editing replies; missing ids make thread parentage ambiguous.',
					details: {
						kind: 'missing-threaded-comment-id',
						partPath,
						commentIndex: index,
					},
				})
			}
			if (!comment.parentId && comment.ref) {
				let rootRefs = rootRefsByPart.get(partPath)
				if (!rootRefs) {
					rootRefs = new Map()
					rootRefsByPart.set(partPath, rootRefs)
				}
				const existing = rootRefs.get(comment.ref.toUpperCase())
				if (existing) {
					issues.push({
						rule: 'threaded-comment-integrity',
						severity: 'warning',
						message: `Duplicate root threaded comment ref "${comment.ref}" in "${partPath}"`,
						refs: [`${existing.sheetName}!${existing.ref}`, `${sheet.name}!${comment.ref}`],
						suggestedFix:
							'Merge duplicate thread roots or give each threaded comment root a distinct cell ref before writing.',
						details: {
							kind: 'duplicate-threaded-comment-root-ref',
							partPath,
							ref: comment.ref,
							firstCommentIndex: existing.index,
							duplicateCommentIndex: index,
						},
					})
				} else {
					rootRefs.set(comment.ref.toUpperCase(), {
						sheetName: sheet.name,
						ref: comment.ref,
						index,
						partPath,
						...(comment.id ? { id: comment.id } : {}),
					})
				}
			}
			if (comment.personId && !comment.author) {
				threadedCommentsWithPersonIds++
				issues.push({
					rule: 'threaded-comment-integrity',
					severity: 'warning',
					message: `Threaded comment at ${sheet.name}!${comment.ref} references unknown person id "${comment.personId}"`,
					refs: [`${sheet.name}!${comment.ref}`],
					suggestedFix:
						'Preserve or repair the threaded comment persons part before author-sensitive edits.',
					details: {
						kind: 'threaded-comment-unknown-person-id',
						partPath,
						commentIndex: index,
						personId: comment.personId,
						...(comment.id ? { id: comment.id } : {}),
					},
				})
			} else if (comment.personId) {
				threadedCommentsWithPersonIds++
				const author = comment.author
				if (!author) continue
				const existingAuthor = authorsByPersonId.get(comment.personId)
				if (!existingAuthor) {
					authorsByPersonId.set(comment.personId, {
						sheetName: sheet.name,
						ref: comment.ref,
						index,
						partPath,
						...(comment.id ? { id: comment.id } : {}),
						personId: comment.personId,
						author,
					})
				} else if (existingAuthor.author !== author) {
					issues.push({
						rule: 'threaded-comment-integrity',
						severity: 'warning',
						message: `Threaded comments bind person id "${comment.personId}" to conflicting authors`,
						refs: [
							`${existingAuthor.sheetName}!${existingAuthor.ref}`,
							`${sheet.name}!${comment.ref}`,
						],
						suggestedFix:
							'Repair the threaded comment persons binding before author-sensitive edits; one personId must identify one display name.',
						details: {
							kind: 'threaded-comment-person-author-conflict',
							personId: comment.personId,
							firstAuthor: existingAuthor.author,
							duplicateAuthor: author,
							firstPartPath: existingAuthor.partPath,
							duplicatePartPath: partPath,
							firstSheetName: existingAuthor.sheetName,
							duplicateSheetName: sheet.name,
							firstCommentIndex: existingAuthor.index,
							duplicateCommentIndex: index,
							...(existingAuthor.id ? { firstId: existingAuthor.id } : {}),
							...(comment.id ? { duplicateId: comment.id } : {}),
						},
					})
				}
			}
		}
	}
	for (const [partPath, sheetNames] of claimedPartsBySheet) {
		if (sheetNames.size <= 1) continue
		issues.push({
			rule: 'threaded-comment-integrity',
			severity: 'warning',
			message: `Threaded comments part "${partPath}" is claimed by multiple sheets`,
			refs: [...sheetNames],
			suggestedFix:
				'Give each worksheet its own threadedComments relationship before writing threaded comment parts.',
			details: {
				kind: 'threaded-comment-part-multiple-sheet-owners',
				partPath,
				sheetNames: [...sheetNames],
			},
		})
	}
	for (const sheet of wb.sheets) {
		for (let index = 0; index < sheet.threadedComments.length; index++) {
			const comment = sheet.threadedComments[index]
			if (!comment?.parentId) continue
			const partPath = comment.partPath ?? '(unknown threaded comment part)'
			if (comment.id && comment.parentId === comment.id) {
				issues.push({
					rule: 'threaded-comment-integrity',
					severity: 'warning',
					message: `Threaded comment at ${sheet.name}!${comment.ref} references itself as parent id "${comment.parentId}"`,
					refs: [`${sheet.name}!${comment.ref}`],
					suggestedFix:
						'Repair the threadedComment parentId before editing replies; self-parented replies cannot form a valid thread.',
					details: {
						kind: 'self-parented-threaded-comment',
						partPath,
						commentIndex: index,
						id: comment.id,
						parentId: comment.parentId,
					},
				})
				continue
			}
			if (idsByPart.get(partPath)?.has(comment.parentId)) continue
			issues.push({
				rule: 'threaded-comment-integrity',
				severity: 'warning',
				message: `Threaded comment at ${sheet.name}!${comment.ref} references missing parent id "${comment.parentId}"`,
				refs: [`${sheet.name}!${comment.ref}`],
				suggestedFix:
					'Inspect the threadedComments part before editing replies; the parent thread id is missing.',
				details: {
					kind: 'missing-threaded-comment-parent-id',
					partPath,
					commentIndex: index,
					parentId: comment.parentId,
					...(comment.id ? { id: comment.id } : {}),
				},
			})
		}
	}
	issues.push(...checkThreadedCommentParentCycles(wb, idsByPart))
	if (!packageGraph) return issues

	const personParts = packageGraph.parts.filter(isThreadedPersonPart)
	for (const part of personParts) {
		const personsById = new Map<string, VerifyThreadedCommentPersonEntry>()
		for (const person of part.threadedCommentPersons ?? []) {
			const existing = personsById.get(person.id)
			if (!existing) {
				personsById.set(person.id, person)
				continue
			}
			issues.push({
				rule: 'threaded-comment-integrity',
				severity: 'warning',
				message: `Threaded comment persons part "${part.path}" contains duplicate person id "${person.id}"`,
				refs: [part.path],
				suggestedFix:
					'Repair duplicate person ids before author-sensitive threaded comment edits; personId binding must identify one author.',
				details: {
					kind: 'duplicate-threaded-comment-person-id',
					partPath: part.path,
					personId: person.id,
					firstPersonIndex: existing.index,
					duplicatePersonIndex: person.index,
					...(existing.displayName ? { firstDisplayName: existing.displayName } : {}),
					...(person.displayName ? { duplicateDisplayName: person.displayName } : {}),
				},
			})
		}
	}
	if (threadedCommentsWithPersonIds > 0 && personParts.length === 0) {
		issues.push({
			rule: 'threaded-comment-integrity',
			severity: 'warning',
			message: 'Threaded comments use person ids but the package graph has no persons part',
			refs: [...claimedPartsBySheet.keys()],
			suggestedFix:
				'Restore the threaded comment persons part before author-sensitive threaded comment edits.',
			details: {
				kind: 'missing-threaded-comment-persons-part',
				threadedCommentPartPaths: [...claimedPartsBySheet.keys()],
			},
		})
	}
	if (threadedCommentsWithPersonIds > 0 && personParts.length > 1) {
		issues.push({
			rule: 'threaded-comment-integrity',
			severity: 'warning',
			message: 'Threaded comments use person ids but the package graph has multiple persons parts',
			refs: personParts.map((part) => part.path),
			suggestedFix:
				'Inspect the threaded comment persons sidecars before author-sensitive edits; duplicate persons parts make personId binding ambiguous.',
			details: {
				kind: 'ambiguous-threaded-comment-persons-parts',
				threadedCommentPartPaths: [...claimedPartsBySheet.keys()],
				personParts: personParts.map((part) => ({
					partPath: part.path,
					ownerScope: part.ownerScope,
					contentType: part.contentType,
					incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
						queryTableRelationshipDetails,
					),
				})),
			},
		})
	}
	if (threadedCommentCount > 0 && threadedCommentsWithPersonIds === 0) {
		for (const part of personParts) {
			issues.push({
				rule: 'threaded-comment-integrity',
				severity: 'warning',
				message: `Threaded comment persons package part "${part.path}" is not referenced by any threaded comments`,
				refs: [part.path],
				suggestedFix:
					'Remove the unclaimed persons sidecar or restore threaded comment personId bindings before writing.',
				details: {
					kind: 'orphan-threaded-comment-persons-part',
					partPath: part.path,
					ownerScope: part.ownerScope,
					contentType: part.contentType,
					threadedCommentsWithPersonIds,
					incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
						queryTableRelationshipDetails,
					),
				},
			})
		}
	}
	if (threadedCommentCount === 0) {
		for (const part of personParts) {
			issues.push({
				rule: 'threaded-comment-integrity',
				severity: 'warning',
				message: `Threaded comment persons package part "${part.path}" is not claimed by any threaded comment model`,
				refs: [part.path],
				suggestedFix:
					'Reconnect the persons part to sheet.threadedComments or remove the orphan persons sidecar before writing.',
				details: {
					kind: 'orphan-threaded-comment-persons-part',
					partPath: part.path,
					ownerScope: part.ownerScope,
					contentType: part.contentType,
					incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
						queryTableRelationshipDetails,
					),
				},
			})
		}
	}
	for (const [partPath, sheetNames] of claimedPartsBySheet) {
		const graphPart = graphPartByPath.get(partPath)
		if (!graphPart) {
			issues.push({
				rule: 'threaded-comment-integrity',
				severity: 'error',
				message: `Threaded comments model references missing package part "${partPath}"`,
				refs: [partPath],
				suggestedFix:
					'Restore the threadedComments package part or clear stale threaded comment metadata before writing.',
				details: {
					kind: 'missing-threaded-comment-part',
					partPath,
					sheetNames: [...sheetNames],
				},
			})
			continue
		}
		const incomingRelationships = graphRelationshipsByTarget.get(partPath) ?? []
		const threadedRelationships = incomingRelationships.filter((relationship) =>
			isThreadedCommentRelationshipType(relationship.type),
		)
		if (threadedRelationships.length === 0) {
			issues.push({
				rule: 'threaded-comment-integrity',
				severity: 'error',
				message: `Threaded comments part "${partPath}" has no worksheet threadedComment relationship`,
				refs: [partPath],
				suggestedFix:
					'Restore the worksheet threadedComment relationship before writing threaded comment metadata.',
				details: {
					kind: 'threaded-comment-relationship-binding-missing',
					partPath,
					graphPart,
					incomingRelationships: incomingRelationships.map(queryTableRelationshipDetails),
				},
			})
			continue
		}
		const worksheetOwners = new Set(
			threadedRelationships
				.filter((relationship) => isWorksheetPartPath(relationship.sourcePartPath))
				.map((relationship) => relationship.sourcePartPath),
		)
		if (worksheetOwners.size === 0) {
			issues.push({
				rule: 'threaded-comment-integrity',
				severity: 'warning',
				message: `Threaded comments part "${partPath}" is not owned by a worksheet relationship`,
				refs: [partPath],
				suggestedFix:
					'Bind the threadedComments part from its owning worksheet before writing threaded comments.',
				details: {
					kind: 'threaded-comment-part-non-worksheet-owner',
					partPath,
					relationships: threadedRelationships.map(queryTableRelationshipDetails),
				},
			})
		}
		if (worksheetOwners.size > 1) {
			issues.push({
				rule: 'threaded-comment-integrity',
				severity: 'warning',
				message: `Threaded comments part "${partPath}" has multiple worksheet relationship owners`,
				refs: [...worksheetOwners, partPath],
				suggestedFix:
					'Give each worksheet a distinct threadedComments part before writing threaded comment metadata.',
				details: {
					kind: 'threaded-comment-part-ambiguous-package-owner',
					partPath,
					worksheetPartPaths: [...worksheetOwners],
				},
			})
		}
		if (worksheetOwners.size === 1 && sheetNames.size === 1) {
			const sheetName = [...sheetNames][0]
			const sheetPartPath = sheetName
				? sheetsByName.get(sheetName)?.preservedXml?.partPath
				: undefined
			const ownerPartPath = [...worksheetOwners][0]
			if (sheetName && sheetPartPath && ownerPartPath && ownerPartPath !== sheetPartPath) {
				issues.push({
					rule: 'threaded-comment-integrity',
					severity: 'error',
					message: `Threaded comments part "${partPath}" is modeled on sheet "${sheetName}" but owned by worksheet part "${ownerPartPath}"`,
					refs: [sheetName, sheetPartPath, ownerPartPath, partPath],
					suggestedFix:
						'Restore the threadedComments relationship on the worksheet that owns these threaded comments before writing.',
					details: {
						kind: 'threaded-comment-sheet-owner-mismatch',
						partPath,
						sheetName,
						expectedWorksheetPartPath: sheetPartPath,
						actualWorksheetPartPath: ownerPartPath,
						relationships: threadedRelationships.map(queryTableRelationshipDetails),
					},
				})
			}
		}
	}
	for (const relationship of packageGraph.relationships) {
		if (!isThreadedCommentRelationshipType(relationship.type)) continue
		if (relationship.resolvedTarget && graphPartByPath.has(relationship.resolvedTarget)) continue
		issues.push({
			rule: 'threaded-comment-integrity',
			severity: 'error',
			message: `Threaded comment relationship "${relationship.id}" resolves to missing threadedComments part "${relationship.resolvedTarget ?? relationship.rawTarget}"`,
			refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
			suggestedFix:
				'Repair the threadedComment relationship target or restore the referenced threadedComments part before writing.',
			details: {
				kind: 'threaded-comment-relationship-missing-target',
				relationship: queryTableRelationshipDetails(relationship),
			},
		})
	}
	for (const part of packageGraph.parts) {
		if (!isThreadedCommentPart(part)) continue
		if (claimedPartsBySheet.has(part.path)) continue
		issues.push({
			rule: 'threaded-comment-integrity',
			severity: 'warning',
			message: `Threaded comments package part "${part.path}" is not claimed by any sheet model`,
			refs: [part.path],
			suggestedFix:
				'Reconnect the threadedComments part to sheet.threadedComments or remove the orphan sidecar before writing.',
			details: {
				kind: 'orphan-threaded-comment-part',
				partPath: part.path,
				ownerScope: part.ownerScope,
				contentType: part.contentType,
				incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
					queryTableRelationshipDetails,
				),
			},
		})
	}
	return issues
}

function checkThreadedCommentParentCycles(
	wb: Workbook,
	idsByPart: ReadonlyMap<string, ReadonlyMap<string, ThreadedCommentIntegrityEntry>>,
): CheckIssue[] {
	const issues: CheckIssue[] = []
	const reportedCycles = new Set<string>()
	for (const sheet of wb.sheets) {
		for (let index = 0; index < sheet.threadedComments.length; index++) {
			const comment = sheet.threadedComments[index]
			if (!comment?.id || !comment.parentId || comment.parentId === comment.id) continue
			const partPath = comment.partPath ?? '(unknown threaded comment part)'
			const partIds = idsByPart.get(partPath)
			if (!partIds) continue
			const cycle = threadedCommentParentCycle(partIds, comment.id, comment.parentId)
			if (!cycle) continue
			const cycleKey = `${partPath}#${[...cycle.ids].sort().join('\u0000')}`
			if (reportedCycles.has(cycleKey)) continue
			reportedCycles.add(cycleKey)
			issues.push({
				rule: 'threaded-comment-integrity',
				severity: 'warning',
				message: `Threaded comments in "${partPath}" contain a parentId cycle: ${cycle.ids.join(' -> ')}`,
				refs: cycle.entries.map((entry) => `${entry.sheetName}!${entry.ref}`),
				suggestedFix:
					'Repair threadedComment parentId chains before editing replies; cyclic threads cannot be represented as a valid root/reply tree.',
				details: {
					kind: 'threaded-comment-parent-cycle',
					partPath,
					cycleIds: cycle.ids,
					commentIndexes: cycle.entries.map((entry) => entry.index),
				},
			})
		}
	}
	return issues
}

function threadedCommentParentCycle(
	partIds: ReadonlyMap<string, ThreadedCommentIntegrityEntry>,
	startId: string,
	parentId: string,
): {
	readonly ids: readonly string[]
	readonly entries: readonly ThreadedCommentIntegrityEntry[]
} | null {
	const seen = new Map<string, number>([[startId, 0]])
	const chain = [startId]
	let currentParentId: string | undefined = parentId
	while (currentParentId) {
		const cycleStart = seen.get(currentParentId)
		if (cycleStart !== undefined) {
			const ids = chain.slice(cycleStart)
			const entries = ids
				.map((id) => partIds.get(id))
				.filter((entry): entry is ThreadedCommentIntegrityEntry => entry !== undefined)
			return { ids, entries }
		}
		const parent = partIds.get(currentParentId)
		if (!parent) return null
		seen.set(currentParentId, chain.length)
		chain.push(currentParentId)
		currentParentId = parent.parentId
	}
	return null
}

function styleVisibility(style: string | undefined): boolean | undefined {
	if (!style) return undefined
	const match = /(?:^|;)\s*visibility\s*:\s*(visible|hidden)\s*(?:;|$)/i.exec(style)
	if (!match) return undefined
	return match[1]?.toLowerCase() === 'visible'
}

function checkLegacyCommentDrawingIntegrity(
	wb: Workbook,
	packageGraph?: VerifyPackageGraph,
): CheckIssue[] {
	const issues: CheckIssue[] = []
	const graphRelationshipsByTarget = new Map<string, VerifyPackageGraphRelationship[]>()
	for (const relationship of packageGraph?.relationships ?? []) {
		if (!relationship.resolvedTarget) continue
		const relationships = graphRelationshipsByTarget.get(relationship.resolvedTarget)
		if (relationships) relationships.push(relationship)
		else graphRelationshipsByTarget.set(relationship.resolvedTarget, [relationship])
	}
	let workbookCommentCount = 0
	let workbookLegacyDrawingCount = 0
	for (const sheet of wb.sheets) {
		const shapeIds = new Map<string, string>()
		for (const [ref, comment] of sheet.comments) {
			workbookCommentCount++
			const sheetRef = `${sheet.name}!${ref}`
			const cellRef = parseWorksheetCellRef(ref)
			if (!cellRef) {
				issues.push({
					rule: 'legacy-comment-drawing-integrity',
					severity: 'warning',
					message: `Legacy comment has invalid cell reference "${ref}"`,
					refs: [sheetRef],
					suggestedFix: 'Repair the comment reference before preserving or editing its VML layout.',
					details: {
						kind: 'legacy-comment-invalid-ref',
						ref,
						...(comment.legacyDrawing?.shapeId ? { shapeId: comment.legacyDrawing.shapeId } : {}),
					},
				})
			}
			const drawing = comment.legacyDrawing
			if (!drawing) continue
			workbookLegacyDrawingCount++
			if (cellRef) {
				if (
					drawing.row === undefined ||
					drawing.column === undefined ||
					drawing.row !== cellRef.row ||
					drawing.column !== cellRef.col
				) {
					issues.push({
						rule: 'legacy-comment-drawing-integrity',
						severity: 'warning',
						message: `Legacy comment VML target for ${sheetRef} points to row ${drawing.row ?? '(missing)'}, column ${drawing.column ?? '(missing)'}`,
						refs: [sheetRef],
						suggestedFix:
							'Repair the VML ClientData Row and Column before relying on comment layout edits.',
						details: {
							kind: 'legacy-comment-vml-target-drift',
							ref,
							expectedRow: cellRef.row,
							expectedColumn: cellRef.col,
							...(drawing.row !== undefined ? { actualRow: drawing.row } : {}),
							...(drawing.column !== undefined ? { actualColumn: drawing.column } : {}),
							...(drawing.shapeId ? { shapeId: drawing.shapeId } : {}),
						},
					})
				}
			}
			if (drawing.shapeId) {
				if (!/^_x0000_s\d+$/i.test(drawing.shapeId)) {
					issues.push({
						rule: 'legacy-comment-drawing-integrity',
						severity: 'warning',
						message: `Legacy comment VML shape id "${drawing.shapeId}" for ${sheetRef} has an unexpected format`,
						refs: [sheetRef],
						suggestedFix:
							'Repair the VML shape id before editing comment drawing layout; note shapes should use Excel VML shape ids such as "_x0000_s1025".',
						details: {
							kind: 'legacy-comment-vml-shape-id-format',
							ref,
							shapeId: drawing.shapeId,
							expectedPattern: '_x0000_s<number>',
						},
					})
				}
				const existingRef = shapeIds.get(drawing.shapeId)
				if (existingRef) {
					issues.push({
						rule: 'legacy-comment-drawing-integrity',
						severity: 'warning',
						message: `Duplicate legacy comment VML shape id "${drawing.shapeId}" on sheet "${sheet.name}"`,
						refs: [`${sheet.name}!${existingRef}`, sheetRef],
						suggestedFix: 'Assign distinct VML shape ids before editing comment drawing layout.',
						details: {
							kind: 'duplicate-legacy-comment-vml-shape-id',
							shapeId: drawing.shapeId,
							firstRef: existingRef,
							duplicateRef: ref,
						},
					})
				} else {
					shapeIds.set(drawing.shapeId, ref)
				}
			} else {
				issues.push({
					rule: 'legacy-comment-drawing-integrity',
					severity: 'warning',
					message: `Legacy comment VML drawing for ${sheetRef} is missing a shape id`,
					refs: [sheetRef],
					suggestedFix:
						'Restore the VML shape id before editing comment drawing layout; missing ids make shape ownership ambiguous.',
					details: {
						kind: 'legacy-comment-vml-missing-shape-id',
						ref,
					},
				})
			}
			const anchor = drawing.anchor as readonly number[] | undefined
			if (
				anchor &&
				(anchor.length !== 8 || anchor.some((value) => !Number.isInteger(value) || value < 0))
			) {
				issues.push({
					rule: 'legacy-comment-drawing-integrity',
					severity: 'warning',
					message: `Legacy comment VML anchor for ${sheetRef} is not eight non-negative integers`,
					refs: [sheetRef],
					suggestedFix: 'Repair the VML Anchor tuple before relying on comment drawing placement.',
					details: {
						kind: 'legacy-comment-vml-anchor-invalid',
						ref,
						anchor: drawing.anchor,
						...(drawing.shapeId ? { shapeId: drawing.shapeId } : {}),
					},
				})
			}
			if (isValidLegacyCommentAnchor(anchor)) {
				const [
					fromCol,
					fromColOffset,
					fromRow,
					fromRowOffset,
					toCol,
					toColOffset,
					toRow,
					toRowOffset,
				] = anchor
				const reversed =
					toCol < fromCol ||
					toRow < fromRow ||
					(toCol === fromCol && toColOffset < fromColOffset) ||
					(toRow === fromRow && toRowOffset < fromRowOffset)
				if (reversed) {
					issues.push({
						rule: 'legacy-comment-drawing-integrity',
						severity: 'warning',
						message: `Legacy comment VML anchor for ${sheetRef} ends before it starts`,
						refs: [sheetRef],
						suggestedFix:
							'Repair the VML Anchor tuple so the end row/column is after the start row/column before editing comment layout.',
						details: {
							kind: 'legacy-comment-vml-anchor-reversed',
							ref,
							anchor,
							...(drawing.shapeId ? { shapeId: drawing.shapeId } : {}),
						},
					})
				}
			}
			const styleVisible = styleVisibility(drawing.style)
			if (
				styleVisible !== undefined &&
				drawing.visible !== undefined &&
				styleVisible !== drawing.visible
			) {
				issues.push({
					rule: 'legacy-comment-drawing-integrity',
					severity: 'warning',
					message: `Legacy comment VML visibility metadata for ${sheetRef} conflicts between style and ClientData`,
					refs: [sheetRef],
					suggestedFix:
						'Make the VML style visibility and ClientData Visible flag agree before editing comment display state.',
					details: {
						kind: 'legacy-comment-vml-visibility-conflict',
						ref,
						styleVisible,
						clientDataVisible: drawing.visible,
						...(drawing.shapeId ? { shapeId: drawing.shapeId } : {}),
					},
				})
			}
		}
	}
	if (!packageGraph) return issues

	const commentsParts = packageGraph.parts.filter(
		(part) => part.featureFamily === 'preservedComments',
	)
	const vmlParts = packageGraph.parts.filter((part) => part.featureFamily === 'preservedVml')
	const commentsRelationships = packageGraph.relationships.filter((relationship) =>
		isCommentsRelationshipType(relationship.type),
	)
	const vmlRelationships = packageGraph.relationships.filter((relationship) =>
		isVmlDrawingRelationshipType(relationship.type),
	)
	const vmlRelationshipsByTarget = relationshipsByTarget(vmlRelationships)
	const rawVmlNoteObjects: Array<{
		readonly sheetName: string
		readonly objectIndex: number
		readonly drawingPartPath: string
		readonly shapeId?: string
	}> = []
	for (const sheet of wb.sheets) {
		for (let objectIndex = 0; objectIndex < sheet.drawingObjectRefs.length; objectIndex++) {
			const object = sheet.drawingObjectRefs[objectIndex]
			if (!object || object.source !== 'vml' || object.vmlObjectType !== 'Note') continue
			rawVmlNoteObjects.push({
				sheetName: sheet.name,
				objectIndex,
				drawingPartPath: object.drawingPartPath,
				...(object.vmlShapeId ? { shapeId: object.vmlShapeId } : {}),
			})
		}
	}
	const commentSheetPartPaths = new Set(
		wb.sheets
			.filter((sheet) => sheet.comments.size > 0 && sheet.preservedXml?.partPath)
			.map((sheet) => sheet.preservedXml?.partPath as string),
	)
	const claimedCommentsPartPaths = new Set(
		commentsRelationships
			.filter(
				(relationship) =>
					relationship.resolvedTarget && commentSheetPartPaths.has(relationship.sourcePartPath),
			)
			.map((relationship) => relationship.resolvedTarget as string),
	)
	const shouldClassifyCommentsPartOrphans =
		workbookCommentCount === 0 || commentSheetPartPaths.size > 0
	if (rawVmlNoteObjects.length > 0 && commentsRelationships.length === 0) {
		issues.push({
			rule: 'legacy-comment-drawing-integrity',
			severity: 'warning',
			message:
				'Legacy comment VML note shapes exist but the package graph has no comments relationship',
			refs: [...new Set(rawVmlNoteObjects.map((object) => object.drawingPartPath))],
			suggestedFix:
				'Restore the worksheet comments relationship and comments part or remove the orphan VML note shapes before writing.',
			details: {
				kind: 'legacy-comment-vml-without-comments-part',
				noteShapeCount: rawVmlNoteObjects.length,
				noteShapes: rawVmlNoteObjects,
			},
		})
	}
	const rawNoteShapeOwners = new Map<string, (typeof rawVmlNoteObjects)[number]>()
	for (const object of rawVmlNoteObjects) {
		if (!object.shapeId) continue
		const key = `${object.drawingPartPath}#${object.shapeId}`
		const first = rawNoteShapeOwners.get(key)
		if (!first) {
			rawNoteShapeOwners.set(key, object)
			continue
		}
		issues.push({
			rule: 'legacy-comment-drawing-integrity',
			severity: 'warning',
			message: `Duplicate raw legacy comment VML shape id "${object.shapeId}" in "${object.drawingPartPath}"`,
			refs: [
				`${first.sheetName}#drawingObject${first.objectIndex}`,
				`${object.sheetName}#drawingObject${object.objectIndex}`,
				object.drawingPartPath,
			],
			suggestedFix:
				'Assign distinct VML note shape ids before editing or regenerating legacy comment drawings.',
			details: {
				kind: 'duplicate-raw-legacy-comment-vml-shape-id',
				drawingPartPath: object.drawingPartPath,
				shapeId: object.shapeId,
				first,
				duplicate: object,
			},
		})
	}
	if (shouldClassifyCommentsPartOrphans) {
		for (const part of commentsParts) {
			if (workbookCommentCount > 0 && claimedCommentsPartPaths.has(part.path)) continue
			issues.push({
				rule: 'legacy-comment-drawing-integrity',
				severity: 'warning',
				message: `Classic comments package part "${part.path}" is not claimed by any sheet comments model`,
				refs: [part.path],
				suggestedFix:
					'Reconnect the comments part to sheet.comments or remove the orphan comments sidecar before writing.',
				details: {
					kind: 'orphan-legacy-comments-part',
					partPath: part.path,
					ownerScope: part.ownerScope,
					contentType: part.contentType,
					incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
						queryTableRelationshipDetails,
					),
				},
			})
		}
	}
	if (workbookCommentCount > 0 && commentsRelationships.length === 0) {
		issues.push({
			rule: 'legacy-comment-drawing-integrity',
			severity: 'warning',
			message: 'Legacy comments exist but the package graph has no comments relationship',
			refs: commentsParts.map((part) => part.path),
			suggestedFix: 'Restore the worksheet comments relationship before writing classic comments.',
			details: {
				kind: 'missing-legacy-comments-relationship',
				commentCount: workbookCommentCount,
				legacyDrawingCount: workbookLegacyDrawingCount,
			},
		})
	}
	if (workbookLegacyDrawingCount > 0 && vmlRelationships.length === 0) {
		issues.push({
			rule: 'legacy-comment-drawing-integrity',
			severity: 'warning',
			message:
				'Legacy comment drawings exist but the package graph has no VML drawing relationship',
			refs: vmlParts.map((part) => part.path),
			suggestedFix:
				'Restore the worksheet legacyDrawing relationship and VML sidecar before writing preserved comment layout.',
			details: {
				kind: 'missing-legacy-comment-vml-relationship',
				commentCount: workbookCommentCount,
				legacyDrawingCount: workbookLegacyDrawingCount,
			},
		})
	}
	for (const relationship of commentsRelationships) {
		if (
			relationship.resolvedTarget &&
			commentsParts.some((part) => part.path === relationship.resolvedTarget)
		) {
			continue
		}
		issues.push({
			rule: 'legacy-comment-drawing-integrity',
			severity: 'error',
			message: `Comments relationship "${relationship.id}" resolves to missing comments part "${relationship.resolvedTarget ?? relationship.rawTarget}"`,
			refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
			suggestedFix:
				'Repair the comments relationship target or restore the referenced comments part before writing.',
			details: {
				kind: 'legacy-comments-relationship-missing-target',
				relationship: queryTableRelationshipDetails(relationship),
			},
		})
	}
	for (const relationship of vmlRelationships) {
		if (
			relationship.resolvedTarget &&
			vmlParts.some((part) => part.path === relationship.resolvedTarget)
		) {
			continue
		}
		issues.push({
			rule: 'legacy-comment-drawing-integrity',
			severity: 'error',
			message: `VML drawing relationship "${relationship.id}" resolves to missing VML part "${relationship.resolvedTarget ?? relationship.rawTarget}"`,
			refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
			suggestedFix:
				'Repair the legacyDrawing relationship target or restore the referenced VML part before writing.',
			details: {
				kind: 'legacy-comment-vml-relationship-missing-target',
				relationship: queryTableRelationshipDetails(relationship),
			},
		})
	}
	for (const part of vmlParts) {
		if (vmlRelationshipsByTarget.has(part.path)) continue
		const hasRawNoteObjects = rawVmlNoteObjects.some(
			(object) => object.drawingPartPath === part.path,
		)
		if (!hasRawNoteObjects) continue
		issues.push({
			rule: 'legacy-comment-drawing-integrity',
			severity: 'warning',
			message: `Legacy comment VML part "${part.path}" has note shapes but no worksheet VML relationship`,
			refs: [part.path],
			suggestedFix:
				'Restore the worksheet legacyDrawing relationship before writing preserved comment VML layout.',
			details: {
				kind: 'orphan-legacy-comment-vml-part',
				partPath: part.path,
				ownerScope: part.ownerScope,
				contentType: part.contentType,
				noteShapes: rawVmlNoteObjects.filter((object) => object.drawingPartPath === part.path),
			},
		})
	}
	return issues
}

function isValidLegacyCommentAnchor(
	anchor: readonly number[] | undefined,
): anchor is readonly [number, number, number, number, number, number, number, number] {
	return anchor?.length === 8 && anchor.every((value) => Number.isInteger(value) && value >= 0)
}

function isExternalLinkRelationshipType(relationshipType: string | undefined): boolean {
	return relationshipType?.toLowerCase().endsWith('/relationships/externallink') ?? false
}

function isExternalLinkPathRelationshipType(relationshipType: string | undefined): boolean {
	const normalized = relationshipType?.toLowerCase()
	if (!normalized) return false
	return (
		normalized.includes('/relationships/externallinkpath') ||
		normalized.includes('/relationships/xlexternallinkpath/')
	)
}

function relationshipKey(
	sourcePartPath: string | undefined,
	id: string | undefined,
): string | null {
	if (!sourcePartPath || !id) return null
	return `${sourcePartPath}#${id}`
}

function pushExternalLinkPathGraphCongruenceIssues(
	issues: CheckIssue[],
	detail: Workbook['externalReferenceDetails'][number],
	pathRelationship: VerifyPackageGraphRelationship,
): void {
	const relRef = `${pathRelationship.relationshipPartPath}#${pathRelationship.id}`
	if (!isExternalLinkPathRelationshipType(pathRelationship.type)) {
		issues.push({
			rule: 'external-link-integrity',
			severity: 'warning',
			message: `ExternalLink path relationship "${pathRelationship.id}" has package graph type "${pathRelationship.type}" that is not an externalLinkPath relationship`,
			refs: [relRef],
			suggestedFix:
				'Repair the externalLinkPath relationship type before preserving this external workbook target.',
			details: {
				kind: 'external-link-package-path-relationship-type-mismatch',
				partPath: detail.partPath,
				linkRelId: pathRelationship.id,
				expectedType: detail.linkRelationshipType,
				actualType: pathRelationship.type,
				actualRawType: pathRelationship.rawType,
			},
		})
	}

	if (
		detail.linkRelationshipRawTarget !== undefined &&
		pathRelationship.rawTarget !== detail.linkRelationshipRawTarget
	) {
		issues.push({
			rule: 'external-link-integrity',
			severity: 'error',
			message: `ExternalLink path relationship "${pathRelationship.id}" package target "${pathRelationship.rawTarget}" does not match metadata target "${detail.linkRelationshipRawTarget}"`,
			refs: [relRef],
			suggestedFix:
				'Repair the externalLinkPath relationship target or refresh externalLink metadata before rewriting external links.',
			details: {
				kind: 'external-link-package-path-target-mismatch',
				partPath: detail.partPath,
				linkRelId: pathRelationship.id,
				expectedRawTarget: detail.linkRelationshipRawTarget,
				actualRawTarget: pathRelationship.rawTarget,
				metadataTarget: detail.target,
			},
		})
	}

	if (
		detail.targetMode !== undefined &&
		pathRelationship.targetMode !== undefined &&
		pathRelationship.targetMode.toLowerCase() !== detail.targetMode.toLowerCase()
	) {
		issues.push({
			rule: 'external-link-integrity',
			severity: 'warning',
			message: `ExternalLink path relationship "${pathRelationship.id}" package TargetMode "${pathRelationship.targetMode}" does not match metadata TargetMode "${detail.targetMode}"`,
			refs: [relRef],
			suggestedFix:
				'Repair the externalLinkPath relationship TargetMode before preserving external link metadata.',
			details: {
				kind: 'external-link-package-path-target-mode-mismatch',
				partPath: detail.partPath,
				linkRelId: pathRelationship.id,
				expectedTargetMode: detail.targetMode,
				actualTargetMode: pathRelationship.targetMode,
				target: detail.target,
			},
		})
	}
}

function externalBookRelationshipRef(
	detail: Workbook['externalReferenceDetails'][number],
	relationshipId: string,
): string {
	return detail.linkRelationshipPart
		? `${detail.linkRelationshipPart}#${relationshipId}`
		: `${detail.partPath}#${relationshipId}`
}

function pushExternalBookRelationshipGraphIssues(
	issues: CheckIssue[],
	detail: Workbook['externalReferenceDetails'][number],
	externalBookRelationship: VerifyPackageGraphRelationship,
): void {
	const relRef = externalBookRelationshipRef(detail, externalBookRelationship.id)
	if (!isExternalLinkPathRelationshipType(externalBookRelationship.type)) {
		issues.push({
			rule: 'external-link-integrity',
			severity: 'warning',
			message: `ExternalBook r:id "${externalBookRelationship.id}" on "${detail.partPath}" has package graph type "${externalBookRelationship.type}" that is not an externalLinkPath relationship`,
			refs: [relRef],
			suggestedFix:
				'Repair the externalBook r:id relationship type before preserving or rewriting external link metadata.',
			details: {
				kind: 'external-book-relationship-type-mismatch',
				partPath: detail.partPath,
				externalBookRelId: externalBookRelationship.id,
				linkRelId: detail.linkRelId,
				expectedType: detail.linkRelationshipType,
				actualType: externalBookRelationship.type,
				actualRawType: externalBookRelationship.rawType,
			},
		})
	}

	if (
		detail.linkRelationshipRawTarget !== undefined &&
		externalBookRelationship.rawTarget !== detail.linkRelationshipRawTarget
	) {
		issues.push({
			rule: 'external-link-integrity',
			severity: 'error',
			message: `ExternalBook r:id "${externalBookRelationship.id}" on "${detail.partPath}" targets "${externalBookRelationship.rawTarget}" but metadata selected "${detail.linkRelationshipRawTarget}"`,
			refs: [relRef],
			suggestedFix:
				'Repair the externalBook r:id target or refresh external link metadata before writing.',
			details: {
				kind: 'external-book-relationship-target-mismatch',
				partPath: detail.partPath,
				externalBookRelId: externalBookRelationship.id,
				linkRelId: detail.linkRelId,
				expectedRawTarget: detail.linkRelationshipRawTarget,
				actualRawTarget: externalBookRelationship.rawTarget,
				metadataTarget: detail.target,
			},
		})
	}

	if (
		detail.targetMode !== undefined &&
		externalBookRelationship.targetMode !== undefined &&
		externalBookRelationship.targetMode.toLowerCase() !== detail.targetMode.toLowerCase()
	) {
		issues.push({
			rule: 'external-link-integrity',
			severity: 'warning',
			message: `ExternalBook r:id "${externalBookRelationship.id}" on "${detail.partPath}" has package TargetMode "${externalBookRelationship.targetMode}" but metadata selected "${detail.targetMode}"`,
			refs: [relRef],
			suggestedFix:
				'Repair the externalBook r:id TargetMode before preserving external link metadata.',
			details: {
				kind: 'external-book-relationship-target-mode-mismatch',
				partPath: detail.partPath,
				externalBookRelId: externalBookRelationship.id,
				linkRelId: detail.linkRelId,
				expectedTargetMode: detail.targetMode,
				actualTargetMode: externalBookRelationship.targetMode,
				target: detail.target,
			},
		})
	}
}

function checkExternalLinkIntegrity(wb: Workbook, packageGraph?: VerifyPackageGraph): CheckIssue[] {
	const issues: CheckIssue[] = []
	const externalReferenceParts = new Set(wb.externalReferences)
	const detailsByPart = new Map<string, Workbook['externalReferenceDetails'][number]>()
	const claimedPathRelationshipKeys = new Set<string>()

	for (const partPath of wb.externalReferences) {
		const details = wb.externalReferenceDetails.filter((entry) => entry.partPath === partPath)
		if (details.length === 0) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'warning',
				message: `External reference "${partPath}" has no externalLink metadata detail`,
				refs: [partPath],
				suggestedFix:
					'Re-read or repair the workbook externalReferences metadata before preserving external links.',
				details: { kind: 'external-reference-missing-detail', partPath },
			})
		}
		if (details.length > 1) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'warning',
				message: `External reference "${partPath}" has duplicate externalLink metadata entries`,
				refs: [partPath],
				suggestedFix:
					'Collapse duplicate externalLink metadata entries before writing workbook relationships.',
				details: { kind: 'duplicate-external-reference-detail', partPath, count: details.length },
			})
		}
	}

	for (const detail of wb.externalReferenceDetails) {
		detailsByPart.set(detail.partPath, detail)
		const pathRelationshipKey = relationshipKey(detail.partPath, detail.linkRelId)
		if (pathRelationshipKey) claimedPathRelationshipKeys.add(pathRelationshipKey)

		if (!externalReferenceParts.has(detail.partPath)) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'warning',
				message: `ExternalLink metadata for "${detail.partPath}" is not listed by workbook externalReferences`,
				refs: [detail.partPath],
				suggestedFix:
					'Reconnect the workbook externalReference entry or remove the orphan externalLink metadata before writing.',
				details: {
					kind: 'orphan-external-link-metadata',
					partPath: detail.partPath,
					relId: detail.relId,
				},
			})
		}

		if (
			detail.sourceRelationshipType !== undefined &&
			!isExternalLinkRelationshipType(detail.sourceRelationshipType)
		) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'warning',
				message: `External reference "${detail.partPath}" uses unexpected workbook relationship type "${detail.sourceRelationshipType}"`,
				refs: [
					detail.sourceRelationshipPart && detail.relId
						? `${detail.sourceRelationshipPart}#${detail.relId}`
						: detail.partPath,
				],
				suggestedFix: 'Repair the workbook relationship type before preserving externalLink parts.',
				details: {
					kind: 'external-link-source-relationship-type-mismatch',
					partPath: detail.partPath,
					relationshipId: detail.relId,
					relationshipType: detail.sourceRelationshipType,
					relationshipRawType: detail.sourceRelationshipRawType,
				},
			})
		}

		if (detail.linkBindingStatus === 'fallbackPathRelationship' && detail.externalBookRelId) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'warning',
				message: `ExternalLink part "${detail.partPath}" has fallbackPathRelationship binding because externalBook r:id "${detail.externalBookRelId}" does not bind to the selected path relationship "${detail.linkRelId ?? '(missing)'}"`,
				refs: [
					detail.partPath,
					detail.linkRelationshipPart && detail.linkRelId
						? `${detail.linkRelationshipPart}#${detail.linkRelId}`
						: detail.partPath,
				],
				suggestedFix:
					'Repair the externalBook r:id or externalLinkPath relationship before relying on this external workbook target.',
				details: {
					kind: 'external-link-binding-risk',
					bindingKind: 'external-book-relationship-binding-mismatch',
					partPath: detail.partPath,
					externalBookRelId: detail.externalBookRelId,
					linkRelId: detail.linkRelId,
					linkBindingStatus: detail.linkBindingStatus,
					linkRelationshipType: detail.linkRelationshipType,
					linkRelationshipRawTarget: detail.linkRelationshipRawTarget,
					target: detail.target,
				},
			})
		}

		if (detail.linkBindingStatus === 'missingPathRelationship') {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'error',
				message: `ExternalLink part "${detail.partPath}" has no external workbook path relationship`,
				refs: [detail.partPath],
				suggestedFix:
					'Restore an externalLinkPath relationship for the externalBook r:id before writing external link metadata.',
				details: {
					kind: 'missing-external-link-path-relationship',
					partPath: detail.partPath,
					externalBookRelId: detail.externalBookRelId,
					linkBindingStatus: detail.linkBindingStatus,
				},
			})
		}

		if (
			detail.linkRelationshipType &&
			!isExternalLinkPathRelationshipType(detail.linkRelationshipType)
		) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'warning',
				message: `ExternalLink part "${detail.partPath}" path relationship "${detail.linkRelId}" has unexpected type "${detail.linkRelationshipType}"`,
				refs: [
					detail.linkRelationshipPart && detail.linkRelId
						? `${detail.linkRelationshipPart}#${detail.linkRelId}`
						: detail.partPath,
				],
				suggestedFix:
					'Inspect the externalLinkPath relationship dialect before preserving this external link.',
				details: {
					kind: 'external-link-path-relationship-type-mismatch',
					partPath: detail.partPath,
					linkRelId: detail.linkRelId,
					linkRelationshipType: detail.linkRelationshipType,
					linkRelationshipRawType: detail.linkRelationshipRawType,
				},
			})
		}

		if (
			detail.linkRelId &&
			detail.targetMode !== undefined &&
			detail.targetMode.toLowerCase() !== 'external'
		) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'warning',
				message: `ExternalLink part "${detail.partPath}" path relationship "${detail.linkRelId}" is not marked TargetMode="External"`,
				refs: [
					detail.linkRelationshipPart
						? `${detail.linkRelationshipPart}#${detail.linkRelId}`
						: detail.partPath,
				],
				suggestedFix:
					'Set the externalLinkPath relationship target mode to External before writing.',
				details: {
					kind: 'external-link-path-target-mode-mismatch',
					partPath: detail.partPath,
					linkRelId: detail.linkRelId,
					targetMode: detail.targetMode,
					target: detail.target,
				},
			})
		}
	}

	if (!packageGraph) return issues

	const graphPartByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
	const graphRelationshipBySourceAndId = new Map(
		packageGraph.relationships.map((relationship) => [
			`${relationship.sourcePartPath}#${relationship.id}`,
			relationship,
		]),
	)
	const graphRelationshipsByTarget = new Map<string, VerifyPackageGraphRelationship[]>()
	for (const relationship of packageGraph.relationships) {
		if (!relationship.resolvedTarget) continue
		const relationships = graphRelationshipsByTarget.get(relationship.resolvedTarget)
		if (relationships) relationships.push(relationship)
		else graphRelationshipsByTarget.set(relationship.resolvedTarget, [relationship])
	}

	for (const detail of wb.externalReferenceDetails) {
		const graphPart = graphPartByPath.get(detail.partPath)
		if (!graphPart) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'error',
				message: `External reference metadata points at missing externalLink package part "${detail.partPath}"`,
				refs: [detail.partPath],
				suggestedFix:
					'Restore the externalLink package part or remove the stale workbook externalReference before writing.',
				details: {
					kind: 'missing-external-link-part',
					partPath: detail.partPath,
					relId: detail.relId,
				},
			})
			continue
		}

		const sourceRelationshipKey = relationshipKey(detail.sourcePartPath, detail.relId)
		const sourceRelationship = sourceRelationshipKey
			? graphRelationshipBySourceAndId.get(sourceRelationshipKey)
			: undefined
		if (
			sourceRelationshipKey &&
			(!sourceRelationship || sourceRelationship.resolvedTarget !== detail.partPath)
		) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'error',
				message: `Workbook externalReference relationship "${detail.relId ?? '(missing)'}" does not bind to "${detail.partPath}" in the package graph`,
				refs: [
					detail.sourceRelationshipPart && detail.relId
						? `${detail.sourceRelationshipPart}#${detail.relId}`
						: detail.partPath,
				],
				suggestedFix:
					'Repair the workbook externalReference relationship id/target binding before writing external links.',
				details: {
					kind: 'external-link-source-relationship-binding-mismatch',
					partPath: detail.partPath,
					graphPart,
					relationshipId: detail.relId,
					incomingRelationships: (graphRelationshipsByTarget.get(detail.partPath) ?? []).map(
						queryTableRelationshipDetails,
					),
				},
			})
		}

		const expectedPathRelId = detail.linkRelId ?? detail.externalBookRelId
		if (expectedPathRelId) {
			const pathRelationship = graphRelationshipBySourceAndId.get(
				`${detail.partPath}#${expectedPathRelId}`,
			)
			if (!pathRelationship) {
				issues.push({
					rule: 'external-link-integrity',
					severity: 'error',
					message: `ExternalLink path relationship "${expectedPathRelId}" is missing from "${detail.partPath}" in the package graph`,
					refs: [
						detail.linkRelationshipPart && detail.linkRelId
							? `${detail.linkRelationshipPart}#${detail.linkRelId}`
							: `${detail.partPath}#${expectedPathRelId}`,
					],
					suggestedFix:
						'Restore the externalLinkPath relationship or clear the stale externalLink metadata before writing.',
					details: {
						kind:
							expectedPathRelId === detail.externalBookRelId && !detail.linkRelId
								? 'external-book-relationship-missing'
								: 'external-link-path-relationship-binding-mismatch',
						partPath: detail.partPath,
						linkRelId: expectedPathRelId,
						linkRelationshipPart: detail.linkRelationshipPart,
					},
				})
			} else {
				pushExternalLinkPathGraphCongruenceIssues(issues, detail, pathRelationship)
			}
		}

		if (
			detail.externalBookRelId &&
			detail.linkRelId &&
			detail.externalBookRelId !== detail.linkRelId
		) {
			const externalBookRelationship = graphRelationshipBySourceAndId.get(
				`${detail.partPath}#${detail.externalBookRelId}`,
			)
			if (!externalBookRelationship) {
				issues.push({
					rule: 'external-link-integrity',
					severity: 'error',
					message: `ExternalBook r:id "${detail.externalBookRelId}" is missing from "${detail.partPath}" even though metadata fell back to path relationship "${detail.linkRelId}"`,
					refs: [externalBookRelationshipRef(detail, detail.externalBookRelId)],
					suggestedFix:
						'Repair the externalBook r:id relationship before writing external link metadata.',
					details: {
						kind: 'external-book-relationship-missing',
						partPath: detail.partPath,
						externalBookRelId: detail.externalBookRelId,
						linkRelId: detail.linkRelId,
						linkRelationshipPart: detail.linkRelationshipPart,
						linkBindingStatus: detail.linkBindingStatus,
					},
				})
			} else {
				pushExternalBookRelationshipGraphIssues(issues, detail, externalBookRelationship)
			}
		}
	}

	for (const part of packageGraph.parts) {
		if (part.featureFamily !== 'preservedExternalLink') continue
		if (!/(^|\/)externalLinks\/[^/]+\.xml$/i.test(part.path)) continue
		if (detailsByPart.has(part.path) || externalReferenceParts.has(part.path)) continue
		issues.push({
			rule: 'external-link-integrity',
			severity: 'warning',
			message: `ExternalLink package part "${part.path}" is not claimed by workbook externalReferences metadata`,
			refs: [part.path],
			suggestedFix:
				'Inspect workbook externalReferences and reconnect or intentionally remove the orphan externalLink sidecar before writing.',
			details: {
				kind: 'orphan-external-link-part',
				partPath: part.path,
				ownerScope: part.ownerScope,
				contentType: part.contentType,
				incomingRelationships: (graphRelationshipsByTarget.get(part.path) ?? []).map(
					queryTableRelationshipDetails,
				),
			},
		})
	}

	for (const relationship of packageGraph.relationships) {
		if (
			isExternalLinkRelationshipType(relationship.type) &&
			relationship.resolvedTarget &&
			!externalReferenceParts.has(relationship.resolvedTarget) &&
			!detailsByPart.has(relationship.resolvedTarget)
		) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'warning',
				message: `ExternalLink relationship "${relationship.id}" is not claimed by workbook externalReferences metadata`,
				refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
				suggestedFix:
					'Reconnect the workbook externalReference metadata or remove the orphan externalLink relationship before writing.',
				details: {
					kind: 'orphan-external-link-relationship',
					relationship: queryTableRelationshipDetails(relationship),
				},
			})
		}
		if (
			isExternalLinkPathRelationshipType(relationship.type) &&
			detailsByPart.has(relationship.sourcePartPath) &&
			!claimedPathRelationshipKeys.has(`${relationship.sourcePartPath}#${relationship.id}`)
		) {
			issues.push({
				rule: 'external-link-integrity',
				severity: 'warning',
				message: `ExternalLink path relationship "${relationship.id}" is not claimed by externalLink metadata`,
				refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
				suggestedFix:
					'Bind the path relationship from externalBook r:id or remove the orphan externalLinkPath relationship before writing.',
				details: {
					kind: 'orphan-external-link-path-relationship',
					relationship: queryTableRelationshipDetails(relationship),
				},
			})
		}
	}

	return issues
}

function checkDrawingPackageIntegrity(
	wb: Workbook,
	packageGraph?: VerifyPackageGraph,
): CheckIssue[] {
	if (!packageGraph) return []
	const issues: CheckIssue[] = []
	const graphPartByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
	const graphRelationshipBySourceAndId = new Map(
		packageGraph.relationships.map((relationship) => [
			`${relationship.sourcePartPath}#${relationship.id}`,
			relationship,
		]),
	)
	const graphRelationshipsByTarget = relationshipsByTarget(packageGraph.relationships)
	const modelDrawingPartPaths = new Set<string>()
	const drawingObjectNamesBySheet = new Map<string, Map<string, WorkbookDrawingObjectOwner>>()

	for (const sheet of wb.sheets) {
		for (const image of sheet.imageRefs) {
			modelDrawingPartPaths.add(image.drawingPartPath)
			pushDrawingAnchorIssues(issues, {
				anchor: image.anchor,
				ownerRef: `${sheet.name}#${image.relId}`,
				kindPrefix: 'image',
				details: {
					sheetName: sheet.name,
					drawingPartPath: image.drawingPartPath,
					relationshipId: image.relId,
					...(image.name ? { imageName: image.name } : {}),
				},
			})
			const drawingPart = graphPartByPath.get(image.drawingPartPath)
			if (!drawingPart) {
				issues.push({
					rule: 'drawing-integrity',
					severity: 'error',
					message: `Image "${image.relId}" on sheet "${sheet.name}" references missing drawing package part "${image.drawingPartPath}"`,
					refs: [`${sheet.name}#${image.relId}`, image.drawingPartPath],
					suggestedFix:
						'Restore the drawing package part or remove the stale sheet image reference before writing.',
					details: {
						kind: 'image-missing-drawing-part',
						sheetName: sheet.name,
						drawingPartPath: image.drawingPartPath,
						relationshipId: image.relId,
						targetPath: image.targetPath,
					},
				})
				continue
			}
			if (!isDrawingPackagePart(drawingPart)) {
				issues.push({
					rule: 'drawing-integrity',
					severity: 'error',
					message: `Image "${image.relId}" on sheet "${sheet.name}" is attached to non-DrawingML part "${image.drawingPartPath}"`,
					refs: [`${sheet.name}#${image.relId}`, image.drawingPartPath],
					suggestedFix:
						'Bind sheet image references to a DrawingML worksheet drawing part before writing.',
					details: {
						kind: 'image-drawing-part-type-mismatch',
						sheetName: sheet.name,
						drawingPartPath: image.drawingPartPath,
						relationshipId: image.relId,
						graphPart: drawingPart,
					},
				})
			}
			const imageRelationship = graphRelationshipBySourceAndId.get(
				`${image.drawingPartPath}#${image.relId}`,
			)
			if (!imageRelationship) {
				issues.push({
					rule: 'drawing-integrity',
					severity: 'error',
					message: `Image relationship "${image.relId}" is missing from drawing part "${image.drawingPartPath}"`,
					refs: [`${sheet.name}#${image.relId}`, image.drawingPartPath],
					suggestedFix:
						'Restore the drawing image relationship or remove the stale image anchor before writing.',
					details: {
						kind: 'image-relationship-missing',
						sheetName: sheet.name,
						drawingPartPath: image.drawingPartPath,
						relationshipId: image.relId,
						targetPath: image.targetPath,
					},
				})
				continue
			}
			pushDrawingMediaRelationshipIssues(issues, {
				relationship: imageRelationship,
				expectedTargetPath: image.targetPath,
				refs: [
					`${sheet.name}#${image.relId}`,
					`${imageRelationship.relationshipPartPath}#${image.relId}`,
				],
				details: {
					sheetName: sheet.name,
					drawingPartPath: image.drawingPartPath,
					relationshipId: image.relId,
					...(image.name ? { imageName: image.name } : {}),
				},
				graphPartByPath,
			})
		}

		for (let index = 0; index < sheet.drawingObjectRefs.length; index++) {
			const object = sheet.drawingObjectRefs[index]
			if (!object) continue
			modelDrawingPartPaths.add(object.drawingPartPath)
			pushDrawingAnchorIssues(issues, {
				anchor: object.anchor,
				ownerRef: `${sheet.name}#drawingObject${index}`,
				kindPrefix: 'drawing-object',
				details: {
					sheetName: sheet.name,
					objectIndex: index,
					drawingPartPath: object.drawingPartPath,
					source: object.source,
					objectKind: object.kind,
				},
			})
			const ownerKey = drawingObjectOwnerKey(object)
			if (ownerKey) {
				let owners = drawingObjectNamesBySheet.get(sheet.name)
				if (!owners) {
					owners = new Map()
					drawingObjectNamesBySheet.set(sheet.name, owners)
				}
				const existing = owners.get(ownerKey)
				if (existing && existing.source !== object.source) {
					issues.push({
						rule: 'drawing-integrity',
						severity: 'warning',
						message: `Drawing object "${ownerKey}" on sheet "${sheet.name}" is claimed by both DrawingML and VML metadata`,
						refs: [
							`${sheet.name}#drawingObject${existing.index}`,
							`${sheet.name}#drawingObject${index}`,
						],
						suggestedFix:
							'Inspect DrawingML and VML ownership before editing shape or image layout metadata.',
						details: {
							kind: 'vml-drawingml-ownership-ambiguity',
							sheetName: sheet.name,
							ownerKey,
							first: existing,
							duplicate: {
								index,
								source: object.source,
								drawingPartPath: object.drawingPartPath,
								kind: object.kind,
							},
						},
					})
				} else {
					owners.set(ownerKey, {
						index,
						source: object.source,
						drawingPartPath: object.drawingPartPath,
						kind: object.kind,
					})
				}
			}
			for (const relId of object.relIds ?? []) {
				const relationship = graphRelationshipBySourceAndId.get(
					`${object.drawingPartPath}#${relId}`,
				)
				if (!relationship) {
					issues.push({
						rule: 'drawing-integrity',
						severity: 'error',
						message: `Drawing object relationship "${relId}" is missing from "${object.drawingPartPath}"`,
						refs: [`${sheet.name}#drawingObject${index}`, `${object.drawingPartPath}#${relId}`],
						suggestedFix:
							'Restore the drawing relationship before editing drawing objects or their linked content.',
						details: {
							kind: 'drawing-object-relationship-missing',
							sheetName: sheet.name,
							objectIndex: index,
							drawingPartPath: object.drawingPartPath,
							relationshipId: relId,
							source: object.source,
						},
					})
					continue
				}
				if (isImageRelationshipType(relationship.type)) {
					pushDrawingMediaRelationshipIssues(issues, {
						relationship,
						refs: [
							`${sheet.name}#drawingObject${index}`,
							`${relationship.relationshipPartPath}#${relId}`,
						],
						details: {
							sheetName: sheet.name,
							objectIndex: index,
							drawingPartPath: object.drawingPartPath,
							relationshipId: relId,
							source: object.source,
						},
						graphPartByPath,
					})
				}
				if (isChartRelationshipType(relationship.type)) {
					pushDrawingChartRelationshipIssues(issues, {
						relationship,
						refs: [
							`${sheet.name}#drawingObject${index}`,
							`${relationship.relationshipPartPath}#${relId}`,
						],
						details: {
							sheetName: sheet.name,
							objectIndex: index,
							drawingPartPath: object.drawingPartPath,
							relationshipId: relId,
							source: object.source,
						},
						graphPartByPath,
					})
				}
			}
		}
	}

	for (const relationship of packageGraph.relationships) {
		if (
			isDrawingRelationshipType(relationship.type) ||
			isVmlDrawingRelationshipType(relationship.type)
		) {
			const targetPart = relationship.resolvedTarget
				? graphPartByPath.get(relationship.resolvedTarget)
				: undefined
			if (
				isDrawingRelationshipType(relationship.type) &&
				targetPart &&
				!isDrawingPackagePart(targetPart)
			) {
				issues.push({
					rule: 'drawing-integrity',
					severity: 'error',
					message: `Drawing relationship "${relationship.id}" targets non-DrawingML part "${relationship.resolvedTarget}"`,
					refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
					suggestedFix:
						'Repair worksheet drawing relationships so DrawingML and VML sidecars use their distinct relationship types.',
					details: {
						kind: 'drawing-relationship-target-type-mismatch',
						relationship: packageRelationshipDetails(relationship),
						targetPart,
					},
				})
			}
			if (
				isVmlDrawingRelationshipType(relationship.type) &&
				targetPart &&
				!isVmlPackagePart(targetPart)
			) {
				issues.push({
					rule: 'drawing-integrity',
					severity: 'error',
					message: `VML drawing relationship "${relationship.id}" targets non-VML part "${relationship.resolvedTarget}"`,
					refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
					suggestedFix:
						'Repair worksheet VML drawing relationships so VML and DrawingML sidecars use their distinct relationship types.',
					details: {
						kind: 'vml-drawing-relationship-target-type-mismatch',
						relationship: packageRelationshipDetails(relationship),
						targetPart,
					},
				})
			}
			if (
				!isWorksheetPartPath(relationship.sourcePartPath) &&
				!isChartSheetPartPath(relationship.sourcePartPath)
			) {
				issues.push({
					rule: 'drawing-integrity',
					severity: 'warning',
					message: `Drawing relationship "${relationship.id}" is not owned by a worksheet or chartsheet`,
					refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
					suggestedFix:
						'Bind drawing sidecars from their worksheet or chartsheet owner before writing.',
					details: {
						kind: 'drawing-missing-worksheet-chartsheet-owner',
						relationship: packageRelationshipDetails(relationship),
					},
				})
			}
		}
		if (isChartRelationshipType(relationship.type)) {
			pushDrawingChartRelationshipIssues(issues, {
				relationship,
				refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
				details: { sourcePartPath: relationship.sourcePartPath },
				graphPartByPath,
			})
		}
		if (isImageRelationshipType(relationship.type)) {
			pushDrawingMediaRelationshipIssues(issues, {
				relationship,
				refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
				details: { sourcePartPath: relationship.sourcePartPath },
				graphPartByPath,
			})
		}
	}

	for (const part of packageGraph.parts) {
		if (!isDrawingPackagePart(part) && !isVmlPackagePart(part)) continue
		if (modelDrawingPartPaths.has(part.path)) continue
		const incomingRelationships = graphRelationshipsByTarget.get(part.path) ?? []
		const ownerRelationships = incomingRelationships.filter(
			(relationship) =>
				isDrawingRelationshipType(relationship.type) ||
				isVmlDrawingRelationshipType(relationship.type),
		)
		if (ownerRelationships.length === 0) {
			issues.push({
				rule: 'drawing-integrity',
				severity: 'warning',
				message: `Drawing package part "${part.path}" is not owned by a worksheet or chartsheet relationship`,
				refs: [part.path],
				suggestedFix:
					'Reconnect the drawing part to its worksheet/chartsheet or remove the orphan drawing sidecar before writing.',
				details: {
					kind: isVmlPackagePart(part) ? 'orphan-vml-drawing-part' : 'orphan-drawing-part',
					partPath: part.path,
					ownerScope: part.ownerScope,
					contentType: part.contentType,
					incomingRelationships: incomingRelationships.map(packageRelationshipDetails),
				},
			})
			continue
		}
		const hasDrawingOwner = ownerRelationships.some((relationship) =>
			isDrawingRelationshipType(relationship.type),
		)
		const hasVmlOwner = ownerRelationships.some((relationship) =>
			isVmlDrawingRelationshipType(relationship.type),
		)
		if (hasDrawingOwner && hasVmlOwner) {
			issues.push({
				rule: 'drawing-integrity',
				severity: 'warning',
				message: `Drawing package part "${part.path}" is claimed by both DrawingML and VML relationships`,
				refs: [
					part.path,
					...ownerRelationships.map((rel) => `${rel.relationshipPartPath}#${rel.id}`),
				],
				suggestedFix:
					'Separate DrawingML and VML sidecars before editing drawing or legacy shape metadata.',
				details: {
					kind: 'vml-drawingml-package-owner-ambiguity',
					partPath: part.path,
					relationships: ownerRelationships.map(packageRelationshipDetails),
				},
			})
		}
	}

	return issues
}

interface WorkbookDrawingObjectOwner {
	readonly index: number
	readonly source: 'drawingml' | 'vml' | undefined
	readonly drawingPartPath: string
	readonly kind: string
}

function pushDrawingAnchorIssues(
	issues: CheckIssue[],
	params: {
		readonly anchor: Workbook['sheets'][number]['imageRefs'][number]['anchor']
		readonly ownerRef: string
		readonly kindPrefix: 'image' | 'drawing-object'
		readonly details: Readonly<Record<string, unknown>>
	},
): void {
	if (!params.anchor) return
	if (!drawingAnchorIsWellFormed(params.anchor)) {
		issues.push({
			rule: 'drawing-integrity',
			severity: 'error',
			message: `Drawing anchor for "${params.ownerRef}" is malformed`,
			refs: [params.ownerRef],
			suggestedFix: 'Repair the drawing anchor coordinates before editing drawing layout metadata.',
			details: {
				kind: `${params.kindPrefix}-anchor-invalid`,
				anchor: params.anchor,
				...params.details,
			},
		})
		return
	}
	if (
		params.anchor.kind === 'twoCell' &&
		(params.anchor.to.row < params.anchor.from.row ||
			(params.anchor.to.row === params.anchor.from.row &&
				params.anchor.to.col < params.anchor.from.col))
	) {
		issues.push({
			rule: 'drawing-integrity',
			severity: 'warning',
			message: `Drawing anchor for "${params.ownerRef}" ends before it starts`,
			refs: [params.ownerRef],
			suggestedFix:
				'Repair the drawing anchor so the end marker is after the start marker before editing layout metadata.',
			details: {
				kind: `${params.kindPrefix}-anchor-reversed`,
				anchor: params.anchor,
				...params.details,
			},
		})
	}
}

function drawingAnchorIsWellFormed(
	anchor: Workbook['sheets'][number]['imageRefs'][number]['anchor'],
): boolean {
	if (!anchor) return true
	if (anchor.kind === 'absolute') {
		return (
			isNonNegativeFiniteNumber(anchor.x) &&
			isNonNegativeFiniteNumber(anchor.y) &&
			(anchor.cx === undefined || isNonNegativeFiniteNumber(anchor.cx)) &&
			(anchor.cy === undefined || isNonNegativeFiniteNumber(anchor.cy))
		)
	}
	if (!drawingAnchorMarkerIsWellFormed(anchor.from)) return false
	if (anchor.kind === 'oneCell') {
		return (
			(anchor.cx === undefined || isNonNegativeFiniteNumber(anchor.cx)) &&
			(anchor.cy === undefined || isNonNegativeFiniteNumber(anchor.cy))
		)
	}
	return drawingAnchorMarkerIsWellFormed(anchor.to)
}

function drawingAnchorMarkerIsWellFormed(
	marker: Workbook['sheets'][number]['imageRefs'][number]['anchor'] extends infer Anchor
		? Anchor extends { readonly from: infer Marker }
			? Marker
			: never
		: never,
): boolean {
	return (
		typeof marker === 'object' &&
		marker !== null &&
		isNonNegativeInteger((marker as { readonly row?: unknown }).row) &&
		isNonNegativeInteger((marker as { readonly col?: unknown }).col) &&
		((marker as { readonly rowOff?: unknown }).rowOff === undefined ||
			isNonNegativeFiniteNumber((marker as { readonly rowOff?: unknown }).rowOff)) &&
		((marker as { readonly colOff?: unknown }).colOff === undefined ||
			isNonNegativeFiniteNumber((marker as { readonly colOff?: unknown }).colOff))
	)
}

function isNonNegativeInteger(value: unknown): boolean {
	return typeof value === 'number' && Number.isInteger(value) && value >= 0
}

function isNonNegativeFiniteNumber(value: unknown): boolean {
	return typeof value === 'number' && Number.isFinite(value) && value >= 0
}

function pushDrawingChartRelationshipIssues(
	issues: CheckIssue[],
	params: {
		readonly relationship: VerifyPackageGraphRelationship
		readonly refs: readonly string[]
		readonly details: Readonly<Record<string, unknown>>
		readonly graphPartByPath: ReadonlyMap<string, VerifyPackageGraphPart>
	},
): void {
	if (!params.relationship.resolvedTarget) return
	const targetPart = params.graphPartByPath.get(params.relationship.resolvedTarget)
	if (!targetPart) {
		issues.push({
			rule: 'drawing-integrity',
			severity: 'error',
			message: `Drawing chart relationship "${params.relationship.id}" resolves to missing chart part "${params.relationship.resolvedTarget}"`,
			refs: params.refs,
			suggestedFix:
				'Repair the drawing chart relationship target or restore the referenced chart part before writing.',
			details: {
				kind: 'drawing-chart-relationship-missing-target',
				relationship: packageRelationshipDetails(params.relationship),
				...params.details,
			},
		})
		return
	}
	if (isChartPackagePart(targetPart)) return
	issues.push({
		rule: 'drawing-integrity',
		severity: 'error',
		message: `Drawing chart relationship "${params.relationship.id}" targets non-chart part "${params.relationship.resolvedTarget}"`,
		refs: params.refs,
		suggestedFix: 'Repair the drawing chart relationship target before editing embedded charts.',
		details: {
			kind: 'drawing-chart-target-type-mismatch',
			relationship: packageRelationshipDetails(params.relationship),
			targetPart,
			...params.details,
		},
	})
}

function pushDrawingMediaRelationshipIssues(
	issues: CheckIssue[],
	params: {
		readonly relationship: VerifyPackageGraphRelationship
		readonly refs: readonly string[]
		readonly expectedTargetPath?: string
		readonly details: Readonly<Record<string, unknown>>
		readonly graphPartByPath: ReadonlyMap<string, VerifyPackageGraphPart>
	},
): void {
	if (!isImageRelationshipType(params.relationship.type)) {
		issues.push({
			rule: 'drawing-integrity',
			severity: 'error',
			message: `Image relationship "${params.relationship.id}" has unexpected type "${params.relationship.type}"`,
			refs: params.refs,
			suggestedFix: 'Repair the image relationship type before writing drawing media references.',
			details: {
				kind: 'image-relationship-type-mismatch',
				relationship: packageRelationshipDetails(params.relationship),
				...params.details,
			},
		})
		return
	}
	if (
		params.expectedTargetPath &&
		params.relationship.resolvedTarget &&
		params.relationship.resolvedTarget !== params.expectedTargetPath
	) {
		issues.push({
			rule: 'drawing-integrity',
			severity: 'error',
			message: `Image relationship "${params.relationship.id}" resolves to "${params.relationship.resolvedTarget}" but sheet image metadata expects "${params.expectedTargetPath}"`,
			refs: params.refs,
			suggestedFix:
				'Make the sheet image targetPath match the drawing relationship target before writing.',
			details: {
				kind: 'image-media-target-mismatch',
				relationship: packageRelationshipDetails(params.relationship),
				expectedTargetPath: params.expectedTargetPath,
				actualTargetPath: params.relationship.resolvedTarget,
				...params.details,
			},
		})
	}
	if (!params.relationship.resolvedTarget) return
	const targetPart = params.graphPartByPath.get(params.relationship.resolvedTarget)
	if (!targetPart) {
		issues.push({
			rule: 'drawing-integrity',
			severity: 'error',
			message: `Image relationship "${params.relationship.id}" resolves to missing media part "${params.relationship.resolvedTarget}"`,
			refs: params.refs,
			suggestedFix:
				'Repair the image relationship target or restore the referenced media part before writing.',
			details: {
				kind: 'image-media-relationship-missing-target',
				relationship: packageRelationshipDetails(params.relationship),
				...params.details,
			},
		})
		return
	}
	if (isMediaPackagePart(targetPart)) return
	issues.push({
		rule: 'drawing-integrity',
		severity: 'error',
		message: `Image relationship "${params.relationship.id}" targets non-media part "${params.relationship.resolvedTarget}"`,
		refs: params.refs,
		suggestedFix: 'Repair the image relationship target so it points at an xl/media package part.',
		details: {
			kind: 'image-media-target-type-mismatch',
			relationship: packageRelationshipDetails(params.relationship),
			targetPart,
			...params.details,
		},
	})
}

function checkChartStyleColorIntegrity(packageGraph?: VerifyPackageGraph): CheckIssue[] {
	if (!packageGraph) return []
	const issues: CheckIssue[] = []
	const graphPartByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
	const graphRelationshipsByTarget = relationshipsByTarget(packageGraph.relationships)
	for (const relationship of packageGraph.relationships) {
		if (
			!isChartStyleRelationshipType(relationship.type) &&
			!isChartColorRelationshipType(relationship.type)
		) {
			continue
		}
		const expected = isChartStyleRelationshipType(relationship.type)
			? 'preservedChartStyle'
			: 'preservedChartColor'
		const targetPart = relationship.resolvedTarget
			? graphPartByPath.get(relationship.resolvedTarget)
			: undefined
		if (relationship.resolvedTarget && !targetPart) {
			issues.push({
				rule: 'chart-package-integrity',
				severity: 'error',
				message: `Chart style/color relationship "${relationship.id}" resolves to missing part "${relationship.resolvedTarget}"`,
				refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
				suggestedFix:
					'Repair the chart style/color relationship target or restore the referenced sidecar before preserving chart formatting.',
				details: {
					kind: 'chart-style-color-relationship-missing-target',
					expectedFeatureFamily: expected,
					relationship: packageRelationshipDetails(relationship),
				},
			})
			continue
		}
		if (targetPart && targetPart.featureFamily !== expected) {
			issues.push({
				rule: 'chart-package-integrity',
				severity: 'error',
				message: `Chart relationship "${relationship.id}" targets ${targetPart.featureFamily ?? 'unknown'} part "${relationship.resolvedTarget}" where ${expected} is expected`,
				refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
				suggestedFix:
					'Repair the chart style/color relationship target before preserving chart formatting sidecars.',
				details: {
					kind: 'chart-style-color-target-mismatch',
					expectedFeatureFamily: expected,
					relationship: packageRelationshipDetails(relationship),
					targetPart,
				},
			})
		}
		if (!isChartPartPath(relationship.sourcePartPath)) {
			issues.push({
				rule: 'chart-package-integrity',
				severity: 'warning',
				message: `Chart style/color relationship "${relationship.id}" is not owned by a chart part`,
				refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
				suggestedFix: 'Bind chart style/color sidecars from their chart part before writing.',
				details: {
					kind: 'chart-style-color-owner-mismatch',
					relationship: packageRelationshipDetails(relationship),
				},
			})
		}
	}
	for (const part of packageGraph.parts) {
		if (!isChartStylePackagePart(part) && !isChartColorPackagePart(part)) continue
		const incoming = graphRelationshipsByTarget.get(part.path) ?? []
		const expectedRelationship = isChartStylePackagePart(part)
			? isChartStyleRelationshipType
			: isChartColorRelationshipType
		if (incoming.some((relationship) => expectedRelationship(relationship.type))) continue
		issues.push({
			rule: 'chart-package-integrity',
			severity: 'warning',
			message: `Chart ${isChartStylePackagePart(part) ? 'style' : 'color'} package part "${part.path}" is not claimed by any chart relationship`,
			refs: [part.path],
			suggestedFix:
				'Reconnect the chart style/color sidecar to its chart part or remove the orphan sidecar before writing.',
			details: {
				kind: isChartStylePackagePart(part) ? 'orphan-chart-style-part' : 'orphan-chart-color-part',
				partPath: part.path,
				ownerScope: part.ownerScope,
				contentType: part.contentType,
				incomingRelationships: incoming.map(packageRelationshipDetails),
			},
		})
	}
	return issues
}

function relationshipsByTarget(
	relationships: readonly VerifyPackageGraphRelationship[],
): Map<string, VerifyPackageGraphRelationship[]> {
	const byTarget = new Map<string, VerifyPackageGraphRelationship[]>()
	for (const relationship of relationships) {
		if (!relationship.resolvedTarget) continue
		const entries = byTarget.get(relationship.resolvedTarget)
		if (entries) entries.push(relationship)
		else byTarget.set(relationship.resolvedTarget, [relationship])
	}
	return byTarget
}

function drawingObjectOwnerKey(
	object: Workbook['sheets'][number]['drawingObjectRefs'][number],
): string | null {
	if (object.name) return `name:${object.name}`
	if (object.id !== undefined) return `id:${object.id}`
	return null
}

function isDrawingRelationshipType(relationshipType: string | undefined): boolean {
	return relationshipType?.toLowerCase().endsWith('/relationships/drawing') ?? false
}

function isChartRelationshipType(relationshipType: string | undefined): boolean {
	return relationshipType?.toLowerCase().endsWith('/relationships/chart') ?? false
}

function isImageRelationshipType(relationshipType: string | undefined): boolean {
	return relationshipType?.toLowerCase().endsWith('/relationships/image') ?? false
}

function isChartStyleRelationshipType(relationshipType: string | undefined): boolean {
	return /\/relationships\/chartstyle$/i.test(relationshipType ?? '')
}

function isChartColorRelationshipType(relationshipType: string | undefined): boolean {
	return /\/relationships\/chartcolorstyle$/i.test(relationshipType ?? '')
}

function isChartSheetPartPath(partPath: string): boolean {
	return /(^|\/)chartsheets\/sheet\d+\.xml$/i.test(partPath)
}

function isChartPartPath(partPath: string): boolean {
	return /(^|\/)(charts\/chart\d+|chartEx\/chartEx\d+)\.xml$/i.test(partPath)
}

function isChartPackagePart(part: VerifyPackageGraphPart): boolean {
	return part.featureFamily === 'preservedChart' && isChartPartPath(part.path)
}

function isChartStylePackagePart(part: VerifyPackageGraphPart): boolean {
	return part.featureFamily === 'preservedChartStyle'
}

function isChartColorPackagePart(part: VerifyPackageGraphPart): boolean {
	return part.featureFamily === 'preservedChartColor'
}

function isDrawingPackagePart(part: VerifyPackageGraphPart): boolean {
	return (
		part.featureFamily === 'preservedDrawing' &&
		/(^|\/)drawings\/drawing[^/]*\.xml$/i.test(part.path)
	)
}

function isVmlPackagePart(part: VerifyPackageGraphPart): boolean {
	return part.featureFamily === 'preservedVml' || /(^|\/)drawings\/[^/]+\.vml$/i.test(part.path)
}

function isMediaPackagePart(part: VerifyPackageGraphPart): boolean {
	return part.featureFamily === 'preservedMedia' || /(^|\/)media\//i.test(part.path)
}

function checkPackageGraphIntegrity(packageGraph?: VerifyPackageGraph): CheckIssue[] {
	if (!packageGraph) return []
	const issues: CheckIssue[] = []
	const partByPath = new Map(packageGraph.parts.map((part) => [part.path, part]))
	const partPaths = new Set(partByPath.keys())
	const reportedMissingSourceSidecars = new Set<string>()

	for (const duplicate of duplicateRelationshipIds(packageGraph.relationships)) {
		const first = duplicate[0] as VerifyPackageGraphRelationship
		issues.push({
			rule: 'package-graph-integrity',
			severity: 'error',
			message: `Package relationship part ${first.relationshipPartPath} contains duplicate relationship id "${first.id}"`,
			refs: [`${first.relationshipPartPath}#${first.id}`],
			suggestedFix:
				'Repair duplicate OPC relationship ids before resolving or writing the package; relationship ids must be unique within each .rels part.',
			details: {
				code: 'package_relationship_duplicate_id',
				sourcePartPath: first.sourcePartPath,
				relationshipPartPath: first.relationshipPartPath,
				relationshipId: first.id,
				featureFamily: first.featureFamily,
				expected: 'unique relationship id within the relationship part',
				actual: duplicate.map((relationship) => ({
					type: relationship.rawType ?? relationship.type,
					target: relationship.rawTarget,
					resolvedTarget: relationship.resolvedTarget,
					targetMode: relationship.targetMode,
				})),
			},
		})
	}

	for (const relationship of packageGraph.relationships) {
		if (relationship.sourcePartPath !== '' && !partPaths.has(relationship.sourcePartPath)) {
			issues.push({
				rule: 'package-graph-integrity',
				severity: 'error',
				message: `Package relationship sidecar ${relationship.relationshipPartPath} belongs to missing source part "${relationship.sourcePartPath}"`,
				refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
				suggestedFix:
					'Remove the orphan relationship sidecar or restore the source package part before writing.',
				details: {
					code: 'package_relationship_source',
					sourcePartPath: relationship.sourcePartPath,
					relationshipPartPath: relationship.relationshipPartPath,
					relationshipId: relationship.id,
					featureFamily: relationship.featureFamily,
					expected: relationship.sourcePartPath,
					actual: undefined,
				},
			})
			reportedMissingSourceSidecars.add(relationship.relationshipPartPath)
		}
		if (relationship.targetMode?.toLowerCase() === 'external') continue
		if (relationship.resolvedTarget && partPaths.has(relationship.resolvedTarget)) continue
		issues.push({
			rule: 'package-graph-integrity',
			severity: 'error',
			message: `Package relationship ${relationship.relationshipPartPath}#${relationship.id} resolves to missing target "${relationship.resolvedTarget ?? relationship.rawTarget}"`,
			refs: [`${relationship.relationshipPartPath}#${relationship.id}`],
			suggestedFix:
				'Repair the relationship target or restore the referenced package part before writing.',
			details: {
				code: 'package_relationship_target',
				sourcePartPath: relationship.sourcePartPath,
				relationshipPartPath: relationship.relationshipPartPath,
				relationshipId: relationship.id,
				featureFamily: relationship.featureFamily,
				expected: relationship.rawTarget,
				actual: relationship.resolvedTarget,
			},
		})
	}
	for (const part of packageGraph.parts) {
		if (part.ownerScope !== 'relationship-part') continue
		if (reportedMissingSourceSidecars.has(part.path)) continue
		const sourcePartPath = sourcePartFromRelationshipPartPath(part.path)
		if (sourcePartPath === null || sourcePartPath === '' || partPaths.has(sourcePartPath)) continue
		issues.push({
			rule: 'package-graph-integrity',
			severity: 'error',
			message: `Package relationship sidecar ${part.path} belongs to missing source part "${sourcePartPath}"`,
			refs: [part.path],
			suggestedFix:
				'Remove the orphan relationship sidecar or restore the source package part before writing.',
			details: {
				code: 'package_relationship_source',
				sourcePartPath,
				relationshipPartPath: part.path,
				featureFamily: part.featureFamily,
				ownerScope: part.ownerScope,
				expected: sourcePartPath,
				actual: undefined,
			},
		})
	}
	for (const override of packageGraph.contentTypeOverrides ?? []) {
		const part = partByPath.get(override.partPath)
		if (!part) {
			issues.push({
				rule: 'package-graph-integrity',
				severity: 'error',
				message: `Content type override points to missing package part "${override.partPath}"`,
				refs: ['[Content_Types].xml', override.partPath],
				suggestedFix:
					'Remove the stale content type override or restore the referenced package part before writing.',
				details: {
					code: 'package_content_type_override_target',
					partPath: override.partPath,
					contentType: override.contentType,
					expected: override.partPath,
					actual: undefined,
				},
			})
			continue
		}
		if (part.contentType && part.contentType !== override.contentType) {
			issues.push({
				rule: 'package-graph-integrity',
				severity: 'error',
				message: `Content type override for "${override.partPath}" declares "${override.contentType}" but package graph resolved "${part.contentType}"`,
				refs: ['[Content_Types].xml', override.partPath],
				suggestedFix:
					'Make the content type override agree with the actual package part type before writing.',
				details: {
					code: 'package_content_type_override_mismatch',
					partPath: override.partPath,
					expected: override.contentType,
					actual: part.contentType,
					featureFamily: part.featureFamily,
				},
			})
		}
	}

	return issues
}

function duplicateRelationshipIds(
	relationships: readonly VerifyPackageGraphRelationship[],
): VerifyPackageGraphRelationship[][] {
	const byRelationshipId = new Map<string, VerifyPackageGraphRelationship[]>()
	for (const relationship of relationships) {
		const key = `${relationship.relationshipPartPath}\u0000${relationship.id}`
		const group = byRelationshipId.get(key)
		if (group) group.push(relationship)
		else byRelationshipId.set(key, [relationship])
	}
	return [...byRelationshipId.values()].filter((group) => group.length > 1)
}

function sourcePartFromRelationshipPartPath(path: string): string | null {
	if (path === '_rels/.rels') return ''
	const match = /^(.*)\/_rels\/([^/]+)\.rels$/i.exec(path)
	if (!match) return null
	const fileName = match[2]
	if (!fileName) return null
	return match[1] ? `${match[1]}/${fileName}` : fileName
}

export function check(workbook: Workbook, analysis?: CheckAnalysis): CheckResult {
	const formulas = analysis?.formulas ?? analyzeWorkbookFormulas(workbook)
	const dependencies = analysis?.dependencies ?? analyzeWorkbookDependencies(workbook)
	const sheetNames = workbook.sheets.map((s) => s.name)
	const issues = [
		...checkBrokenRefs(workbook, formulas, sheetNames),
		...checkExternalRefs(workbook, formulas),
		...checkCalcFreshness(workbook, formulas, analysis?.packageGraph),
		...checkChartSeriesReferences(workbook, sheetNames),
		...checkChartPartOwnership(workbook, sheetNames, analysis?.packageGraph),
		...checkCircularRefs(workbook, dependencies),
		...checkFormulaErrors(workbook, formulas),
		...checkBlockedSpills(workbook, formulas),
		...checkOrphanedNames(workbook),
		...checkMergeOverlaps(workbook),
		...checkTableIntegrity(workbook),
		...checkTablePackageGraphIntegrity(workbook, analysis?.packageGraph),
		...checkTableQueryTableIntegrity(workbook, analysis?.packageGraph),
		...checkExternalLinkIntegrity(workbook, analysis?.packageGraph),
		...checkPivotSlicerTimelineIntegrity(workbook, analysis?.packageGraph),
		...checkX14DataValidationIntegrity(workbook, sheetNames),
		...checkConditionalFormatIntegrity(workbook),
		...checkThreadedCommentIntegrity(workbook, analysis?.packageGraph),
		...checkLegacyCommentDrawingIntegrity(workbook, analysis?.packageGraph),
		...checkDrawingPackageIntegrity(workbook, analysis?.packageGraph),
		...checkChartStyleColorIntegrity(analysis?.packageGraph),
		...checkPackageGraphIntegrity(analysis?.packageGraph),
	]
	return { passed: issues.length === 0, issues }
}
