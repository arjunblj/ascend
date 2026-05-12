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
import { extractRefs, type FormulaRef, parseFormula } from '@ascend/formulas'
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
}

export interface VerifyPackageGraphPart {
	readonly path: string
	readonly featureFamily?: string
	readonly ownerScope?: string
	readonly contentType?: string
	readonly preservationPolicy?: string
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

function checkExternalRefs(analysis: WorkbookFormulaAnalysis): CheckIssue[] {
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

	return issues
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
				if (!parsed.sheetName) continue
				if (sheetNameSet.has(parsed.sheetName.toLowerCase())) continue
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
			}
		}
	}
	return issues
}

function checkChartPartOwnership(wb: Workbook, sheetNames: readonly string[]): CheckIssue[] {
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

function checkBlockedSpills(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	const unknownSpills: Array<{ sheetIndex: number; row: number; col: number }> = []
	for (let sheetIndex = 0; sheetIndex < wb.sheets.length; sheetIndex++) {
		const sheet = wb.sheets[sheetIndex]
		if (!sheet) continue
		for (const [row, col, cell] of sheet.cells.iterate()) {
			if (!cellHasFormula(cell) || !isError(cell.value) || cell.value.value !== '#SPILL!') continue
			const binding = cell.formulaInfo
			if (binding?.kind === 'blockedSpill') {
				issues.push(blockedSpillIssue(sheet.name, row, col, binding))
				continue
			}
			unknownSpills.push({ sheetIndex, row, col })
		}
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
		const bang = ref.indexOf('!')
		if (bang !== -1) {
			const sheetPart = ref.substring(0, bang).replace(/^'|'$/g, '')
			if (!sheetNames.has(sheetPart.toLowerCase())) {
				const closest = findClosestSheetName(sheetPart, sheetNameList)
				issues.push({
					rule: 'orphaned-names',
					severity: 'warning',
					message: `Defined name "${name}" references non-existent sheet "${sheetPart}"`,
					refs: [ref],
					...(closest ? { suggestedFix: `Did you mean sheet "${closest}"?` } : {}),
				})
			}
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

function isQueryTableRelationshipType(relationshipType: string): boolean {
	return relationshipType.toLowerCase().endsWith('/relationships/querytable')
}

function queryTableRelationshipDetails(
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

interface ExternalMetadataReference {
	readonly field: string
	readonly reference: string
	readonly sheetName: string
	readonly externalTarget: string
	readonly token?: string
}

function sqrefTokens(sqref: string): string[] {
	return sqref.trim().split(/\s+/).filter(Boolean)
}

function x14EntryRefs(sheetName: string, sqref: string, fallback: string): string[] {
	const refs = sqrefTokens(sqref).map((token) =>
		chartReferenceSheetName(token).sheetName ? token : `${sheetName}!${token}`,
	)
	return refs.length > 0 ? refs : [`${sheetName}#${fallback}`]
}

function missingSheetReferencesInSqref(
	sqref: string,
	sheetNameSet: ReadonlySet<string>,
): MissingSheetReference[] {
	const missing: MissingSheetReference[] = []
	for (const token of sqrefTokens(sqref)) {
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
	for (const ref of extractRefs(parsed.value)) {
		for (const sheetName of sheetNamesForFormulaRef(ref)) {
			if (sheetName.startsWith('[')) continue
			if (sheetNameSet.has(sheetName.toLowerCase())) continue
			missing.push({
				field: entry.field,
				reference: entry.formula,
				sheetName,
			})
		}
	}
	return missing
}

function externalReferencesInFormula(entry: X14FormulaReferenceEntry): ExternalMetadataReference[] {
	const parsed = parseFormulaText(entry.formula)
	if (!parsed.ok) return []
	const external: ExternalMetadataReference[] = []
	for (const ref of extractRefs(parsed.value)) {
		for (const sheetName of sheetNamesForFormulaRef(ref)) {
			if (!sheetName.startsWith('[')) continue
			external.push({
				field: entry.field,
				reference: entry.formula,
				sheetName,
				externalTarget: externalTargetFromSheetName(sheetName),
			})
		}
	}
	return external
}

function sheetNamesForFormulaRef(ref: FormulaRef): string[] {
	if (ref.kind === 'sheetSpan') return [ref.startSheet, ref.endSheet]
	return ref.sheet ? [ref.sheet] : []
}

function formulaHasDetectableReferences(formula: string | undefined): boolean {
	if (!formula) return false
	const parsed = parseFormulaText(formula)
	return parsed.ok && extractRefs(parsed.value).length > 0
}

function x14FormulaLiveFields(entries: readonly X14FormulaReferenceEntry[]): string[] {
	const fields: string[] = []
	for (const entry of entries) {
		if (formulaHasDetectableReferences(entry.formula)) fields.push(entry.field)
	}
	return fields
}

function pushX14MissingSheetIssues(
	issues: CheckIssue[],
	params: {
		readonly rule: 'conditional-format-integrity' | 'data-validation-integrity'
		readonly source: 'x14ConditionalFormat' | 'x14DataValidation'
		readonly sheetName: string
		readonly index: number
		readonly sqref: string
		readonly missing: readonly MissingSheetReference[]
		readonly sheetNames: readonly string[]
		readonly fallbackRef: string
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
				: 'Repair the x14 reference before writing preserved extension metadata.',
			details: {
				kind: missing.field === 'sqref' ? 'x14-sqref-missing-sheet' : 'x14-formula-missing-sheet',
				source: params.source,
				sheetName: params.sheetName,
				index: params.index,
				field: missing.field,
				reference: missing.reference,
				missingSheet: missing.sheetName,
				...(missing.token ? { token: missing.token } : {}),
			},
		})
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

	for (const sheet of wb.sheets) {
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
			pushExternalMetadataReferenceIssues(issues, {
				rule: 'data-validation-integrity',
				source: 'dataValidation',
				sheetName: sheet.name,
				index,
				references: [
					...externalReferencesInSqref(validation.sqref),
					...formulaEntries.flatMap(externalReferencesInFormula),
				],
				refs: x14EntryRefs(sheet.name, validation.sqref, `dataValidation[${index}]`),
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
	for (const sheet of wb.sheets) {
		const entries: ConditionalFormatPriorityEntry[] = []
		for (let formatIndex = 0; formatIndex < sheet.conditionalFormats.length; formatIndex++) {
			const format = sheet.conditionalFormats[formatIndex]
			if (!format) continue
			const ranges = parseSqrefRanges(format.sqref)
			pushExternalMetadataReferenceIssues(issues, {
				rule: 'conditional-format-integrity',
				source: 'conditionalFormat',
				sheetName: sheet.name,
				index: formatIndex,
				references: externalReferencesInSqref(format.sqref),
				refs: x14EntryRefs(sheet.name, format.sqref, `conditionalFormat[${formatIndex}]`),
				suggestedFix:
					'Replace conditional-format ranges with local workbook ranges or verify the external link metadata before editing conditional formats.',
				details: {
					sqref: format.sqref,
				},
			})
			for (let ruleIndex = 0; ruleIndex < format.rules.length; ruleIndex++) {
				const rule = format.rules[ruleIndex]
				if (!rule) continue
				pushExternalMetadataReferenceIssues(issues, {
					rule: 'conditional-format-integrity',
					source: 'conditionalFormat',
					sheetName: sheet.name,
					index: formatIndex,
					references: conditionalFormatRuleFormulaEntries(rule, ruleIndex).flatMap(
						externalReferencesInFormula,
					),
					refs: x14EntryRefs(sheet.name, format.sqref, `conditionalFormat[${formatIndex}]`),
					suggestedFix:
						'Replace conditional-format formulas with local workbook ranges or verify the external link metadata before editing conditional formats.',
					details: {
						sqref: format.sqref,
						ruleIndex,
						ruleType: rule.type,
					},
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
		for (const format of sheet.x14ConditionalFormats) {
			const fallbackRef = `x14ConditionalFormat[${format.index}]`
			const formulaEntries = x14ConditionalFormatFormulaEntries(format)
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
	const rootRefsByPart = new Map<string, Map<string, ThreadedCommentIntegrityEntry>>()
	const claimedPartsBySheet = new Map<string, Set<string>>()
	let threadedCommentsWithPersonIds = 0

	for (const sheet of wb.sheets) {
		for (let index = 0; index < sheet.threadedComments.length; index++) {
			const comment = sheet.threadedComments[index]
			if (!comment) continue
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
				try {
					parseA1(comment.ref)
				} catch {
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
						partPath,
						commentIndex: index,
						personId: comment.personId,
						...(comment.id ? { id: comment.id } : {}),
					},
				})
			} else if (comment.personId) {
				threadedCommentsWithPersonIds++
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
					partPath,
					commentIndex: index,
					parentId: comment.parentId,
					...(comment.id ? { id: comment.id } : {}),
				},
			})
		}
	}
	if (!packageGraph) return issues

	const personParts = packageGraph.parts.filter(isThreadedPersonPart)
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
			const drawing = comment.legacyDrawing
			if (!drawing) continue
			workbookLegacyDrawingCount++
			const sheetRef = `${sheet.name}!${ref}`
			try {
				const cellRef = parseA1(ref)
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
							ref,
							expectedRow: cellRef.row,
							expectedColumn: cellRef.col,
							...(drawing.row !== undefined ? { actualRow: drawing.row } : {}),
							...(drawing.column !== undefined ? { actualColumn: drawing.column } : {}),
							...(drawing.shapeId ? { shapeId: drawing.shapeId } : {}),
						},
					})
				}
			} catch {
				issues.push({
					rule: 'legacy-comment-drawing-integrity',
					severity: 'warning',
					message: `Legacy comment has invalid cell reference "${ref}"`,
					refs: [sheetRef],
					suggestedFix: 'Repair the comment reference before preserving or editing its VML layout.',
					details: {
						ref,
						...(drawing.shapeId ? { shapeId: drawing.shapeId } : {}),
					},
				})
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
	if (workbookCommentCount === 0) {
		for (const part of commentsParts) {
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
	if (workbookLegacyDrawingCount > 0 && commentsRelationships.length === 0) {
		issues.push({
			rule: 'legacy-comment-drawing-integrity',
			severity: 'warning',
			message: 'Legacy comment drawings exist but the package graph has no comments relationship',
			refs: commentsParts.map((part) => part.path),
			suggestedFix:
				'Restore the worksheet comments relationship before writing classic comments with preserved VML layout.',
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

function checkPackageGraphIntegrity(packageGraph?: VerifyPackageGraph): CheckIssue[] {
	if (!packageGraph) return []
	const issues: CheckIssue[] = []
	const partPaths = new Set(packageGraph.parts.map((part) => part.path))

	for (const relationship of packageGraph.relationships) {
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

	return issues
}

export function check(workbook: Workbook, analysis?: CheckAnalysis): CheckResult {
	const formulas = analysis?.formulas ?? analyzeWorkbookFormulas(workbook)
	const dependencies = analysis?.dependencies ?? analyzeWorkbookDependencies(workbook)
	const sheetNames = workbook.sheets.map((s) => s.name)
	const issues = [
		...checkBrokenRefs(workbook, formulas, sheetNames),
		...checkExternalRefs(formulas),
		...checkChartSeriesReferences(workbook, sheetNames),
		...checkChartPartOwnership(workbook, sheetNames),
		...checkCircularRefs(workbook, dependencies),
		...checkFormulaErrors(workbook, formulas),
		...checkBlockedSpills(workbook),
		...checkOrphanedNames(workbook),
		...checkMergeOverlaps(workbook),
		...checkTableIntegrity(workbook),
		...checkTableQueryTableIntegrity(workbook, analysis?.packageGraph),
		...checkExternalLinkIntegrity(workbook, analysis?.packageGraph),
		...checkX14DataValidationIntegrity(workbook, sheetNames),
		...checkConditionalFormatIntegrity(workbook),
		...checkThreadedCommentIntegrity(workbook, analysis?.packageGraph),
		...checkLegacyCommentDrawingIntegrity(workbook, analysis?.packageGraph),
		...checkPackageGraphIntegrity(analysis?.packageGraph),
	]
	return { passed: issues.length === 0, issues }
}
