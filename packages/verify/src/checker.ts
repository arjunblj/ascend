import type { CellFormulaBinding, RangeRef, Workbook } from '@ascend/core'
import { indexToColumn, parseA1, parseRange, toA1 } from '@ascend/core'
import {
	analyzeWorkbookDependencies,
	analyzeWorkbookFormulas,
	cellHasFormula,
	defaultCalcContext,
	parseCellKey,
	recalculate,
	type WorkbookDependencyAnalysis,
	type WorkbookFormulaAnalysis,
} from '@ascend/engine'
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

function checkChartSeriesReferences(wb: Workbook, sheetNames: readonly string[]): CheckIssue[] {
	const issues: CheckIssue[] = []
	const sheetNameSet = new Set(sheetNames.map((name) => name.toLowerCase()))
	for (const chart of wb.chartParts) {
		for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
			const series = chart.series[seriesIndex]
			if (!series) continue
			for (const entry of chartSeriesReferenceEntries(series)) {
				const parsed = chartReferenceSheetName(entry.reference)
				if (parsed.external || !parsed.sheetName) continue
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

	for (const sheet of wb.sheets) {
		for (const table of sheet.tables) {
			const rangeWidth = table.ref.end.col - table.ref.start.col + 1
			if (table.columns.length !== rangeWidth) {
				const rangeStr = `${indexToColumn(table.ref.start.col)}${table.ref.start.row + 1}:${indexToColumn(table.ref.end.col)}${table.ref.end.row + 1}`
				issues.push({
					rule: 'table-integrity',
					severity: 'error',
					message: `Table "${table.name}" has ${table.columns.length} columns but range spans ${rangeWidth}`,
					refs: [`${sheet.name}!${rangeStr}`],
				})
			}
		}
	}

	return issues
}

interface ConditionalFormatPriorityEntry {
	readonly source: 'conditionalFormat' | 'x14ConditionalFormat'
	readonly sheetName: string
	readonly priority: number
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
	for (const sheet of wb.sheets) {
		const entries: ConditionalFormatPriorityEntry[] = []
		for (let formatIndex = 0; formatIndex < sheet.conditionalFormats.length; formatIndex++) {
			const format = sheet.conditionalFormats[formatIndex]
			if (!format) continue
			const ranges = parseSqrefRanges(format.sqref)
			for (let ruleIndex = 0; ruleIndex < format.rules.length; ruleIndex++) {
				const rule = format.rules[ruleIndex]
				if (!rule || rule.priority === undefined) continue
				entries.push({
					source: 'conditionalFormat',
					sheetName: sheet.name,
					priority: rule.priority,
					sqref: format.sqref,
					formatIndex,
					ruleIndex,
					ruleType: rule.type,
					ranges,
				})
			}
		}
		for (const format of sheet.x14ConditionalFormats) {
			if (format.priority === undefined) continue
			entries.push({
				source: 'x14ConditionalFormat',
				sheetName: sheet.name,
				priority: format.priority,
				sqref: format.sqref,
				formatIndex: format.index,
				ranges: parseSqrefRanges(format.sqref),
				...(format.type ? { ruleType: format.type } : {}),
			})
		}
		for (const entry of entries) {
			if (entry.priority <= 0) {
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
				if (!right || left.priority !== right.priority) continue
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
	}
	return issues
}

function checkThreadedCommentIntegrity(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	for (const sheet of wb.sheets) {
		const idsByPart = new Map<
			string,
			Map<string, { readonly ref: string; readonly index: number }>
		>()
		for (let index = 0; index < sheet.threadedComments.length; index++) {
			const comment = sheet.threadedComments[index]
			if (!comment) continue
			const partPath = comment.partPath ?? '(unknown threaded comment part)'
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
						refs: [`${sheet.name}!${existing.ref}`, `${sheet.name}!${comment.ref}`],
						suggestedFix:
							'Inspect the threadedComments part before editing this thread; duplicate ids make replies ambiguous.',
						details: {
							partPath,
							id: comment.id,
							firstCommentIndex: existing.index,
							duplicateCommentIndex: index,
						},
					})
				} else {
					ids.set(comment.id, { ref: comment.ref, index })
				}
			}
			if (comment.personId && !comment.author) {
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
			}
		}
		for (let index = 0; index < sheet.threadedComments.length; index++) {
			const comment = sheet.threadedComments[index]
			if (!comment?.parentId) continue
			const partPath = comment.partPath ?? '(unknown threaded comment part)'
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
	return issues
}

function checkLegacyCommentDrawingIntegrity(wb: Workbook): CheckIssue[] {
	const issues: CheckIssue[] = []
	for (const sheet of wb.sheets) {
		const shapeIds = new Map<string, string>()
		for (const [ref, comment] of sheet.comments) {
			const drawing = comment.legacyDrawing
			if (!drawing) continue
			const sheetRef = `${sheet.name}!${ref}`
			try {
				const cellRef = parseA1(ref)
				if (
					(drawing.row !== undefined && drawing.row !== cellRef.row) ||
					(drawing.column !== undefined && drawing.column !== cellRef.col)
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
			}
			if (drawing.anchor?.some((value) => !Number.isInteger(value) || value < 0)) {
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
		}
	}
	return issues
}

export function check(
	workbook: Workbook,
	analysis?: {
		readonly formulas?: WorkbookFormulaAnalysis
		readonly dependencies?: WorkbookDependencyAnalysis
	},
): CheckResult {
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
		...checkConditionalFormatIntegrity(workbook),
		...checkThreadedCommentIntegrity(workbook),
		...checkLegacyCommentDrawingIntegrity(workbook),
	]
	return { passed: issues.length === 0, issues }
}
