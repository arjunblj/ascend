import type { Cell, RangeRef, Sheet, Table, Workbook } from '@ascend/core'
import { parseRange, toA1 } from '@ascend/core'
import type { FormulaCellRef, FormulaNode } from '@ascend/formulas'
import { cachedParseFormula, printFormula } from '@ascend/formulas'
import { shiftIndex, shiftRangeBounds } from './ref-shift.ts'

type ConditionalFormatRule = Sheet['conditionalFormats'][number]['rules'][number]
type X14ConditionalFormat = Sheet['x14ConditionalFormats'][number]
type FormulaTextRewriter = (formula: string | undefined) => string | undefined
const REF_ERROR_NODE: FormulaNode = { type: 'error', value: '#REF!' }

export interface PartialFormulaMoveReference {
	readonly ownerKind: 'cell-formula' | 'defined-name' | 'worksheet-metadata'
	readonly owner: string
	readonly formula: string
	readonly reference: string
}

export interface FormulaMoveSkipRange {
	readonly sheetName: string
	readonly range: RangeRef
}

export function rewriteWorkbookFormulasForShift(
	workbook: Workbook,
	targetSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	for (const sheet of workbook.sheets) {
		const isTarget = sheet.name === targetSheet
		const updates: [number, number, Cell][] = []
		for (const [row, col, existing] of sheet.cells.iterate()) {
			if (existing.formula === null) continue
			const parsed = cachedParseFormula(existing.formula)
			if (!parsed.ok) continue
			if (!isTarget && !formulaAstReferencesSheet(parsed.value, targetSheet)) continue
			const rewritten = rewriteFormulaAstForShift(
				parsed.value,
				targetSheet,
				sheet.name,
				axis,
				at,
				delta,
			)
			if (rewritten === parsed.value) continue
			const nextFormula = printFormula(rewritten)
			if (nextFormula === existing.formula) continue
			updates.push([
				row,
				col,
				{
					value: existing.value,
					formula: nextFormula,
					styleId: existing.styleId,
				},
			])
		}
		for (const [row, col, updated] of updates) {
			sheet.cells.set(row, col, updated)
		}
	}
}

export function rewriteDefinedNameFormulasForShift(
	workbook: Workbook,
	targetSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	const entries = [...workbook.definedNames.list()]
	for (const entry of entries) {
		const scope = entry.scope.kind === 'sheet' ? entry.scope : undefined
		const scopeSheet = scope
			? workbook.sheets.find((sheet) => sheet.id === scope.sheetId)?.name
			: undefined
		const parsed = cachedParseFormula(entry.formula)
		if (!parsed.ok) continue
		const formulaSheet = scopeSheet ?? targetSheet
		if (formulaSheet !== targetSheet && !formulaAstReferencesSheet(parsed.value, targetSheet))
			continue
		const rewritten = rewriteFormulaAstForShift(
			parsed.value,
			targetSheet,
			formulaSheet,
			axis,
			at,
			delta,
		)
		if (rewritten === parsed.value) continue
		const formula = printFormula(rewritten)
		if (formula === entry.formula) continue
		workbook.definedNames.set(entry.name, formula, entry.scope, definedNameOptions(entry))
	}
}

export function rewriteWorkbookFormulasForMove(
	workbook: Workbook,
	sourceSheet: string,
	targetSheet: string,
	sourceRange: RangeRef,
	targetRange: RangeRef,
): void {
	for (const sheet of workbook.sheets) {
		const updates: [number, number, Cell][] = []
		for (const [row, col, existing] of sheet.cells.iterate()) {
			if (existing.formula === null) continue
			if (sheet.name === targetSheet && rangeContainsCell(targetRange, row, col)) continue
			const rewrittenFormula = rewriteFormulaTextForMove(
				existing.formula,
				sourceSheet,
				targetSheet,
				sheet.name,
				sourceRange,
				targetRange,
			)
			if (rewrittenFormula === undefined || rewrittenFormula === existing.formula) continue
			updates.push([
				row,
				col,
				{
					value: existing.value,
					formula: rewrittenFormula,
					styleId: existing.styleId,
				},
			])
		}
		for (const [row, col, updated] of updates) sheet.cells.set(row, col, updated)
	}
}

export function rewriteDefinedNameFormulasForMove(
	workbook: Workbook,
	sourceSheet: string,
	targetSheet: string,
	sourceRange: RangeRef,
	targetRange: RangeRef,
): void {
	const entries = [...workbook.definedNames.list()]
	for (const entry of entries) {
		const scope = entry.scope.kind === 'sheet' ? entry.scope : undefined
		const scopeSheet = scope
			? workbook.sheets.find((sheet) => sheet.id === scope.sheetId)?.name
			: undefined
		const formulaSheet = scopeSheet ?? sourceSheet
		const formula = rewriteFormulaTextForMove(
			entry.formula,
			sourceSheet,
			targetSheet,
			formulaSheet,
			sourceRange,
			targetRange,
		)
		if (formula === undefined || formula === entry.formula) continue
		workbook.definedNames.set(entry.name, formula, entry.scope, definedNameOptions(entry))
	}
}

export function rewriteWorkbookMetadataFormulasForMove(
	workbook: Workbook,
	sourceSheet: string,
	targetSheet: string,
	sourceRange: RangeRef,
	targetRange: RangeRef,
): void {
	for (const sheet of workbook.sheets) {
		rewriteSheetMetadataFormulasForMove(
			sheet,
			sourceSheet,
			targetSheet,
			sheet.name,
			sourceRange,
			targetRange,
		)
	}
}

export function findPartialFormulaMoveReference(
	workbook: Workbook,
	sourceSheet: string,
	sourceRange: RangeRef,
	options: { readonly skipCells?: readonly FormulaMoveSkipRange[] } = {},
): PartialFormulaMoveReference | null {
	for (const sheet of workbook.sheets) {
		for (const [row, col, existing] of sheet.cells.iterate()) {
			if (existing.formula === null) continue
			if (shouldSkipFormulaCell(sheet.name, row, col, options.skipCells)) continue
			const reference = findPartialMoveReferenceInFormula(
				existing.formula,
				sourceSheet,
				sheet.name,
				sourceRange,
			)
			if (!reference) continue
			return {
				ownerKind: 'cell-formula',
				owner: `${sheet.name}!${toA1({ row, col })}`,
				formula: existing.formula,
				reference,
			}
		}
	}

	for (const entry of workbook.definedNames.list()) {
		const nameScope = entry.scope
		const scope =
			nameScope.kind === 'sheet'
				? workbook.sheets.find((sheet) => sheet.id === nameScope.sheetId)?.name
				: undefined
		const formulaSheet = scope ?? sourceSheet
		const reference = findPartialMoveReferenceInFormula(
			entry.formula,
			sourceSheet,
			formulaSheet,
			sourceRange,
		)
		if (!reference) continue
		return {
			ownerKind: 'defined-name',
			owner: scope ? `${scope}!${entry.name}` : entry.name,
			formula: entry.formula,
			reference,
		}
	}

	for (const sheet of workbook.sheets) {
		const issue = findPartialMoveReferenceInSheetMetadata(
			sheet,
			sourceSheet,
			sheet.name,
			sourceRange,
		)
		if (issue) return issue
	}
	return null
}

function shouldSkipFormulaCell(
	sheetName: string,
	row: number,
	col: number,
	skipCells: readonly FormulaMoveSkipRange[] | undefined,
): boolean {
	return (
		skipCells?.some(
			(skip) => skip.sheetName === sheetName && rangeContainsCell(skip.range, row, col),
		) ?? false
	)
}

export function retargetExplicitFormulaSheetRefsInRange(
	formula: string | undefined,
	sourceSheet: string,
	targetSheet: string,
	range: RangeRef,
): string | undefined {
	if (!formula || sourceSheet === targetSheet) return formula
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(
		retargetExplicitSheetRefsInRangeAst(parsed.value, sourceSheet, targetSheet, range),
	)
}

function rewriteFormulaTextForMove(
	formula: string | undefined,
	sourceSheet: string,
	targetSheet: string,
	formulaSheet: string,
	sourceRange: RangeRef,
	targetRange: RangeRef,
): string | undefined {
	if (!formula) return formula
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(
		rewriteFormulaAstForMove(
			parsed.value,
			sourceSheet,
			targetSheet,
			formulaSheet,
			sourceRange,
			targetRange,
		),
	)
}

function findPartialMoveReferenceInFormula(
	formula: string | undefined,
	sourceSheet: string,
	formulaSheet: string,
	sourceRange: RangeRef,
): string | null {
	if (!formula) return null
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return null
	const issue = findPartialMoveReferenceAst(parsed.value, sourceSheet, formulaSheet, sourceRange)
	return issue ? printFormula(issue) : null
}

export function rewriteFormulaTextForShift(
	formula: string | undefined,
	targetSheet: string,
	formulaSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): string | undefined {
	if (!formula) return formula
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(
		rewriteFormulaAstForShift(parsed.value, targetSheet, formulaSheet, axis, at, delta),
	)
}

export function rewriteSheetNameInFormulas(
	workbook: Workbook,
	oldName: string,
	newName: string,
): void {
	for (const sheet of workbook.sheets) {
		const updates: [number, number, Cell][] = []
		for (const [row, col, existing] of sheet.cells.iterate()) {
			if (existing.formula === null) continue
			const parsed = cachedParseFormula(existing.formula)
			if (!parsed.ok) continue
			const rewritten = rewriteSheetName(parsed.value, oldName, newName)
			const nextFormula = printFormula(rewritten)
			if (nextFormula === existing.formula) continue
			updates.push([
				row,
				col,
				{
					value: existing.value,
					formula: nextFormula,
					styleId: existing.styleId,
				},
			])
		}
		for (const [row, col, updated] of updates) {
			sheet.cells.set(row, col, updated)
		}
	}
}

export function rewriteSheetNameInDefinedNames(
	workbook: Workbook,
	oldName: string,
	newName: string,
): void {
	const entries = [...workbook.definedNames.list()]
	for (const entry of entries) {
		const parsed = cachedParseFormula(entry.formula)
		if (!parsed.ok) continue
		const rewritten = rewriteSheetName(parsed.value, oldName, newName)
		const formula = printFormula(rewritten)
		if (formula === entry.formula) continue
		workbook.definedNames.set(entry.name, formula, entry.scope, definedNameOptions(entry))
	}
}

export function rewriteTableNameInFormulas(
	workbook: Workbook,
	oldName: string,
	newName: string,
): void {
	for (const sheet of workbook.sheets) {
		const updates: [number, number, Cell][] = []
		for (const [row, col, existing] of sheet.cells.iterate()) {
			if (existing.formula === null) continue
			const rewrittenFormula = rewriteFormulaTextForTableRename(existing.formula, oldName, newName)
			if (rewrittenFormula === undefined) continue
			if (rewrittenFormula === existing.formula) continue
			updates.push([
				row,
				col,
				{
					value: existing.value,
					formula: rewrittenFormula,
					styleId: existing.styleId,
				},
			])
		}
		for (const [row, col, updated] of updates) {
			sheet.cells.set(row, col, updated)
		}
		rewriteSheetMetadataFormulasForTableRename(sheet, oldName, newName)
	}
}

export function rewriteTableNameInDefinedNames(
	workbook: Workbook,
	oldName: string,
	newName: string,
): void {
	const entries = [...workbook.definedNames.list()]
	for (const entry of entries) {
		const formula = rewriteFormulaTextForTableRename(entry.formula, oldName, newName)
		if (formula === undefined) continue
		if (formula === entry.formula) continue
		workbook.definedNames.set(entry.name, formula, entry.scope, definedNameOptions(entry))
	}
}

export function rewriteTableColumnInFormulas(
	workbook: Workbook,
	targetTable: Table,
	oldColumn: string,
	newColumn: string,
): void {
	for (const sheet of workbook.sheets) {
		const updates: [number, number, Cell][] = []
		for (const [row, col, existing] of sheet.cells.iterate()) {
			if (existing.formula === null) continue
			const includeLocalRefs =
				sheet.id === targetTable.sheetId && cellWithinTable(targetTable, row, col)
			const rewrittenFormula = rewriteFormulaTextForTableColumnRename(
				existing.formula,
				targetTable.name,
				oldColumn,
				newColumn,
				includeLocalRefs,
			)
			if (rewrittenFormula === undefined) continue
			if (rewrittenFormula === existing.formula) continue
			updates.push([
				row,
				col,
				{
					value: existing.value,
					formula: rewrittenFormula,
					styleId: existing.styleId,
				},
			])
		}
		for (const [row, col, updated] of updates) {
			sheet.cells.set(row, col, updated)
		}
		rewriteSheetMetadataFormulasForTableColumnRename(sheet, targetTable, oldColumn, newColumn)
	}
}

export function rewriteTableColumnInDefinedNames(
	workbook: Workbook,
	tableName: string,
	oldColumn: string,
	newColumn: string,
): void {
	const entries = [...workbook.definedNames.list()]
	for (const entry of entries) {
		const formula = rewriteFormulaTextForTableColumnRename(
			entry.formula,
			tableName,
			oldColumn,
			newColumn,
			false,
		)
		if (formula === undefined) continue
		if (formula === entry.formula) continue
		workbook.definedNames.set(entry.name, formula, entry.scope, definedNameOptions(entry))
	}
}

function definedNameOptions(entry: {
	readonly hidden?: boolean
	readonly extraAttributes?: readonly { readonly name: string; readonly value: string }[]
}): {
	readonly hidden?: boolean
	readonly extraAttributes?: readonly { readonly name: string; readonly value: string }[]
} {
	return {
		...(entry.hidden !== undefined ? { hidden: entry.hidden } : {}),
		...(entry.extraAttributes && entry.extraAttributes.length > 0
			? { extraAttributes: entry.extraAttributes }
			: {}),
	}
}

export function rewriteFormulaTextForRename(
	formula: string | undefined,
	oldName: string,
	newName: string,
): string | undefined {
	if (!formula) return formula
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(rewriteSheetName(parsed.value, oldName, newName))
}

export function rewriteFormulaTextForTableRename(
	formula: string | undefined,
	oldName: string,
	newName: string,
): string | undefined {
	if (!formula) return formula
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(rewriteTableName(parsed.value, oldName, newName))
}

export function rewriteFormulaTextForTableColumnRename(
	formula: string | undefined,
	tableName: string,
	oldColumn: string,
	newColumn: string,
	includeLocalRefs = false,
): string | undefined {
	if (!formula) return formula
	const parsed = cachedParseFormula(formula)
	if (!parsed.ok) return formula
	return printFormula(
		rewriteTableColumn(parsed.value, tableName, oldColumn, newColumn, includeLocalRefs),
	)
}

export function rewriteSheetMetadataFormulasForShift(
	sheet: Sheet,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	rewriteSheetMetadataFormulasForShiftTarget(sheet, sheet.name, sheet.name, axis, at, delta)
}

export function rewriteWorkbookMetadataFormulasForShift(
	workbook: Workbook,
	targetSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	for (const sheet of workbook.sheets) {
		if (sheet.name === targetSheet) continue
		rewriteSheetMetadataFormulasForShiftTarget(sheet, targetSheet, sheet.name, axis, at, delta)
	}
}

function rewriteSheetMetadataFormulasForShiftTarget(
	sheet: Sheet,
	targetSheet: string,
	formulaSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): void {
	for (let i = 0; i < sheet.dataValidations.length; i++) {
		const validation = sheet.dataValidations[i]
		if (!validation) continue
		const formula1 = rewriteFormulaTextForShift(
			validation.formula1,
			targetSheet,
			formulaSheet,
			axis,
			at,
			delta,
		)
		const formula2 = rewriteFormulaTextForShift(
			validation.formula2,
			targetSheet,
			formulaSheet,
			axis,
			at,
			delta,
		)
		sheet.dataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.conditionalFormats.length; i++) {
		const format = sheet.conditionalFormats[i]
		if (!format) continue
		const rewrite = (formula: string | undefined): string | undefined =>
			rewriteFormulaTextForShift(formula, targetSheet, formulaSheet, axis, at, delta)
		sheet.conditionalFormats[i] = {
			...format,
			rules: format.rules.map((rule) => rewriteConditionalFormatRuleWith(rule, rewrite)),
		}
	}
	for (let i = 0; i < sheet.x14DataValidations.length; i++) {
		const validation = sheet.x14DataValidations[i]
		if (!validation || validation.deleted) continue
		const formula1 = rewriteFormulaTextForShift(
			validation.formula1,
			targetSheet,
			formulaSheet,
			axis,
			at,
			delta,
		)
		const formula2 = rewriteFormulaTextForShift(
			validation.formula2,
			targetSheet,
			formulaSheet,
			axis,
			at,
			delta,
		)
		sheet.x14DataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.x14ConditionalFormats.length; i++) {
		const format = sheet.x14ConditionalFormats[i]
		if (!format || format.deleted) continue
		sheet.x14ConditionalFormats[i] = {
			...format,
			formulas: format.formulas.map(
				(formula) =>
					rewriteFormulaTextForShift(formula, targetSheet, formulaSheet, axis, at, delta) ??
					formula,
			),
			...(format.dataBar
				? {
						dataBar: {
							...format.dataBar,
							cfvo: format.dataBar.cfvo.map((entry) =>
								rewriteConditionalFormatValueObject(
									entry,
									targetSheet,
									formulaSheet,
									axis,
									at,
									delta,
								),
							),
						},
					}
				: {}),
			...(format.iconSet
				? {
						iconSet: {
							...format.iconSet,
							cfvo: format.iconSet.cfvo.map((entry) =>
								rewriteConditionalFormatValueObject(
									entry,
									targetSheet,
									formulaSheet,
									axis,
									at,
									delta,
								),
							),
						},
					}
				: {}),
		}
	}
	for (let i = 0; i < sheet.tables.length; i++) {
		const table = sheet.tables[i]
		if (!table) continue
		const columns = table.columns.map((column) => {
			const formula = rewriteFormulaTextForShift(
				column.formula,
				targetSheet,
				formulaSheet,
				axis,
				at,
				delta,
			)
			const totalsRowFormula = rewriteFormulaTextForShift(
				column.totalsRowFormula,
				targetSheet,
				formulaSheet,
				axis,
				at,
				delta,
			)
			return {
				...column,
				...(formula !== undefined ? { formula } : {}),
				...(totalsRowFormula !== undefined ? { totalsRowFormula } : {}),
			}
		})
		sheet.tables[i] = { ...table, columns }
	}
}

function rewriteConditionalFormatValueObject<T extends { readonly value?: string }>(
	entry: T,
	targetSheet: string,
	formulaSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): T {
	if (entry.value === undefined) return entry
	return {
		...entry,
		value:
			rewriteFormulaTextForShift(entry.value, targetSheet, formulaSheet, axis, at, delta) ??
			entry.value,
	}
}

function rewriteConditionalFormatValueObjectForMove<T extends { readonly value?: string }>(
	entry: T,
	sourceSheet: string,
	targetSheet: string,
	formulaSheet: string,
	sourceRange: RangeRef,
	targetRange: RangeRef,
): T {
	if (entry.value === undefined) return entry
	return {
		...entry,
		value:
			rewriteFormulaTextForMove(
				entry.value,
				sourceSheet,
				targetSheet,
				formulaSheet,
				sourceRange,
				targetRange,
			) ?? entry.value,
	}
}

function rewriteConditionalFormatRuleForMove(
	rule: ConditionalFormatRule,
	sourceSheet: string,
	targetSheet: string,
	formulaSheet: string,
	sourceRange: RangeRef,
	targetRange: RangeRef,
): ConditionalFormatRule {
	return {
		...rule,
		formulas: rule.formulas.map(
			(formula) =>
				rewriteFormulaTextForMove(
					formula,
					sourceSheet,
					targetSheet,
					formulaSheet,
					sourceRange,
					targetRange,
				) ?? formula,
		),
		...(rule.colorScale
			? {
					colorScale: {
						...rule.colorScale,
						cfvo: rule.colorScale.cfvo.map((entry) =>
							rewriteConditionalFormatValueObjectForMove(
								entry,
								sourceSheet,
								targetSheet,
								formulaSheet,
								sourceRange,
								targetRange,
							),
						),
						colors: rule.colorScale.colors.map((color) => ({ ...color })),
					},
				}
			: {}),
		...(rule.dataBar
			? {
					dataBar: {
						...rule.dataBar,
						cfvo: rule.dataBar.cfvo.map((entry) =>
							rewriteConditionalFormatValueObjectForMove(
								entry,
								sourceSheet,
								targetSheet,
								formulaSheet,
								sourceRange,
								targetRange,
							),
						),
						...(rule.dataBar.color ? { color: { ...rule.dataBar.color } } : {}),
					},
				}
			: {}),
		...(rule.iconSet
			? {
					iconSet: {
						...rule.iconSet,
						cfvo: rule.iconSet.cfvo.map((entry) =>
							rewriteConditionalFormatValueObjectForMove(
								entry,
								sourceSheet,
								targetSheet,
								formulaSheet,
								sourceRange,
								targetRange,
							),
						),
					},
				}
			: {}),
	}
}

function rewriteConditionalFormatValueObjectWith<T extends { readonly value?: string }>(
	entry: T,
	rewrite: FormulaTextRewriter,
): T {
	if (entry.value === undefined) return entry
	return {
		...entry,
		value: rewrite(entry.value) ?? entry.value,
	}
}

function rewriteConditionalFormatRuleWith(
	rule: ConditionalFormatRule,
	rewrite: FormulaTextRewriter,
): ConditionalFormatRule {
	return {
		...rule,
		formulas: rule.formulas.map((formula) => rewrite(formula) ?? formula),
		...(rule.colorScale
			? {
					colorScale: {
						...rule.colorScale,
						cfvo: rule.colorScale.cfvo.map((entry) =>
							rewriteConditionalFormatValueObjectWith(entry, rewrite),
						),
						colors: rule.colorScale.colors.map((color) => ({ ...color })),
					},
				}
			: {}),
		...(rule.dataBar
			? {
					dataBar: {
						...rule.dataBar,
						cfvo: rule.dataBar.cfvo.map((entry) =>
							rewriteConditionalFormatValueObjectWith(entry, rewrite),
						),
						...(rule.dataBar.color ? { color: { ...rule.dataBar.color } } : {}),
					},
				}
			: {}),
		...(rule.iconSet
			? {
					iconSet: {
						...rule.iconSet,
						cfvo: rule.iconSet.cfvo.map((entry) =>
							rewriteConditionalFormatValueObjectWith(entry, rewrite),
						),
					},
				}
			: {}),
	}
}

function rewriteX14ConditionalFormatWith(
	format: X14ConditionalFormat,
	rewrite: FormulaTextRewriter,
): X14ConditionalFormat {
	return {
		...format,
		formulas: format.formulas.map((formula) => rewrite(formula) ?? formula),
		...(format.dataBar
			? {
					dataBar: {
						...format.dataBar,
						cfvo: format.dataBar.cfvo.map((entry) =>
							rewriteConditionalFormatValueObjectWith(entry, rewrite),
						),
					},
				}
			: {}),
		...(format.iconSet
			? {
					iconSet: {
						...format.iconSet,
						cfvo: format.iconSet.cfvo.map((entry) =>
							rewriteConditionalFormatValueObjectWith(entry, rewrite),
						),
					},
				}
			: {}),
	}
}

export function rewriteSheetMetadataFormulasForMove(
	sheet: Sheet,
	sourceSheet: string,
	targetSheet: string,
	formulaSheet: string,
	sourceRange: RangeRef,
	targetRange: RangeRef,
): void {
	for (let i = 0; i < sheet.dataValidations.length; i++) {
		const validation = sheet.dataValidations[i]
		if (!validation) continue
		const formula1 = rewriteFormulaTextForMove(
			validation.formula1,
			sourceSheet,
			targetSheet,
			formulaSheet,
			sourceRange,
			targetRange,
		)
		const formula2 = rewriteFormulaTextForMove(
			validation.formula2,
			sourceSheet,
			targetSheet,
			formulaSheet,
			sourceRange,
			targetRange,
		)
		sheet.dataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.conditionalFormats.length; i++) {
		const format = sheet.conditionalFormats[i]
		if (!format) continue
		sheet.conditionalFormats[i] = {
			...format,
			rules: format.rules.map((rule) =>
				rewriteConditionalFormatRuleForMove(
					rule,
					sourceSheet,
					targetSheet,
					formulaSheet,
					sourceRange,
					targetRange,
				),
			),
		}
	}
	for (let i = 0; i < sheet.x14DataValidations.length; i++) {
		const validation = sheet.x14DataValidations[i]
		if (!validation || validation.deleted) continue
		const formula1 = rewriteFormulaTextForMove(
			validation.formula1,
			sourceSheet,
			targetSheet,
			formulaSheet,
			sourceRange,
			targetRange,
		)
		const formula2 = rewriteFormulaTextForMove(
			validation.formula2,
			sourceSheet,
			targetSheet,
			formulaSheet,
			sourceRange,
			targetRange,
		)
		sheet.x14DataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.x14ConditionalFormats.length; i++) {
		const format = sheet.x14ConditionalFormats[i]
		if (!format || format.deleted) continue
		sheet.x14ConditionalFormats[i] = {
			...format,
			formulas: format.formulas.map(
				(formula) =>
					rewriteFormulaTextForMove(
						formula,
						sourceSheet,
						targetSheet,
						formulaSheet,
						sourceRange,
						targetRange,
					) ?? formula,
			),
			...(format.dataBar
				? {
						dataBar: {
							...format.dataBar,
							cfvo: format.dataBar.cfvo.map((entry) =>
								rewriteConditionalFormatValueObjectForMove(
									entry,
									sourceSheet,
									targetSheet,
									formulaSheet,
									sourceRange,
									targetRange,
								),
							),
						},
					}
				: {}),
			...(format.iconSet
				? {
						iconSet: {
							...format.iconSet,
							cfvo: format.iconSet.cfvo.map((entry) =>
								rewriteConditionalFormatValueObjectForMove(
									entry,
									sourceSheet,
									targetSheet,
									formulaSheet,
									sourceRange,
									targetRange,
								),
							),
						},
					}
				: {}),
		}
	}
	for (let i = 0; i < sheet.tables.length; i++) {
		const table = sheet.tables[i]
		if (!table) continue
		const columns = table.columns.map((column) => {
			const formula = rewriteFormulaTextForMove(
				column.formula,
				sourceSheet,
				targetSheet,
				formulaSheet,
				sourceRange,
				targetRange,
			)
			const totalsRowFormula = rewriteFormulaTextForMove(
				column.totalsRowFormula,
				sourceSheet,
				targetSheet,
				formulaSheet,
				sourceRange,
				targetRange,
			)
			return {
				...column,
				...(formula !== undefined ? { formula } : {}),
				...(totalsRowFormula !== undefined ? { totalsRowFormula } : {}),
			}
		})
		sheet.tables[i] = { ...table, columns }
	}
}

function findPartialMoveReferenceInSheetMetadata(
	sheet: Sheet,
	sourceSheet: string,
	formulaSheet: string,
	sourceRange: RangeRef,
): PartialFormulaMoveReference | null {
	const check = (
		formula: string | undefined,
		owner: string,
	): PartialFormulaMoveReference | null => {
		if (!formula) return null
		const reference = findPartialMoveReferenceInFormula(
			formula,
			sourceSheet,
			formulaSheet,
			sourceRange,
		)
		if (!reference) return null
		return { ownerKind: 'worksheet-metadata', owner, formula, reference }
	}

	for (const validation of sheet.dataValidations) {
		const owner = `${sheet.name}!dataValidation(${validation.sqref})`
		const formula1 = check(validation.formula1, `${owner}.formula1`)
		if (formula1) return formula1
		const formula2 = check(validation.formula2, `${owner}.formula2`)
		if (formula2) return formula2
	}
	for (const format of sheet.conditionalFormats) {
		for (const [ruleIndex, rule] of format.rules.entries()) {
			const owner = `${sheet.name}!conditionalFormat(${format.sqref}).rules[${ruleIndex}]`
			for (const [formulaIndex, formula] of rule.formulas.entries()) {
				const issue = check(formula, `${owner}.formulas[${formulaIndex}]`)
				if (issue) return issue
			}
			const colorScale = findPartialMoveReferenceInValueObjects(
				rule.colorScale?.cfvo,
				check,
				`${owner}.colorScale.cfvo`,
			)
			if (colorScale) return colorScale
			const dataBar = findPartialMoveReferenceInValueObjects(
				rule.dataBar?.cfvo,
				check,
				`${owner}.dataBar.cfvo`,
			)
			if (dataBar) return dataBar
			const iconSet = findPartialMoveReferenceInValueObjects(
				rule.iconSet?.cfvo,
				check,
				`${owner}.iconSet.cfvo`,
			)
			if (iconSet) return iconSet
		}
	}
	for (const validation of sheet.x14DataValidations) {
		if (validation.deleted) continue
		const owner = `${sheet.name}!x14DataValidation(${validation.sqref})`
		const formula1 = check(validation.formula1, `${owner}.formula1`)
		if (formula1) return formula1
		const formula2 = check(validation.formula2, `${owner}.formula2`)
		if (formula2) return formula2
	}
	for (const format of sheet.x14ConditionalFormats) {
		if (format.deleted) continue
		const owner = `${sheet.name}!x14ConditionalFormat(${format.sqref})`
		for (const [formulaIndex, formula] of format.formulas.entries()) {
			const issue = check(formula, `${owner}.formulas[${formulaIndex}]`)
			if (issue) return issue
		}
		const dataBar = findPartialMoveReferenceInValueObjects(
			format.dataBar?.cfvo,
			check,
			`${owner}.dataBar.cfvo`,
		)
		if (dataBar) return dataBar
		const iconSet = findPartialMoveReferenceInValueObjects(
			format.iconSet?.cfvo,
			check,
			`${owner}.iconSet.cfvo`,
		)
		if (iconSet) return iconSet
	}
	for (const table of sheet.tables) {
		for (const [columnIndex, column] of table.columns.entries()) {
			const owner = `${sheet.name}!table(${table.name}).columns[${columnIndex}]`
			const formula = check(column.formula, `${owner}.formula`)
			if (formula) return formula
			const totalsRowFormula = check(column.totalsRowFormula, `${owner}.totalsRowFormula`)
			if (totalsRowFormula) return totalsRowFormula
		}
	}
	return null
}

function findPartialMoveReferenceInValueObjects<T extends { readonly value?: string }>(
	entries: readonly T[] | undefined,
	check: (formula: string | undefined, owner: string) => PartialFormulaMoveReference | null,
	owner: string,
): PartialFormulaMoveReference | null {
	if (!entries) return null
	for (const [index, entry] of entries.entries()) {
		const issue = check(entry.value, `${owner}[${index}]`)
		if (issue) return issue
	}
	return null
}

export function rewriteSheetMetadataFormulasForRename(
	sheet: Sheet,
	oldName: string,
	newName: string,
): void {
	for (let i = 0; i < sheet.dataValidations.length; i++) {
		const validation = sheet.dataValidations[i]
		if (!validation) continue
		const formula1 = rewriteFormulaTextForRename(validation.formula1, oldName, newName)
		const formula2 = rewriteFormulaTextForRename(validation.formula2, oldName, newName)
		sheet.dataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.conditionalFormats.length; i++) {
		const format = sheet.conditionalFormats[i]
		if (!format) continue
		sheet.conditionalFormats[i] = {
			...format,
			rules: format.rules.map((rule) => ({
				...rule,
				formulas: rule.formulas.map(
					(formula) => rewriteFormulaTextForRename(formula, oldName, newName) ?? formula,
				),
			})),
		}
	}
	for (let i = 0; i < sheet.tables.length; i++) {
		const table = sheet.tables[i]
		if (!table) continue
		const columns = table.columns.map((column) => {
			const formula = rewriteFormulaTextForRename(column.formula, oldName, newName)
			const totalsRowFormula = rewriteFormulaTextForRename(
				column.totalsRowFormula,
				oldName,
				newName,
			)
			return {
				...column,
				...(formula !== undefined ? { formula } : {}),
				...(totalsRowFormula !== undefined ? { totalsRowFormula } : {}),
			}
		})
		sheet.tables[i] = { ...table, columns }
	}
}

export function rewriteSheetMetadataFormulasForTableRename(
	sheet: Sheet,
	oldName: string,
	newName: string,
): void {
	const rewrite = (formula: string | undefined): string | undefined =>
		rewriteFormulaTextForTableRename(formula, oldName, newName)
	for (let i = 0; i < sheet.dataValidations.length; i++) {
		const validation = sheet.dataValidations[i]
		if (!validation) continue
		const formula1 = rewrite(validation.formula1)
		const formula2 = rewrite(validation.formula2)
		sheet.dataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.conditionalFormats.length; i++) {
		const format = sheet.conditionalFormats[i]
		if (!format) continue
		sheet.conditionalFormats[i] = {
			...format,
			rules: format.rules.map((rule) => rewriteConditionalFormatRuleWith(rule, rewrite)),
		}
	}
	for (let i = 0; i < sheet.x14DataValidations.length; i++) {
		const validation = sheet.x14DataValidations[i]
		if (!validation || validation.deleted) continue
		const formula1 = rewrite(validation.formula1)
		const formula2 = rewrite(validation.formula2)
		sheet.x14DataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.x14ConditionalFormats.length; i++) {
		const format = sheet.x14ConditionalFormats[i]
		if (!format || format.deleted) continue
		sheet.x14ConditionalFormats[i] = rewriteX14ConditionalFormatWith(format, rewrite)
	}
	for (let i = 0; i < sheet.tables.length; i++) {
		const table = sheet.tables[i]
		if (!table) continue
		const columns = table.columns.map((column) => {
			const formula = rewriteFormulaTextForTableRename(column.formula, oldName, newName)
			const totalsRowFormula = rewriteFormulaTextForTableRename(
				column.totalsRowFormula,
				oldName,
				newName,
			)
			return {
				...column,
				...(formula !== undefined ? { formula } : {}),
				...(totalsRowFormula !== undefined ? { totalsRowFormula } : {}),
			}
		})
		sheet.tables[i] = { ...table, columns }
	}
}

export function rewriteSheetMetadataFormulasForTableColumnRename(
	sheet: Sheet,
	targetTable: Table,
	oldColumn: string,
	newColumn: string,
): void {
	const tableName = targetTable.name
	const rewriteForSqref = (sqref: string): FormulaTextRewriter => {
		const includeLocalRefs =
			sheet.id === targetTable.sheetId && sqrefIntersectsRange(sqref, targetTable.ref)
		return (formula) =>
			rewriteFormulaTextForTableColumnRename(
				formula,
				tableName,
				oldColumn,
				newColumn,
				includeLocalRefs,
			)
	}
	for (let i = 0; i < sheet.dataValidations.length; i++) {
		const validation = sheet.dataValidations[i]
		if (!validation) continue
		const rewrite = rewriteForSqref(validation.sqref)
		const formula1 = rewrite(validation.formula1)
		const formula2 = rewrite(validation.formula2)
		sheet.dataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.conditionalFormats.length; i++) {
		const format = sheet.conditionalFormats[i]
		if (!format) continue
		const rewrite = rewriteForSqref(format.sqref)
		sheet.conditionalFormats[i] = {
			...format,
			rules: format.rules.map((rule) => rewriteConditionalFormatRuleWith(rule, rewrite)),
		}
	}
	for (let i = 0; i < sheet.x14DataValidations.length; i++) {
		const validation = sheet.x14DataValidations[i]
		if (!validation || validation.deleted) continue
		const rewrite = rewriteForSqref(validation.sqref)
		const formula1 = rewrite(validation.formula1)
		const formula2 = rewrite(validation.formula2)
		sheet.x14DataValidations[i] = {
			...validation,
			...(formula1 !== undefined ? { formula1 } : {}),
			...(formula2 !== undefined ? { formula2 } : {}),
		}
	}
	for (let i = 0; i < sheet.x14ConditionalFormats.length; i++) {
		const format = sheet.x14ConditionalFormats[i]
		if (!format || format.deleted) continue
		const rewrite = rewriteForSqref(format.sqref)
		sheet.x14ConditionalFormats[i] = rewriteX14ConditionalFormatWith(format, rewrite)
	}
	for (let i = 0; i < sheet.tables.length; i++) {
		const table = sheet.tables[i]
		if (!table) continue
		const includeLocalRefs = table.id === targetTable.id
		const columns = table.columns.map((column) => {
			const formula = rewriteFormulaTextForTableColumnRename(
				column.formula,
				tableName,
				oldColumn,
				newColumn,
				includeLocalRefs,
			)
			const totalsRowFormula = rewriteFormulaTextForTableColumnRename(
				column.totalsRowFormula,
				tableName,
				oldColumn,
				newColumn,
				includeLocalRefs,
			)
			return {
				...column,
				...(formula !== undefined ? { formula } : {}),
				...(totalsRowFormula !== undefined ? { totalsRowFormula } : {}),
			}
		})
		sheet.tables[i] = { ...table, columns }
	}
}

export function rewriteFormulaAstForShift(
	node: FormulaNode,
	targetSheet: string,
	formulaSheet: string,
	axis: 'row' | 'col',
	at: number,
	delta: number,
): FormulaNode {
	const onTarget = (refSheet: string | undefined) => (refSheet ?? formulaSheet) === targetSheet

	const shiftRef = (ref: FormulaCellRef, isOnTarget: boolean): FormulaCellRef | null => {
		if (!isOnTarget) return ref
		const next = axis === 'row' ? shiftIndex(ref.row, at, delta) : shiftIndex(ref.col, at, delta)
		if (next === null) return null
		if (axis === 'row') return next === ref.row ? ref : { ...ref, row: next }
		return next === ref.col ? ref : { ...ref, col: next }
	}

	switch (node.type) {
		case 'cellRef': {
			const ref = shiftRef(node.ref, onTarget(node.sheet))
			if (!ref) return REF_ERROR_NODE
			return ref === node.ref
				? node
				: node.sheet !== undefined
					? { type: 'cellRef', ref, sheet: node.sheet }
					: { type: 'cellRef', ref }
		}
		case 'rangeRef': {
			const hit = onTarget(node.sheet)
			if (!hit) return node
			const startIndex = axis === 'row' ? node.start.row : node.start.col
			const endIndex = axis === 'row' ? node.end.row : node.end.col
			const shifted = shiftRangeBounds(startIndex, endIndex, at, delta)
			if (!shifted) return REF_ERROR_NODE
			if (shifted.start === startIndex && shifted.end === endIndex) return node
			const start =
				axis === 'row'
					? { ...node.start, row: shifted.start }
					: { ...node.start, col: shifted.start }
			const end =
				axis === 'row' ? { ...node.end, row: shifted.end } : { ...node.end, col: shifted.end }
			return start === node.start && end === node.end
				? node
				: node.sheet !== undefined
					? { type: 'rangeRef', start, end, sheet: node.sheet }
					: { type: 'rangeRef', start, end }
		}
		case 'wholeRowRange': {
			if (!onTarget(node.sheet) || axis !== 'row') return node
			const shifted = shiftRangeBounds(node.startRow, node.endRow, at, delta)
			if (!shifted) return REF_ERROR_NODE
			const startRow = shifted.start
			const endRow = shifted.end
			return startRow === node.startRow && endRow === node.endRow
				? node
				: node.sheet !== undefined
					? { type: 'wholeRowRange', startRow, endRow, sheet: node.sheet }
					: { type: 'wholeRowRange', startRow, endRow }
		}
		case 'wholeColumnRange': {
			if (!onTarget(node.sheet) || axis !== 'col') return node
			const shifted = shiftRangeBounds(node.startCol, node.endCol, at, delta)
			if (!shifted) return REF_ERROR_NODE
			const startCol = shifted.start
			const endCol = shifted.end
			return startCol === node.startCol && endCol === node.endCol
				? node
				: node.sheet !== undefined
					? { type: 'wholeColumnRange', startCol, endCol, sheet: node.sheet }
					: { type: 'wholeColumnRange', startCol, endCol }
		}
		case 'binary': {
			const left = rewriteFormulaAstForShift(node.left, targetSheet, formulaSheet, axis, at, delta)
			const right = rewriteFormulaAstForShift(
				node.right,
				targetSheet,
				formulaSheet,
				axis,
				at,
				delta,
			)
			return left === node.left && right === node.right
				? node
				: { type: 'binary', op: node.op, left, right }
		}
		case 'dynamicRangeRef': {
			const start = rewriteFormulaAstForShift(
				node.start,
				targetSheet,
				formulaSheet,
				axis,
				at,
				delta,
			)
			const end = rewriteFormulaAstForShift(node.end, targetSheet, formulaSheet, axis, at, delta)
			return start === node.start && end === node.end
				? node
				: { type: 'dynamicRangeRef', start, end }
		}
		case 'unary': {
			const operand = rewriteFormulaAstForShift(
				node.operand,
				targetSheet,
				formulaSheet,
				axis,
				at,
				delta,
			)
			return operand === node.operand ? node : { type: 'unary', op: node.op, operand }
		}
		case 'spillRef': {
			const target = rewriteFormulaAstForShift(
				node.target,
				targetSheet,
				formulaSheet,
				axis,
				at,
				delta,
			)
			return target === node.target ? node : { type: 'spillRef', target }
		}
		case 'function': {
			const args = node.args.map((arg) =>
				rewriteFormulaAstForShift(arg, targetSheet, formulaSheet, axis, at, delta),
			)
			return args.every((arg, i) => arg === node.args[i])
				? node
				: { type: 'function', name: node.name, args }
		}
		case 'array': {
			const rows = node.rows.map((row) =>
				row.map((cell) =>
					rewriteFormulaAstForShift(cell, targetSheet, formulaSheet, axis, at, delta),
				),
			)
			return rows.every((row, ri) => row.every((cell, ci) => cell === node.rows[ri]?.[ci]))
				? node
				: { type: 'array', rows }
		}
		default:
			return node
	}
}

function rewriteFormulaAstForMove(
	node: FormulaNode,
	sourceSheet: string,
	targetSheet: string,
	formulaSheet: string,
	sourceRange: RangeRef,
	targetRange: RangeRef,
): FormulaNode {
	const rowDelta = targetRange.start.row - sourceRange.start.row
	const colDelta = targetRange.start.col - sourceRange.start.col
	const refOnSource = (refSheet: string | undefined) => (refSheet ?? formulaSheet) === sourceSheet
	const targetQualifier = (originalSheet: string | undefined): string | undefined => {
		if (originalSheet !== undefined) return targetSheet
		return targetSheet === formulaSheet ? undefined : targetSheet
	}
	const moveRef = (ref: FormulaCellRef, refSheet: string | undefined): FormulaCellRef | null => {
		if (!refOnSource(refSheet)) return null
		if (!rangeContainsCell(sourceRange, ref.row, ref.col)) return null
		return { ...ref, row: ref.row + rowDelta, col: ref.col + colDelta }
	}

	switch (node.type) {
		case 'cellRef': {
			const ref = moveRef(node.ref, node.sheet)
			if (!ref) return node
			const sheet = targetQualifier(node.sheet)
			return sheet !== undefined ? { type: 'cellRef', ref, sheet } : { type: 'cellRef', ref }
		}
		case 'rangeRef': {
			if (!refOnSource(node.sheet)) return node
			if (
				!rangeContainsCell(sourceRange, node.start.row, node.start.col) ||
				!rangeContainsCell(sourceRange, node.end.row, node.end.col)
			) {
				return node
			}
			const start = {
				...node.start,
				row: node.start.row + rowDelta,
				col: node.start.col + colDelta,
			}
			const end = { ...node.end, row: node.end.row + rowDelta, col: node.end.col + colDelta }
			const sheet = targetQualifier(node.sheet)
			return sheet !== undefined
				? { type: 'rangeRef', start, end, sheet }
				: { type: 'rangeRef', start, end }
		}
		case 'binary': {
			const left = rewriteFormulaAstForMove(
				node.left,
				sourceSheet,
				targetSheet,
				formulaSheet,
				sourceRange,
				targetRange,
			)
			const right = rewriteFormulaAstForMove(
				node.right,
				sourceSheet,
				targetSheet,
				formulaSheet,
				sourceRange,
				targetRange,
			)
			return left === node.left && right === node.right
				? node
				: { type: 'binary', op: node.op, left, right }
		}
		case 'dynamicRangeRef': {
			const start = rewriteFormulaAstForMove(
				node.start,
				sourceSheet,
				targetSheet,
				formulaSheet,
				sourceRange,
				targetRange,
			)
			const end = rewriteFormulaAstForMove(
				node.end,
				sourceSheet,
				targetSheet,
				formulaSheet,
				sourceRange,
				targetRange,
			)
			return start === node.start && end === node.end
				? node
				: { type: 'dynamicRangeRef', start, end }
		}
		case 'unary': {
			const operand = rewriteFormulaAstForMove(
				node.operand,
				sourceSheet,
				targetSheet,
				formulaSheet,
				sourceRange,
				targetRange,
			)
			return operand === node.operand ? node : { type: 'unary', op: node.op, operand }
		}
		case 'spillRef': {
			const target = rewriteFormulaAstForMove(
				node.target,
				sourceSheet,
				targetSheet,
				formulaSheet,
				sourceRange,
				targetRange,
			)
			return target === node.target ? node : { type: 'spillRef', target }
		}
		case 'function': {
			const args = node.args.map((arg) =>
				rewriteFormulaAstForMove(
					arg,
					sourceSheet,
					targetSheet,
					formulaSheet,
					sourceRange,
					targetRange,
				),
			)
			return args.every((arg, i) => arg === node.args[i])
				? node
				: { type: 'function', name: node.name, args }
		}
		case 'array': {
			const rows = node.rows.map((row) =>
				row.map((cell) =>
					rewriteFormulaAstForMove(
						cell,
						sourceSheet,
						targetSheet,
						formulaSheet,
						sourceRange,
						targetRange,
					),
				),
			)
			return rows.every((row, ri) => row.every((cell, ci) => cell === node.rows[ri]?.[ci]))
				? node
				: { type: 'array', rows }
		}
		case 'sheetSpanRef': {
			if (node.startSheet !== sourceSheet && node.endSheet !== sourceSheet) return node
			const target = rewriteFormulaAstForMove(
				node.target,
				sourceSheet,
				targetSheet,
				formulaSheet,
				sourceRange,
				targetRange,
			)
			return target === node.target ? node : { ...node, target }
		}
		default:
			return node
	}
}

function findPartialMoveReferenceAst(
	node: FormulaNode,
	sourceSheet: string,
	formulaSheet: string,
	sourceRange: RangeRef,
): FormulaNode | null {
	const refOnSource = (refSheet: string | undefined) => (refSheet ?? formulaSheet) === sourceSheet
	switch (node.type) {
		case 'rangeRef': {
			if (!refOnSource(node.sheet)) return null
			const range = formulaRangeFromRefs(node.start, node.end)
			return rangesOverlap(range, sourceRange) && !rangeContainsRange(sourceRange, range)
				? node
				: null
		}
		case 'wholeRowRange':
			if (!refOnSource(node.sheet)) return null
			return node.endRow >= sourceRange.start.row && node.startRow <= sourceRange.end.row
				? node
				: null
		case 'wholeColumnRange':
			if (!refOnSource(node.sheet)) return null
			return node.endCol >= sourceRange.start.col && node.startCol <= sourceRange.end.col
				? node
				: null
		case 'binary':
			return (
				findPartialMoveReferenceAst(node.left, sourceSheet, formulaSheet, sourceRange) ??
				findPartialMoveReferenceAst(node.right, sourceSheet, formulaSheet, sourceRange)
			)
		case 'dynamicRangeRef':
			return (
				findPartialMoveReferenceAst(node.start, sourceSheet, formulaSheet, sourceRange) ??
				findPartialMoveReferenceAst(node.end, sourceSheet, formulaSheet, sourceRange)
			)
		case 'unary':
			return findPartialMoveReferenceAst(node.operand, sourceSheet, formulaSheet, sourceRange)
		case 'spillRef':
			return findPartialMoveReferenceAst(node.target, sourceSheet, formulaSheet, sourceRange)
		case 'function':
			for (const arg of node.args) {
				const issue = findPartialMoveReferenceAst(arg, sourceSheet, formulaSheet, sourceRange)
				if (issue) return issue
			}
			return null
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) {
					const issue = findPartialMoveReferenceAst(cell, sourceSheet, formulaSheet, sourceRange)
					if (issue) return issue
				}
			}
			return null
		case 'sheetSpanRef':
			if (node.startSheet !== sourceSheet && node.endSheet !== sourceSheet) return null
			return findPartialMoveReferenceAst(node.target, sourceSheet, sourceSheet, sourceRange)
		default:
			return null
	}
}

function formulaRangeFromRefs(start: FormulaCellRef, end: FormulaCellRef): RangeRef {
	return {
		start: {
			row: Math.min(start.row, end.row),
			col: Math.min(start.col, end.col),
		},
		end: {
			row: Math.max(start.row, end.row),
			col: Math.max(start.col, end.col),
		},
	}
}

function rangeContainsCell(range: RangeRef, row: number, col: number): boolean {
	return (
		row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col
	)
}

function rangeContainsRange(outer: RangeRef, inner: RangeRef): boolean {
	return (
		rangeContainsCell(outer, inner.start.row, inner.start.col) &&
		rangeContainsCell(outer, inner.end.row, inner.end.col)
	)
}

function rangesOverlap(a: RangeRef, b: RangeRef): boolean {
	return !(
		a.end.row < b.start.row ||
		a.start.row > b.end.row ||
		a.end.col < b.start.col ||
		a.start.col > b.end.col
	)
}

function retargetExplicitSheetRefsInRangeAst(
	node: FormulaNode,
	sourceSheet: string,
	targetSheet: string,
	range: RangeRef,
): FormulaNode {
	switch (node.type) {
		case 'cellRef':
			return node.sheet === sourceSheet && rangeContainsCell(range, node.ref.row, node.ref.col)
				? { type: 'cellRef', ref: node.ref, sheet: targetSheet }
				: node
		case 'rangeRef':
			return node.sheet === sourceSheet &&
				rangeContainsCell(range, node.start.row, node.start.col) &&
				rangeContainsCell(range, node.end.row, node.end.col)
				? { type: 'rangeRef', start: node.start, end: node.end, sheet: targetSheet }
				: node
		case 'binary': {
			const left = retargetExplicitSheetRefsInRangeAst(node.left, sourceSheet, targetSheet, range)
			const right = retargetExplicitSheetRefsInRangeAst(node.right, sourceSheet, targetSheet, range)
			return left === node.left && right === node.right
				? node
				: { type: 'binary', op: node.op, left, right }
		}
		case 'dynamicRangeRef': {
			const start = retargetExplicitSheetRefsInRangeAst(node.start, sourceSheet, targetSheet, range)
			const end = retargetExplicitSheetRefsInRangeAst(node.end, sourceSheet, targetSheet, range)
			return start === node.start && end === node.end
				? node
				: { type: 'dynamicRangeRef', start, end }
		}
		case 'unary': {
			const operand = retargetExplicitSheetRefsInRangeAst(
				node.operand,
				sourceSheet,
				targetSheet,
				range,
			)
			return operand === node.operand ? node : { type: 'unary', op: node.op, operand }
		}
		case 'spillRef': {
			const target = retargetExplicitSheetRefsInRangeAst(
				node.target,
				sourceSheet,
				targetSheet,
				range,
			)
			return target === node.target ? node : { type: 'spillRef', target }
		}
		case 'function': {
			const args = node.args.map((arg) =>
				retargetExplicitSheetRefsInRangeAst(arg, sourceSheet, targetSheet, range),
			)
			return args.every((arg, i) => arg === node.args[i])
				? node
				: { type: 'function', name: node.name, args }
		}
		case 'array': {
			const rows = node.rows.map((row) =>
				row.map((cell) =>
					retargetExplicitSheetRefsInRangeAst(cell, sourceSheet, targetSheet, range),
				),
			)
			return rows.every((row, ri) => row.every((cell, ci) => cell === node.rows[ri]?.[ci]))
				? node
				: { type: 'array', rows }
		}
		case 'sheetSpanRef': {
			const target = retargetExplicitSheetRefsInRangeAst(
				node.target,
				sourceSheet,
				targetSheet,
				range,
			)
			return target === node.target ? node : { ...node, target }
		}
		default:
			return node
	}
}

export function formulaAstReferencesSheet(node: FormulaNode, sheetName: string): boolean {
	switch (node.type) {
		case 'cellRef':
		case 'rangeRef':
		case 'wholeRowRange':
		case 'wholeColumnRange':
		case 'name':
			return node.sheet === sheetName
		case 'sheetSpanRef':
			return node.startSheet === sheetName || node.endSheet === sheetName
		case 'binary':
			return (
				formulaAstReferencesSheet(node.left, sheetName) ||
				formulaAstReferencesSheet(node.right, sheetName)
			)
		case 'dynamicRangeRef':
			return (
				formulaAstReferencesSheet(node.start, sheetName) ||
				formulaAstReferencesSheet(node.end, sheetName)
			)
		case 'unary':
			return formulaAstReferencesSheet(node.operand, sheetName)
		case 'spillRef':
			return formulaAstReferencesSheet(node.target, sheetName)
		case 'function':
			return node.args.some((arg) => formulaAstReferencesSheet(arg, sheetName))
		case 'array':
			return node.rows.some((row) => row.some((cell) => formulaAstReferencesSheet(cell, sheetName)))
		default:
			return false
	}
}

function rewriteSheetName(node: FormulaNode, oldName: string, newName: string): FormulaNode {
	switch (node.type) {
		case 'cellRef':
			return node.sheet === oldName ? { type: 'cellRef', ref: node.ref, sheet: newName } : node
		case 'rangeRef':
			return node.sheet === oldName
				? { type: 'rangeRef', start: node.start, end: node.end, sheet: newName }
				: node
		case 'wholeRowRange':
			return node.sheet === oldName
				? { type: 'wholeRowRange', startRow: node.startRow, endRow: node.endRow, sheet: newName }
				: node
		case 'wholeColumnRange':
			return node.sheet === oldName
				? { type: 'wholeColumnRange', startCol: node.startCol, endCol: node.endCol, sheet: newName }
				: node
		case 'sheetSpanRef':
			return {
				type: 'sheetSpanRef',
				startSheet: node.startSheet === oldName ? newName : node.startSheet,
				endSheet: node.endSheet === oldName ? newName : node.endSheet,
				target: rewriteSheetName(node.target, oldName, newName),
			}
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteSheetName(node.left, oldName, newName),
				right: rewriteSheetName(node.right, oldName, newName),
			}
		case 'dynamicRangeRef':
			return {
				type: 'dynamicRangeRef',
				start: rewriteSheetName(node.start, oldName, newName),
				end: rewriteSheetName(node.end, oldName, newName),
			}
		case 'unary':
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteSheetName(node.operand, oldName, newName),
			}
		case 'function':
			return {
				type: 'function',
				name: node.name,
				args: node.args.map((arg) => rewriteSheetName(arg, oldName, newName)),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((row) => row.map((cell) => rewriteSheetName(cell, oldName, newName))),
			}
		default:
			return node
	}
}

function rewriteTableName(node: FormulaNode, oldName: string, newName: string): FormulaNode {
	switch (node.type) {
		case 'structuredRef':
			return node.table === oldName ? { ...node, table: newName } : node
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteTableName(node.left, oldName, newName),
				right: rewriteTableName(node.right, oldName, newName),
			}
		case 'dynamicRangeRef':
			return {
				type: 'dynamicRangeRef',
				start: rewriteTableName(node.start, oldName, newName),
				end: rewriteTableName(node.end, oldName, newName),
			}
		case 'unary':
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteTableName(node.operand, oldName, newName),
			}
		case 'function':
			return {
				type: 'function',
				name: node.name,
				args: node.args.map((arg) => rewriteTableName(arg, oldName, newName)),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((row) => row.map((cell) => rewriteTableName(cell, oldName, newName))),
			}
		case 'sheetSpanRef':
			return {
				...node,
				target: rewriteTableName(node.target, oldName, newName),
			}
		case 'spillRef':
			return {
				type: 'spillRef',
				target: rewriteTableName(node.target, oldName, newName),
			}
		default:
			return node
	}
}

function rewriteTableColumn(
	node: FormulaNode,
	tableName: string,
	oldColumn: string,
	newColumn: string,
	includeLocalRefs: boolean,
): FormulaNode {
	switch (node.type) {
		case 'structuredRef': {
			const tableMatches = node.table.toLowerCase() === tableName.toLowerCase()
			const localMatches = includeLocalRefs && node.table.length === 0
			if (!tableMatches && !localMatches) return node
			const oldColumnLower = oldColumn.toLowerCase()
			const column = node.column?.toLowerCase() === oldColumnLower ? newColumn : node.column
			const endColumn =
				node.endColumn?.toLowerCase() === oldColumnLower ? newColumn : node.endColumn
			return {
				...node,
				...(column !== undefined ? { column } : {}),
				...(endColumn !== undefined ? { endColumn } : {}),
			}
		}
		case 'binary':
			return {
				type: 'binary',
				op: node.op,
				left: rewriteTableColumn(node.left, tableName, oldColumn, newColumn, includeLocalRefs),
				right: rewriteTableColumn(node.right, tableName, oldColumn, newColumn, includeLocalRefs),
			}
		case 'dynamicRangeRef':
			return {
				type: 'dynamicRangeRef',
				start: rewriteTableColumn(node.start, tableName, oldColumn, newColumn, includeLocalRefs),
				end: rewriteTableColumn(node.end, tableName, oldColumn, newColumn, includeLocalRefs),
			}
		case 'unary':
			return {
				type: 'unary',
				op: node.op,
				operand: rewriteTableColumn(
					node.operand,
					tableName,
					oldColumn,
					newColumn,
					includeLocalRefs,
				),
			}
		case 'function':
			return {
				type: 'function',
				name: node.name,
				args: node.args.map((arg) =>
					rewriteTableColumn(arg, tableName, oldColumn, newColumn, includeLocalRefs),
				),
			}
		case 'array':
			return {
				type: 'array',
				rows: node.rows.map((row) =>
					row.map((cell) =>
						rewriteTableColumn(cell, tableName, oldColumn, newColumn, includeLocalRefs),
					),
				),
			}
		case 'sheetSpanRef':
			return {
				...node,
				target: rewriteTableColumn(node.target, tableName, oldColumn, newColumn, includeLocalRefs),
			}
		case 'spillRef':
			return {
				type: 'spillRef',
				target: rewriteTableColumn(node.target, tableName, oldColumn, newColumn, includeLocalRefs),
			}
		default:
			return node
	}
}

function sqrefIntersectsRange(sqref: string, range: RangeRef): boolean {
	for (const part of sqref.split(/\s+/)) {
		if (!part) continue
		try {
			if (rangesIntersect(parseRange(part), range)) return true
		} catch {}
	}
	return false
}

function rangesIntersect(a: RangeRef, b: RangeRef): boolean {
	const aStartRow = Math.min(a.start.row, a.end.row)
	const aEndRow = Math.max(a.start.row, a.end.row)
	const aStartCol = Math.min(a.start.col, a.end.col)
	const aEndCol = Math.max(a.start.col, a.end.col)
	const bStartRow = Math.min(b.start.row, b.end.row)
	const bEndRow = Math.max(b.start.row, b.end.row)
	const bStartCol = Math.min(b.start.col, b.end.col)
	const bEndCol = Math.max(b.start.col, b.end.col)
	return (
		aStartRow <= bEndRow && aEndRow >= bStartRow && aStartCol <= bEndCol && aEndCol >= bStartCol
	)
}

function cellWithinTable(table: Table, row: number, col: number): boolean {
	return (
		row >= table.ref.start.row &&
		row <= table.ref.end.row &&
		col >= table.ref.start.col &&
		col <= table.ref.end.col
	)
}
