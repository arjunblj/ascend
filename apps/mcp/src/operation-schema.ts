import { z } from 'zod'

const cellUpdateSchema = z.object({
	ref: z.string(),
	value: z.union([z.string(), z.number(), z.boolean(), z.null(), z.date()]),
})

const sortSpecSchema = z.object({
	column: z.union([z.string(), z.number()]),
	descending: z.boolean().optional(),
})

const operationSchema = z.discriminatedUnion('op', [
	z.object({
		op: z.literal('setCells'),
		sheet: z.string(),
		updates: z.array(cellUpdateSchema),
	}),
	z.object({
		op: z.literal('setFormula'),
		sheet: z.string(),
		ref: z.string(),
		formula: z.string(),
	}),
	z.object({
		op: z.literal('fillFormula'),
		sheet: z.string(),
		range: z.string(),
		formula: z.string(),
	}),
	z.object({
		op: z.literal('clearRange'),
		sheet: z.string(),
		range: z.string(),
		what: z.enum(['values', 'formulas', 'styles', 'all']),
	}),
	z.object({
		op: z.literal('insertRows'),
		sheet: z.string(),
		at: z.number().int().nonnegative(),
		count: z.number().int().positive(),
	}),
	z.object({
		op: z.literal('deleteRows'),
		sheet: z.string(),
		at: z.number().int().nonnegative(),
		count: z.number().int().positive(),
	}),
	z.object({
		op: z.literal('insertCols'),
		sheet: z.string(),
		at: z.number().int().nonnegative(),
		count: z.number().int().positive(),
	}),
	z.object({
		op: z.literal('deleteCols'),
		sheet: z.string(),
		at: z.number().int().nonnegative(),
		count: z.number().int().positive(),
	}),
	z.object({
		op: z.literal('addSheet'),
		name: z.string(),
		position: z.number().int().nonnegative().optional(),
	}),
	z.object({
		op: z.literal('deleteSheet'),
		sheet: z.string(),
	}),
	z.object({
		op: z.literal('renameSheet'),
		sheet: z.string(),
		newName: z.string(),
	}),
	z.object({
		op: z.literal('moveSheet'),
		sheet: z.string(),
		position: z.number().int().nonnegative(),
	}),
	z.object({
		op: z.literal('createTable'),
		sheet: z.string(),
		ref: z.string(),
		name: z.string(),
		hasHeaders: z.boolean(),
	}),
	z.object({
		op: z.literal('appendRows'),
		table: z.string(),
		rows: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null(), z.date()]))),
	}),
	z.object({
		op: z.literal('sortRange'),
		sheet: z.string(),
		range: z.string(),
		by: z.array(sortSpecSchema),
	}),
	z.object({
		op: z.literal('mergeCells'),
		sheet: z.string(),
		range: z.string(),
	}),
	z.object({
		op: z.literal('unmergeCells'),
		sheet: z.string(),
		range: z.string(),
	}),
	z.object({
		op: z.literal('setColWidth'),
		sheet: z.string(),
		col: z.number().int().nonnegative(),
		width: z.number(),
	}),
	z.object({
		op: z.literal('setRowHeight'),
		sheet: z.string(),
		row: z.number().int().nonnegative(),
		height: z.number(),
	}),
	z.object({
		op: z.literal('setComment'),
		sheet: z.string(),
		ref: z.string(),
		text: z.string(),
		author: z.string().optional(),
	}),
	z.object({
		op: z.literal('setHyperlink'),
		sheet: z.string(),
		ref: z.string(),
		url: z.string(),
		display: z.string().optional(),
	}),
	z.object({
		op: z.literal('setNumberFormat'),
		sheet: z.string(),
		range: z.string(),
		format: z.string(),
	}),
	z.object({
		op: z.literal('setDefinedName'),
		name: z.string(),
		ref: z.string(),
		scope: z.string().optional(),
	}),
	z.object({
		op: z.literal('deleteDefinedName'),
		name: z.string(),
		scope: z.string().optional(),
	}),
	z.object({
		op: z.literal('setStyle'),
		sheet: z.string(),
		range: z.string(),
		style: z.record(z.string(), z.unknown()),
	}),
	z.object({
		op: z.literal('freezePane'),
		sheet: z.string(),
		row: z.number().int().nonnegative(),
		col: z.number().int().nonnegative(),
	}),
	z.object({
		op: z.literal('deleteComment'),
		sheet: z.string(),
		ref: z.string(),
	}),
	z.object({
		op: z.literal('deleteHyperlink'),
		sheet: z.string(),
		ref: z.string(),
	}),
	z.object({
		op: z.literal('setDataValidation'),
		sheet: z.string(),
		range: z.string(),
		rule: z.record(z.string(), z.unknown()),
	}),
	z.object({
		op: z.literal('deleteDataValidation'),
		sheet: z.string(),
		range: z.string(),
	}),
	z.object({
		op: z.literal('setAutoFilter'),
		sheet: z.string(),
		range: z.string(),
	}),
	z.object({
		op: z.literal('clearAutoFilter'),
		sheet: z.string(),
	}),
	z.object({
		op: z.literal('setSheetProtection'),
		sheet: z.string(),
		password: z.string().optional(),
		options: z.record(z.string(), z.unknown()).optional(),
	}),
	z.object({
		op: z.literal('setTabColor'),
		sheet: z.string(),
		color: z.string(),
	}),
	z.object({
		op: z.literal('hideSheet'),
		sheet: z.string(),
		hidden: z.boolean().optional(),
	}),
	z.object({
		op: z.literal('hideRows'),
		sheet: z.string(),
		at: z.number().int().nonnegative(),
		count: z.number().int().positive(),
		hidden: z.boolean().optional(),
	}),
	z.object({
		op: z.literal('hideCols'),
		sheet: z.string(),
		at: z.number().int().nonnegative(),
		count: z.number().int().positive(),
		hidden: z.boolean().optional(),
	}),
	z.object({
		op: z.literal('copySheet'),
		sheet: z.string(),
		newName: z.string(),
		position: z.number().int().nonnegative().optional(),
	}),
	z.object({
		op: z.literal('setConditionalFormat'),
		sheet: z.string(),
		range: z.string(),
		rule: z.record(z.string(), z.unknown()),
	}),
	z.object({
		op: z.literal('deleteConditionalFormat'),
		sheet: z.string(),
		range: z.string(),
	}),
	z.object({
		op: z.literal('setPageSetup'),
		sheet: z.string(),
		setup: z.record(z.string(), z.unknown()),
	}),
	z.object({
		op: z.literal('setPrintArea'),
		sheet: z.string(),
		range: z.string(),
	}),
	z.object({
		op: z.literal('copyRange'),
		sheet: z.string(),
		source: z.string(),
		target: z.string(),
	}),
	z.object({
		op: z.literal('moveRange'),
		sheet: z.string(),
		source: z.string(),
		target: z.string(),
	}),
	z.object({
		op: z.literal('groupRows'),
		sheet: z.string(),
		from: z.number().int().nonnegative(),
		to: z.number().int().nonnegative(),
		collapsed: z.boolean().optional(),
		summaryBelow: z.boolean().optional(),
	}),
	z.object({
		op: z.literal('groupCols'),
		sheet: z.string(),
		from: z.number().int().nonnegative(),
		to: z.number().int().nonnegative(),
		collapsed: z.boolean().optional(),
		summaryRight: z.boolean().optional(),
	}),
	z.object({
		op: z.literal('setRichText'),
		sheet: z.string(),
		ref: z.string(),
		runs: z.array(
			z.object({
				text: z.string(),
				bold: z.boolean().optional(),
				italic: z.boolean().optional(),
				underline: z.boolean().optional(),
				color: z.string().optional(),
				size: z.number().optional(),
			}),
		),
	}),
	z.object({
		op: z.literal('setWorkbookProtection'),
		protection: z.record(z.string(), z.unknown()),
	}),
	z.object({
		op: z.literal('deleteTable'),
		table: z.string(),
	}),
	z.object({
		op: z.literal('renameTable'),
		table: z.string(),
		newName: z.string(),
	}),
	z.object({
		op: z.literal('resizeTable'),
		table: z.string(),
		ref: z.string(),
	}),
	z.object({
		op: z.literal('replaceImage'),
		sheet: z.string(),
		contentBase64: z.string(),
		contentType: z.string(),
		targetPath: z.string().optional(),
		relId: z.string().optional(),
		name: z.string().optional(),
		imageIndex: z.number().int().nonnegative().optional(),
	}),
	z.object({
		op: z.literal('setPivotCache'),
		cacheId: z.number().int().nonnegative().optional(),
		partPath: z.string().optional(),
		pivotTable: z.string().optional(),
		sourceSheet: z.string().optional(),
		sourceRef: z.string().optional(),
		refreshOnLoad: z.boolean().optional(),
		enableRefresh: z.boolean().optional(),
		invalid: z.boolean().optional(),
		saveData: z.boolean().optional(),
	}),
	z.object({
		op: z.literal('rewriteExternalLink'),
		partPath: z.string().optional(),
		relId: z.string().optional(),
		linkRelId: z.string().optional(),
		target: z.string().optional(),
		newTarget: z.string(),
		targetMode: z.string().optional(),
	}),
])

export const operationsSchema = z.array(operationSchema)

export function parseOperations(
	input: unknown,
): { ok: true; value: unknown[] } | { ok: false; error: string } {
	const result = operationsSchema.safeParse(input)
	if (result.success) {
		return { ok: true, value: result.data }
	}
	const first = result.error.issues[0]
	const path = first?.path?.join('.') ?? 'ops'
	const msg = first?.message ?? result.error.message
	return {
		ok: false,
		error: `Invalid operations: ${path}: ${msg}`,
	}
}
