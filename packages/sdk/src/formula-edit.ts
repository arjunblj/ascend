import {
	type FunctionDef,
	functionRegistry,
	normalizeFormulaInput,
	parseFormula,
	type Token,
	TokenType,
	tokenize,
} from '@ascend/formulas'

export type FormulaTokenClass =
	| 'reference'
	| 'name'
	| 'function'
	| 'literal'
	| 'operator'
	| 'punctuation'
	| 'whitespace'

export interface FormulaTokenRange {
	readonly type: TokenType
	readonly text: string
	readonly start: number
	readonly end: number
	readonly className: FormulaTokenClass
}

export interface FormulaReferenceRange {
	readonly text: string
	readonly start: number
	readonly end: number
	readonly kind:
		| 'cell'
		| 'range'
		| 'sheet-cell'
		| 'sheet-range'
		| 'sheet-3d-cell'
		| 'sheet-3d-range'
		| 'structured'
		| 'spill'
}

export type FormulaBindingRoleKind =
	| 'let-binding-declaration'
	| 'let-binding-use'
	| 'table-name-use'
	| 'table-column-use'
	| 'unresolved-name'

export interface FormulaBindingRole {
	readonly role: FormulaBindingRoleKind
	readonly text: string
	readonly start: number
	readonly end: number
	readonly bindingStart?: number
	readonly bindingEnd?: number
	readonly reference?: FormulaReferenceRange
}

export interface CycleReferenceResult {
	readonly formula: string
	readonly cursor: number
	readonly changed: boolean
	readonly reference?: FormulaReferenceRange
}

export interface InsertFormulaReferenceOptions {
	readonly replaceReferenceAtCursor?: boolean
}

export interface InsertFormulaReferenceResult {
	readonly formula: string
	readonly cursor: number
	readonly inserted: string
	readonly replaced?: FormulaReferenceRange
}

export interface FormulaDiagnostic {
	readonly code:
		| 'formula-parse-error'
		| 'formula-structured-reference-error'
		| 'formula-reference-qualifier-error'
	readonly severity: 'error'
	readonly message: string
	readonly start: number
	readonly end: number
}

export interface FormulaDiagnosticsResult {
	readonly parseOk: boolean
	readonly diagnostics: readonly FormulaDiagnostic[]
}

export interface FormulaSignatureParameter {
	readonly label: string
	readonly index: number
	readonly required: boolean
}

export interface FormulaFunctionSignature {
	readonly name: string
	readonly minArgs: number
	readonly maxArgs: number
	readonly volatile: boolean
	readonly label: string
	readonly parameters: readonly FormulaSignatureParameter[]
	readonly variadic: boolean
}

export interface FormulaFunctionCompletion {
	readonly name: string
	readonly minArgs: number
	readonly maxArgs: number
	readonly volatile: boolean
	readonly signature: FormulaFunctionSignature
}

export interface FormulaFunctionCompletionOptions {
	readonly limit?: number
}

export interface FormulaAssistOptions {
	readonly cursor?: number
	readonly prefix?: string
	readonly completionLimit?: number
	readonly functionName?: string
	readonly reference?: string
	readonly replaceReferenceAtCursor?: boolean
	readonly cycleReference?: boolean
}

export interface FormulaAssistResult {
	readonly formula: string
	readonly cursor: number | null
	readonly diagnostics: FormulaDiagnosticsResult
	readonly tokens: readonly FormulaTokenRange[]
	readonly references: readonly FormulaReferenceRange[]
	readonly bindings: readonly FormulaBindingRole[]
	readonly activeReference: FormulaReferenceRange | null
	readonly hover: FormulaHoverInfo | null
	readonly renameTarget: FormulaPrepareRenameResult | null
	readonly completions: readonly FormulaFunctionCompletion[]
	readonly signature: FormulaFunctionSignature | null
	readonly signatureHelp: FormulaFunctionSignatureHelp | null
	readonly codeActions: readonly FormulaCodeAction[]
	readonly cycle: CycleReferenceResult | null
	readonly insertion: InsertFormulaReferenceResult | null
}

export type FormulaHoverKind = 'diagnostic' | 'reference' | 'function' | 'token'

export interface FormulaHoverInfo {
	readonly kind: FormulaHoverKind
	readonly label: string
	readonly contents: readonly string[]
	readonly start: number
	readonly end: number
	readonly reference?: FormulaReferenceRange
	readonly signature?: FormulaFunctionSignature
	readonly diagnostic?: FormulaDiagnostic
	readonly token?: FormulaTokenRange
}

export interface FormulaCodeActionOptions {
	readonly reference?: string
	readonly replaceReferenceAtCursor?: boolean
	readonly cycleReference?: boolean
}

export interface FormulaCodeAction {
	readonly title: string
	readonly kind: 'quickfix' | 'refactor.rewrite' | 'source.insert'
	readonly start: number
	readonly end: number
	readonly edit: {
		readonly formula: string
		readonly cursor: number
	}
	readonly reference?: FormulaReferenceRange
	readonly diagnosticCodes?: readonly FormulaDiagnostic['code'][]
}

export type FormulaPrepareRenameBlockReason =
	| 'no-symbol-at-cursor'
	| 'workbook-context-required'
	| 'reference-target-not-renameable'

export interface FormulaPrepareRenameRange {
	readonly start: number
	readonly end: number
}

export interface FormulaPrepareRenameResult {
	readonly ok: boolean
	readonly reason?: FormulaPrepareRenameBlockReason
	readonly placeholder?: string
	readonly range?: FormulaPrepareRenameRange
	readonly occurrences: readonly FormulaPrepareRenameRange[]
	readonly role?: FormulaBindingRole
	readonly reference?: FormulaReferenceRange
	readonly boundary: string
}

export interface FormulaFunctionSignatureHelp {
	readonly signature: FormulaFunctionSignature
	readonly activeParameter: number
	readonly callStart: number
	readonly callEnd: number
	readonly argumentListStart: number
	readonly argumentListEnd: number
}

interface TokenSpan {
	readonly token: Token
	readonly start: number
	readonly end: number
	readonly text: string
}

interface FunctionCallFrame {
	readonly signature: FormulaFunctionSignature
	readonly callStart: number
	readonly openStart: number
	readonly separatorEnds: number[]
	closeStart?: number
	closeEnd?: number
}

type NestingFrame =
	| { readonly kind: 'function'; readonly call: FunctionCallFrame }
	| { readonly kind: 'paren' | 'brace' }

export function formulaTokenRanges(formula: string): FormulaTokenRange[] {
	const spans = tokenSpans(formula)
	return spans
		.filter((span) => span.token.type !== TokenType.EOF)
		.map((span) => ({
			type: span.token.type,
			text: span.text,
			start: span.start,
			end: span.end,
			className: classifyFormulaToken(span.token),
		}))
}

export function referenceAtCursor(formula: string, cursor: number): FormulaReferenceRange | null {
	const references = formulaReferenceRanges(formula)
	const clampedCursor = Math.max(0, Math.min(formula.length, cursor))
	for (const reference of references) {
		if (clampedCursor >= reference.start && clampedCursor <= reference.end) return reference
	}
	return null
}

export function formulaReferenceRanges(formula: string): FormulaReferenceRange[] {
	return collectFormulaReferenceRanges(formula, tokenSpans(formula))
}

export function formulaBindingRoles(formula: string): FormulaBindingRole[] {
	const spans = tokenSpans(formula)
	const references = collectFormulaReferenceRanges(formula, spans)
	const roles: FormulaBindingRole[] = []
	const claimed = new Set<string>()
	for (const reference of references) {
		if (reference.kind !== 'structured') continue
		for (const role of structuredReferenceBindingRoles(formula, spans, reference)) {
			roles.push(role)
			claimed.add(bindingRoleKey(role))
		}
	}
	for (const role of letBindingRoles(spans)) {
		roles.push(role)
		claimed.add(bindingRoleKey(role))
	}
	for (const span of spans) {
		if (
			span.token.type !== TokenType.Name ||
			isClaimedBindingSpan(span, claimed) ||
			references.some(
				(reference) =>
					reference.kind === 'structured' &&
					span.start >= reference.start &&
					span.end <= reference.end,
			)
		) {
			continue
		}
		roles.push({
			role: 'unresolved-name',
			text: span.text,
			start: span.start,
			end: span.end,
		})
	}
	return roles.sort((left, right) => left.start - right.start || left.end - right.end)
}

export function formulaPrepareRename(formula: string, cursor: number): FormulaPrepareRenameResult {
	const clampedCursor = Math.max(0, Math.min(formula.length, Math.floor(cursor)))
	const bindings = formulaBindingRoles(formula)
	const role = bindings.find(
		(binding) => clampedCursor >= binding.start && clampedCursor <= binding.end,
	)
	if (!role) {
		const reference = referenceAtCursor(formula, clampedCursor)
		return reference
			? {
					ok: false,
					reason: 'reference-target-not-renameable',
					occurrences: [],
					reference,
					boundary:
						'Cell, range, sheet, and external references require workbook operations, not formula-local rename.',
				}
			: {
					ok: false,
					reason: 'no-symbol-at-cursor',
					occurrences: [],
					boundary: 'No renameable formula symbol was found at the cursor.',
				}
	}
	if (role.role === 'let-binding-declaration' || role.role === 'let-binding-use') {
		const bindingStart = role.role === 'let-binding-declaration' ? role.start : role.bindingStart
		const bindingEnd = role.role === 'let-binding-declaration' ? role.end : role.bindingEnd
		if (bindingStart !== undefined && bindingEnd !== undefined) {
			const occurrences = bindings
				.filter(
					(candidate) =>
						(candidate.role === 'let-binding-declaration' &&
							candidate.start === bindingStart &&
							candidate.end === bindingEnd) ||
						(candidate.role === 'let-binding-use' &&
							candidate.bindingStart === bindingStart &&
							candidate.bindingEnd === bindingEnd),
				)
				.map(({ start, end }) => ({ start, end }))
			return {
				ok: true,
				placeholder: formula.slice(bindingStart, bindingEnd),
				range: { start: bindingStart, end: bindingEnd },
				occurrences,
				role,
				boundary:
					'Only formula-local LET bindings are rename-ready; callers must still apply edits explicitly.',
			}
		}
	}
	return {
		ok: false,
		reason: 'workbook-context-required',
		occurrences: [],
		role,
		boundary:
			'Workbook names, table names, and table columns require workbook-context resolution before rename.',
	}
}

export function cycleFormulaReferenceMode(formula: string, cursor: number): CycleReferenceResult {
	const spans = tokenSpans(formula)
	const clampedCursor = Math.max(0, Math.min(formula.length, cursor))
	const target = cycleTargetCellSpan(formula, spans, clampedCursor)
	if (!target) return { formula, cursor, changed: false }
	const cycled = cycleA1Reference(target.text)
	if (cycled === target.text) return { formula, cursor, changed: false }
	const nextFormula = `${formula.slice(0, target.start)}${cycled}${formula.slice(target.end)}`
	const reference = referenceAtCursor(nextFormula, target.start + cycled.length)
	const result: CycleReferenceResult = {
		formula: nextFormula,
		cursor: target.start + cycled.length,
		changed: true,
	}
	return reference ? { ...result, reference } : result
}

function cycleTargetCellSpan(
	formula: string,
	spans: readonly TokenSpan[],
	cursor: number,
): TokenSpan | undefined {
	for (const span of spans) {
		if (span.token.type === TokenType.CellRef && cursor >= span.start && cursor <= span.end) {
			return span
		}
	}
	const activeReference = collectFormulaReferenceRanges(formula, spans).find(
		(reference) => cursor >= reference.start && cursor <= reference.end,
	)
	if (
		!activeReference ||
		(activeReference.kind !== 'sheet-cell' &&
			activeReference.kind !== 'sheet-range' &&
			activeReference.kind !== 'sheet-3d-cell' &&
			activeReference.kind !== 'sheet-3d-range')
	) {
		return undefined
	}
	const firstCell = spans.find(
		(span) =>
			span.token.type === TokenType.CellRef &&
			span.start >= activeReference.start &&
			span.end <= activeReference.end,
	)
	if (!firstCell || cursor >= firstCell.start) return undefined
	return firstCell
}

export function insertFormulaReference(
	formula: string,
	cursor: number,
	reference: string,
	options: InsertFormulaReferenceOptions = {},
): InsertFormulaReferenceResult {
	const clampedCursor = Math.max(0, Math.min(formula.length, cursor))
	const replaced = options.replaceReferenceAtCursor
		? referenceAtCursor(formula, clampedCursor)
		: null
	const start = replaced?.start ?? clampedCursor
	const end = replaced?.end ?? clampedCursor
	return {
		formula: `${formula.slice(0, start)}${reference}${formula.slice(end)}`,
		cursor: start + reference.length,
		inserted: reference,
		...(replaced ? { replaced } : {}),
	}
}

export function formulaAssist(
	formula: string,
	options: FormulaAssistOptions = {},
): FormulaAssistResult {
	const cursor =
		typeof options.cursor === 'number' && Number.isFinite(options.cursor)
			? Math.max(0, Math.min(formula.length, Math.floor(options.cursor)))
			: null
	const completionLimit =
		typeof options.completionLimit === 'number' && Number.isFinite(options.completionLimit)
			? Math.max(0, Math.floor(options.completionLimit))
			: undefined
	return {
		formula,
		cursor,
		diagnostics: formulaDiagnostics(formula),
		tokens: formulaTokenRanges(formula),
		references: formulaReferenceRanges(formula),
		bindings: formulaBindingRoles(formula),
		activeReference: cursor === null ? null : referenceAtCursor(formula, cursor),
		hover: cursor === null ? null : formulaHover(formula, cursor),
		renameTarget: cursor === null ? null : formulaPrepareRename(formula, cursor),
		completions:
			options.prefix === undefined
				? []
				: formulaFunctionCompletions(options.prefix, {
						...(completionLimit !== undefined ? { limit: completionLimit } : {}),
					}),
		signature: options.functionName ? formulaFunctionSignature(options.functionName) : null,
		signatureHelp: cursor === null ? null : formulaFunctionSignatureHelp(formula, cursor),
		codeActions:
			cursor === null
				? []
				: formulaCodeActions(formula, cursor, {
						...(options.reference ? { reference: options.reference } : {}),
						replaceReferenceAtCursor: options.replaceReferenceAtCursor === true,
						cycleReference: options.cycleReference === true,
					}),
		cycle:
			cursor !== null && options.cycleReference === true
				? cycleFormulaReferenceMode(formula, cursor)
				: null,
		insertion:
			cursor !== null && options.reference
				? insertFormulaReference(formula, cursor, options.reference, {
						replaceReferenceAtCursor: options.replaceReferenceAtCursor === true,
					})
				: null,
	}
}

export function formulaHover(formula: string, cursor: number): FormulaHoverInfo | null {
	const clampedCursor = Math.max(0, Math.min(formula.length, cursor))
	const diagnostic = formulaDiagnostics(formula).diagnostics.find(
		(entry) => clampedCursor >= entry.start && clampedCursor <= entry.end,
	)
	if (diagnostic) {
		return {
			kind: 'diagnostic',
			label: diagnostic.code,
			contents: [diagnostic.message],
			start: diagnostic.start,
			end: diagnostic.end,
			diagnostic,
		}
	}
	const reference = referenceAtCursor(formula, clampedCursor)
	if (reference) {
		return {
			kind: 'reference',
			label: reference.text,
			contents: [`${reference.kind} reference`, `Range: ${reference.start}-${reference.end}`],
			start: reference.start,
			end: reference.end,
			reference,
		}
	}
	const signatureHelp = formulaFunctionSignatureHelp(formula, clampedCursor)
	if (signatureHelp) {
		const parameter = signatureHelp.signature.parameters[signatureHelp.activeParameter]
		return {
			kind: 'function',
			label: signatureHelp.signature.name,
			contents: [
				signatureHelp.signature.label,
				parameter ? `Active parameter: ${parameter.label}` : 'Active parameter: variadic',
			],
			start: signatureHelp.callStart,
			end: signatureHelp.callEnd,
			signature: signatureHelp.signature,
		}
	}
	const token = formulaTokenRanges(formula).find(
		(entry) => clampedCursor >= entry.start && clampedCursor <= entry.end,
	)
	if (!token || token.className === 'whitespace') return null
	const signature =
		token.type === TokenType.Function || token.className === 'function'
			? formulaFunctionSignature(token.text)
			: null
	return {
		kind: signature ? 'function' : 'token',
		label: token.text,
		contents: signature ? [signature.label] : [`${token.className} token`],
		start: token.start,
		end: token.end,
		token,
		...(signature ? { signature } : {}),
	}
}

export function formulaCodeActions(
	formula: string,
	cursor: number,
	options: FormulaCodeActionOptions = {},
): FormulaCodeAction[] {
	const clampedCursor = Math.max(0, Math.min(formula.length, cursor))
	const actions: FormulaCodeAction[] = []
	const activeReference = referenceAtCursor(formula, clampedCursor)
	if (options.cycleReference === true) {
		const cycled = cycleFormulaReferenceMode(formula, clampedCursor)
		if (cycled.changed) {
			actions.push({
				title: 'Cycle reference absolute/relative mode',
				kind: 'refactor.rewrite',
				start: activeReference?.start ?? clampedCursor,
				end: activeReference?.end ?? clampedCursor,
				edit: { formula: cycled.formula, cursor: cycled.cursor },
				...(cycled.reference ? { reference: cycled.reference } : {}),
			})
		}
	}
	if (options.reference) {
		const inserted = insertFormulaReference(formula, clampedCursor, options.reference, {
			replaceReferenceAtCursor: options.replaceReferenceAtCursor === true,
		})
		actions.push({
			title: inserted.replaced
				? `Replace reference with ${options.reference}`
				: `Insert reference ${options.reference}`,
			kind: inserted.replaced ? 'quickfix' : 'source.insert',
			start: inserted.replaced?.start ?? clampedCursor,
			end: inserted.replaced?.end ?? clampedCursor,
			edit: { formula: inserted.formula, cursor: inserted.cursor },
			...(inserted.replaced ? { reference: inserted.replaced } : {}),
		})
	}
	return actions
}

export function formulaFunctionSignature(name: string): FormulaFunctionSignature | null {
	const def = functionRegistry.get(name.toUpperCase())
	return def ? buildFormulaFunctionSignature(def) : null
}

export function formulaFunctionSignatureHelp(
	formula: string,
	cursor: number,
): FormulaFunctionSignatureHelp | null {
	const spans = tokenSpans(formula)
	const clampedCursor = Math.max(0, Math.min(formula.length, cursor))
	const calls: FunctionCallFrame[] = []
	const stack: NestingFrame[] = []

	for (let index = 0; index < spans.length; index++) {
		const span = spans[index]
		if (!span || span.token.type === TokenType.EOF) continue
		switch (span.token.type) {
			case TokenType.OpenParen: {
				const previous = previousNonWhitespace(spans, index - 1)
				if (previous?.span.token.type === TokenType.Function) {
					const signature = formulaFunctionSignature(previous.span.text)
					if (signature) {
						const call: FunctionCallFrame = {
							signature,
							callStart: previous.span.start,
							openStart: span.start,
							separatorEnds: [],
						}
						calls.push(call)
						stack.push({ kind: 'function', call })
						break
					}
				}
				stack.push({ kind: 'paren' })
				break
			}
			case TokenType.OpenBrace:
				stack.push({ kind: 'brace' })
				break
			case TokenType.CloseBrace:
				popNesting(stack, 'brace')
				break
			case TokenType.CloseParen: {
				const frame = popNesting(stack, 'function', 'paren')
				if (frame?.kind === 'function') {
					frame.call.closeStart = span.start
					frame.call.closeEnd = span.end
				}
				break
			}
			case TokenType.Comma:
			case TokenType.Semicolon: {
				const frame = stack[stack.length - 1]
				if (frame?.kind === 'function') frame.call.separatorEnds.push(span.end)
				break
			}
		}
	}

	let active: FunctionCallFrame | null = null
	for (const call of calls) {
		const contentStart = call.openStart + 1
		const contentEnd = call.closeStart ?? formula.length
		if (clampedCursor < contentStart || clampedCursor > contentEnd) continue
		if (!active || call.openStart >= active.openStart) active = call
	}
	if (!active) return null
	const activeParameter = active.separatorEnds.filter((end) => end <= clampedCursor).length
	return {
		signature: active.signature,
		activeParameter,
		callStart: active.callStart,
		callEnd: active.closeEnd ?? formula.length,
		argumentListStart: active.openStart,
		argumentListEnd: active.closeEnd ?? formula.length,
	}
}

export function formulaFunctionCompletions(
	prefix = '',
	options: FormulaFunctionCompletionOptions = {},
): FormulaFunctionCompletion[] {
	const normalizedPrefix = prefix.toUpperCase()
	const limit = Math.max(0, Math.floor(options.limit ?? 50))
	if (limit === 0) return []
	return Array.from(functionRegistry.values())
		.filter((def) => def.name.startsWith(normalizedPrefix))
		.sort((left, right) => left.name.localeCompare(right.name))
		.slice(0, limit)
		.map((def) => ({
			name: def.name,
			minArgs: def.minArgs,
			maxArgs: def.maxArgs,
			volatile: def.volatile === true,
			signature: buildFormulaFunctionSignature(def),
		}))
}

export function formulaDiagnostics(formula: string): FormulaDiagnosticsResult {
	const lexicalDiagnostics = formulaLexicalDiagnostics(formula)
	const normalized = normalizeFormulaInput(formula)
	const parsed = parseFormula(normalized)
	if (parsed.ok && lexicalDiagnostics.length === 0) return { parseOk: true, diagnostics: [] }
	if (parsed.ok) return { parseOk: false, diagnostics: lexicalDiagnostics }
	const offset = formula.startsWith('=') ? 1 : 0
	const position = parseDiagnosticPosition(parsed.error.message)
	const start =
		position === null ? formula.length : Math.max(0, Math.min(formula.length, position + offset))
	return {
		parseOk: false,
		diagnostics: [
			...lexicalDiagnostics,
			{
				code: 'formula-parse-error',
				severity: 'error',
				message: parsed.error.message,
				start,
				end: start < formula.length ? start + 1 : start,
			},
		],
	}
}

function formulaLexicalDiagnostics(formula: string): FormulaDiagnostic[] {
	const spans = tokenSpans(formula)
	const diagnostics: FormulaDiagnostic[] = []
	for (let index = 0; index < spans.length; index++) {
		const span = spans[index]
		if (!span || span.token.type !== TokenType.Name) continue
		if (span.text.startsWith('[') && bracketBalance(span.text) > 0) {
			diagnostics.push({
				code: 'formula-reference-qualifier-error',
				severity: 'error',
				message: 'Unterminated external workbook or bracketed reference',
				start: span.start,
				end: span.end,
			})
			continue
		}
		if (span.text.startsWith("'") && !quotedNameClosed(span.text)) {
			diagnostics.push({
				code: 'formula-reference-qualifier-error',
				severity: 'error',
				message: 'Unterminated quoted sheet or workbook reference',
				start: span.start,
				end: span.end,
			})
		}
		const next = nextNonWhitespace(spans, index + 1)
		if (next?.span.token.type !== TokenType.Name || !next.span.text.startsWith('[')) continue
		const balance = bracketBalance(next.span.text)
		if (balance === 0) continue
		diagnostics.push({
			code: 'formula-structured-reference-error',
			severity: 'error',
			message: balance > 0 ? 'Unterminated structured reference' : 'Malformed structured reference',
			start: next.span.start,
			end: next.span.end,
		})
		index = next.index
	}
	return diagnostics
}

function tokenSpans(formula: string): TokenSpan[] {
	const tokens = tokenize(formula)
	return tokens.map((token, index) => {
		const next = tokens[index + 1]
		const end = next ? next.position : formula.length
		return {
			token,
			start: token.position,
			end,
			text: formula.slice(token.position, end),
		}
	})
}

function collectFormulaReferenceRanges(
	formula: string,
	spans: readonly TokenSpan[],
): FormulaReferenceRange[] {
	const references: FormulaReferenceRange[] = []
	for (let i = 0; i < spans.length; i++) {
		const span = spans[i]
		if (!span || span.token.type === TokenType.EOF) continue
		if (span.token.type === TokenType.CellRef) {
			const range = collectCellReferenceRange(formula, spans, i)
			references.push(range)
			i = advanceReferenceIndex(spans, i, range.end)
			continue
		}
		if (span.token.type !== TokenType.Name) continue
		const next = nextNonWhitespace(spans, i + 1)
		if (span.text.startsWith('[') && next?.span.token.type === TokenType.Name) {
			const separator = nextNonWhitespace(spans, next.index + 1)
			const toSheet = separator ? nextNonWhitespace(spans, separator.index + 1) : undefined
			const bang = toSheet ? nextNonWhitespace(spans, toSheet.index + 1) : undefined
			const cell = bang ? nextNonWhitespace(spans, bang.index + 1) : undefined
			if (
				separator?.span.token.type === TokenType.Colon &&
				toSheet?.span.token.type === TokenType.Name &&
				bang?.span.token.type === TokenType.Bang &&
				cell?.span.token.type === TokenType.CellRef
			) {
				const local = collectCellReferenceRange(formula, spans, cell.index)
				references.push({
					text: formula.slice(span.start, local.end),
					start: span.start,
					end: local.end,
					kind: local.kind === 'range' ? 'sheet-3d-range' : 'sheet-3d-cell',
				})
				i = advanceReferenceIndex(spans, cell.index, local.end)
				continue
			}
		}
		if (next?.span.token.type === TokenType.Colon) {
			const toSheet = nextNonWhitespace(spans, next.index + 1)
			const bang = toSheet ? nextNonWhitespace(spans, toSheet.index + 1) : undefined
			const cell = bang ? nextNonWhitespace(spans, bang.index + 1) : undefined
			if (
				toSheet?.span.token.type === TokenType.Name &&
				bang?.span.token.type === TokenType.Bang &&
				cell?.span.token.type === TokenType.CellRef
			) {
				const local = collectCellReferenceRange(formula, spans, cell.index)
				references.push({
					text: formula.slice(span.start, local.end),
					start: span.start,
					end: local.end,
					kind: local.kind === 'range' ? 'sheet-3d-range' : 'sheet-3d-cell',
				})
				i = advanceReferenceIndex(spans, cell.index, local.end)
				continue
			}
		}
		if (span.text.startsWith('[') && next?.span.token.type === TokenType.Name) {
			const bang = nextNonWhitespace(spans, next.index + 1)
			if (bang?.span.token.type === TokenType.Bang) {
				const cell = nextNonWhitespace(spans, bang.index + 1)
				if (cell?.span.token.type === TokenType.CellRef) {
					const local = collectCellReferenceRange(formula, spans, cell.index)
					references.push({
						text: formula.slice(span.start, local.end),
						start: span.start,
						end: local.end,
						kind: local.kind === 'range' ? 'sheet-range' : 'sheet-cell',
					})
					i = advanceReferenceIndex(spans, cell.index, local.end)
					continue
				}
			}
		}
		if (next?.span.token.type === TokenType.Name && next.span.text.startsWith('[')) {
			references.push({
				text: formula.slice(span.start, next.span.end),
				start: span.start,
				end: next.span.end,
				kind: 'structured',
			})
			i = next.index
			continue
		}
		if (next?.span.token.type === TokenType.Bang) {
			const cell = nextNonWhitespace(spans, next.index + 1)
			if (cell?.span.token.type === TokenType.CellRef) {
				const local = collectCellReferenceRange(formula, spans, cell.index)
				const sheetSpanKind = sheetSpanReferenceKind(span.text, local.kind)
				references.push({
					text: formula.slice(span.start, local.end),
					start: span.start,
					end: local.end,
					kind: sheetSpanKind ?? (local.kind === 'range' ? 'sheet-range' : 'sheet-cell'),
				})
				i = advanceReferenceIndex(spans, cell.index, local.end)
			}
		}
	}
	return references
}

function advanceReferenceIndex(
	spans: readonly TokenSpan[],
	index: number,
	referenceEnd: number,
): number {
	let nextIndex = index
	while (nextIndex + 1 < spans.length) {
		const next = spans[nextIndex + 1]
		if (!next || next.start >= referenceEnd) break
		nextIndex++
	}
	return nextIndex
}

function structuredReferenceBindingRoles(
	formula: string,
	spans: readonly TokenSpan[],
	reference: FormulaReferenceRange,
): FormulaBindingRole[] {
	const roles: FormulaBindingRole[] = []
	const table = spans.find(
		(span) =>
			span.token.type === TokenType.Name &&
			span.start === reference.start &&
			span.end <= reference.end,
	)
	if (table) {
		roles.push({
			role: 'table-name-use',
			text: table.text,
			start: table.start,
			end: table.end,
			reference,
		})
	}
	const text = formula.slice(reference.start, reference.end)
	const tableLength = table?.text.length ?? 0
	const bodyStart = reference.start + tableLength
	for (const match of text.slice(tableLength).matchAll(/\[([^[\]#][^\]]*)\]/g)) {
		const raw = match[1]
		if (!raw) continue
		let trimStart = 0
		while (raw[trimStart] === '@' || raw[trimStart] === '[') trimStart++
		const column = raw.slice(trimStart).trim()
		if (!column || column.includes(':') || column.startsWith('#')) continue
		const leadingWhitespace = raw.slice(trimStart).length - raw.slice(trimStart).trimStart().length
		const start = bodyStart + (match.index ?? 0) + 1 + trimStart + leadingWhitespace
		roles.push({
			role: 'table-column-use',
			text: column,
			start,
			end: start + column.length,
			reference,
		})
	}
	return roles
}

interface LetBindingDeclaration {
	readonly name: string
	readonly span: TokenSpan
}

function letBindingRoles(spans: readonly TokenSpan[]): FormulaBindingRole[] {
	const roles: FormulaBindingRole[] = []
	collectLetBindingRoles(spans, 0, spans.length, [], roles)
	return roles
}

function collectLetBindingRoles(
	spans: readonly TokenSpan[],
	start: number,
	end: number,
	env: readonly LetBindingDeclaration[],
	roles: FormulaBindingRole[],
): void {
	for (let i = start; i < end; i++) {
		const span = spans[i]
		if (!span || span.token.type === TokenType.EOF) break
		if (span.token.type !== TokenType.Function || span.text.toUpperCase() !== 'LET') {
			if (span.token.type === TokenType.Name) {
				const binding = env.find((declaration) => namesEqual(declaration.name, span.text))
				if (binding) {
					roles.push({
						role: 'let-binding-use',
						text: span.text,
						start: span.start,
						end: span.end,
						bindingStart: binding.span.start,
						bindingEnd: binding.span.end,
					})
				}
			}
			continue
		}
		const open = nextNonWhitespace(spans, i + 1)
		if (open?.span.token.type !== TokenType.OpenParen) continue
		const args = functionArgumentRanges(spans, open.index)
		if (args.length < 3) continue
		const closeIndex = matchingCloseParenIndex(spans, open.index)
		let scopedEnv = env
		for (let argIndex = 0; argIndex < args.length - 1; argIndex += 2) {
			const declaration = singleNameArgument(spans, args[argIndex])
			if (declaration) {
				roles.push({
					role: 'let-binding-declaration',
					text: declaration.text,
					start: declaration.start,
					end: declaration.end,
				})
			}
			collectLetBindingRoles(
				spans,
				args[argIndex + 1]?.start ?? 0,
				args[argIndex + 1]?.end ?? 0,
				scopedEnv,
				roles,
			)
			if (declaration) scopedEnv = [{ name: declaration.text, span: declaration }, ...scopedEnv]
		}
		const finalArg = args[args.length - 1]
		if (finalArg) collectLetBindingRoles(spans, finalArg.start, finalArg.end, scopedEnv, roles)
		if (closeIndex !== undefined) i = closeIndex
	}
}

function functionArgumentRanges(
	spans: readonly TokenSpan[],
	openIndex: number,
): Array<{ readonly start: number; readonly end: number }> {
	const ranges: Array<{ readonly start: number; readonly end: number }> = []
	let depth = 0
	let start = openIndex + 1
	for (let i = openIndex + 1; i < spans.length; i++) {
		const span = spans[i]
		if (!span || span.token.type === TokenType.EOF) break
		if (span.token.type === TokenType.OpenParen) {
			depth++
			continue
		}
		if (span.token.type === TokenType.CloseParen) {
			if (depth === 0) {
				ranges.push({ start, end: i })
				return ranges
			}
			depth--
			continue
		}
		if (span.token.type === TokenType.Comma && depth === 0) {
			ranges.push({ start, end: i })
			start = i + 1
		}
	}
	return ranges
}

function matchingCloseParenIndex(
	spans: readonly TokenSpan[],
	openIndex: number,
): number | undefined {
	let depth = 0
	for (let i = openIndex + 1; i < spans.length; i++) {
		const span = spans[i]
		if (!span || span.token.type === TokenType.EOF) return undefined
		if (span.token.type === TokenType.OpenParen) {
			depth++
			continue
		}
		if (span.token.type === TokenType.CloseParen) {
			if (depth === 0) return i
			depth--
		}
	}
	return undefined
}

function singleNameArgument(
	spans: readonly TokenSpan[],
	range: { readonly start: number; readonly end: number } | undefined,
): TokenSpan | undefined {
	const names = nameSpansInRange(spans, range ?? null)
	return names.length === 1 ? names[0] : undefined
}

function nameSpansInRange(
	spans: readonly TokenSpan[],
	range: { readonly start: number; readonly end: number } | null,
): TokenSpan[] {
	if (!range) return []
	const names: TokenSpan[] = []
	for (let i = range.start; i < range.end; i++) {
		const span = spans[i]
		if (span?.token.type === TokenType.Name) names.push(span)
	}
	return names
}

function namesEqual(left: string, right: string): boolean {
	return left.toLowerCase() === right.toLowerCase()
}

function bindingRoleKey(role: Pick<FormulaBindingRole, 'start' | 'end'>): string {
	return `${role.start}:${role.end}`
}

function isClaimedBindingSpan(span: TokenSpan, claimed: ReadonlySet<string>): boolean {
	return claimed.has(bindingRoleKey(span))
}

function collectCellReferenceRange(
	formula: string,
	spans: readonly TokenSpan[],
	index: number,
): FormulaReferenceRange {
	const start = spans[index]
	if (!start) throw new Error('Missing cell reference token')
	let end = start.end
	let kind: FormulaReferenceRange['kind'] = 'cell'
	const afterCell = nextNonWhitespace(spans, index + 1)
	if (afterCell?.span.token.type === TokenType.Colon) {
		const rangeEnd = nextNonWhitespace(spans, afterCell.index + 1)
		if (rangeEnd?.span.token.type === TokenType.CellRef) {
			end = rangeEnd.span.end
			kind = 'range'
		}
	}
	const afterRange = nextNonWhitespace(spans, index + 1)
	if (
		kind === 'cell' &&
		afterRange?.span.token.type === TokenType.Operator &&
		afterRange.span.text === '#'
	) {
		end = afterRange.span.end
		kind = 'spill'
	}
	return {
		text: formula.slice(start.start, end),
		start: start.start,
		end,
		kind,
	}
}

function sheetSpanReferenceKind(
	qualifierText: string,
	localKind: FormulaReferenceRange['kind'],
): FormulaReferenceRange['kind'] | null {
	return sheetSpanColonIndex(unquoteSheetQualifierText(qualifierText)) >= 0
		? localKind === 'range'
			? 'sheet-3d-range'
			: 'sheet-3d-cell'
		: null
}

function unquoteSheetQualifierText(text: string): string {
	return text.startsWith("'") && text.endsWith("'") && text.length >= 2
		? text.slice(1, -1).replace(/''/g, "'")
		: text
}

function sheetSpanColonIndex(text: string): number {
	const workbookOpen = text.indexOf('[')
	if (workbookOpen < 0) return text.indexOf(':')
	const workbookClose = text.indexOf(']', workbookOpen + 1)
	if (workbookClose < 0) return -1
	return text.indexOf(':', workbookClose + 1)
}

function nextNonWhitespace(
	spans: readonly TokenSpan[],
	start: number,
): { readonly span: TokenSpan; readonly index: number } | undefined {
	for (let index = start; index < spans.length; index++) {
		const span = spans[index]
		if (!span || span.token.type === TokenType.EOF) return undefined
		if (span.token.type !== TokenType.Whitespace) return { span, index }
	}
	return undefined
}

function previousNonWhitespace(
	spans: readonly TokenSpan[],
	start: number,
): { readonly span: TokenSpan; readonly index: number } | undefined {
	for (let index = start; index >= 0; index--) {
		const span = spans[index]
		if (!span) return undefined
		if (span.token.type !== TokenType.Whitespace) return { span, index }
	}
	return undefined
}

function popNesting(
	stack: NestingFrame[],
	...kinds: readonly NestingFrame['kind'][]
): NestingFrame | undefined {
	const frame = stack[stack.length - 1]
	return frame && kinds.includes(frame.kind) ? stack.pop() : undefined
}

function classifyFormulaToken(token: Token): FormulaTokenClass {
	switch (token.type) {
		case TokenType.CellRef:
			return 'reference'
		case TokenType.Name:
			return 'name'
		case TokenType.Function:
			return 'function'
		case TokenType.Number:
		case TokenType.String:
		case TokenType.Boolean:
		case TokenType.Error:
			return 'literal'
		case TokenType.Operator:
			return 'operator'
		case TokenType.Whitespace:
			return 'whitespace'
		default:
			return 'punctuation'
	}
}

function cycleA1Reference(ref: string): string {
	const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/.exec(ref)
	if (!match) return ref
	const [, colAbs = '', col = '', rowAbs = '', row = ''] = match
	if (!colAbs && !rowAbs) return `$${col.toUpperCase()}$${row}`
	if (colAbs && rowAbs) return `${col.toUpperCase()}$${row}`
	if (!colAbs && rowAbs) return `$${col.toUpperCase()}${row}`
	return `${col.toUpperCase()}${row}`
}

function buildFormulaFunctionSignature(def: FunctionDef): FormulaFunctionSignature {
	const showAllOptional = def.maxArgs - def.minArgs <= 3
	const shownOptionalCount = showAllOptional
		? def.maxArgs - def.minArgs
		: def.maxArgs > def.minArgs
			? 1
			: 0
	const shownCount = def.minArgs + shownOptionalCount
	const parameters: FormulaSignatureParameter[] = []
	const labels: string[] = []
	for (let index = 0; index < shownCount; index++) {
		const required = index < def.minArgs
		const label = `arg${index + 1}`
		parameters.push({ label, index, required })
		labels.push(required ? label : `[${label}]`)
	}
	const variadic = shownCount < def.maxArgs
	if (variadic) labels.push('...')
	return {
		name: def.name,
		minArgs: def.minArgs,
		maxArgs: def.maxArgs,
		volatile: def.volatile === true,
		label: `${def.name}(${labels.join(', ')})`,
		parameters,
		variadic,
	}
}

function bracketBalance(text: string): number {
	let balance = 0
	let escaped = false
	for (const char of text) {
		if (escaped) {
			escaped = false
			continue
		}
		if (char === "'") {
			escaped = true
			continue
		}
		if (char === '[') balance += 1
		if (char === ']') balance -= 1
		if (balance < 0) return balance
	}
	return balance
}

function quotedNameClosed(text: string): boolean {
	if (!text.startsWith("'")) return true
	for (let index = 1; index < text.length; index++) {
		if (text[index] !== "'") continue
		if (text[index + 1] === "'") {
			index += 1
			continue
		}
		return index === text.length - 1
	}
	return false
}

function parseDiagnosticPosition(message: string): number | null {
	const match = /\bat position (-?\d+)/.exec(message)
	if (!match) return null
	const position = Number(match[1])
	return Number.isSafeInteger(position) && position >= 0 ? position : null
}
