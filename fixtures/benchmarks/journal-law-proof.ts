import { basename } from 'node:path'
import { AscendWorkbook, type Operation } from '../../packages/sdk/src/index.ts'

export interface JournalLawProofOptions {
	readonly seed?: number
	readonly exactCaseCount?: number
	readonly sequenceLength?: number
}

export interface JournalLawProofCaseResult {
	readonly name: string
	readonly kind: 'exact-sequence' | 'lossy-boundary'
	readonly operationCount: number
	readonly journalExact: boolean | null
	readonly inverseRestored: boolean | null
	readonly issueReasons: readonly string[]
	readonly passed: boolean
	readonly failure?: string
}

export interface JournalLawProofResult {
	readonly generatedAt: string
	readonly seed: number
	readonly exactCaseCount: number
	readonly sequenceLength: number
	readonly exactChecked: number
	readonly lossyChecked: number
	readonly failureCount: number
	readonly operationFamilies: Readonly<Record<string, number>>
	readonly issueReasons: Readonly<Record<string, number>>
	readonly cases: readonly JournalLawProofCaseResult[]
	readonly passed: boolean
}

interface JournalLawCase {
	readonly name: string
	readonly kind: JournalLawProofCaseResult['kind']
	readonly setup?: (workbook: AscendWorkbook) => void
	readonly ops: readonly Operation[]
	readonly expectedIssue?: readonly {
		readonly surface: string
		readonly reason: string
	}[]
}

interface OperationFactory {
	readonly family: string
	readonly build: (ordinal: number, seed: number) => Operation
}

const DEFAULT_SEED = 0x5eed_cafe
const DEFAULT_EXACT_CASE_COUNT = 48
const DEFAULT_SEQUENCE_LENGTH = 5

const EXACT_OPERATION_FACTORIES: readonly OperationFactory[] = [
	{
		family: 'setCells',
		build: (ordinal, seed) => ({
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: cellRef(ordinal, 0), value: `v-${seed % 997}` }],
		}),
	},
	{
		family: 'setFormula',
		build: (ordinal, seed) => ({
			op: 'setFormula',
			sheet: 'Sheet1',
			ref: cellRef(ordinal, 1),
			formula: `${cellRef(ordinal, 0)}+${(seed % 9) + 1}`,
		}),
	},
	{
		family: 'setComment',
		build: (ordinal) => ({
			op: 'setComment',
			sheet: 'Sheet1',
			ref: cellRef(ordinal, 4),
			text: `comment-${ordinal}`,
		}),
	},
	{
		family: 'setHyperlink',
		build: (ordinal) => ({
			op: 'setHyperlink',
			sheet: 'Sheet1',
			ref: cellRef(ordinal, 5),
			url: `https://example.com/${ordinal}`,
		}),
	},
	{
		family: 'freezePane',
		build: (ordinal) => ({
			op: 'freezePane',
			sheet: 'Sheet1',
			row: ordinal % 4,
			col: (ordinal + 1) % 4,
		}),
	},
	{
		family: 'setDataValidation',
		build: (ordinal) => ({
			op: 'setDataValidation',
			sheet: 'Sheet1',
			range: `${cellRef(ordinal, 6)}:${cellRef(ordinal, 6)}`,
			rule: {
				type: 'whole',
				operator: 'greaterThan',
				formula1: '0',
				allowBlank: true,
				showErrorMessage: true,
			},
		}),
	},
	{
		family: 'setConditionalFormat',
		build: (ordinal) => ({
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: `${cellRef(ordinal, 7)}:${cellRef(ordinal, 7)}`,
			rule: { type: 'expression', formula: `${cellRef(ordinal, 7)}>0`, priority: 1 },
		}),
	},
]

export function runJournalLawProof(options: JournalLawProofOptions = {}): JournalLawProofResult {
	const seed = options.seed ?? DEFAULT_SEED
	const exactCaseCount = positiveInteger(options.exactCaseCount, DEFAULT_EXACT_CASE_COUNT)
	const sequenceLength = positiveInteger(options.sequenceLength, DEFAULT_SEQUENCE_LENGTH)
	const cases = [
		...generatedExactCases(seed, exactCaseCount, sequenceLength),
		...metadataExactCases(),
		...metadataBoundaryCases(),
	]
	const operationFamilies: Record<string, number> = {}
	const issueReasons: Record<string, number> = {}
	const results: JournalLawProofCaseResult[] = []
	for (const entry of cases) {
		for (const op of entry.ops) operationFamilies[op.op] = (operationFamilies[op.op] ?? 0) + 1
		const result = runJournalLawCase(entry)
		for (const reason of result.issueReasons) issueReasons[reason] = (issueReasons[reason] ?? 0) + 1
		results.push(result)
	}
	const failureCount = results.filter((result) => !result.passed).length
	return {
		generatedAt: new Date().toISOString(),
		seed,
		exactCaseCount,
		sequenceLength,
		exactChecked: results.filter((result) => result.kind === 'exact-sequence').length,
		lossyChecked: results.filter((result) => result.kind === 'lossy-boundary').length,
		failureCount,
		operationFamilies,
		issueReasons,
		cases: results,
		passed: failureCount === 0,
	}
}

export function journalLawProofMarkdown(result: JournalLawProofResult): string {
	return [
		'# Journal Law Proof',
		'',
		`Generated: ${result.generatedAt}`,
		`Seed: ${result.seed}`,
		`Generated exact budget: ${result.exactCaseCount} x ${result.sequenceLength}`,
		`Exact law cases: ${result.exactChecked}`,
		`Lossy boundary cases: ${result.lossyChecked}`,
		`Failures: ${result.failureCount}`,
		'',
		'Boundary: deterministic generation exercises journal inverse laws and known lossy metadata selectors. It is not shrinkable property testing; add fast-check when dependency policy allows it.',
		'',
		`Operation families: ${formatCounts(result.operationFamilies)}`,
		`Issue reasons: ${formatCounts(result.issueReasons)}`,
		'',
		'| Case | Kind | Ops | Journal exact | Inverse restored | Issues | Passed | Failure |',
		'| --- | --- | ---: | --- | --- | --- | --- | --- |',
		...result.cases.map(markdownRow),
	].join('\n')
}

function runJournalLawCase(entry: JournalLawCase): JournalLawProofCaseResult {
	const wb = AscendWorkbook.create()
	entry.setup?.(wb)
	const before = workbookEvidence(wb)
	const changed = wb.apply(entry.ops, { journal: true })
	if (changed.errors.length > 0) {
		return failedCase(entry, `apply errors: ${JSON.stringify(changed.errors)}`)
	}
	const journal = changed.journal
	if (!journal?.supported) return failedCase(entry, 'missing or unsupported journal')
	const issueReasons = journal.issues.map(
		(issue) => `${issue.surface ?? '<unknown>'}:${issue.reason ?? '<unknown>'}`,
	)
	if (entry.kind === 'lossy-boundary') {
		const missingIssue = (entry.expectedIssue ?? []).find(
			(expected) => !issueReasons.includes(`${expected.surface}:${expected.reason}`),
		)
		return {
			name: entry.name,
			kind: entry.kind,
			operationCount: entry.ops.length,
			journalExact: journal.exact,
			inverseRestored: null,
			issueReasons,
			passed: journal.exact === false && missingIssue === undefined,
			...(journal.exact !== false
				? { failure: 'expected lossy journal but journal was exact' }
				: missingIssue
					? { failure: `missing issue ${missingIssue.surface}:${missingIssue.reason}` }
					: {}),
		}
	}
	if (!journal.exact) {
		return failedCase(entry, `expected exact journal, got issues ${issueReasons.join(', ')}`)
	}
	const undo = wb.apply(journal.inverseOps, { transaction: true })
	if (undo.errors.length > 0) {
		return failedCase(
			entry,
			`inverse apply errors: ${JSON.stringify(undo.errors)}`,
			true,
			issueReasons,
		)
	}
	const inverseRestored = workbookEvidence(wb) === before
	return {
		name: entry.name,
		kind: entry.kind,
		operationCount: entry.ops.length,
		journalExact: journal.exact,
		inverseRestored,
		issueReasons,
		passed: inverseRestored,
		...(inverseRestored ? {} : { failure: 'inverse did not restore workbook evidence' }),
	}
}

function failedCase(
	entry: JournalLawCase,
	failure: string,
	journalExact: boolean | null = null,
	issueReasons: readonly string[] = [],
): JournalLawProofCaseResult {
	return {
		name: entry.name,
		kind: entry.kind,
		operationCount: entry.ops.length,
		journalExact,
		inverseRestored: null,
		issueReasons,
		passed: false,
		failure,
	}
}

function generatedExactCases(
	seed: number,
	exactCaseCount: number,
	sequenceLength: number,
): JournalLawCase[] {
	const cases: JournalLawCase[] = []
	let state = seed >>> 0
	for (let caseIndex = 0; caseIndex < exactCaseCount; caseIndex++) {
		const ops: Operation[] = []
		for (let step = 0; step < sequenceLength; step++) {
			state = nextSeed(state)
			const factory =
				EXACT_OPERATION_FACTORIES[
					(seed + caseIndex * sequenceLength + step) % EXACT_OPERATION_FACTORIES.length
				]
			if (!factory) throw new Error('missing operation factory')
			ops.push(factory.build(caseIndex * sequenceLength + step, state))
		}
		cases.push({ name: `generated-exact-${caseIndex + 1}`, kind: 'exact-sequence', ops })
	}
	return cases
}

function metadataBoundaryCases(): JournalLawCase[] {
	return [
		{
			name: 'data-validation-non-suffix-delete',
			kind: 'lossy-boundary',
			setup: setupOrderedDataValidations,
			ops: [{ op: 'deleteDataValidation', sheet: 'Sheet1', range: 'A1:A1' }],
			expectedIssue: [{ surface: 'data-validations', reason: 'metadata-order' }],
		},
		{
			name: 'data-validation-duplicate-delete',
			kind: 'lossy-boundary',
			setup: (wb) => {
				const sheet = wb.getWorkbookModel().getSheet('Sheet1')
				if (!sheet) throw new Error('missing Sheet1')
				sheet.dataValidations.push(
					{ sqref: 'A1:A1', type: 'whole', formula1: '1', allowBlank: true },
					{ sqref: 'A1:A1', type: 'whole', formula1: '2', allowBlank: true },
				)
			},
			ops: [{ op: 'deleteDataValidation', sheet: 'Sheet1', range: 'A1:A1' }],
			expectedIssue: [{ surface: 'data-validations', reason: 'metadata-duplicate' }],
		},
		{
			name: 'conditional-format-non-tail-delete',
			kind: 'lossy-boundary',
			setup: setupOrderedConditionalFormats,
			ops: [{ op: 'deleteConditionalFormat', sheet: 'Sheet1', range: 'A1:A1' }],
			expectedIssue: [{ surface: 'conditional-formats', reason: 'metadata-order' }],
		},
		{
			name: 'conditional-format-non-tail-replace',
			kind: 'lossy-boundary',
			setup: setupOrderedConditionalFormats,
			ops: [
				{
					op: 'setConditionalFormat',
					sheet: 'Sheet1',
					range: 'A1:A1',
					rule: { type: 'expression', formula: 'A1>5', priority: 1 },
				},
			],
			expectedIssue: [{ surface: 'conditional-formats', reason: 'metadata-order' }],
		},
		{
			name: 'conditional-format-duplicate-delete',
			kind: 'lossy-boundary',
			setup: (wb) => {
				const sheet = wb.getWorkbookModel().getSheet('Sheet1')
				if (!sheet) throw new Error('missing Sheet1')
				sheet.conditionalFormats.push(
					{ sqref: 'A1:A1', rules: [{ type: 'expression', formulas: ['A1>0'] }] },
					{ sqref: 'A1:A1', rules: [{ type: 'expression', formulas: ['A1<0'] }] },
				)
			},
			ops: [{ op: 'deleteConditionalFormat', sheet: 'Sheet1', range: 'A1:A1' }],
			expectedIssue: [{ surface: 'conditional-formats', reason: 'metadata-duplicate' }],
		},
	]
}

function metadataExactCases(): JournalLawCase[] {
	return [
		{
			name: 'existing-row-layout-replacement',
			kind: 'exact-sequence',
			setup: (wb) => {
				applyOrThrow(wb, [
					{ op: 'setRowHeight', sheet: 'Sheet1', row: 1, height: 20 },
					{ op: 'hideRows', sheet: 'Sheet1', at: 1, count: 1, hidden: true },
				])
			},
			ops: [
				{ op: 'setRowHeight', sheet: 'Sheet1', row: 1, height: 25 },
				{ op: 'hideRows', sheet: 'Sheet1', at: 1, count: 1, hidden: false },
			],
		},
		{
			name: 'existing-column-layout-replacement',
			kind: 'exact-sequence',
			setup: (wb) => {
				applyOrThrow(wb, [
					{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 12 },
					{ op: 'hideCols', sheet: 'Sheet1', at: 1, count: 1, hidden: true },
				])
			},
			ops: [
				{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 18 },
				{ op: 'hideCols', sheet: 'Sheet1', at: 1, count: 1, hidden: false },
			],
		},
		{
			name: 'sheet-protection-replacement',
			kind: 'exact-sequence',
			setup: (wb) => {
				applyOrThrow(wb, [
					{
						op: 'setSheetProtection',
						sheet: 'Sheet1',
						password: 'before',
						options: { selectLockedCells: false },
					},
				])
			},
			ops: [
				{
					op: 'setSheetProtection',
					sheet: 'Sheet1',
					password: 'after',
					options: { selectUnlockedCells: false },
				},
			],
		},
		{
			name: 'tab-color-replacement',
			kind: 'exact-sequence',
			setup: (wb) => {
				applyOrThrow(wb, [{ op: 'setTabColor', sheet: 'Sheet1', color: 'FF0000' }])
			},
			ops: [{ op: 'setTabColor', sheet: 'Sheet1', color: '00FF00' }],
		},
		{
			name: 'page-setup-print-area-replacement',
			kind: 'exact-sequence',
			setup: (wb) => {
				applyOrThrow(wb, [
					{
						op: 'setPageSetup',
						sheet: 'Sheet1',
						setup: { orientation: 'portrait', paperSize: 9 },
					},
					{ op: 'setPrintArea', sheet: 'Sheet1', range: 'A1:B2' },
				])
			},
			ops: [
				{
					op: 'setPageSetup',
					sheet: 'Sheet1',
					setup: { orientation: 'landscape', paperSize: 9 },
				},
				{ op: 'setPrintArea', sheet: 'Sheet1', range: 'C1:D2' },
			],
		},
	]
}

function setupOrderedDataValidations(wb: AscendWorkbook): void {
	applyOrThrow(wb, [
		{
			op: 'setDataValidation',
			sheet: 'Sheet1',
			range: 'A1:A1',
			rule: { type: 'whole', formula1: '1', allowBlank: true, showErrorMessage: true },
		},
		{
			op: 'setDataValidation',
			sheet: 'Sheet1',
			range: 'B1:B1',
			rule: { type: 'whole', formula1: '2', allowBlank: true, showErrorMessage: true },
		},
	])
}

function setupOrderedConditionalFormats(wb: AscendWorkbook): void {
	applyOrThrow(wb, [
		{
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A1',
			rule: { type: 'expression', formula: 'A1>0', priority: 1 },
		},
		{
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'B1:B1',
			rule: { type: 'expression', formula: 'B1>0', priority: 2 },
		},
	])
}

function applyOrThrow(wb: AscendWorkbook, ops: readonly Operation[]): void {
	const result = wb.apply(ops)
	if (result.errors.length > 0) throw new Error(`setup failed: ${JSON.stringify(result.errors)}`)
}

function workbookEvidence(wb: AscendWorkbook): string {
	const { timestamp: _timestamp, ...snapshot } = wb.snapshot()
	const model = wb.getWorkbookModel()
	return stableJson({
		snapshot,
		styles: Array.from({ length: model.styles.size }, (_, index) => model.styles.get(index)),
		workbookProperties: model.workbookProperties,
		documentProperties: wb.inspect().documentProperties,
		workbookViews: wb.inspect().workbookViews,
		calcSettings: model.calcSettings,
		themeSummary: wb.inspect().themeSummary,
		sheets: model.sheets.map((sheet) => ({
			name: sheet.name,
			merges: sheet.merges,
			dataValidations: sheet.dataValidations,
			conditionalFormats: sheet.conditionalFormats,
			comments: [...sheet.comments.entries()],
			hyperlinks: [...sheet.hyperlinks.entries()],
			tabColor: sheet.tabColor,
			protection: sheet.protection,
			state: sheet.state,
			rowHeights: [...sheet.rowHeights.entries()],
			colWidths: [...sheet.colWidths.entries()],
			rowDefs: [...sheet.rowDefs.entries()],
			colDefs: sheet.colDefs,
			pageSetup: sheet.pageSetup,
			printArea: sheet.printArea,
			frozenRows: sheet.frozenRows,
			frozenCols: sheet.frozenCols,
		})),
	})
}

function cellRef(ordinal: number, offset: number): string {
	const row = Math.floor((ordinal + offset) / 10) + 1
	const col = (ordinal + offset) % 10
	return `${columnName(col)}${row}`
}

function columnName(index: number): string {
	let value = index + 1
	let name = ''
	while (value > 0) {
		const remainder = (value - 1) % 26
		name = String.fromCharCode(65 + remainder) + name
		value = Math.floor((value - 1) / 26)
	}
	return name
}

function nextSeed(value: number): number {
	return (Math.imul(value, 1664525) + 1013904223) >>> 0
}

function positiveInteger(value: number | undefined, fallback: number): number {
	return Number.isInteger(value) && value !== undefined && value > 0 ? value : fallback
}

function stableJson(value: unknown): string {
	if (Array.isArray(value)) return `[${value.map(stableJson).join(',')}]`
	if (value && typeof value === 'object') {
		return `{${Object.entries(value as Record<string, unknown>)
			.sort(([left], [right]) => left.localeCompare(right))
			.map(([key, entry]) => `${JSON.stringify(key)}:${stableJson(entry)}`)
			.join(',')}}`
	}
	return JSON.stringify(value)
}

function formatCounts(counts: Readonly<Record<string, number>>): string {
	const entries = Object.entries(counts).sort(([left], [right]) => left.localeCompare(right))
	return entries.length === 0 ? 'none' : entries.map(([key, value]) => `${key}=${value}`).join(', ')
}

function markdownRow(row: JournalLawProofCaseResult): string {
	return [
		row.name,
		row.kind,
		String(row.operationCount),
		row.journalExact === null ? 'n/a' : String(row.journalExact),
		row.inverseRestored === null ? 'n/a' : String(row.inverseRestored),
		row.issueReasons.length > 0 ? row.issueReasons.join(', ') : 'none',
		String(row.passed),
		row.failure ?? '',
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function readNumberFlag(name: string): number | undefined {
	const index = process.argv.indexOf(name)
	if (index === -1) return undefined
	const raw = process.argv[index + 1]
	if (raw === undefined) return undefined
	const parsed = Number(raw)
	return Number.isFinite(parsed) ? parsed : undefined
}

if (import.meta.main) {
	const result = runJournalLawProof({
		seed: readNumberFlag('--seed'),
		exactCaseCount: readNumberFlag('--exact-cases'),
		sequenceLength: readNumberFlag('--sequence-length'),
	})
	if (process.argv.includes('--json')) {
		console.log(JSON.stringify(result, null, 2))
	} else {
		console.log(journalLawProofMarkdown(result))
		console.error(`Generated journal law proof over ${result.cases.length} cases.`)
		console.error(`Run with --json for machine-readable output from ${basename(import.meta.path)}.`)
	}
	if (!result.passed) process.exitCode = 1
}
