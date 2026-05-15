import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { makeXlsx } from '../../packages/io-xlsx/test/helpers.ts'
import {
	AscendWorkbook,
	commitAgentPlan,
	createAgentCommitPackageActionProof,
	createAgentPlan,
	createPackageActionProof,
	type Operation,
	type PackageActionKind,
	type PackageActionProof,
} from '../../packages/sdk/src/index.ts'

export interface PackageActionProofCase {
	readonly name: string
	readonly sourceKind: PackageActionProofSourceKind
	readonly fixture: string
	readonly ops: readonly Operation[]
	readonly allowLoss?: readonly string[] | 'all'
	readonly prepareInput: (path: string) => Promise<Uint8Array>
	readonly expectedCommitActions: readonly ExpectedPackageAction[]
}

export type PackageActionProofSourceKind =
	| 'public-fixture'
	| 'generated-workbook'
	| 'generated-edge-package'

export interface ExpectedPackageAction {
	readonly action: PackageActionKind
	readonly partPathIncludes?: string
}

export interface PackageActionProofOptions {
	readonly includeTimings?: boolean
}

export interface PackageActionProofCaseResult {
	readonly name: string
	readonly sourceKind: PackageActionProofSourceKind
	readonly fixture: string
	readonly inputBytes: number
	readonly inputSha256: string
	readonly outputBytes: number
	readonly planActionCounts: Readonly<Record<PackageActionKind, number>>
	readonly commitActionCounts: Readonly<Record<PackageActionKind, number>>
	readonly commitCoverage: PackageActionProof['coverage']
	readonly commitJournalExact: boolean | null
	readonly commitJournalPackageIssueCount: number
	readonly commitJournalPackageIssueRefs: readonly string[]
	readonly expectedActionsPresent: boolean
	readonly proofJsonBytes: number
	readonly proofMedianMs?: number
	readonly postWriteAuditsPassed: boolean
	readonly issueCount: number
	readonly exampleActions: readonly string[]
}

export interface PackageActionProofResult {
	readonly generatedAt: string
	readonly cases: readonly PackageActionProofCaseResult[]
	readonly combinedCommitActionCounts: Readonly<Record<PackageActionKind, number>>
}

const ACTIONS: readonly PackageActionKind[] = ['passthrough', 'regenerate', 'add', 'drop', 'error']

export async function runPackageActionProof(
	options: PackageActionProofOptions = {},
): Promise<PackageActionProofResult> {
	const cases = defaultPackageActionProofCases()
	const results: PackageActionProofCaseResult[] = []
	for (const proofCase of cases) results.push(await runPackageActionProofCase(proofCase, options))
	return {
		generatedAt: new Date().toISOString(),
		cases: results,
		combinedCommitActionCounts: combineCounts(results.map((result) => result.commitActionCounts)),
	}
}

export function defaultPackageActionProofCases(): PackageActionProofCase[] {
	return [
		{
			name: 'docprops-passthrough',
			sourceKind: 'generated-edge-package',
			fixture: 'synthetic docProps package',
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'docprops' }] }],
			prepareInput: writeBytes(docPropsWorkbook),
			expectedCommitActions: [{ action: 'passthrough', partPathIncludes: 'docProps/core.xml' }],
		},
		{
			name: 'regenerate-existing-sheet',
			sourceKind: 'generated-workbook',
			fixture: 'new Ascend workbook',
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'regen' }] }],
			prepareInput: writeNewWorkbook,
			expectedCommitActions: [
				{ action: 'regenerate', partPathIncludes: 'xl/worksheets/sheet1.xml' },
			],
		},
		{
			name: 'add-sheet-part',
			sourceKind: 'generated-workbook',
			fixture: 'new Ascend workbook',
			ops: [{ op: 'addSheet', name: 'Added' }],
			prepareInput: writeNewWorkbook,
			expectedCommitActions: [{ action: 'add', partPathIncludes: 'xl/worksheets/sheet' }],
		},
		{
			name: 'calc-chain-drop',
			sourceKind: 'generated-edge-package',
			fixture: 'synthetic calcChain package',
			ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1+A1' }],
			prepareInput: writeBytes(calcChainWorkbook),
			expectedCommitActions: [{ action: 'drop', partPathIncludes: 'xl/calcChain.xml' }],
		},
		{
			name: 'signature-invalidation-drop',
			sourceKind: 'generated-edge-package',
			fixture: 'synthetic digital-signature package',
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'signed' }] }],
			allowLoss: 'all',
			prepareInput: writeBytes(signedWorkbook),
			expectedCommitActions: [{ action: 'drop', partPathIncludes: '_xmlsignatures/' }],
		},
		{
			name: 'macro-passthrough',
			sourceKind: 'public-fixture',
			fixture: 'fixtures/xlsx/calamine/vba.xlsm',
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'macro' }] }],
			allowLoss: 'all',
			prepareInput: writeFixture('fixtures/xlsx/calamine/vba.xlsm'),
			expectedCommitActions: [{ action: 'passthrough', partPathIncludes: 'vbaProject.bin' }],
		},
		{
			name: 'chart-sidecar-accounting',
			sourceKind: 'public-fixture',
			fixture: 'fixtures/xlsx/poi/WithChart.xlsx',
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'chart' }] }],
			prepareInput: writeFixture('fixtures/xlsx/poi/WithChart.xlsx'),
			expectedCommitActions: [
				{ action: 'regenerate', partPathIncludes: 'charts/' },
				{ action: 'passthrough', partPathIncludes: 'drawings/' },
			],
		},
		{
			name: 'unknown-part-error',
			sourceKind: 'generated-edge-package',
			fixture: 'synthetic unknown package part',
			ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'unknown' }] }],
			allowLoss: 'all',
			prepareInput: writeBytes(unknownPartWorkbook),
			expectedCommitActions: [{ action: 'error', partPathIncludes: 'custom/custom1.xml' }],
		},
	]
}

export function packageActionProofMarkdown(result: PackageActionProofResult): string {
	return [
		'# Package Action Proof Report',
		'',
		`Generated: ${result.generatedAt}`,
		'Boundary: this is local package-part action evidence. It is not signed provenance, Excel recalculation equivalence, or a guarantee that unsupported package features are semantically understood.',
		'',
		'| Case | Fixture | Input bytes | Output bytes | Commit actions | Source graph | Digest pairs | Journal package issues | Proof issues | Proof JSON bytes | Proof ms | Expected action present | Post-write audits | Examples |',
		'| --- | --- | ---: | ---: | --- | --- | ---: | ---: | ---: | ---: | ---: | --- | --- | --- |',
		...result.cases.map(markdownRow),
		'',
		`Combined commit actions: ${formatCounts(result.combinedCommitActionCounts)}`,
	].join('\n')
}

async function runPackageActionProofCase(
	proofCase: PackageActionProofCase,
	options: PackageActionProofOptions,
): Promise<PackageActionProofCaseResult> {
	const dir = await mkdtemp(join(tmpdir(), `ascend-package-action-proof-${proofCase.name}-`))
	try {
		const input = join(dir, proofCase.name.endsWith('.xlsm') ? 'input.xlsm' : 'input.xlsx')
		const output = join(dir, 'output.xlsx')
		const inputBytes = await proofCase.prepareInput(input)
		const plan = await createAgentPlan(input, proofCase.ops)
		const planProof = createPackageActionProof(plan.preservation, {
			sourceBytes: inputBytes,
			writePolicy: plan.writePolicy,
			packageGraphAudit: plan.packageGraphAudit,
		})
		const committed = await commitAgentPlan(input, proofCase.ops, {
			output,
			approvals: 'all',
			...(proofCase.allowLoss ? { allowLoss: proofCase.allowLoss } : {}),
		})
		const outputBytes = readFileSync(output)
		const measured = measureProof(() => createAgentCommitPackageActionProof(committed), options)
		const commitProof = measured.value
		const journalPackageIssues = (committed.apply.journal?.issues ?? []).filter(
			(issue) => issue.surface === 'package-parts' && issue.reason === 'package-part-preservation',
		)
		assertExpectedActions(proofCase, commitProof)
		return {
			name: proofCase.name,
			sourceKind: proofCase.sourceKind,
			fixture: proofCase.fixture,
			inputBytes: inputBytes.byteLength,
			inputSha256: sha256Bytes(inputBytes),
			outputBytes: outputBytes.byteLength,
			planActionCounts: planProof.byAction,
			commitActionCounts: commitProof.byAction,
			commitCoverage: commitProof.coverage,
			commitJournalExact: committed.apply.journal?.exact ?? null,
			commitJournalPackageIssueCount: journalPackageIssues.length,
			commitJournalPackageIssueRefs: Array.from(
				new Set(journalPackageIssues.flatMap((issue) => issue.refs ?? [])),
			),
			expectedActionsPresent: true,
			proofJsonBytes: new TextEncoder().encode(JSON.stringify(commitProof)).byteLength,
			...(measured.ms !== undefined ? { proofMedianMs: measured.ms } : {}),
			postWriteAuditsPassed: committed.postWrite.auditsPassed,
			issueCount: commitProof.issues.length,
			exampleActions: exampleActions(commitProof),
		}
	} finally {
		await rm(dir, { recursive: true, force: true })
	}
}

function assertExpectedActions(proofCase: PackageActionProofCase, proof: PackageActionProof): void {
	for (const expected of proofCase.expectedCommitActions) {
		const found = proof.actions.some(
			(action) =>
				action.action === expected.action &&
				(expected.partPathIncludes === undefined ||
					action.partPath?.includes(expected.partPathIncludes)),
		)
		if (!found) {
			throw new Error(
				`${proofCase.name} missing ${expected.action} action for ${expected.partPathIncludes ?? 'any part'}`,
			)
		}
	}
}

function measureProof(
	fn: () => PackageActionProof,
	options: PackageActionProofOptions,
): { readonly value: PackageActionProof; readonly ms?: number } {
	if (options.includeTimings === false) return { value: fn() }
	const samples: number[] = []
	let value = fn()
	for (let index = 0; index < 5; index++) {
		const start = performance.now()
		value = fn()
		samples.push(performance.now() - start)
	}
	return { value, ms: roundMs(median(samples)) }
}

async function writeNewWorkbook(path: string): Promise<Uint8Array> {
	const wb = AscendWorkbook.create()
	await wb.save(path)
	return readFileSync(path)
}

function writeBytes(factory: () => Uint8Array): (path: string) => Promise<Uint8Array> {
	return async (path) => {
		const bytes = factory()
		await Bun.write(path, bytes)
		return bytes
	}
}

function writeFixture(fixture: string): (path: string) => Promise<Uint8Array> {
	return async (path) => {
		if (!existsSync(fixture)) throw new Error(`Missing package action proof fixture ${fixture}`)
		const bytes = readFileSync(fixture)
		await Bun.write(path, bytes)
		return bytes
	}
}

function docPropsWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': contentTypes(`
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
`),
		'_rels/.rels': relationships(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rIdApp" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
`),
		'xl/_rels/workbook.xml.rels': workbookRels(''),
		'xl/workbook.xml': workbookXml(),
		'xl/worksheets/sheet1.xml': worksheetXml(
			'<row r="1"><c r="A1" t="inlineStr"><is><t>source</t></is></c></row>',
		),
		'docProps/core.xml':
			'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>',
		'docProps/app.xml':
			'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Ascend Fixture</Application></Properties>',
		'docProps/custom.xml':
			'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"/>',
	})
}

function calcChainWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': contentTypes(`
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
`),
		'_rels/.rels': rootRels(),
		'xl/_rels/workbook.xml.rels': workbookRels(`
  <Relationship Id="rIdCalcChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
`),
		'xl/workbook.xml': workbookXml(),
		'xl/worksheets/sheet1.xml': worksheetXml(
			'<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1</f><v>1</v></c></row>',
		),
		'xl/calcChain.xml':
			'<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>',
	})
}

function signedWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': contentTypes(`
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/_xmlsignatures/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>
  <Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>
`),
		'_rels/.rels': relationships(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rIdSignatureOrigin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/>
`),
		'_xmlsignatures/_rels/origin.sigs.rels': relationships(`
  <Relationship Id="rIdSignature" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="sig1.xml"/>
`),
		'xl/_rels/workbook.xml.rels': workbookRels(''),
		'xl/workbook.xml': workbookXml(),
		'xl/worksheets/sheet1.xml': worksheetXml(''),
		'_xmlsignatures/origin.sigs': '',
		'_xmlsignatures/sig1.xml':
			'<?xml version="1.0"?><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>',
	})
}

function unknownPartWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': contentTypes(`
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/custom/custom1.xml" ContentType="application/vnd.example.opaque+xml"/>
`),
		'_rels/.rels': rootRels(),
		'xl/_rels/workbook.xml.rels': workbookRels(`
  <Relationship Id="rIdCustom" Type="http://schemas.example.invalid/relationships/opaque" Target="custom/custom1.xml"/>
`),
		'xl/workbook.xml': workbookXml(),
		'xl/worksheets/sheet1.xml': worksheetXml(''),
		'xl/custom/custom1.xml': '<opaque/>',
	})
}

function contentTypes(extra: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
${extra}
</Types>`
}

function rootRels(): string {
	return relationships(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`)
}

function workbookRels(extra: string): string {
	return relationships(`
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
${extra}
`)
}

function relationships(extra: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${extra}
</Relationships>`
}

function workbookXml(): string {
	return `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`
}

function worksheetXml(rows: string): string {
	return `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${rows}</sheetData>
</worksheet>`
}

function combineCounts(
	counts: readonly Readonly<Record<PackageActionKind, number>>[],
): Readonly<Record<PackageActionKind, number>> {
	const combined = emptyCounts()
	for (const entry of counts) {
		for (const action of ACTIONS) combined[action] += entry[action]
	}
	return combined
}

function emptyCounts(): Record<PackageActionKind, number> {
	return { passthrough: 0, regenerate: 0, add: 0, drop: 0, error: 0 }
}

function exampleActions(proof: PackageActionProof): string[] {
	const examples: string[] = []
	for (const action of ACTIONS) {
		const entry = proof.actions.find((candidate) => candidate.action === action)
		if (entry) examples.push(`${entry.action}:${entry.partPath ?? 'workbook'}`)
	}
	return examples
}

function formatCounts(counts: Readonly<Record<PackageActionKind, number>>): string {
	return ACTIONS.map((action) => `${action}=${counts[action]}`).join(', ')
}

function sha256Bytes(bytes: Uint8Array): string {
	return createHash('sha256').update(bytes).digest('hex')
}

function markdownRow(row: PackageActionProofCaseResult): string {
	return [
		row.name,
		`\`${row.fixture}\``,
		String(row.inputBytes),
		String(row.outputBytes),
		formatCounts(row.commitActionCounts),
		String(row.commitCoverage.sourceGraphIncluded),
		String(
			row.commitCoverage.matchingByteDigestCount + row.commitCoverage.mismatchedByteDigestCount,
		),
		String(row.commitJournalPackageIssueCount),
		String(row.issueCount),
		String(row.proofJsonBytes),
		row.proofMedianMs?.toFixed(3) ?? 'n/a',
		String(row.expectedActionsPresent),
		row.postWriteAuditsPassed ? 'passed' : 'needs review',
		row.exampleActions.join('; '),
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const mid = Math.floor(sorted.length / 2)
	return sorted.length % 2 === 0
		? ((sorted[mid - 1] ?? 0) + (sorted[mid] ?? 0)) / 2
		: (sorted[mid] ?? 0)
}

function roundMs(value: number): number {
	return Number(value.toFixed(3))
}

if (import.meta.main) {
	const result = await runPackageActionProof({
		includeTimings: !process.argv.includes('--no-timings'),
	})
	if (process.argv.includes('--json')) {
		console.log(JSON.stringify(result, null, 2))
	} else {
		console.log(packageActionProofMarkdown(result))
		console.error(`Generated package action proof over ${result.cases.length} cases.`)
	}
}
