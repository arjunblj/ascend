import { existsSync, readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { makeXlsx } from '../../packages/io-xlsx/test/helpers.ts'
import {
	AscendWorkbook,
	inspectWorkbookOpenPlan,
	type WorkbookOpenPlan,
} from '../../packages/sdk/src/index.ts'

export type SafeOpenProofCaseKind = 'file' | 'synthetic' | 'malformed'

export interface SafeOpenProofCase {
	readonly name: string
	readonly kind: SafeOpenProofCaseKind
	readonly fixture: string
	readonly bytes: Uint8Array
	readonly fullOpenExpected: boolean
	readonly expectedMode?: WorkbookOpenPlan['recommendedMode']
	readonly expectedReviewBeforeHydration?: boolean
	readonly expectedRiskFamilies?: readonly string[]
}

export interface SafeOpenProofOptions {
	readonly repeat?: number
	readonly warmup?: number
	readonly includeTimings?: boolean
}

export interface SafeOpenProofCaseResult {
	readonly name: string
	readonly kind: SafeOpenProofCaseKind
	readonly fixture: string
	readonly bytes: number
	readonly status: 'ok' | 'rejected'
	readonly recommendedMode?: WorkbookOpenPlan['recommendedMode']
	readonly reviewBeforeHydration?: boolean
	readonly riskFamilies: readonly string[]
	readonly partCount?: number
	readonly worksheetPartCount?: number
	readonly relationshipCount?: number
	readonly openPlanMedianMs?: number
	readonly fullOpenMedianMs?: number
	readonly fullOpenRatio?: number
	readonly boundary: string
}

export interface SafeOpenProofResult {
	readonly repeat: number
	readonly warmup: number
	readonly generatedAt: string
	readonly cases: readonly SafeOpenProofCaseResult[]
}

interface Timed<T> {
	readonly value: T
	readonly ms: number
}

const DEFAULT_REPEAT = 7
const DEFAULT_WARMUP = 2

export async function runSafeOpenProof(
	options: SafeOpenProofOptions = {},
): Promise<SafeOpenProofResult> {
	const repeat = positiveInteger(options.repeat, DEFAULT_REPEAT)
	const warmup = positiveInteger(options.warmup, DEFAULT_WARMUP)
	const includeTimings = options.includeTimings ?? true
	const cases = defaultSafeOpenProofCases()
	const results: SafeOpenProofCaseResult[] = []
	for (const proofCase of cases) {
		results.push(await runSafeOpenProofCase(proofCase, { repeat, warmup, includeTimings }))
	}
	return {
		repeat,
		warmup,
		generatedAt: new Date().toISOString(),
		cases: results,
	}
}

export function defaultSafeOpenProofCases(): SafeOpenProofCase[] {
	return [
		fileCase('clean', 'fixtures/xlsx/poi/SampleSS.xlsx', {
			expectedMode: 'formula',
			expectedReviewBeforeHydration: false,
		}),
		fileCase('formula-heavy', 'fixtures/xlsx/poi/formula_stress_test.xlsx', {
			expectedMode: 'formula',
			expectedReviewBeforeHydration: false,
		}),
		fileCase('macro', 'fixtures/xlsx/calamine/vba.xlsm', {
			expectedMode: 'metadata-only',
			expectedReviewBeforeHydration: true,
			expectedRiskFamilies: ['preservedMacro'],
		}),
		fileCase('pivot', 'fixtures/xlsx/poi/ExcelPivotTableSample.xlsx', {
			expectedMode: 'formula',
			expectedReviewBeforeHydration: false,
		}),
		fileCase('activex', 'fixtures/xlsx/libreoffice/activex_checkbox.xlsx', {
			expectedMode: 'metadata-only',
			expectedReviewBeforeHydration: true,
			expectedRiskFamilies: ['preservedActiveX'],
		}),
		fileCase('chart', 'fixtures/xlsx/poi/WithChart.xlsx', {
			expectedMode: 'formula',
			expectedReviewBeforeHydration: false,
		}),
		{
			name: 'signed',
			kind: 'synthetic',
			fixture: 'synthetic digital-signature package',
			bytes: signedWorkbook(),
			fullOpenExpected: true,
			expectedMode: 'metadata-only',
			expectedReviewBeforeHydration: true,
			expectedRiskFamilies: ['preservedSignature'],
		},
		{
			name: 'unknown-part',
			kind: 'synthetic',
			fixture: 'synthetic unknown package part',
			bytes: unknownPartWorkbook(),
			fullOpenExpected: true,
			expectedMode: 'metadata-only',
			expectedReviewBeforeHydration: true,
			expectedRiskFamilies: ['preservedOther'],
		},
		{
			name: 'malformed',
			kind: 'malformed',
			fixture: 'synthetic malformed bytes',
			bytes: new TextEncoder().encode('not a zip'),
			fullOpenExpected: false,
		},
	]
}

export function safeOpenProofMarkdown(result: SafeOpenProofResult): string {
	return [
		'# Safe Unknown Workbook Opening Proof',
		'',
		`Generated: ${result.generatedAt}`,
		`Samples: repeat ${result.repeat}, warmup ${result.warmup}`,
		'',
		'Boundary: this proves pre-hydration package-feature routing, not malware scanning, sandboxing, active-content execution, or malformed-package recovery.',
		'',
		'| Case | Fixture | Bytes | Status | Mode | Review before hydration | Risk families | Parts | Worksheets | Relationships | Median open-plan ms | Median full-open ms | Full/open-plan ratio | Boundary |',
		'| --- | --- | ---: | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |',
		...result.cases.map(markdownRow),
		'',
		'Allowed claim: Ascend can recommend a load mode and review branch from XLSX/XLSM package features before hydrating workbook cells.',
	].join('\n')
}

async function runSafeOpenProofCase(
	proofCase: SafeOpenProofCase,
	options: Required<SafeOpenProofOptions>,
): Promise<SafeOpenProofCaseResult> {
	const openPlanSamples: number[] = []
	const fullOpenSamples: number[] = []
	let plan: WorkbookOpenPlan | undefined
	let boundary = 'ok'
	let status: SafeOpenProofCaseResult['status'] = 'ok'
	const openPlanIterations = options.includeTimings ? options.warmup + options.repeat : 1
	for (let index = 0; index < openPlanIterations; index++) {
		try {
			const measured = timeMs(() =>
				inspectWorkbookOpenPlan(proofCase.bytes, { intent: 'edit-plan' }),
			)
			plan = measured.value
			if (index >= options.warmup && options.includeTimings) openPlanSamples.push(measured.ms)
		} catch (error) {
			status = 'rejected'
			boundary = `open-plan rejected: ${errorMessage(error)}`
			break
		}
	}

	if (status === 'ok') {
		assertExpectedPlan(proofCase, plan)
		const fullOpenIterations = options.includeTimings ? options.warmup + options.repeat : 0
		for (let index = 0; index < fullOpenIterations; index++) {
			const measured = await timeMsAsync(() =>
				AscendWorkbook.open(proofCase.bytes, { mode: 'full' }),
			)
			if (index >= options.warmup) fullOpenSamples.push(measured.ms)
		}
	} else if (proofCase.fullOpenExpected) {
		throw new Error(`${proofCase.name} open-plan rejected unexpectedly: ${boundary}`)
	}

	const riskFamilies = plan?.riskFeatures.map((feature) => feature.featureFamily) ?? []
	return {
		name: proofCase.name,
		kind: proofCase.kind,
		fixture: proofCase.fixture,
		bytes: proofCase.bytes.byteLength,
		status,
		...(plan
			? {
					recommendedMode: plan.recommendedMode,
					reviewBeforeHydration: plan.reviewBeforeHydration,
					partCount: plan.partCount,
					worksheetPartCount: plan.worksheetPartCount,
					relationshipCount: plan.relationshipCount,
				}
			: {}),
		riskFamilies,
		...(openPlanSamples.length > 0 ? { openPlanMedianMs: roundMs(median(openPlanSamples)) } : {}),
		...(fullOpenSamples.length > 0 ? { fullOpenMedianMs: roundMs(median(fullOpenSamples)) } : {}),
		...(openPlanSamples.length > 0 && fullOpenSamples.length > 0
			? {
					fullOpenRatio: roundRatio(
						median(fullOpenSamples) / Math.max(median(openPlanSamples), Number.EPSILON),
					),
				}
			: {}),
		boundary,
	}
}

function fileCase(
	name: string,
	fixture: string,
	options: Pick<
		SafeOpenProofCase,
		'expectedMode' | 'expectedReviewBeforeHydration' | 'expectedRiskFamilies'
	>,
): SafeOpenProofCase {
	if (!existsSync(fixture)) throw new Error(`Missing proof fixture ${fixture}`)
	return {
		name,
		kind: 'file',
		fixture,
		bytes: readFileSync(fixture),
		fullOpenExpected: true,
		...options,
	}
}

function assertExpectedPlan(
	proofCase: SafeOpenProofCase,
	plan: WorkbookOpenPlan | undefined,
): void {
	if (!plan) throw new Error(`${proofCase.name} did not produce an open plan`)
	if (proofCase.expectedMode !== undefined && plan.recommendedMode !== proofCase.expectedMode) {
		throw new Error(
			`${proofCase.name} expected mode ${proofCase.expectedMode}, got ${plan.recommendedMode}`,
		)
	}
	if (
		proofCase.expectedReviewBeforeHydration !== undefined &&
		plan.reviewBeforeHydration !== proofCase.expectedReviewBeforeHydration
	) {
		throw new Error(
			`${proofCase.name} expected reviewBeforeHydration ${proofCase.expectedReviewBeforeHydration}, got ${plan.reviewBeforeHydration}`,
		)
	}
	for (const expectedFamily of proofCase.expectedRiskFamilies ?? []) {
		if (!plan.riskFeatures.some((feature) => feature.featureFamily === expectedFamily)) {
			throw new Error(`${proofCase.name} missing risk family ${expectedFamily}`)
		}
	}
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
		'xl/_rels/workbook.xml.rels': relationships(`
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
`),
		'xl/workbook.xml': workbookXml('Signed'),
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
		'_rels/.rels': relationships(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
		'xl/_rels/workbook.xml.rels': relationships(`
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdCustom" Type="http://schemas.example.invalid/relationships/opaque" Target="custom/custom1.xml"/>
`),
		'xl/workbook.xml': workbookXml('Unknown'),
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

function relationships(extra: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${extra}
</Relationships>`
}

function workbookXml(sheetName: string): string {
	return `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="${sheetName}" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`
}

function worksheetXml(rows: string): string {
	return `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${rows}</sheetData>
</worksheet>`
}

function markdownRow(row: SafeOpenProofCaseResult): string {
	return [
		row.name,
		`\`${row.fixture}\``,
		String(row.bytes),
		row.status,
		row.recommendedMode ?? 'n/a',
		row.reviewBeforeHydration === undefined ? 'n/a' : String(row.reviewBeforeHydration),
		row.riskFamilies.length > 0 ? row.riskFamilies.join(', ') : 'none',
		row.partCount?.toString() ?? 'n/a',
		row.worksheetPartCount?.toString() ?? 'n/a',
		row.relationshipCount?.toString() ?? 'n/a',
		row.openPlanMedianMs?.toFixed(3) ?? 'n/a',
		row.fullOpenMedianMs?.toFixed(3) ?? 'n/a',
		row.fullOpenRatio?.toFixed(2) ?? 'n/a',
		row.boundary,
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

function positiveInteger(value: number | undefined, fallback: number): number {
	return Number.isInteger(value) && value !== undefined && value > 0 ? value : fallback
}

function timeMs<T>(fn: () => T): Timed<T> {
	const start = performance.now()
	const value = fn()
	return { value, ms: performance.now() - start }
}

async function timeMsAsync<T>(fn: () => Promise<T>): Promise<Timed<T>> {
	const start = performance.now()
	const value = await fn()
	return { value, ms: performance.now() - start }
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const mid = Math.floor(sorted.length / 2)
	const value =
		sorted.length % 2 === 0 ? ((sorted[mid - 1] ?? 0) + (sorted[mid] ?? 0)) / 2 : (sorted[mid] ?? 0)
	return value
}

function roundMs(value: number): number {
	return Number(value.toFixed(3))
}

function roundRatio(value: number): number {
	return Number(value.toFixed(2))
}

function errorMessage(error: unknown): string {
	return error instanceof Error ? error.message : String(error)
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	if (index === -1) return undefined
	return process.argv[index + 1]
}

if (import.meta.main) {
	const json = process.argv.includes('--json')
	const result = await runSafeOpenProof({
		repeat: Number(readFlag('--repeat')) || undefined,
		warmup: Number(readFlag('--warmup')) || undefined,
		includeTimings: !process.argv.includes('--no-timings'),
	})
	console.log(json ? JSON.stringify(result, null, 2) : safeOpenProofMarkdown(result))
	if (!json) {
		console.error(`Generated safe-open proof over ${result.cases.length} cases.`)
		console.error(`Run with --json for machine-readable output from ${basename(import.meta.path)}.`)
	}
}
