import { createHash } from 'node:crypto'
import { readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { AscendWorkbook } from '@ascend/sdk'
import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'

const CORPUS_DIR = resolve(import.meta.dir, '../../research/excel-corpus')
const MANIFEST_PATH = resolve(CORPUS_DIR, 'manifest.json')
const REPORT_PATH = resolve(CORPUS_DIR, 'audit-report.json')
const PROBE_VALUE = '__ascend_probe__'

interface CorpusManifestEntry {
	readonly file: string
	readonly size_bytes: number
	readonly features: Record<string, boolean>
	readonly counts: Record<string, number>
}

interface PackageSummary {
	readonly workbookContentType?: string
	readonly partCount: number
	readonly families: Record<string, number>
}

interface WorkbookSemanticSummary {
	readonly sheetCount: number
	readonly tableCount: number
	readonly commentCount: number
	readonly conditionalFormatCount: number
	readonly dataValidationCount: number
	readonly imageCount: number
	readonly pivotTableCount: number
	readonly pivotCacheCount: number
	readonly slicerCount: number
	readonly slicerCacheCount: number
	readonly externalReferenceCount: number
	readonly compatibilityStatus: string
}

interface ProbeTarget {
	readonly sheet: string
	readonly ref: string
}

interface AuditResult {
	readonly file: string
	readonly sourceSha256: string
	readonly sourceBytes: number
	readonly noOpByteIdentical: boolean
	readonly probe: ProbeTarget
	readonly sourcePackage: PackageSummary
	readonly dirtyPackage: PackageSummary
	readonly sourceSemantic: WorkbookSemanticSummary
	readonly dirtySemantic: WorkbookSemanticSummary
	readonly packageRegressions: readonly string[]
	readonly semanticRegressions: readonly string[]
	readonly risk: 'low' | 'medium' | 'high'
}

async function main(): Promise<void> {
	const manifest = await loadManifest()
	const requestedFile = readFlagValue('--file')
	const selected = requestedFile
		? manifest.filter((entry) => entry.file === requestedFile)
		: manifest
	const results: AuditResult[] = []

	for (const entry of selected) {
		results.push(await auditEntry(entry))
		runGc()
	}

	if (process.argv.includes('--write')) {
		await writeFile(REPORT_PATH, `${JSON.stringify(results, null, 2)}\n`, 'utf-8')
	}

	if (process.argv.includes('--json')) {
		console.log(JSON.stringify(results, null, 2))
		return
	}

	renderSummary(results)
}

async function loadManifest(): Promise<CorpusManifestEntry[]> {
	const raw = await readFile(MANIFEST_PATH, 'utf-8')
	return JSON.parse(raw) as CorpusManifestEntry[]
}

async function auditEntry(entry: CorpusManifestEntry): Promise<AuditResult> {
	const sourcePath = resolve(CORPUS_DIR, entry.file)
	const sourceBytes = new Uint8Array(await readFile(sourcePath))
	const sourceSha256 = sha256(sourceBytes)
	const sourcePackage = summarizePackage(sourceBytes)
	const { sourceSemantic, noOpByteIdentical, dirtyBytes, probe } =
		await inspectAndBuildDirtyWorkbook(entry.file, sourceBytes)
	runGc()
	const dirtyPackage = summarizePackage(dirtyBytes)
	const dirtySemantic = await inspectDirtyWorkbook(dirtyBytes)
	runGc()

	const packageRegressions = diffPackageSummary(sourcePackage, dirtyPackage)
	const semanticRegressions = diffSemanticSummary(sourceSemantic, dirtySemantic)

	return {
		file: entry.file,
		sourceSha256,
		sourceBytes: sourceBytes.byteLength,
		noOpByteIdentical,
		probe,
		sourcePackage,
		dirtyPackage,
		sourceSemantic,
		dirtySemantic,
		packageRegressions,
		semanticRegressions,
		risk: classifyRisk(noOpByteIdentical, packageRegressions, semanticRegressions),
	}
}

async function inspectAndBuildDirtyWorkbook(
	file: string,
	sourceBytes: Uint8Array,
): Promise<{
	sourceSemantic: WorkbookSemanticSummary
	noOpByteIdentical: boolean
	dirtyBytes: Uint8Array
	probe: ProbeTarget
}> {
	const workbook = await AscendWorkbook.open(sourceBytes)
	const sourceSemantic = summarizeWorkbook(workbook)
	const noOpByteIdentical = sha256(workbook.toBytes()) === sha256(sourceBytes)
	const probe = pickProbeTarget(workbook)
	const apply = workbook.apply([
		{
			op: 'setCells',
			sheet: probe.sheet,
			updates: [{ ref: probe.ref, value: PROBE_VALUE }],
		},
	])
	if (apply.errors.length > 0) {
		throw new Error(`${file}: failed to apply probe edit: ${apply.errors[0]?.message}`)
	}
	if (apply.recalcRequired) {
		const recalc = workbook.recalc()
		if (recalc.errors.length > 0) {
			throw new Error(`${file}: recalc failed after probe edit`)
		}
	}
	return { sourceSemantic, noOpByteIdentical, dirtyBytes: workbook.toBytes(), probe }
}

async function inspectDirtyWorkbook(dirtyBytes: Uint8Array): Promise<WorkbookSemanticSummary> {
	const workbook = await AscendWorkbook.open(dirtyBytes)
	return summarizeWorkbook(workbook)
}

function summarizeWorkbook(workbook: AscendWorkbook): WorkbookSemanticSummary {
	const info = workbook.inspect()
	const totalTables = info.sheets.reduce((sum, sheet) => sum + (sheet.tableCount ?? 0), 0)
	return {
		sheetCount: info.sheetCount,
		tableCount: totalTables,
		commentCount: info.commentCount ?? 0,
		conditionalFormatCount: info.conditionalFormatCount ?? 0,
		dataValidationCount: info.dataValidationCount ?? 0,
		imageCount: info.imageCount ?? 0,
		pivotTableCount: info.pivotTableCount,
		pivotCacheCount: info.pivotCacheCount,
		slicerCount: info.slicerCount,
		slicerCacheCount: info.slicerCacheCount,
		externalReferenceCount: info.externalReferenceCount,
		compatibilityStatus: info.compatibility.status,
	}
}

function pickProbeTarget(workbook: AscendWorkbook): ProbeTarget {
	const rankedSheets = workbook.sheets
		.map((name) => workbook.sheet(name))
		.filter((sheet): sheet is NonNullable<typeof sheet> => sheet !== undefined)
		.sort((left, right) => scoreSheet(right) - scoreSheet(left))

	const sheet = rankedSheets[0]
	if (!sheet) throw new Error('Workbook has no sheets')
	const used = sheet.usedRange()
	if (!used) return { sheet: sheet.name, ref: 'A1' }
	return {
		sheet: sheet.name,
		ref: toA1(used.end.row + 1, used.end.col + 1),
	}
}

function scoreSheet(sheet: NonNullable<ReturnType<AscendWorkbook['sheet']>>): number {
	let score = 0
	if (sheet.state === 'visible') score += 10
	if (!sheet.protection?.sheet) score += 5
	if (!sheet.autoFilter) score += 1
	return score
}

function toA1(row: number, col: number): string {
	return `${columnLabel(col)}${row + 1}`
}

function columnLabel(col: number): string {
	let n = col
	let label = ''
	while (n >= 0) {
		label = String.fromCharCode(65 + (n % 26)) + label
		n = Math.floor(n / 26) - 1
	}
	return label
}

function summarizePackage(bytes: Uint8Array): PackageSummary {
	const archive = extractZip(bytes)
	const paths = [...archive.entries()].map((entry) => entry.path)
	const contentTypes = archive.readText('[Content_Types].xml') ?? ''
	const workbookContentType = readWorkbookContentType(contentTypes)
	return {
		workbookContentType,
		partCount: paths.length,
		families: {
			charts: countPaths(paths, /^xl\/(charts|chartEx)\//),
			drawings: countPaths(paths, /^xl\/drawings\/(?!.*\.vml$)/),
			vml: countPaths(paths, /^xl\/drawings\/.*\.vml$/),
			media: countPaths(paths, /^xl\/media\//),
			tables: countPaths(paths, /^xl\/tables\//),
			comments: countPaths(paths, /^xl\/comments\d+\.xml$/),
			threadedComments: countPaths(paths, /^xl\/threadedComments\//),
			pivotTables: countPaths(paths, /^xl\/pivotTables\//),
			pivotCaches: countPaths(paths, /^xl\/pivotCache/),
			slicers: countPaths(paths, /^xl\/slicers\//),
			slicerCaches: countPaths(paths, /^xl\/slicerCaches\//),
			macros: countPaths(paths, /^xl\/vbaProject/i),
			customXml: countPaths(paths, /^customXml\//),
			externalLinks: countPaths(paths, /^xl\/externalLinks\//),
			connections: countPaths(paths, /^xl\/connections\.xml$/),
			calcChain: countPaths(paths, /^xl\/calcChain\.xml$/),
		},
	}
}

function countPaths(paths: readonly string[], pattern: RegExp): number {
	return paths.filter((path) => pattern.test(path)).length
}

function readWorkbookContentType(xml: string): string | undefined {
	const match =
		/<Override\s+PartName="\/xl\/workbook\.xml"\s+ContentType="([^"]+)"/.exec(xml) ??
		/<Override\s+ContentType="([^"]+)"\s+PartName="\/xl\/workbook\.xml"/.exec(xml)
	return match?.[1]
}

function diffPackageSummary(source: PackageSummary, dirty: PackageSummary): string[] {
	const regressions: string[] = []
	if (source.workbookContentType && source.workbookContentType !== dirty.workbookContentType) {
		regressions.push(
			`workbook content type changed from ${source.workbookContentType} to ${dirty.workbookContentType ?? 'missing'}`,
		)
	}
	for (const family of Object.keys(source.families).sort()) {
		const before = source.families[family] ?? 0
		const after = dirty.families[family] ?? 0
		if (after < before) {
			regressions.push(`${family} parts dropped: ${before} -> ${after}`)
		}
	}
	return regressions
}

function diffSemanticSummary(
	source: WorkbookSemanticSummary,
	dirty: WorkbookSemanticSummary,
): string[] {
	const regressions: string[] = []
	const numericKeys = [
		'sheetCount',
		'tableCount',
		'commentCount',
		'conditionalFormatCount',
		'dataValidationCount',
		'imageCount',
		'pivotTableCount',
		'pivotCacheCount',
		'slicerCount',
		'slicerCacheCount',
		'externalReferenceCount',
	] as const
	for (const key of numericKeys) {
		if (dirty[key] < source[key]) {
			regressions.push(`${key} dropped: ${source[key]} -> ${dirty[key]}`)
		}
	}
	if (
		source.compatibilityStatus === 'clean' &&
		dirty.compatibilityStatus !== source.compatibilityStatus
	) {
		regressions.push(
			`compatibility status changed from ${source.compatibilityStatus} to ${dirty.compatibilityStatus}`,
		)
	}
	return regressions
}

function classifyRisk(
	noOpByteIdentical: boolean,
	packageRegressions: readonly string[],
	semanticRegressions: readonly string[],
): 'low' | 'medium' | 'high' {
	if (packageRegressions.length > 0 || semanticRegressions.length > 0) return 'high'
	if (!noOpByteIdentical) return 'medium'
	return 'low'
}

function renderSummary(results: readonly AuditResult[]): void {
	for (const result of results) {
		console.log(`${riskBadge(result.risk)} ${result.file}`)
		console.log(
			`  no-op=${result.noOpByteIdentical ? 'identical' : 'changed'} probe=${result.probe.sheet}!${result.probe.ref}`,
		)
		if (result.packageRegressions.length === 0 && result.semanticRegressions.length === 0) {
			console.log('  regressions: none detected')
			continue
		}
		for (const regression of [...result.packageRegressions, ...result.semanticRegressions]) {
			console.log(`  - ${regression}`)
		}
	}
}

function riskBadge(risk: AuditResult['risk']): string {
	switch (risk) {
		case 'high':
			return '[high]'
		case 'medium':
			return '[medium]'
		default:
			return '[low]'
	}
}

function sha256(bytes: Uint8Array): string {
	return createHash('sha256').update(bytes).digest('hex')
}

function readFlagValue(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	if (index === -1) return undefined
	return process.argv[index + 1]
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		// Best effort only.
	}
}

void main()
