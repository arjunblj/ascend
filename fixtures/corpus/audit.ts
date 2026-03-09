import { createHash } from 'node:crypto'
import { readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { AscendWorkbook } from '@ascend/sdk'
import { readXlsx } from '../../packages/io-xlsx/src/reader/index.ts'
import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'
import {
	type CorpusAssertionClass,
	type CorpusBenchmarkTier,
	type CorpusManifestEntry,
	type CorpusRiskClass,
	type NormalizedCorpusManifestEntry,
	normalizeManifest,
	selectManifestEntries,
} from './manifest.ts'

const CORPUS_DIR = resolve(import.meta.dir, '../../research/excel-corpus')
const MANIFEST_PATH = resolve(CORPUS_DIR, 'manifest.json')
const REPORT_PATH = resolve(CORPUS_DIR, 'audit-report.json')
const PROBE_VALUE = '__ascend_probe__'

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
	readonly hyperlinkCount: number
	readonly ignoredErrorCount: number
	readonly mergeCount: number
	readonly definedNameCount: number
	readonly workbookViewCount: number
	readonly hasWorkbookProtection: boolean
	readonly pivotTableCount: number
	readonly pivotCacheCount: number
	readonly slicerCount: number
	readonly slicerCacheCount: number
	readonly externalReferenceCount: number
	readonly compatibilityStatus: string
	readonly compatibilityFeatures: Record<string, { tier: string; count: number }>
	readonly styleSummary: {
		readonly numFmtCount: number
		readonly fontCount: number
		readonly fillCount: number
		readonly borderCount: number
		readonly cellXfCount: number
		readonly dxfCount: number
		readonly tableStyleCount: number
	}
	readonly themeSummary: {
		readonly hasThemePart: boolean
		readonly colorCount: number
	}
	readonly sheets: Array<{
		readonly name: string
		readonly cellCount: number
		readonly tableCount: number
		readonly commentCount: number
		readonly conditionalFormatCount: number
		readonly dataValidationCount: number
		readonly imageCount: number
		readonly hyperlinkCount: number
		readonly ignoredErrorCount: number
		readonly hasAutoFilter: boolean
		readonly hasDrawingRefs: boolean
		readonly hasProtection: boolean
		readonly hasPageMetadata: boolean
		readonly hasFrozenPanes: boolean
	}>
}

interface ProbeTarget {
	readonly sheet: string
	readonly ref: string
}

interface AuditResult {
	readonly file: string
	readonly benchmarkTier: CorpusBenchmarkTier
	readonly assertionClass: CorpusAssertionClass
	readonly riskClass: CorpusRiskClass
	readonly featureTags: readonly string[]
	readonly vendorable: boolean
	readonly writePlanSummary: {
		readonly totalParts: number
		readonly byOrigin: Readonly<{
			generated: number
			'preserved-inline': number
			'preserved-source': number
			capsule: number
		}>
		readonly byOwnerKind: Readonly<{
			package: number
			workbook: number
			sheet: number
		}>
		readonly sheetPartCounts: Readonly<Record<string, number>>
	}
	readonly sourceSha256: string
	readonly sourceBytes: number
	readonly noOpByteIdentical: boolean
	readonly probe: ProbeTarget
	readonly sourcePackage: PackageSummary
	readonly dirtyPackage: PackageSummary
	readonly sourceSemantic: WorkbookSemanticSummary
	readonly dirtySemantic: WorkbookSemanticSummary
	readonly probeValuePersisted: boolean
	readonly probeRecalcError?: string
	readonly packageRegressions: readonly string[]
	readonly semanticRegressions: readonly string[]
	readonly risk: 'low' | 'medium' | 'high'
}

async function main(): Promise<void> {
	const manifest = await loadManifest()
	const requestedFile = readFlagValue('--file')
	const tags = readFlagValues('--tag')
	const tiers = readFlagValues('--tier') as CorpusBenchmarkTier[]
	const risks = readFlagValues('--risk') as CorpusRiskClass[]
	const assertionClasses = readFlagValues('--assertion-class') as CorpusAssertionClass[]
	const selected = selectManifestEntries(manifest, {
		...(requestedFile ? { file: requestedFile } : {}),
		...(tags.length > 0 ? { tags } : {}),
		...(tiers.length > 0 ? { tiers } : {}),
		...(risks.length > 0 ? { risks } : {}),
		...(assertionClasses.length > 0 ? { assertionClasses } : {}),
		...(process.argv.includes('--vendorable-only') ? { vendorableOnly: true } : {}),
	})
	if (selected.length === 0) {
		throw new Error('No corpus entries matched the requested audit filters')
	}
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

async function loadManifest(): Promise<readonly NormalizedCorpusManifestEntry[]> {
	const raw = await readFile(MANIFEST_PATH, 'utf-8')
	return normalizeManifest(JSON.parse(raw) as CorpusManifestEntry[])
}

async function auditEntry(entry: NormalizedCorpusManifestEntry): Promise<AuditResult> {
	const sourcePath = resolve(CORPUS_DIR, entry.file)
	const sourceBytes = new Uint8Array(await readFile(sourcePath))
	const sourceSha256 = sha256(sourceBytes)
	const sourcePackage = summarizePackage(sourceBytes)
	const {
		sourceSemantic,
		noOpByteIdentical,
		dirtyBytes,
		probe,
		probeRecalcError,
		writePlanSummary,
	} = await inspectAndBuildDirtyWorkbook(entry.file, sourceBytes)
	runGc()
	const dirtyPackage = summarizePackage(dirtyBytes)
	const { summary: dirtySemantic, probeValuePersisted } = await inspectDirtyWorkbook(
		dirtyBytes,
		probe,
	)
	runGc()

	const packageRegressions = diffPackageSummary(sourcePackage, dirtyPackage)
	const semanticRegressions = diffSemanticSummary(sourceSemantic, dirtySemantic)

	return {
		file: entry.file,
		benchmarkTier: entry.benchmarkTier,
		assertionClass: entry.assertionClass,
		riskClass: entry.riskClass,
		featureTags: entry.featureTags,
		vendorable: entry.vendorable,
		writePlanSummary,
		sourceSha256,
		sourceBytes: sourceBytes.byteLength,
		noOpByteIdentical,
		probe,
		sourcePackage,
		dirtyPackage,
		sourceSemantic,
		dirtySemantic,
		probeValuePersisted,
		...(probeRecalcError ? { probeRecalcError } : {}),
		packageRegressions,
		semanticRegressions,
		risk: classifyRisk(
			noOpByteIdentical,
			probeValuePersisted,
			probeRecalcError,
			packageRegressions,
			semanticRegressions,
		),
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
	probeRecalcError?: string
	writePlanSummary: AuditResult['writePlanSummary']
}> {
	const workbook = await AscendWorkbook.open(sourceBytes)
	const raw = readXlsx(sourceBytes)
	if (!raw.ok) throw new Error(`${file}: failed to read source workbook for audit summary`)
	const sourceSemantic = summarizeWorkbook(workbook, raw.value.report)
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
			return {
				sourceSemantic,
				noOpByteIdentical,
				dirtyBytes: workbook.toBytes(),
				probe,
				probeRecalcError: `${file}: recalc failed after probe edit`,
				writePlanSummary: workbook.writePlanSummary(),
			}
		}
	}
	return {
		sourceSemantic,
		noOpByteIdentical,
		dirtyBytes: workbook.toBytes(),
		probe,
		writePlanSummary: workbook.writePlanSummary(),
	}
}

async function inspectDirtyWorkbook(
	dirtyBytes: Uint8Array,
	probe: ProbeTarget,
): Promise<{ summary: WorkbookSemanticSummary; probeValuePersisted: boolean }> {
	const workbook = await AscendWorkbook.open(dirtyBytes)
	const raw = readXlsx(dirtyBytes)
	if (!raw.ok) throw new Error('Failed to read dirty workbook for audit summary')
	const probeValuePersisted =
		workbook.sheet(probe.sheet)?.cell(probe.ref)?.value.kind === 'string' &&
		workbook.sheet(probe.sheet)?.cell(probe.ref)?.value.value === PROBE_VALUE
	return { summary: summarizeWorkbook(workbook, raw.value.report), probeValuePersisted }
}

function summarizeWorkbook(
	workbook: AscendWorkbook,
	report: { status: string; features: readonly { feature: string; tier: string; count: number }[] },
): WorkbookSemanticSummary {
	const info = workbook.inspect()
	const totalTables = info.sheets.reduce((sum, sheet) => sum + (sheet.tableCount ?? 0), 0)
	const totalHyperlinks = info.sheets.reduce((sum, sheet) => sum + (sheet.hyperlinkCount ?? 0), 0)
	const totalIgnoredErrors = info.sheets.reduce(
		(sum, sheet) => sum + (sheet.ignoredErrorCount ?? 0),
		0,
	)
	const totalMerges = workbook.sheets.reduce(
		(sum, name) => sum + (workbook.sheet(name)?.merges.length ?? 0),
		0,
	)
	const compatibilityFeatures = Object.fromEntries(
		report.features.map((feature) => [
			feature.feature,
			{ tier: feature.tier, count: feature.count },
		]),
	)
	return {
		sheetCount: info.sheetCount,
		tableCount: totalTables,
		commentCount: info.commentCount ?? 0,
		conditionalFormatCount: info.conditionalFormatCount ?? 0,
		dataValidationCount: info.dataValidationCount ?? 0,
		imageCount: info.imageCount ?? 0,
		hyperlinkCount: totalHyperlinks,
		ignoredErrorCount: totalIgnoredErrors,
		mergeCount: totalMerges,
		definedNameCount: info.definedNames.length,
		workbookViewCount: info.workbookViewCount,
		hasWorkbookProtection: info.hasWorkbookProtection,
		pivotTableCount: info.pivotTableCount,
		pivotCacheCount: info.pivotCacheCount,
		slicerCount: info.slicerCount,
		slicerCacheCount: info.slicerCacheCount,
		externalReferenceCount: info.externalReferenceCount,
		compatibilityStatus: report.status,
		compatibilityFeatures,
		styleSummary: info.styleSummary,
		themeSummary: {
			hasThemePart: info.themeSummary.hasThemePart,
			colorCount: info.themeSummary.colorCount,
		},
		sheets: info.sheets.map((sheet) => ({
			name: sheet.name,
			cellCount: sheet.cellCount ?? 0,
			tableCount: sheet.tableCount ?? 0,
			commentCount: sheet.commentCount ?? 0,
			conditionalFormatCount: sheet.conditionalFormatCount ?? 0,
			dataValidationCount: sheet.dataValidationCount ?? 0,
			imageCount: sheet.imageCount ?? 0,
			hyperlinkCount: sheet.hyperlinkCount ?? 0,
			ignoredErrorCount: sheet.ignoredErrorCount ?? 0,
			hasAutoFilter: sheet.hasAutoFilter ?? false,
			hasDrawingRefs: sheet.hasDrawingRefs ?? false,
			hasProtection: sheet.hasProtection ?? false,
			hasPageMetadata: sheet.hasPageMetadata ?? false,
			hasFrozenPanes: sheet.hasFrozenPanes ?? false,
		})),
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
		'hyperlinkCount',
		'ignoredErrorCount',
		'mergeCount',
		'definedNameCount',
		'workbookViewCount',
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
	if (dirty.hasWorkbookProtection !== source.hasWorkbookProtection) {
		regressions.push(
			`workbook protection changed: ${source.hasWorkbookProtection} -> ${dirty.hasWorkbookProtection}`,
		)
	}
	for (const key of Object.keys(source.compatibilityFeatures).sort()) {
		const before = source.compatibilityFeatures[key]
		const after = dirty.compatibilityFeatures[key]
		if (!before) continue
		if (!after) {
			regressions.push(`compatibility feature missing after edit: ${key}`)
			continue
		}
		if (before.tier !== 'unsupported' && after.count < before.count) {
			regressions.push(
				`compatibility feature count dropped for ${key}: ${before.count} -> ${after.count}`,
			)
		}
		const tierRank = { exact: 3, normalized: 2, preserved: 1, unsupported: 0 } as const
		if (
			(tierRank[after.tier as keyof typeof tierRank] ?? -1) <
			(tierRank[before.tier as keyof typeof tierRank] ?? -1)
		) {
			regressions.push(
				`compatibility feature tier worsened for ${key}: ${before.tier} -> ${after.tier}`,
			)
		}
	}
	for (const key of Object.keys(source.styleSummary) as Array<
		keyof WorkbookSemanticSummary['styleSummary']
	>) {
		if (dirty.styleSummary[key] < source.styleSummary[key]) {
			regressions.push(
				`style summary dropped for ${key}: ${source.styleSummary[key]} -> ${dirty.styleSummary[key]}`,
			)
		}
	}
	if (dirty.themeSummary.hasThemePart !== source.themeSummary.hasThemePart) {
		regressions.push(
			`theme presence changed: ${source.themeSummary.hasThemePart} -> ${dirty.themeSummary.hasThemePart}`,
		)
	}
	if (dirty.themeSummary.colorCount < source.themeSummary.colorCount) {
		regressions.push(
			`theme color count dropped: ${source.themeSummary.colorCount} -> ${dirty.themeSummary.colorCount}`,
		)
	}
	for (const beforeSheet of source.sheets) {
		const afterSheet = dirty.sheets.find((sheet) => sheet.name === beforeSheet.name)
		if (!afterSheet) {
			regressions.push(`sheet missing after edit: ${beforeSheet.name}`)
			continue
		}
		const keys: Array<keyof typeof beforeSheet> = [
			'cellCount',
			'tableCount',
			'commentCount',
			'conditionalFormatCount',
			'dataValidationCount',
			'imageCount',
			'hyperlinkCount',
			'ignoredErrorCount',
		]
		for (const key of keys) {
			if (afterSheet[key] < beforeSheet[key]) {
				regressions.push(
					`sheet ${beforeSheet.name} ${key} dropped: ${beforeSheet[key]} -> ${afterSheet[key]}`,
				)
			}
		}
		const flagKeys: Array<
			'hasAutoFilter' | 'hasDrawingRefs' | 'hasProtection' | 'hasPageMetadata' | 'hasFrozenPanes'
		> = ['hasAutoFilter', 'hasDrawingRefs', 'hasProtection', 'hasPageMetadata', 'hasFrozenPanes']
		for (const key of flagKeys) {
			if (beforeSheet[key] && !afterSheet[key]) {
				regressions.push(`sheet ${beforeSheet.name} lost ${key}`)
			}
		}
	}
	return regressions
}

function classifyRisk(
	noOpByteIdentical: boolean,
	probeValuePersisted: boolean,
	probeRecalcError: string | undefined,
	packageRegressions: readonly string[],
	semanticRegressions: readonly string[],
): 'low' | 'medium' | 'high' {
	if (probeRecalcError) return 'high'
	if (!probeValuePersisted) return 'high'
	if (packageRegressions.length > 0 || semanticRegressions.length > 0) return 'high'
	if (!noOpByteIdentical) return 'medium'
	return 'low'
}

function renderSummary(results: readonly AuditResult[]): void {
	const byTier = summarizeBy(results, (result) => result.benchmarkTier)
	const byRisk = summarizeBy(results, (result) => result.riskClass)
	console.log(
		`summary: files=${results.length} tiers=${renderCounts(byTier)} risks=${renderCounts(byRisk)}`,
	)
	for (const result of results) {
		console.log(`${riskBadge(result.risk)} ${result.file}`)
		console.log(
			`  tier=${result.benchmarkTier} risk=${result.riskClass} assertion=${result.assertionClass} tags=${result.featureTags.join(', ')}`,
		)
		console.log(
			`  no-op=${result.noOpByteIdentical ? 'identical' : 'changed'} probe=${result.probe.sheet}!${result.probe.ref} persisted=${result.probeValuePersisted ? 'yes' : 'no'} vendorable=${result.vendorable ? 'yes' : 'no'}`,
		)
		console.log(
			`  write-plan: total=${result.writePlanSummary.totalParts} generated=${result.writePlanSummary.byOrigin.generated} preserved-inline=${result.writePlanSummary.byOrigin['preserved-inline']} preserved-source=${result.writePlanSummary.byOrigin['preserved-source']} capsule=${result.writePlanSummary.byOrigin.capsule}`,
		)
		if (result.probeRecalcError) {
			console.log(`  probe-recalc-error: ${result.probeRecalcError}`)
		}
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

function summarizeBy<Key extends string>(
	results: readonly AuditResult[],
	getKey: (result: AuditResult) => Key,
): Record<Key, number> {
	const counts = {} as Record<Key, number>
	for (const result of results) {
		const key = getKey(result)
		counts[key] = (counts[key] ?? 0) + 1
	}
	return counts
}

function renderCounts(counts: Record<string, number>): string {
	return Object.entries(counts)
		.sort(([left], [right]) => left.localeCompare(right))
		.map(([key, value]) => `${key}:${value}`)
		.join(', ')
}

function sha256(bytes: Uint8Array): string {
	return createHash('sha256').update(bytes).digest('hex')
}

function readFlagValue(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	if (index === -1) return undefined
	return process.argv[index + 1]
}

function readFlagValues(name: string): string[] {
	const values: string[] = []
	for (let i = 0; i < process.argv.length; i++) {
		if (process.argv[i] !== name) continue
		const value = process.argv[i + 1]
		if (!value || value.startsWith('--')) continue
		values.push(value)
	}
	return values
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		// Best effort only.
	}
}

void main()
