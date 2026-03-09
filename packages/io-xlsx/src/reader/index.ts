import type { CellStyle, StyleId } from '@ascend/core'
import { Workbook } from '@ascend/core'
import type {
	AscendError,
	CompatibilityReport,
	CompatibilityStatus,
	CompatibilityTier,
	FeatureReport,
	Result,
} from '@ascend/schema'
import { ascendError, emptyReport, err, ok } from '@ascend/schema'
import type { PreservationCapsule } from '../preserve.ts'
import { parseCommentsXml } from './comments.ts'
import { type ContentTypes, parseContentTypes } from './content-types.ts'
import { parseDrawingImageRefs } from './drawing.ts'
import {
	getRelsPath,
	parseRelationships,
	REL_COMMENTS,
	REL_DRAWING,
	REL_OFFICE_DOC,
	REL_SHARED_STRINGS,
	REL_STYLES,
	REL_TABLE,
	REL_THEME,
	type Relationship,
	resolvePath,
} from './relationships.ts'
import { parseSharedStrings } from './shared-strings.ts'
import { parseSheet } from './sheet.ts'
import { parseStyles } from './styles.ts'
import { parseTable } from './table.ts'
import { parseThemeXml } from './theme.ts'
import { parseWorkbookXml } from './workbook.ts'
import { extractZip, type ZipArchive } from './zip.ts'

export interface ReadXlsxResult {
	readonly workbook: Workbook
	readonly report: CompatibilityReport
	readonly capsules: PreservationCapsule[]
	readonly loadInfo: ReadXlsxLoadInfo
}

export interface ReadXlsxOptions {
	readonly mode?: 'full' | 'metadata-only'
	readonly sheets?: readonly string[]
}

export interface ReadXlsxLoadInfo {
	readonly mode: 'full' | 'metadata-only' | 'selective'
	readonly isPartial: boolean
	readonly cellsHydrated: boolean
	readonly hasAllSheets: boolean
	readonly sourceSheetNames: readonly string[]
	readonly loadedSheetNames: readonly string[]
}

interface FormulaFeatureSummary {
	sharedFormulaSheets: string[]
	arrayFormulaSheets: string[]
}

export function readXlsx(
	bytes: Uint8Array,
	options: ReadXlsxOptions = {},
): Result<ReadXlsxResult, AscendError> {
	const mode = options.mode ?? 'full'
	const selectedSheets = options.sheets
		? new Set(options.sheets.map((name) => name.toLowerCase()))
		: null

	let archive: ZipArchive
	try {
		archive = extractZip(bytes)
	} catch (e) {
		return err(
			ascendError('CORRUPT_FILE', `Invalid ZIP: ${e instanceof Error ? e.message : 'unknown'}`),
		)
	}

	const contentTypesXml = readPart(archive, '[Content_Types].xml')
	if (!contentTypesXml) {
		return err(ascendError('CORRUPT_FILE', 'Missing [Content_Types].xml'))
	}
	const contentTypes = parseContentTypes(contentTypesXml)

	const rootRelsXml = readPart(archive, '_rels/.rels')
	if (!rootRelsXml) {
		return err(ascendError('CORRUPT_FILE', 'Missing _rels/.rels'))
	}
	const rootRels = parseRelationships(rootRelsXml)

	const docRel = rootRels.find((r) => r.type === REL_OFFICE_DOC)
	if (!docRel) {
		return err(ascendError('CORRUPT_FILE', 'No officeDocument relationship found'))
	}
	const workbookPath = docRel.target.replace(/^\//, '')

	const wbXml = readPart(archive, workbookPath)
	if (!wbXml) {
		return err(ascendError('CORRUPT_FILE', `Missing workbook: ${workbookPath}`))
	}
	const wbInfo = parseWorkbookXml(wbXml)

	const wbRelsPath = getRelsPath(workbookPath)
	const wbRelsXml = readPart(archive, wbRelsPath)
	const wbRels = wbRelsXml ? parseRelationships(wbRelsXml) : []
	const relMap = new Map(wbRels.map((r) => [r.id, r]))

	const workbook = new Workbook()
	workbook.calcSettings = wbInfo.calcSettings
	workbook.workbookProperties = { ...wbInfo.workbookProperties }
	workbook.workbookProtection = wbInfo.workbookProtection ? { ...wbInfo.workbookProtection } : null
	workbook.workbookViews.push(...wbInfo.workbookViews.map((view) => ({ ...view })))
	workbook.preservedXml = {
		workbookXml: wbXml,
		...(wbRelsXml ? { workbookRelsXml: wbRelsXml } : {}),
	}
	for (const relId of wbInfo.externalReferenceRelIds) {
		const rel = relMap.get(relId)
		if (!rel) continue
		workbook.externalReferences.push(resolvePath(workbookPath, rel.target))
	}

	const consumed = new Set<string>()
	consumed.add('[Content_Types].xml')
	consumed.add('_rels/.rels')
	consumed.add(workbookPath)
	consumed.add(wbRelsPath)

	const ssPart = wbRels.find((r) => r.type === REL_SHARED_STRINGS)
	const ssPath = ssPart ? resolvePath(workbookPath, ssPart.target) : undefined
	if (ssPath) consumed.add(ssPath)

	const stylesPart = wbRels.find((r) => r.type === REL_STYLES)
	const stylesPath = stylesPart ? resolvePath(workbookPath, stylesPart.target) : undefined
	if (stylesPath) consumed.add(stylesPath)

	const themePart = wbRels.find((r) => r.type === REL_THEME)
	const themePath = themePart ? resolvePath(workbookPath, themePart.target) : undefined
	if (themePath) consumed.add(themePath)
	if (themePath) {
		const themeXml = readPart(archive, themePath)
		if (themeXml) {
			workbook.themeMetadata = parseThemeXml(themeXml)
			workbook.preservedTheme = {
				path: themePath,
				contentType: resolveContentType(themePath, contentTypes),
				xml: themeXml,
			}
		}
	}

	const sheetPathToName = new Map<string, string>()
	const formulaFeatures: FormulaFeatureSummary = {
		sharedFormulaSheets: [],
		arrayFormulaSheets: [],
	}
	const sheetsToParse: Array<{
		name: string
		path: string
		state: (typeof wbInfo.sheets)[number]['state']
	}> = []

	for (const entry of wbInfo.sheets) {
		const rel = relMap.get(entry.rId)
		if (!rel) continue

		const sheetPath = resolvePath(workbookPath, rel.target)
		sheetPathToName.set(sheetPath, entry.name)
		consumed.add(sheetPath)
		consumed.add(getRelsPath(sheetPath))

		if (selectedSheets && !selectedSheets.has(entry.name.toLowerCase())) continue
		sheetsToParse.push({ name: entry.name, path: sheetPath, state: entry.state })
	}

	if (mode === 'metadata-only') {
		for (const entry of sheetsToParse) {
			const sheet = workbook.addSheet(entry.name)
			sheet.state = entry.state
		}
	} else {
		const ssXml = ssPath ? readPart(archive, ssPath) : undefined
		const sharedStrings = ssXml ? parseSharedStrings(ssXml) : []

		const stylesXml = stylesPath ? readPart(archive, stylesPath) : undefined
		const parsedStyles = stylesXml
			? parseStyles(stylesXml)
			: {
					cellStyles: [{}],
					differentialStyles: [],
					isDateFormat: [false],
					metadata: {
						numFmtCount: 0,
						fontCount: 0,
						fillCount: 0,
						borderCount: 0,
						cellXfCount: 1,
						dxfCount: 0,
						tableStyleCount: 0,
					},
				}
		if (stylesXml) {
			workbook.preservedStyles = { xml: stylesXml, xfByStyleId: {} }
		}
		const styleIds = registerStyles(workbook, parsedStyles.cellStyles)
		workbook.styleMetadata = parsedStyles.metadata
		workbook.differentialStyles.push(...parsedStyles.differentialStyles)

		for (const entry of sheetsToParse) {
			const sheetXml = readPart(archive, entry.path)
			if (!sheetXml) continue
			const sheetRelsXml = readPart(archive, getRelsPath(entry.path))
			const sheetRelationships = sheetRelsXml ? parseRelationships(sheetRelsXml) : []
			recordFormulaFeatures(sheetXml, entry.name, formulaFeatures)
			const sheet = parseSheet(entry.name, sheetXml, {
				sharedStrings,
				styleIds,
				isDateFormat: parsedStyles.isDateFormat,
				differentialStyles: parsedStyles.differentialStyles,
				relationships: sheetRelationships,
			})
			attachComments(archive, entry.path, sheet, sheetRelationships)
			attachDrawingImages(archive, entry.path, sheet, sheetRelationships)
			attachTables(archive, entry.path, sheet, sheetRelationships)
			sheet.state = entry.state
			sheet.preservedXml = {
				xml: sheetXml,
				...(sheetRelsXml ? { relsXml: sheetRelsXml } : {}),
			}
			workbook.sheets.push(sheet)
		}
	}

	for (const dn of wbInfo.definedNames) {
		if (dn.localSheetId !== undefined) {
			const sheet = workbook.sheets[dn.localSheetId]
			if (sheet) {
				workbook.definedNames.set(dn.name, dn.formula, { kind: 'sheet', sheetId: sheet.id })
				continue
			}
		}
		workbook.definedNames.set(dn.name, dn.formula)
	}

	const capsules = collectCapsules(
		archive,
		consumed,
		contentTypes,
		wbRels,
		sheetPathToName,
		workbookPath,
	)
	const sourceSheetNames = wbInfo.sheets.map((sheet) => sheet.name)
	const loadedSheetNames = sheetsToParse.map((sheet) => sheet.name)
	const hasAllSheets = loadedSheetNames.length === sourceSheetNames.length
	const cellsHydrated = mode !== 'metadata-only'
	const loadInfo: ReadXlsxLoadInfo = {
		mode: selectedSheets ? 'selective' : mode,
		isPartial: !hasAllSheets || !cellsHydrated,
		cellsHydrated,
		hasAllSheets,
		sourceSheetNames,
		loadedSheetNames,
	}
	const report = buildReport(contentTypes, formulaFeatures, workbook, capsules, loadInfo)

	return ok({ workbook, report, capsules, loadInfo })
}

function readPart(archive: ZipArchive, path: string): string | undefined {
	return archive.readText(path)
}

function registerStyles(workbook: Workbook, cellStyles: CellStyle[]): StyleId[] {
	const styleIds = cellStyles.map((style) => workbook.styles.register(style))
	const xfByStyleId: Record<number, number> = {}
	for (let xfIndex = 0; xfIndex < styleIds.length; xfIndex++) {
		const styleId = styleIds[xfIndex]
		if (styleId === undefined) continue
		if (xfByStyleId[styleId] === undefined) xfByStyleId[styleId] = xfIndex
	}
	if (xfByStyleId[0] === undefined) xfByStyleId[0] = 0
	if (workbook.preservedStyles) {
		workbook.preservedStyles = {
			...workbook.preservedStyles,
			xfByStyleId,
		}
	}
	return styleIds
}

function collectCapsules(
	archive: ZipArchive,
	consumed: Set<string>,
	contentTypes: ContentTypes,
	wbRels: Relationship[],
	sheetPathToName: Map<string, string>,
	workbookPath: string,
): PreservationCapsule[] {
	const capsules: PreservationCapsule[] = []

	const wbRelByTarget = new Map<string, Relationship>()
	for (const rel of wbRels) {
		wbRelByTarget.set(resolvePath(workbookPath, rel.target), rel)
	}

	const sheetRelByTarget = new Map<string, { sheetName: string; rel: Relationship }>()
	for (const [sheetPath, sheetName] of sheetPathToName) {
		const sheetRelsXml = readPart(archive, getRelsPath(sheetPath))
		if (!sheetRelsXml) continue
		for (const rel of parseRelationships(sheetRelsXml)) {
			sheetRelByTarget.set(resolvePath(sheetPath, rel.target), { sheetName, rel })
		}
	}

	for (const entry of archive.entries()) {
		const partPath = entry.path
		if (consumed.has(partPath)) continue
		if (partPath.endsWith('.rels')) continue
		if (partPath.startsWith('_rels/')) continue
		if (partPath.startsWith('docProps/')) continue

		const ct = resolveContentType(partPath, contentTypes)

		let anchor: PreservationCapsule['anchor'] = { kind: 'workbook' }
		let relType: string | undefined

		const wbRef = wbRelByTarget.get(partPath)
		if (wbRef) relType = wbRef.type

		const sheetRef = sheetRelByTarget.get(partPath)
		if (sheetRef) {
			anchor = { kind: 'sheet', sheetName: sheetRef.sheetName }
			relType = sheetRef.rel.type
		}

		const capsuleRelsXml = readPart(archive, getRelsPath(partPath))
		const relationships = capsuleRelsXml
			? parseRelationships(capsuleRelsXml).map((r) => ({
					id: r.id,
					type: r.type,
					target: r.target,
				}))
			: []

		const capsule: PreservationCapsule = {
			partPath,
			contentType: ct,
			relationships,
			content: archive.readBytes(partPath) ?? new Uint8Array(),
			anchor,
		}
		if (relType) capsule.relType = relType
		capsules.push(capsule)
	}

	return capsules
}

function resolveContentType(partPath: string, contentTypes: ContentTypes): string {
	const normalized = partPath.startsWith('/') ? partPath.substring(1) : partPath
	const ct = contentTypes.overrides.get(normalized)
	if (ct) return ct

	const ext = partPath.split('.').pop() ?? ''
	return contentTypes.defaults.get(ext) ?? 'application/octet-stream'
}

function buildReport(
	contentTypes: {
		overrides: ReadonlyMap<string, string>
	},
	formulaFeatures: FormulaFeatureSummary,
	workbook: Workbook,
	capsules: readonly PreservationCapsule[],
	loadInfo: ReadXlsxLoadInfo,
): CompatibilityReport {
	const features: FeatureReport[] = []

	const unsupportedTypes: [string, string][] = [
		['chart', 'chart+xml'],
		['pivotTable', 'pivotTable+xml'],
		['drawing', 'drawing+xml'],
		['vbaProject', 'vbaProject'],
	]

	for (const [feature, pattern] of unsupportedTypes) {
		const locations: string[] = []
		for (const [path, ct] of contentTypes.overrides) {
			if (ct.includes(pattern)) locations.push(path)
		}
		if (locations.length > 0) {
			features.push({
				feature,
				tier: 'unsupported' as CompatibilityTier,
				count: locations.length,
				locations,
			})
		}
	}

	if (formulaFeatures.sharedFormulaSheets.length > 0) {
		features.push({
			feature: 'sharedFormula',
			tier: 'unsupported',
			count: formulaFeatures.sharedFormulaSheets.length,
			locations: formulaFeatures.sharedFormulaSheets,
			note: 'Shared formulas are detected but not yet modeled as first-class semantics.',
		})
	}

	if (formulaFeatures.arrayFormulaSheets.length > 0) {
		features.push({
			feature: 'arrayFormula',
			tier: 'unsupported',
			count: formulaFeatures.arrayFormulaSheets.length,
			locations: formulaFeatures.arrayFormulaSheets,
			note: 'Array formulas are detected but not yet modeled as first-class semantics.',
		})
	}

	const tableLocations = workbook.sheets
		.filter((sheet) => sheet.tables.length > 0)
		.map((sheet) => sheet.name)
	if (tableLocations.length > 0) {
		features.push({
			feature: 'table',
			tier: 'normalized',
			count: tableLocations.length,
			locations: tableLocations,
			note: 'Tables are imported for read access, but full round-trip table semantics are not complete.',
		})
	}

	if (workbook.calcSettings.calcMode === 'manual' || workbook.calcSettings.fullCalcOnLoad) {
		const reasons: string[] = []
		if (workbook.calcSettings.calcMode === 'manual') reasons.push('manual calculation mode')
		if (workbook.calcSettings.fullCalcOnLoad) reasons.push('full recalculation requested on load')
		features.push({
			feature: 'formulaFreshness',
			tier: 'normalized',
			count: 1,
			locations: ['workbook'],
			note: `Workbook indicates ${reasons.join(' and ')}.`,
		})
	}

	const calcChainLocations: string[] = []
	for (const [path, ct] of contentTypes.overrides) {
		if (ct.includes('calcChain+xml')) calcChainLocations.push(path)
	}
	if (calcChainLocations.length > 0) {
		features.push({
			feature: 'calcChain',
			tier: 'normalized',
			count: calcChainLocations.length,
			locations: calcChainLocations,
			note: 'Calc chain is treated as an optimization hint and not as calculation truth.',
		})
	}

	if (capsules.length > 0) {
		features.push({
			feature: 'preservedPart',
			tier: 'preserved',
			count: capsules.length,
			locations: capsules.map((capsule) => capsule.partPath),
			note: 'Extra OOXML parts are preserved outside the semantic workbook model.',
		})
	}

	if (loadInfo.isPartial) {
		const reasons: string[] = []
		if (!loadInfo.hasAllSheets) reasons.push('only selected sheets are loaded')
		if (!loadInfo.cellsHydrated) reasons.push('sheet cells are not hydrated')
		features.push({
			feature: 'partialLoad',
			tier: 'normalized',
			count: 1,
			locations: loadInfo.loadedSheetNames,
			note: `Workbook is being inspected through a partial view because ${reasons.join(' and ')}.`,
		})
	}

	if (features.length === 0) return emptyReport('xlsx')

	const summary = { exact: 0, normalized: 0, preserved: 0, unsupported: 0 }
	for (const f of features) {
		summary[f.tier] += f.count
	}

	const status: CompatibilityStatus =
		summary.unsupported > 0 ? 'has-unsupported' : summary.preserved > 0 ? 'has-preserved' : 'clean'

	return {
		status,
		features,
		summary,
		sourceFormat: 'xlsx',
	}
}

function attachTables(
	archive: ZipArchive,
	sheetPath: string,
	sheet: Workbook['sheets'][number],
	sheetRelationships: readonly Relationship[],
): void {
	if (!sheet) return
	const tableRels = sheetRelationships.filter((rel) => rel.type === REL_TABLE)
	for (const rel of tableRels) {
		const tablePath = resolvePath(sheetPath, rel.target)
		const tableXml = readPart(archive, tablePath)
		if (!tableXml) continue
		const table = parseTable(tableXml, sheet.id)
		if (!table) continue
		sheet.tables.push(table)
	}
}

function attachComments(
	archive: ZipArchive,
	sheetPath: string,
	sheet: Workbook['sheets'][number],
	sheetRelationships: readonly Relationship[],
): void {
	if (!sheet) return
	const commentsRel = sheetRelationships.find((rel) => rel.type === REL_COMMENTS)
	if (!commentsRel) return
	const commentsPath = resolvePath(sheetPath, commentsRel.target)
	const commentsXml = readPart(archive, commentsPath)
	if (!commentsXml) return
	for (const [ref, comment] of parseCommentsXml(commentsXml)) {
		sheet.comments.set(ref, comment)
	}
}

function attachDrawingImages(
	archive: ZipArchive,
	sheetPath: string,
	sheet: Workbook['sheets'][number],
	sheetRelationships: readonly Relationship[],
): void {
	if (!sheet) return
	for (const drawingRel of sheetRelationships.filter((rel) => rel.type === REL_DRAWING)) {
		const drawingPath = resolvePath(sheetPath, drawingRel.target)
		const drawingXml = readPart(archive, drawingPath)
		const drawingRelsXml = readPart(archive, getRelsPath(drawingPath))
		if (!drawingXml || !drawingRelsXml) continue
		sheet.imageRefs.push(
			...parseDrawingImageRefs(drawingXml, drawingPath, parseRelationships(drawingRelsXml)),
		)
	}
}

function recordFormulaFeatures(
	sheetXml: string,
	sheetName: string,
	summary: FormulaFeatureSummary,
): void {
	if (/<f\b[^>]*\bt="shared"/.test(sheetXml)) {
		summary.sharedFormulaSheets.push(sheetName)
	}
	if (/<f\b[^>]*\bt="array"/.test(sheetXml)) {
		summary.arrayFormulaSheets.push(sheetName)
	}
}
