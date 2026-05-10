import type {
	ActiveContentInfo,
	CellStyle,
	SheetId,
	StyleId,
	WorkbookConnectionPartInfo,
	WorkbookDataModelPartInfo,
} from '@ascend/core'
import { DEFAULT_STYLE_ID, Workbook } from '@ascend/core'
import type {
	AscendError,
	CompatibilityReport,
	CompatibilityStatus,
	FeatureReport,
	Result,
} from '@ascend/schema'
import { ascendError, emptyReport, err, ok } from '@ascend/schema'
import { normalizeStoredFormulaText } from '../formula-storage.ts'
import type { PreservationCapsule } from '../preserve.ts'
import { parseChartXml } from './charts.ts'
import {
	parseCommentsXml,
	parseThreadedCommentPersonsXml,
	parseThreadedCommentsXml,
} from './comments.ts'
import { parseConnectionPartInfos } from './connections.ts'
import { type ContentTypes, parseContentTypes } from './content-types.ts'
import { parseDataModelPartInfo } from './data-model.ts'
import { parseDrawingImageRefs, parseDrawingObjectRefs } from './drawing.ts'
import { maybeDecryptOoxmlPackage } from './encryption.ts'
import { parseMacroSheetInfo } from './macro-sheet.ts'
import { parseMetadataXml } from './metadata.ts'
import {
	parsePivotCacheDefinitionXml,
	parsePivotTableXml,
	parseSlicerCacheXml,
	parseSlicerXml,
	parseTimelineCacheXml,
	parseTimelineXml,
} from './pivots.ts'
import {
	getRelsPath,
	parseRelationships,
	REL_CHART,
	REL_CHARTSHEET,
	REL_COMMENTS,
	REL_DRAWING,
	REL_MACROSHEET,
	REL_OFFICE_DOC,
	REL_PIVOT_TABLE,
	REL_SHARED_STRINGS,
	REL_SHEET_METADATA,
	REL_SLICER_CACHE,
	REL_STYLES,
	REL_TABLE,
	REL_THEME,
	REL_THREADED_COMMENT,
	type Relationship,
	resolvePath,
} from './relationships.ts'
import { emptySharedStrings, parseSharedStrings } from './shared-strings.ts'
import {
	parseSheet,
	parseSheetValuesOnlyBytes,
	type SheetFormulaFeatures,
	ValueInternPool,
} from './sheet.ts'
import { parseStyles, parseStylesLite } from './styles.ts'
import { parseTable } from './table.ts'
import { parseThemeColorsXml, parseThemeXml } from './theme.ts'
import { summarizeVbaProject } from './vba.ts'
import { parseWorkbookXml } from './workbook.ts'
import { extractZip, type ZipArchive } from './zip.ts'

const XML_DECODER = new TextDecoder('utf-8')
const VALUES_ONLY_BYTE_PARSE_MIN_BYTES = 10_000_000

export interface ReadXlsxResult {
	readonly workbook: Workbook
	readonly report: CompatibilityReport
	readonly capsules: PreservationCapsule[]
	readonly loadInfo: ReadXlsxLoadInfo
}

export interface ReadXlsxOptions {
	readonly mode?: 'full' | 'metadata-only' | 'values' | 'formula'
	readonly sheets?: readonly (string | number)[]
	readonly maxRows?: number
	readonly richMetadata?: boolean
	readonly parseDates?: boolean
	readonly password?: string
}

export interface ReadXlsxLoadInfo {
	readonly mode: 'full' | 'metadata-only' | 'values' | 'formula' | 'selective'
	readonly isPartial: boolean
	readonly cellsHydrated: boolean
	readonly richSheetMetadataHydrated: boolean
	readonly hasAllSheets: boolean
	readonly sourceSheetNames: readonly string[]
	readonly loadedSheetNames: readonly string[]
}

interface FormulaFeatureSummary {
	sharedFormulaSheets: string[]
	arrayFormulaSheets: string[]
	dynamicArraySheets: string[]
}

export function readXlsx(
	bytes: Uint8Array,
	options: ReadXlsxOptions = {},
): Result<ReadXlsxResult, AscendError> {
	const mode = options.mode ?? 'full'
	let archiveBytes = bytes

	let archive: ZipArchive
	try {
		archive = extractZip(archiveBytes)
	} catch (e) {
		const decrypted = maybeDecryptOoxmlPackage(bytes, options.password)
		if (!decrypted.ok) return decrypted
		if (decrypted.value) {
			archiveBytes = decrypted.value
			try {
				archive = extractZip(archiveBytes)
			} catch (decryptedError) {
				return err(
					ascendError(
						'CORRUPT_FILE',
						`Invalid decrypted ZIP: ${decryptedError instanceof Error ? decryptedError.message : 'unknown'}`,
					),
				)
			}
		} else {
			return err(
				ascendError('CORRUPT_FILE', `Invalid ZIP: ${e instanceof Error ? e.message : 'unknown'}`),
			)
		}
	}

	try {
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
		const selectedSheets = resolveSheetFilter(options.sheets, wbInfo.sheets)

		const wbRelsPath = getRelsPath(workbookPath)
		const wbRelsXml = readPart(archive, wbRelsPath)
		const wbRels = wbRelsXml ? parseRelationships(wbRelsXml) : []
		const relMap = new Map(wbRels.map((r) => [r.id, r]))

		const workbook = new Workbook()
		workbook.sourceArchiveBytes = archiveBytes
		workbook.calcSettings = wbInfo.calcSettings
		workbook.workbookProperties = { ...wbInfo.workbookProperties }
		workbook.workbookProtection = wbInfo.workbookProtection
			? { ...wbInfo.workbookProtection }
			: null
		workbook.workbookViews.push(...wbInfo.workbookViews.map((view) => ({ ...view })))
		const workbookContentType = contentTypes.overrides.get(workbookPath)
		workbook.preservedXml = {
			workbookPath,
			...(wbRelsXml ? { workbookRelsPath: wbRelsPath } : {}),
			...(workbookContentType ? { contentType: workbookContentType } : {}),
			contentTypeDefaults: Array.from(contentTypes.defaults, ([extension, contentType]) => ({
				extension,
				contentType,
			})),
			sheetEntries: wbInfo.sheets.map((entry) => ({
				kind: workbookSheetEntryKind(relMap.get(entry.rId)?.type),
				sheetId: entry.sheetId,
				name: entry.name,
			})),
		}
		for (const relId of wbInfo.externalReferenceRelIds) {
			const rel = relMap.get(relId)
			if (!rel) continue
			const partPath = resolvePath(workbookPath, rel.target)
			workbook.externalReferences.push(partPath)
			const relsXml = readPart(archive, getRelsPath(partPath))
			const linkRelationship = relsXml ? parseRelationships(relsXml)[0] : undefined
			workbook.externalReferenceDetails.push({
				partPath,
				relId,
				...(linkRelationship?.id ? { linkRelId: linkRelationship.id } : {}),
				...(linkRelationship?.target ? { target: linkRelationship.target } : {}),
				...(linkRelationship?.targetMode ? { targetMode: linkRelationship.targetMode } : {}),
			})
		}

		const consumed = new Set<string>()
		consumed.add('[Content_Types].xml')
		consumed.add('_rels/.rels')
		consumed.add(workbookPath)
		consumed.add(wbRelsPath)

		if (mode === 'full') {
			for (const entry of wbInfo.pivotCacheEntries) {
				const rel = relMap.get(entry.relId)
				if (!rel) continue
				const partPath = resolvePath(workbookPath, rel.target)
				const xml = readPart(archive, partPath)
				const relsXml = readPart(archive, getRelsPath(partPath))
				const relationships = relsXml ? parseRelationships(relsXml) : []
				const parsed = xml
					? parsePivotCacheDefinitionXml(xml, partPath, entry.cacheId, entry.relId, relationships)
					: null
				if (parsed) workbook.pivotCaches.push(parsed)
				if (relsXml) consumed.add(getRelsPath(partPath))
			}
			for (const rel of wbRels.filter((relationship) => relationship.type === REL_SLICER_CACHE)) {
				const partPath = resolvePath(workbookPath, rel.target)
				const xml = readPart(archive, partPath)
				if (xml) {
					const parsed = parseSlicerCacheXml(xml, partPath)
					if (parsed) workbook.slicerCaches.push(parsed)
				}
			}
			for (const entry of archive.entries()) {
				if (!entry.path.startsWith('xl/slicers/') || !entry.path.endsWith('.xml')) continue
				const xml = readPart(archive, entry.path)
				if (!xml) continue
				workbook.slicers.push(...parseSlicerXml(xml, entry.path))
			}
			for (const entry of archive.entries()) {
				if (!entry.path.startsWith('xl/timelineCaches/') || !entry.path.endsWith('.xml')) {
					continue
				}
				const xml = readPart(archive, entry.path)
				if (!xml) continue
				const parsed = parseTimelineCacheXml(xml, entry.path)
				if (parsed) workbook.timelineCaches.push(parsed)
			}
			for (const entry of archive.entries()) {
				if (!entry.path.startsWith('xl/timelines/') || !entry.path.endsWith('.xml')) continue
				const xml = readPart(archive, entry.path)
				if (!xml) continue
				workbook.timelines.push(...parseTimelineXml(xml, entry.path))
			}
		}

		const ssPart = wbRels.find((r) => r.type === REL_SHARED_STRINGS)
		const ssPath = ssPart ? resolvePath(workbookPath, ssPart.target) : undefined
		if (ssPath) {
			consumed.add(ssPath)
			workbook.preservedSharedStrings = { path: ssPath }
		}

		const stylesPart = wbRels.find((r) => r.type === REL_STYLES)
		const stylesPath = stylesPart ? resolvePath(workbookPath, stylesPart.target) : undefined
		if (stylesPath) consumed.add(stylesPath)

		const themePart = wbRels.find((r) => r.type === REL_THEME)
		const themePath = themePart ? resolvePath(workbookPath, themePart.target) : undefined
		if (themePath) consumed.add(themePath)
		if (themePath && mode === 'full') {
			const themeXml = readPart(archive, themePath)
			if (themeXml) {
				workbook.themeMetadata = parseThemeXml(themeXml)
				workbook.themeColors.push(...parseThemeColorsXml(themeXml))
				workbook.preservedTheme = {
					path: themePath,
					contentType: resolveContentType(themePath, contentTypes),
				}
			}
		}

		const sheetPathToAnchor = new Map<string, { sheetId: string; sheetName: string }>()
		const sheetRelsByPath = new Map<string, Relationship[]>()
		const formulaFeatures: FormulaFeatureSummary = {
			sharedFormulaSheets: [],
			arrayFormulaSheets: [],
			dynamicArraySheets: [],
		}
		const metadataRel = wbRels.find((rel) => rel.type === REL_SHEET_METADATA)
		const metadataPath = metadataRel ? resolvePath(workbookPath, metadataRel.target) : undefined
		const needsMetadata = mode === 'full' || mode === 'formula'
		const metadataXml = needsMetadata && metadataPath ? readPart(archive, metadataPath) : undefined
		const metadata = metadataXml ? parseMetadataXml(metadataXml) : undefined
		if (metadataPath && metadataRel) {
			consumed.add(metadataPath)
			workbook.preservedMetadata = {
				path: metadataPath,
				contentType: resolveContentType(metadataPath, contentTypes),
				...(metadataXml ? { xml: metadataXml } : {}),
				...(metadata
					? {
							dynamicArrayMetadata: [...metadata.dynamicArrayByCellMetadataIndex.values()],
						}
					: {}),
			}
		}
		const sheetsToParse: Array<{
			name: string
			path: string
			sheetId: SheetId
			state: (typeof wbInfo.sheets)[number]['state']
		}> = []
		const sourceWorksheetNames: string[] = []

		for (const entry of wbInfo.sheets) {
			const rel = relMap.get(entry.rId)
			if (!rel) continue

			const sheetPath = resolvePath(workbookPath, rel.target)
			sheetPathToAnchor.set(sheetPath, { sheetId: entry.sheetId, sheetName: entry.name })

			if (rel.type === REL_CHARTSHEET) {
				const chartsheetRelsXml = readPart(archive, getRelsPath(sheetPath))
				const chartsheetRelationships = chartsheetRelsXml
					? parseRelationships(chartsheetRelsXml)
					: []
				sheetRelsByPath.set(sheetPath, chartsheetRelationships)
				if (!selectedSheets || selectedSheets.has(entry.name.toLowerCase())) {
					workbook.chartSheets.push({
						name: entry.name,
						sheetId: entry.sheetId,
						relId: entry.rId,
						partPath: sheetPath,
						state: entry.state,
						chartPartPaths: chartsheetRelationships
							.filter((relationship) => relationship.type === REL_CHART)
							.map((relationship) => resolvePath(sheetPath, relationship.target)),
					})
				}
				continue
			}

			if (rel.type === REL_MACROSHEET) {
				const macroSheetRelsXml = readPart(archive, getRelsPath(sheetPath))
				const macroSheetRelationships = macroSheetRelsXml
					? parseRelationships(macroSheetRelsXml)
					: []
				sheetRelsByPath.set(sheetPath, macroSheetRelationships)
				if (!selectedSheets || selectedSheets.has(entry.name.toLowerCase())) {
					workbook.macroSheets.push(
						parseMacroSheetInfo(readPart(archive, sheetPath), {
							name: entry.name,
							sheetId: entry.sheetId,
							relId: entry.rId,
							partPath: sheetPath,
							state: entry.state,
							relationships: macroSheetRelationships,
						}),
					)
				}
				continue
			}

			consumed.add(sheetPath)
			consumed.add(getRelsPath(sheetPath))
			sourceWorksheetNames.push(entry.name)
			if (selectedSheets && !selectedSheets.has(entry.name.toLowerCase())) continue
			sheetsToParse.push({
				name: entry.name,
				path: sheetPath,
				sheetId: entry.sheetId as SheetId,
				state: entry.state,
			})
		}

		const valuesOnly = mode === 'values'
		const formulaOnly = mode === 'formula'
		const hydrateRichSheetMetadata =
			mode === 'full' || ((valuesOnly || formulaOnly) && options.richMetadata === true)
		if (mode === 'metadata-only') {
			for (const entry of sheetsToParse) {
				const sheet = workbook.addSheet(entry.name, entry.sheetId as SheetId)
				sheet.state = entry.state
			}
		} else {
			const valuePool = valuesOnly ? undefined : new ValueInternPool()
			const ssXml = ssPath ? readPart(archive, ssPath) : undefined
			const sharedStrings = ssXml
				? parseSharedStrings(ssXml, {
						...(valuesOnly || formulaOnly
							? {}
							: { normalize: (value) => valuePool?.internValue(value) ?? value }),
						lazy: valuesOnly || formulaOnly || selectedSheets !== null,
					})
				: emptySharedStrings()

			const parseLiteStyles = !((valuesOnly || formulaOnly) && options.parseDates === false)
			const stylesXml =
				stylesPath && (parseLiteStyles || !(valuesOnly || formulaOnly))
					? readPart(archive, stylesPath)
					: undefined
			let styleIds: StyleId[]
			let isDateFormat: boolean[]
			let differentialStyles: readonly CellStyle[]
			if (valuesOnly || formulaOnly) {
				const parsedStyles =
					stylesXml && parseLiteStyles
						? parseStylesLite(stylesXml)
						: {
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
				styleIds = new Array<StyleId>(parsedStyles.isDateFormat.length).fill(DEFAULT_STYLE_ID)
				isDateFormat = parsedStyles.isDateFormat
				differentialStyles = []
				workbook.styleMetadata = parsedStyles.metadata
			} else {
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
				if (stylesXml && stylesPath) {
					workbook.preservedStyles = { path: stylesPath, xfByStyleId: {} }
				}
				styleIds = registerStyles(workbook, parsedStyles.cellStyles)
				isDateFormat = parsedStyles.isDateFormat
				differentialStyles = parsedStyles.differentialStyles
				workbook.styleMetadata = parsedStyles.metadata
				workbook.differentialStyles.push(...parsedStyles.differentialStyles)
			}
			const hasDateStyles = isDateFormat.some(Boolean)

			for (const entry of sheetsToParse) {
				const canUseValuesOnlyByteParser =
					valuesOnly &&
					!hydrateRichSheetMetadata &&
					(archive.get(entry.path)?.uncompressedSize ?? 0) >= VALUES_ONLY_BYTE_PARSE_MIN_BYTES &&
					options.maxRows === undefined
				const sheetBytes = canUseValuesOnlyByteParser
					? readPartBytes(archive, entry.path)
					: undefined
				let sheetXml =
					sheetBytes === undefined || hydrateRichSheetMetadata
						? readPart(archive, entry.path)
						: undefined
				if (sheetBytes === undefined && !sheetXml) continue
				const sheetRelsXml = hydrateRichSheetMetadata
					? readPart(archive, getRelsPath(entry.path))
					: undefined
				const sheetRelationships = sheetRelsXml ? parseRelationships(sheetRelsXml) : []
				sheetRelsByPath.set(entry.path, sheetRelationships)
				const sheetFormulaFeatures: SheetFormulaFeatures = {
					hasSharedFormula: false,
					hasArrayFormula: false,
					hasDynamicArrayFormula: false,
				}
				const sheetCtx = {
					sharedStrings,
					styleIds,
					isDateFormat,
					hasDateStyles,
					differentialStyles,
					relationships: sheetRelationships,
					valuesOnly,
					formulaOnly,
					richMetadata: hydrateRichSheetMetadata,
					formulaFeatures: sheetFormulaFeatures,
					...(valuePool ? { valuePool } : {}),
					...(metadata ? { metadata } : {}),
					...(options.maxRows !== undefined ? { maxRows: options.maxRows } : {}),
				}
				const resolvedSheetId = (sheetPathToAnchor.get(entry.path)?.sheetId ??
					entry.name) as SheetId
				let sheet = sheetBytes
					? parseSheetValuesOnlyBytes(entry.name, sheetBytes, sheetCtx, resolvedSheetId)
					: null
				if (!sheet) {
					sheetXml ??= sheetBytes ? XML_DECODER.decode(sheetBytes) : readPart(archive, entry.path)
					if (!sheetXml) continue
					sheet = parseSheet(entry.name, sheetXml, sheetCtx, resolvedSheetId)
				}
				if (sheetFormulaFeatures.hasSharedFormula) {
					formulaFeatures.sharedFormulaSheets.push(entry.name)
				}
				if (sheetFormulaFeatures.hasArrayFormula) {
					formulaFeatures.arrayFormulaSheets.push(entry.name)
				}
				if (sheetFormulaFeatures.hasDynamicArrayFormula) {
					formulaFeatures.dynamicArraySheets.push(entry.name)
				}
				if (hydrateRichSheetMetadata) {
					attachComments(archive, entry.path, sheet, sheetRelationships)
					attachDrawingImages(archive, entry.path, sheet, sheetRelationships)
					attachPivotTables(archive, entry.path, entry.name, workbook, sheetRelationships)
				}
				if (hydrateRichSheetMetadata) attachTables(archive, entry.path, sheet, sheetRelationships)
				sheet.state = entry.state
				if (hydrateRichSheetMetadata) {
					sheet.preservedXml = {
						partPath: entry.path,
						...(sheetRelsXml ? { relsPath: getRelsPath(entry.path) } : {}),
					}
				}
				workbook.sheets.push(sheet)
			}
		}

		for (const dn of wbInfo.definedNames) {
			const options = dn.hidden !== undefined ? { hidden: dn.hidden } : {}
			if (dn.localSheetId !== undefined) {
				const sourceSheet = wbInfo.sheets[dn.localSheetId]
				const sheet = sourceSheet ? workbook.getSheet(sourceSheet.name) : undefined
				if (sheet) {
					workbook.definedNames.set(
						dn.name,
						normalizeStoredFormulaText(dn.formula),
						{
							kind: 'sheet',
							sheetId: sheet.id,
						},
						options,
					)
					continue
				}
				continue
			}
			workbook.definedNames.set(dn.name, normalizeStoredFormulaText(dn.formula), undefined, options)
		}

		const sourceSheetNames = sourceWorksheetNames
		const loadedSheetNames = sheetsToParse.map((sheet) => sheet.name)
		const hasAllSheets = loadedSheetNames.length === sourceSheetNames.length
		const cellsHydrated = mode !== 'metadata-only'
		const richSheetMetadataHydrated = hydrateRichSheetMetadata
		const fidelityPartial = mode === 'values' || mode === 'formula' || options.maxRows !== undefined
		const isPartial = !hasAllSheets || !cellsHydrated || fidelityPartial
		const loadInfo: ReadXlsxLoadInfo = {
			mode: selectedSheets ? 'selective' : mode,
			isPartial,
			cellsHydrated,
			richSheetMetadataHydrated,
			hasAllSheets,
			sourceSheetNames,
			loadedSheetNames,
		}
		const capsules = isPartial
			? []
			: collectCapsules(
					archive,
					consumed,
					contentTypes,
					rootRels,
					wbRels,
					sheetPathToAnchor,
					workbookPath,
					sheetRelsByPath,
				)
		if (!isPartial) workbook.activeContent.push(...collectActiveContent(archive, capsules))
		if (!isPartial) workbook.connectionParts.push(...collectConnectionParts(archive, capsules))
		if (!isPartial) workbook.dataModelParts.push(...collectDataModelParts(capsules))
		if (!isPartial) attachChartParts(archive, workbook, capsules)
		const report = buildReport(contentTypes, formulaFeatures, workbook, capsules, loadInfo)
		if (loadInfo.isPartial) {
			workbook.sourceArchiveBytes = null
			return ok({ workbook, report, capsules: [], loadInfo })
		}

		return ok({ workbook, report, capsules, loadInfo })
	} catch (e) {
		if (e instanceof TypeError || e instanceof ReferenceError) throw e
		return err(
			ascendError(
				'CORRUPT_FILE',
				`Invalid workbook payload: ${e instanceof Error ? e.message : 'unknown'}`,
			),
		)
	}
}

function readPart(archive: ZipArchive, path: string): string | undefined {
	return archive.readText(path)
}

function readPartBytes(archive: ZipArchive, path: string): Uint8Array | undefined {
	return archive.readBytes(path)
}

function attachChartParts(
	archive: ZipArchive,
	workbook: Workbook,
	capsules: readonly PreservationCapsule[],
): void {
	const sheetNameByChartPartPath = mapEmbeddedChartsToSheets(capsules)
	for (const capsule of capsules) {
		if (!isChartCapsule(capsule)) continue
		const xml = readPart(archive, capsule.partPath)
		if (!xml) continue
		const sheetName =
			capsule.anchor.kind === 'sheet'
				? capsule.anchor.sheetName
				: sheetNameByChartPartPath.get(capsule.partPath)
		workbook.chartParts.push(parseChartXml(xml, capsule.partPath, sheetName))
	}
}

function mapEmbeddedChartsToSheets(
	capsules: readonly PreservationCapsule[],
): ReadonlyMap<string, string> {
	const sheetNameByChartPartPath = new Map<string, string>()
	for (const capsule of capsules) {
		if (capsule.anchor.kind !== 'sheet' || !capsule.anchor.sheetName) continue
		if (!isDrawingCapsule(capsule)) continue
		for (const rel of capsule.relationships) {
			if (rel.type !== REL_CHART || rel.targetMode === 'External') continue
			const chartPath = resolvePath(capsule.partPath, rel.target)
			if (!sheetNameByChartPartPath.has(chartPath)) {
				sheetNameByChartPartPath.set(chartPath, capsule.anchor.sheetName)
			}
		}
	}
	return sheetNameByChartPartPath
}

function resolveSheetFilter(
	filter: readonly (string | number)[] | undefined,
	sheets: readonly { readonly name: string }[],
): Set<string> | null {
	if (!filter || filter.length === 0) return null
	const selected = new Set<string>()
	for (const selector of filter) {
		if (typeof selector === 'string') {
			selected.add(selector.toLowerCase())
		} else {
			const entry = sheets[selector]
			if (entry) selected.add(entry.name.toLowerCase())
		}
	}
	return selected.size > 0 ? selected : null
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
	rootRels: Relationship[],
	wbRels: Relationship[],
	sheetPathToAnchor: Map<string, { sheetId: string; sheetName: string }>,
	workbookPath: string,
	sheetRelsByPath: Map<string, Relationship[]>,
): PreservationCapsule[] {
	const capsules: PreservationCapsule[] = []

	const rootRelByTarget = new Map<string, Relationship>()
	for (const rel of rootRels) {
		rootRelByTarget.set(resolvePath('', rel.target), rel)
	}

	const wbRelByTarget = new Map<string, Relationship>()
	for (const rel of wbRels) {
		wbRelByTarget.set(resolvePath(workbookPath, rel.target), rel)
	}

	const sheetRelByTarget = new Map<
		string,
		{ sheetName: string; sheetId: string; rel: Relationship }
	>()
	for (const [sheetPath, anchor] of sheetPathToAnchor) {
		const rels = sheetRelsByPath.get(sheetPath)
		if (!rels) continue
		for (const rel of rels) {
			sheetRelByTarget.set(resolvePath(sheetPath, rel.target), { ...anchor, rel })
		}
	}

	for (const entry of archive.entries()) {
		const partPath = entry.path
		if (isIgnorablePackageEntry(partPath)) continue
		if (consumed.has(partPath)) continue
		if (partPath.endsWith('.rels')) continue
		if (partPath.startsWith('_rels/')) continue

		const contentType = resolveContentTypeInfo(partPath, contentTypes)

		let anchor: PreservationCapsule['anchor'] = { kind: 'workbook' }
		let relType: string | undefined
		const directSheetAnchor = sheetPathToAnchor.get(partPath)
		if (directSheetAnchor) {
			anchor = {
				kind: 'sheet',
				sheetId: directSheetAnchor.sheetId,
				sheetName: directSheetAnchor.sheetName,
			}
		}

		const wbRef = wbRelByTarget.get(partPath)
		if (wbRef) relType = wbRef.type
		const rootRef = rootRelByTarget.get(partPath)
		if (rootRef) relType = rootRef.type

		const sheetRef = sheetRelByTarget.get(partPath)
		if (sheetRef) {
			anchor = { kind: 'sheet', sheetId: sheetRef.sheetId, sheetName: sheetRef.sheetName }
			relType = sheetRef.rel.type
		}

		const capsuleRelsXml = readPart(archive, getRelsPath(partPath))
		const relationships = capsuleRelsXml
			? parseRelationships(capsuleRelsXml).map((r) => ({
					id: r.id,
					type: r.type,
					target: r.target,
					...(r.targetMode ? { targetMode: r.targetMode } : {}),
				}))
			: []

		const capsule: PreservationCapsule = {
			partPath,
			contentType: contentType.value,
			contentTypeSource: contentType.source,
			relationships,
			anchor,
		}
		if (relType) capsule.relType = relType
		capsules.push(capsule)
	}

	return capsules
}

function isIgnorablePackageEntry(partPath: string): boolean {
	return (
		partPath.endsWith('/') ||
		partPath === '.DS_Store' ||
		partPath.endsWith('/.DS_Store') ||
		partPath.startsWith('__MACOSX/')
	)
}

function isChartCapsule(capsule: PreservationCapsule): boolean {
	return isStructuredChartPartPath(capsule.partPath)
}

function isStructuredChartPartPath(partPath: string): boolean {
	return (
		/(^|\/)charts\/chart\d+\.xml$/i.test(partPath) ||
		/(^|\/)chartEx\/chartEx\d+\.xml$/i.test(partPath)
	)
}

function isDrawingCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.relType === REL_DRAWING ||
		capsule.contentType.includes('drawing+xml') ||
		capsule.partPath.includes('/drawings/drawing')
	)
}

function resolveContentType(partPath: string, contentTypes: ContentTypes): string {
	return resolveContentTypeInfo(partPath, contentTypes).value
}

function resolveContentTypeInfo(
	partPath: string,
	contentTypes: ContentTypes,
): { value: string; source: 'override' | 'default' | 'fallback' } {
	const normalized = partPath.startsWith('/') ? partPath.substring(1) : partPath
	const ct = contentTypes.overrides.get(normalized)
	if (ct) return { value: ct, source: 'override' }

	const ext = partPath.split('.').pop() ?? ''
	const defaultType = contentTypes.defaults.get(ext)
	if (defaultType) return { value: defaultType, source: 'default' }

	return { value: 'application/octet-stream', source: 'fallback' }
}

function capsuleFamily(path: string): string {
	if (path.startsWith('docProps/')) return 'preservedDocumentProperties'
	if (path.includes('/chartsheets/')) return 'preservedChartSheet'
	if (path.includes('/macrosheets/')) return 'preservedMacroSheet'
	if (path.includes('/charts/') || path.includes('/chartEx/')) return 'preservedChart'
	if (path.includes('/queryTables/')) return 'preservedQueryTable'
	if (path.endsWith('/connections.xml')) return 'preservedConnection'
	if (path.includes('/customData/')) return 'preservedPowerQuery'
	if (path.includes('/model/')) return 'preservedDataModel'
	if (path.includes('/drawings/') && path.endsWith('.vml')) return 'preservedVml'
	if (path.includes('/drawings/')) return 'preservedDrawing'
	if (path.includes('/media/')) return 'preservedMedia'
	if (path.includes('/theme/')) return 'preservedTheme'
	if (path.includes('/activeX/')) return 'preservedActiveX'
	if (path.includes('/vbaProject')) return 'preservedMacro'
	if (path.startsWith('_xmlsignatures/')) return 'preservedSignature'
	if (path.includes('/printerSettings/')) return 'preservedPrinterSettings'
	if (path.startsWith('customXml/')) return 'preservedCustomXml'
	if (path.includes('/ctrlProps/')) return 'preservedControl'
	if (/(^|\/)externalLinks\//.test(path)) return 'preservedExternalLink'
	if (/(^|\/)pivotTables\//.test(path) || /(^|\/)pivotCache\//.test(path)) return 'preservedPivot'
	if (
		path.includes('/slicers/') ||
		path.includes('/slicerCaches/') ||
		path.includes('/timelines/') ||
		path.includes('/timelineCaches/')
	) {
		return 'preservedSlicer'
	}
	if (path.includes('/tables/')) return 'preservedTable'
	if (path.includes('/metadata')) return 'preservedMetadata'
	if (path.endsWith('/calcChain.xml')) return 'preservedCalcChain'
	if (/\/comments\d+\.xml$/i.test(path)) return 'preservedComments'
	if (path.includes('/threadedComments/')) return 'preservedThreadedComments'
	return 'preservedOther'
}

function workbookSheetEntryKind(
	relType: string | undefined,
): 'worksheet' | 'chartsheet' | 'macrosheet' {
	if (relType === REL_CHARTSHEET) return 'chartsheet'
	if (relType === REL_MACROSHEET) return 'macrosheet'
	return 'worksheet'
}

function categorizeCapsules(capsules: readonly PreservationCapsule[]): Map<string, string[]> {
	const families = new Map<string, string[]>()
	for (const capsule of capsules) {
		const family = capsuleFamily(capsule.partPath)
		let paths = families.get(family)
		if (!paths) {
			paths = []
			families.set(family, paths)
		}
		paths.push(capsule.partPath)
	}
	return families
}

function collectActiveContent(
	archive: ZipArchive,
	capsules: readonly PreservationCapsule[],
): ActiveContentInfo[] {
	const activeContent: ActiveContentInfo[] = []
	for (const capsule of capsules) {
		const kind = classifyActiveContent(capsule)
		if (!kind) continue
		const entry = archive.get(capsule.partPath)
		const vbaProject =
			kind === 'vbaProject'
				? summarizeVbaProject(archive.readBytes(capsule.partPath) ?? new Uint8Array())
				: undefined
		activeContent.push({
			kind,
			partPath: capsule.partPath,
			contentType: capsule.contentType,
			anchor: capsule.anchor.kind,
			...(capsule.anchor.kind === 'sheet' ? { sheetName: capsule.anchor.sheetName } : {}),
			...(capsule.relType ? { relType: capsule.relType } : {}),
			relationshipCount: capsule.relationships.length,
			...(kind === 'vbaProject' && entry ? { byteSize: entry.uncompressedSize } : {}),
			...(kind === 'vbaProject' ? { opaque: true, executionPolicy: 'blocked' as const } : {}),
			...(kind === 'macroSheet' ? { opaque: true, executionPolicy: 'blocked' as const } : {}),
			...(kind === 'digitalSignature' || kind === 'vbaSignature'
				? {
						invalidationPolicy: 'invalidatedByPackageEdit' as const,
						resigningPolicy: 'notSupported' as const,
					}
				: {}),
			...(vbaProject ? { vbaProject } : {}),
		})
	}
	return activeContent
}

function collectConnectionParts(
	archive: ZipArchive,
	capsules: readonly PreservationCapsule[],
): WorkbookConnectionPartInfo[] {
	const connectionParts: WorkbookConnectionPartInfo[] = []
	for (const capsule of capsules) {
		connectionParts.push(...parseConnectionPartInfos(capsule, readPart(archive, capsule.partPath)))
	}
	return connectionParts
}

function collectDataModelParts(
	capsules: readonly PreservationCapsule[],
): WorkbookDataModelPartInfo[] {
	const dataModelParts: WorkbookDataModelPartInfo[] = []
	for (const capsule of capsules) {
		const part = parseDataModelPartInfo(capsule)
		if (part) dataModelParts.push(part)
	}
	return dataModelParts
}

function collectThreadedCommentPeople(archive: ZipArchive): Map<string, string> {
	const people = new Map<string, string>()
	for (const entry of archive.entries()) {
		if (!entry.path.startsWith('xl/persons/') || !entry.path.endsWith('.xml')) continue
		const xml = readPart(archive, entry.path)
		if (!xml) continue
		for (const [id, displayName] of parseThreadedCommentPersonsXml(xml)) {
			people.set(id, displayName)
		}
	}
	return people
}

function classifyActiveContent(capsule: PreservationCapsule): ActiveContentInfo['kind'] | null {
	const path = capsule.partPath.toLowerCase()
	const contentType = capsule.contentType.toLowerCase()
	const relType = capsule.relType?.toLowerCase() ?? ''
	if (
		path.includes('/macrosheets/') ||
		contentType.includes('macrosheet') ||
		relType.includes('xlmacrosheet')
	) {
		return 'macroSheet'
	}
	if (
		path.includes('vbaprojectsignature') ||
		(contentType.includes('vba') && path.includes('signature'))
	) {
		return 'vbaSignature'
	}
	if (
		path.startsWith('_xmlsignatures/') ||
		contentType.includes('digital-signature') ||
		relType.includes('digital-signature')
	) {
		return 'digitalSignature'
	}
	if (path.includes('vbaproject') || contentType.includes('vbaproject')) return 'vbaProject'
	if (
		path.includes('/activex/') ||
		contentType.includes('activex') ||
		relType.includes('activex')
	) {
		return 'activeX'
	}
	if (path.includes('/ctrlprops/') || contentType.includes('controlproperties'))
		return 'formControl'
	if (path.startsWith('customui/') || contentType.includes('customui')) return 'customUi'
	if (relType.includes('control') || relType.includes('macro')) return 'unknownActiveContent'
	return null
}

function preservedFeatureNote(feature: string): string | undefined {
	if (feature === 'preservedChartSheet') {
		return 'Chartsheets are inventoried and preserved exactly where possible; they are not modeled as worksheet grids.'
	}
	if (feature === 'preservedMacroSheet') {
		return 'Excel 4 macro sheets are inventoried and preserved exactly where possible; macro formulas are not executed or semantically edited.'
	}
	if (feature === 'preservedChart') {
		return 'Chart parts are inventoried and preserved exactly where possible; chart semantics are not yet editable.'
	}
	if (feature === 'preservedPivot') {
		return 'Pivot table and pivot cache parts are inventoried and preserved exactly where possible; pivot execution is not performed headlessly.'
	}
	if (feature === 'preservedDrawing') {
		return 'Drawing parts are inventoried and preserved exactly where possible; drawing-object semantics are not yet editable.'
	}
	if (feature === 'preservedMacro') {
		return 'Macro project bytes are preserved exactly where possible; writes require explicit loss approval because macro semantics are not editable.'
	}
	if (feature === 'preservedActiveX') {
		return 'ActiveX binaries and descriptors are preserved exactly where possible; writes require explicit loss approval because controls are active content.'
	}
	if (feature === 'preservedControl') {
		return 'Form/control property parts are preserved exactly where possible; linked behavior is not semantically editable.'
	}
	if (feature === 'preservedSignature') {
		return 'Digital signature parts are preserved for clean round-trips; generated workbook edits invalidate existing signatures unless the package is re-signed outside Ascend.'
	}
	if (feature === 'preservedQueryTable') {
		return 'Query table parts are inventoried with connection IDs and refresh flags, then preserved exactly where possible.'
	}
	if (feature === 'preservedConnection') {
		return 'Workbook connection metadata is inventoried with connection IDs and refresh flags, then preserved exactly where possible.'
	}
	if (feature === 'preservedPowerQuery') {
		return 'Power Query mashup/customData parts are inventoried and preserved; query execution is not performed headlessly.'
	}
	if (feature === 'preservedThreadedComments') {
		return 'Threaded comments are inventoried with thread/person metadata and preserved; semantic edits require explicit support.'
	}
	if (feature === 'preservedExternalLink') {
		return 'External link package parts are inventoried and preserved; link target edits should use explicit external-link operations.'
	}
	if (feature === 'preservedTheme') {
		return 'Theme parts are inventoried and preserved exactly where possible; full theme editing is not yet first-class.'
	}
	if (feature === 'preservedComments') {
		return 'Classic comment parts are inventoried and preserved; comment text is inspectable when sheet metadata is hydrated.'
	}
	if (feature === 'preservedCalcChain') {
		return 'Calc chain parts are preserved for compatible value edits but treated as rebuildable calculation hints.'
	}
	if (feature === 'preservedDataModel') {
		return 'Workbook data model parts are inventoried and preserved; Power Pivot/data-model execution is not performed headlessly.'
	}
	if (feature === 'preservedDocumentProperties') {
		return 'Document property parts are preserved exactly where possible; semantic property editing is not yet first-class.'
	}
	return undefined
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

	if (formulaFeatures.sharedFormulaSheets.length > 0) {
		features.push({
			feature: 'sharedFormula',
			tier: 'normalized',
			count: formulaFeatures.sharedFormulaSheets.length,
			locations: formulaFeatures.sharedFormulaSheets,
			note: 'Shared formulas are translated into concrete formulas with binding metadata, but shared groups are not yet first-class editable semantics.',
		})
	}

	if (formulaFeatures.arrayFormulaSheets.length > 0) {
		features.push({
			feature: 'arrayFormula',
			tier: 'normalized',
			count: formulaFeatures.arrayFormulaSheets.length,
			locations: formulaFeatures.arrayFormulaSheets,
			note: 'Array formulas are imported with binding metadata, but legacy CSE semantics are not yet modeled as first-class editable semantics.',
		})
	}

	if (formulaFeatures.dynamicArraySheets.length > 0) {
		features.push({
			feature: 'dynamicArray',
			tier: 'normalized',
			count: formulaFeatures.dynamicArraySheets.length,
			locations: formulaFeatures.dynamicArraySheets,
			note: 'Dynamic-array anchors are imported with metadata-backed bindings and normalized formula syntax.',
		})
	}

	const tableLocations = workbook.sheets
		.filter((sheet) => sheet.tables.length > 0)
		.map((sheet) => sheet.name)
	if (tableLocations.length > 0) {
		features.push({
			feature: 'table',
			tier: 'exact',
			count: tableLocations.length,
			locations: tableLocations,
			note: 'Tables round-trip with style info, auto-filter, sort state, and total row semantics.',
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
		const families = categorizeCapsules(capsules)
		for (const [family, paths] of families) {
			const note = preservedFeatureNote(family)
			features.push({
				feature: family,
				tier: 'preserved',
				count: paths.length,
				locations: paths,
				...(note ? { note } : {}),
			})
		}
	}

	if (loadInfo.isPartial) {
		const reasons: string[] = []
		if (!loadInfo.hasAllSheets) reasons.push('only selected sheets are loaded')
		if (!loadInfo.cellsHydrated) reasons.push('sheet cells are not hydrated')
		if (loadInfo.mode === 'values') {
			reasons.push(
				loadInfo.richSheetMetadataHydrated
					? 'formulas, styles, and preservation capsules are not hydrated'
					: 'only cell values are hydrated',
			)
		}
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
	threadedCommentPeople: ReadonlyMap<string, string> = collectThreadedCommentPeople(archive),
): void {
	if (!sheet) return
	const commentsRel = sheetRelationships.find((rel) => rel.type === REL_COMMENTS)
	if (commentsRel) {
		const commentsPath = resolvePath(sheetPath, commentsRel.target)
		const commentsXml = readPart(archive, commentsPath)
		if (commentsXml) {
			for (const [ref, comment] of parseCommentsXml(commentsXml)) {
				sheet.comments.set(ref, comment)
			}
		}
	}
	for (const rel of sheetRelationships.filter(
		(relationship) => relationship.type === REL_THREADED_COMMENT,
	)) {
		const commentsPath = resolvePath(sheetPath, rel.target)
		const commentsXml = readPart(archive, commentsPath)
		if (!commentsXml) continue
		sheet.threadedComments.push(
			...parseThreadedCommentsXml(commentsXml, commentsPath, threadedCommentPeople),
		)
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
		if (!drawingXml) continue
		const relationships = drawingRelsXml ? parseRelationships(drawingRelsXml) : []
		sheet.imageRefs.push(...parseDrawingImageRefs(drawingXml, drawingPath, relationships))
		sheet.drawingObjectRefs.push(...parseDrawingObjectRefs(drawingXml, drawingPath))
	}
}

function attachPivotTables(
	archive: ZipArchive,
	sheetPath: string,
	sheetName: string,
	workbook: Workbook,
	sheetRelationships: readonly Relationship[],
): void {
	for (const rel of sheetRelationships.filter(
		(relationship) => relationship.type === REL_PIVOT_TABLE,
	)) {
		const partPath = resolvePath(sheetPath, rel.target)
		const xml = readPart(archive, partPath)
		if (!xml) continue
		const parsed = parsePivotTableXml(xml, partPath, sheetName)
		if (parsed) workbook.pivotTables.push(parsed)
	}
}
