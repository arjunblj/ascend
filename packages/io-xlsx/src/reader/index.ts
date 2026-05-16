import type {
	ActiveContentInfo,
	CellStyle,
	DefinedNameOptions,
	SheetId,
	StyleId,
	WorkbookConnectionPartInfo,
	WorkbookDataModelPartInfo,
} from '@ascend/core'
import { DEFAULT_STYLE_ID, Workbook } from '@ascend/core'
import type {
	AscendError,
	CellValue,
	CompatibilityReport,
	CompatibilityStatus,
	FeatureReport,
	Result,
} from '@ascend/schema'
import { ascendError, emptyReport, err, ok } from '@ascend/schema'
import { normalizeStoredFormulaText } from '../formula-storage.ts'
import { classifyPackageFeatureFamily } from '../package-graph.ts'
import type { PreservationCapsule } from '../preserve.ts'
import {
	parseActiveXControlInfo,
	parseCustomUiInfo,
	parseDrawingShapeMacroInfos,
	parseFormControlInfo,
	parseVmlControlInfos,
	parseWorksheetControlInfos,
} from './active-content.ts'
import { parseChartXml } from './charts.ts'
import {
	parseCommentsXml,
	parseCommentVmlXml,
	parseThreadedCommentPersonsXml,
	parseThreadedCommentsXml,
} from './comments.ts'
import { parseConnectionPartInfos } from './connections.ts'
import { type ContentTypes, parseContentTypes } from './content-types.ts'
import { parseDataModelPartInfo } from './data-model.ts'
import { parseDocumentProperties } from './doc-props.ts'
import {
	parseDrawingImageRefs,
	parseDrawingObjectRefs,
	parseVmlDrawingObjectRefs,
} from './drawing.ts'
import { maybeDecryptOoxmlPackage } from './encryption.ts'
import { parseExternalLinkInfo } from './external-links.ts'
import { inferLegacyArrayFormulaBlocks } from './legacy-array-inference.ts'
import { parseMacroSheetInfo } from './macro-sheet.ts'
import { parseMetadataXml } from './metadata.ts'
import {
	parseMaterializedPivotCacheRecordsXml,
	parsePivotCacheDefinitionXml,
	parsePivotTableXml,
	parseSlicerCacheXml,
	parseSlicerXml,
	parseTimelineCacheXml,
	parseTimelineXml,
} from './pivots.ts'
import {
	externalLinkPathRelationshipKind,
	getRelsPath,
	isExternalLinkPathRelationshipType,
	parseRelationships,
	REL_CHART,
	REL_CHARTSHEET,
	REL_COMMENTS,
	REL_DRAWING,
	REL_MACROSHEET,
	REL_OFFICE_DOC,
	REL_PIVOT_TABLE,
	REL_QUERY_TABLE,
	REL_SHARED_STRINGS,
	REL_SHEET_METADATA,
	REL_SLICER_CACHE,
	REL_STYLES,
	REL_TABLE,
	REL_THEME,
	REL_THREADED_COMMENT,
	REL_VML_DRAWING,
	REL_WORKSHEET,
	type Relationship,
	resolvePath,
} from './relationships.ts'
import {
	emptySharedStrings,
	parseSharedStrings,
	parseSharedStringsBytes,
	parseSharedStringsChunks,
} from './shared-strings.ts'
import {
	parseSheet,
	parseSheetFormulaOnlyBytes,
	parseSheetFullScalarBytes,
	parseSheetValuesOnlyByteChunks,
	parseSheetValuesOnlyBytes,
	type SheetFormulaFeatures,
	ValueInternPool,
} from './sheet.ts'
import { parseStyles, parseStylesLite } from './styles.ts'
import { parseTable } from './table.ts'
import { parseThemeColorsXml, parseThemeXml } from './theme.ts'
import { summarizeVbaProject } from './vba.ts'
import { type DefinedNameEntry, parseWorkbookXml, type SheetEntry } from './workbook.ts'
import { extractZip, type ZipArchive } from './zip.ts'

const XML_DECODER = new TextDecoder('utf-8')
const VALUES_ONLY_BYTE_PARSE_MIN_BYTES = 128 * 1024
const STREAMED_MAX_ROWS_COMPRESSED_CHUNK_BYTES = 16 * 1024
const CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
const CT_CHARTSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml'
const CT_SHARED_STRINGS =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
const CT_THEME = 'application/vnd.openxmlformats-officedocument.theme+xml'
const REL_CORE_PROPS =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const REL_EXT_PROPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
const REL_CUSTOM_PROPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'

export interface ReadXlsxResult {
	readonly workbook: Workbook
	readonly report: CompatibilityReport
	readonly capsules: PreservationCapsule[]
	readonly loadInfo: ReadXlsxLoadInfo
	readonly sourceArchive?: ZipArchive
}

export interface ReadXlsxOptions {
	readonly mode?: 'full' | 'metadata-only' | 'values' | 'formula'
	readonly sheets?: readonly (string | number)[]
	readonly maxRows?: number
	readonly richMetadata?: boolean
	readonly formulaModeHydrateValues?: boolean
	readonly parseDates?: boolean
	readonly password?: string
	readonly pivotCacheRecordMaterializeLimit?: number | 'all'
}

export interface ReadXlsxLoadInfo {
	readonly mode: 'full' | 'metadata-only' | 'values' | 'formula' | 'selective'
	readonly isPartial: boolean
	readonly cellsHydrated: boolean
	readonly richSheetMetadataHydrated: boolean
	readonly hasAllSheets: boolean
	readonly maxRows?: number
	readonly partialReasons: readonly string[]
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

	return readXlsxArchive(archive, archiveBytes, options)
}

export function readXlsxArchive(
	archive: ZipArchive,
	archiveBytes: Uint8Array | null,
	options: ReadXlsxOptions = {},
): Result<ReadXlsxResult, AscendError> {
	const mode = options.mode ?? 'full'

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
		const workbookPath = resolvePath('', docRel.target)

		const wbXml = readPart(archive, workbookPath)
		if (!wbXml) {
			return err(ascendError('CORRUPT_FILE', `Missing workbook: ${workbookPath}`))
		}
		const wbInfo = parseWorkbookXml(wbXml)
		const selectedSheets = resolveSheetFilter(options.sheets, wbInfo.sheets)

		const wbRelsPath = getRelsPath(workbookPath)
		const wbRelsXml = readPart(archive, wbRelsPath)
		const wbRels = wbRelsXml ? parseRelationships(wbRelsXml) : []
		const effectiveWbRels = recoverWorkbookRelationships(
			archive,
			contentTypes,
			workbookPath,
			wbInfo.sheets,
			wbRels,
		)
		const relMap = new Map(effectiveWbRels.map((r) => [r.id, r]))

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
			contentTypeOverrides: Array.from(contentTypes.overrides, ([partPath, contentType]) => ({
				partPath,
				contentType,
			})),
			sheetEntries: wbInfo.sheets.map((entry) => ({
				kind: workbookSheetEntryKind(relMap.get(entry.rId)?.type),
				sheetId: entry.sheetId,
				name: entry.name,
			})),
		}
		workbook.documentProperties = parseDocumentProperties({
			coreXml: documentPropertyXml(archive, rootRels, REL_CORE_PROPS),
			appXml: documentPropertyXml(archive, rootRels, REL_EXT_PROPS),
			customXml: documentPropertyXml(archive, rootRels, REL_CUSTOM_PROPS),
		})
		for (const relId of wbInfo.externalReferenceRelIds) {
			const rel = relMap.get(relId)
			if (!rel) continue
			const partPath = resolvePath(workbookPath, rel.target)
			workbook.externalReferences.push(partPath)
			const relsXml = readPart(archive, getRelsPath(partPath))
			const linkXml = readPart(archive, partPath)
			const externalLinkInfo = linkXml ? parseExternalLinkInfo(linkXml) : undefined
			const externalLinkRelId = externalLinkInfo?.relationshipId
			const externalBookRelId =
				externalLinkInfo?.kind === 'externalBook' ? externalLinkRelId : undefined
			const linkRelationships = relsXml ? parseRelationships(relsXml) : []
			const linkedByExternalLinkRelId = externalLinkRelId
				? linkRelationships.find(
						(entry) =>
							entry.id === externalLinkRelId && isExternalLinkPathRelationshipType(entry.type),
					)
				: undefined
			const fallbackLinkRelationship = linkRelationships.find((entry) =>
				isExternalLinkPathRelationshipType(entry.type),
			)
			const linkRelationship = linkedByExternalLinkRelId ?? fallbackLinkRelationship
			const linkBindingStatus = linkedByExternalLinkRelId
				? externalLinkInfo?.kind === 'externalBook'
					? 'externalBookRelId'
					: 'externalLinkRelId'
				: linkRelationship
					? 'fallbackPathRelationship'
					: 'missingPathRelationship'
			workbook.externalReferenceDetails.push({
				partPath,
				relId,
				sourcePartPath: workbookPath,
				sourceRelationshipPart: wbRelsPath,
				sourceRelationshipType: rel.type,
				...(rel.rawType ? { sourceRelationshipRawType: rel.rawType } : {}),
				sourceRelationshipRawTarget: rel.target,
				sourceRelationshipResolvedTarget: partPath,
				...(externalLinkInfo?.kind ? { externalLinkKind: externalLinkInfo.kind } : {}),
				...(externalLinkRelId ? { externalLinkRelId } : {}),
				...(externalBookRelId ? { externalBookRelId } : {}),
				...(externalLinkInfo?.ddeService
					? { externalLinkDdeService: externalLinkInfo.ddeService }
					: {}),
				...(externalLinkInfo?.ddeTopic ? { externalLinkDdeTopic: externalLinkInfo.ddeTopic } : {}),
				...(linkRelationship?.id ? { linkRelId: linkRelationship.id } : {}),
				...(linkRelationship ? { linkRelationshipPart: getRelsPath(partPath) } : {}),
				...(linkRelationship?.type
					? { linkRelationshipKind: externalLinkPathRelationshipKind(linkRelationship.type) }
					: {}),
				linkBindingStatus,
				...(linkRelationship?.type ? { linkRelationshipType: linkRelationship.type } : {}),
				...(linkRelationship?.rawType ? { linkRelationshipRawType: linkRelationship.rawType } : {}),
				...(linkRelationship?.target ? { linkRelationshipRawTarget: linkRelationship.target } : {}),
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
				if (parsed) {
					const recordsXml = parsed.recordsPartPath
						? readPart(archive, parsed.recordsPartPath)
						: null
					const records =
						recordsXml && parsed.recordsPartPath
							? parseMaterializedPivotCacheRecordsXml(
									recordsXml,
									parsed.recordsPartPath,
									pivotCacheRecordMaterializeLimit(options),
								)
							: null
					workbook.pivotCaches.push(records ? { ...parsed, records } : parsed)
				}
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

		const ssPart = effectiveWbRels.find((r) => r.type === REL_SHARED_STRINGS)
		const ssPath = ssPart ? resolvePath(workbookPath, ssPart.target) : undefined
		if (ssPath) {
			consumed.add(ssPath)
			workbook.preservedSharedStrings = { path: ssPath }
		}

		const stylesPart = effectiveWbRels.find((r) => r.type === REL_STYLES)
		const stylesPath = stylesPart ? resolvePath(workbookPath, stylesPart.target) : undefined
		if (stylesPath) consumed.add(stylesPath)

		const themePart = effectiveWbRels.find((r) => r.type === REL_THEME)
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
		const sheetFormulaFeaturesByName = new Map<string, SheetFormulaFeatures>()
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
				const chartsheetRelsXml = readPartWithPath(archive, getRelsPath(sheetPath))?.text
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
						chartPartPaths: chartPartPathsForSheetRelationships(
							archive,
							sheetPath,
							chartsheetRelationships,
						),
					})
				}
				continue
			}

			if (rel.type === REL_MACROSHEET) {
				const macroSheetRelsXml = readPartWithPath(archive, getRelsPath(sheetPath))?.text
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
			const sheetRelsPart = readPartWithPath(archive, getRelsPath(sheetPath))
			const sheetRelationships = sheetRelsPart ? parseRelationships(sheetRelsPart.text) : []
			if (sheetRelationships.length > 0) sheetRelsByPath.set(sheetPath, sheetRelationships)
			consumed.add(sheetRelsPart?.path ?? getRelsPath(sheetPath))
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
			const sharedStringOptions = {
				...(valuesOnly || formulaOnly
					? {}
					: {
							normalize: (value: CellValue) => valuePool?.internValue(value) ?? value,
							...(valuePool
								? { normalizeString: (value: string) => valuePool.internStringValue(value) }
								: {}),
						}),
				lazy: valuesOnly || formulaOnly || mode === 'full' || selectedSheets !== null,
			}
			const canStreamSharedStrings =
				ssPath !== undefined &&
				valuesOnly &&
				!hydrateRichSheetMetadata &&
				options.maxRows !== undefined
			const canUseSharedStringsByteParser =
				ssPath !== undefined &&
				!canStreamSharedStrings &&
				!sharedStringOptions.lazy &&
				(archive.get(ssPath)?.uncompressedSize ?? 0) >= VALUES_ONLY_BYTE_PARSE_MIN_BYTES
			const ssBytes = canUseSharedStringsByteParser ? readPartBytes(archive, ssPath) : undefined
			const ssXml =
				ssPath && !canStreamSharedStrings && !canUseSharedStringsByteParser
					? readPart(archive, ssPath)
					: undefined
			const sharedStrings = ssPath
				? canStreamSharedStrings
					? parseSharedStringsChunks(
							archive.readTextChunks(ssPath, STREAMED_MAX_ROWS_COMPRESSED_CHUNK_BYTES, {
								preferStreaming: true,
							}),
							{
								fallback: () => {
									const xml = readPart(archive, ssPath)
									return xml ? parseSharedStrings(xml, sharedStringOptions) : emptySharedStrings()
								},
							},
						)
					: ssBytes
						? parseSharedStringsBytes(ssBytes, sharedStringOptions)
						: ssXml
							? parseSharedStrings(ssXml, sharedStringOptions)
							: emptySharedStrings()
				: emptySharedStrings()

			const parseLiteStyles = !((valuesOnly || formulaOnly) && options.parseDates === false)
			const stylesXml =
				stylesPath && (parseLiteStyles || !(valuesOnly || formulaOnly))
					? readPart(archive, stylesPath)
					: undefined
			let styleIds: StyleId[]
			let isDateFormat: boolean[]
			let differentialStyles: readonly CellStyle[]
			if (formulaOnly && options.richMetadata === true && stylesXml) {
				const parsedStyles = parseStyles(stylesXml)
				styleIds = registerStyles(workbook, parsedStyles.cellStyles)
				isDateFormat = parsedStyles.isDateFormat
				differentialStyles = parsedStyles.differentialStyles
				workbook.styleMetadata = parsedStyles.metadata
				workbook.differentialStyles.push(...parsedStyles.differentialStyles)
			} else if (valuesOnly || formulaOnly) {
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
				const sheetEntryInfo = archive.get(entry.path)
				const canUseStreamedMaxRowsParser =
					valuesOnly &&
					!hydrateRichSheetMetadata &&
					options.maxRows !== undefined &&
					(sheetEntryInfo?.uncompressedSize ?? 0) >= VALUES_ONLY_BYTE_PARSE_MIN_BYTES
				const canUseValuesOnlyByteParser =
					valuesOnly &&
					!hydrateRichSheetMetadata &&
					!canUseStreamedMaxRowsParser &&
					(sheetEntryInfo?.uncompressedSize ?? 0) >= VALUES_ONLY_BYTE_PARSE_MIN_BYTES
				const canUseFormulaOnlyByteParser =
					formulaOnly &&
					!hydrateRichSheetMetadata &&
					(sheetEntryInfo?.uncompressedSize ?? 0) >= VALUES_ONLY_BYTE_PARSE_MIN_BYTES
				const canUseFullScalarByteParser =
					hydrateRichSheetMetadata &&
					(sheetEntryInfo?.uncompressedSize ?? 0) >= VALUES_ONLY_BYTE_PARSE_MIN_BYTES
				const sheetBytes =
					canUseValuesOnlyByteParser || canUseFormulaOnlyByteParser || canUseFullScalarByteParser
						? readPartBytes(archive, entry.path)
						: undefined
				let sheetXml =
					(sheetBytes === undefined && !canUseStreamedMaxRowsParser) ||
					(hydrateRichSheetMetadata && !canUseFullScalarByteParser)
						? readPart(archive, entry.path)
						: undefined
				if (sheetBytes === undefined && !sheetXml && !canUseStreamedMaxRowsParser) continue
				const sheetRelsPart = hydrateRichSheetMetadata
					? readPartWithPath(archive, getRelsPath(entry.path))
					: undefined
				const sheetRelsXml = sheetRelsPart?.text
				const sheetRelationships = sheetRelsXml
					? parseRelationships(sheetRelsXml)
					: (sheetRelsByPath.get(entry.path) ?? [])
				if (sheetRelationships.length > 0) sheetRelsByPath.set(entry.path, sheetRelationships)
				const sheetFormulaFeatures: SheetFormulaFeatures = {
					hasPlainFormula: false,
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
					formulaModeHydrateValues: options.formulaModeHydrateValues ?? true,
					formulaFeatures: sheetFormulaFeatures,
					fullScalarNumberSpanScratch: [],
					fullScalarCellOutScratch: {
						row: 0,
						col: 0,
						numberValue: undefined,
						sharedStringIndex: -1,
						booleanRaw: -1,
						stringStart: -1,
						stringEnd: -1,
						stringHasEntity: false,
						styleIdx: 0,
					},
					...(valuePool ? { valuePool } : {}),
					...(metadata ? { metadata } : {}),
					...(options.maxRows !== undefined ? { maxRows: options.maxRows } : {}),
				}
				const resolvedSheetId = (sheetPathToAnchor.get(entry.path)?.sheetId ??
					entry.name) as SheetId
				let sheet = canUseStreamedMaxRowsParser
					? parseSheetValuesOnlyByteChunks(
							entry.name,
							archive.readByteChunks(entry.path, STREAMED_MAX_ROWS_COMPRESSED_CHUNK_BYTES, {
								preferStreaming: true,
							}),
							sheetCtx,
							resolvedSheetId,
						)
					: sheetBytes
						? canUseValuesOnlyByteParser
							? parseSheetValuesOnlyBytes(entry.name, sheetBytes, sheetCtx, resolvedSheetId)
							: canUseFormulaOnlyByteParser
								? parseSheetFormulaOnlyBytes(entry.name, sheetBytes, sheetCtx, resolvedSheetId)
								: parseSheetFullScalarBytes(entry.name, sheetBytes, sheetCtx, resolvedSheetId)
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
				sheetFormulaFeaturesByName.set(entry.name, sheetFormulaFeatures)
				if (hydrateRichSheetMetadata) {
					attachComments(archive, entry.path, sheet, sheetRelationships)
					attachDrawingImages(archive, entry.path, sheet, sheetRelationships)
					attachPivotTables(archive, entry.path, entry.name, workbook, sheetRelationships)
				}
				if (hydrateRichSheetMetadata) {
					attachTables(archive, contentTypes, entry.path, sheet, sheetRelationships)
				}
				sheet.state = entry.state
				if (hydrateRichSheetMetadata) {
					sheet.preservedXml = {
						partPath: entry.path,
						...(sheetRelsPart ? { relsPath: sheetRelsPart.path } : {}),
					}
				}
				workbook.sheets.push(sheet)
			}
		}

		for (const dn of wbInfo.definedNames) {
			const options = definedNameOptions(dn)
			if (dn.localSheetId !== undefined) {
				const sourceSheet = wbInfo.sheets[dn.localSheetId]
				const sheet = sourceSheet ? workbook.getSheet(sourceSheet.name) : undefined
				if (sheet) {
					workbook.definedNames.add(
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
			workbook.definedNames.add(dn.name, normalizeStoredFormulaText(dn.formula), undefined, options)
		}
		if (!valuesOnly) {
			const legacyArrayCandidateSheets = workbook.sheets.filter((sheet) => {
				const features = sheetFormulaFeaturesByName.get(sheet.name)
				return features?.hasPlainFormula === true
			})
			for (const sheetName of inferLegacyArrayFormulaBlocks(workbook, legacyArrayCandidateSheets)) {
				if (!formulaFeatures.arrayFormulaSheets.includes(sheetName)) {
					formulaFeatures.arrayFormulaSheets.push(sheetName)
				}
			}
		}

		const sourceSheetNames = sourceWorksheetNames
		const loadedSheetNames = sheetsToParse.map((sheet) => sheet.name)
		const hasAllSheets = loadedSheetNames.length === sourceSheetNames.length
		const cellsHydrated = mode !== 'metadata-only'
		const richSheetMetadataHydrated = hydrateRichSheetMetadata
		const fidelityPartial = mode === 'values' || mode === 'formula' || options.maxRows !== undefined
		const isPartial = !hasAllSheets || !cellsHydrated || fidelityPartial
		const partialReasons = buildPartialLoadReasons({
			mode: selectedSheets ? 'selective' : mode,
			isPartial,
			cellsHydrated,
			richSheetMetadataHydrated,
			hasAllSheets,
			...(options.maxRows !== undefined ? { maxRows: options.maxRows } : {}),
		})
		const loadInfo: ReadXlsxLoadInfo = {
			mode: selectedSheets ? 'selective' : mode,
			isPartial,
			cellsHydrated,
			richSheetMetadataHydrated,
			hasAllSheets,
			...(options.maxRows !== undefined ? { maxRows: options.maxRows } : {}),
			partialReasons,
			sourceSheetNames,
			loadedSheetNames,
		}
		const packageInventory = collectCapsules(
			archive,
			consumed,
			contentTypes,
			rootRels,
			effectiveWbRels,
			sheetPathToAnchor,
			workbookPath,
			sheetRelsByPath,
		)
		const capsules = isPartial ? [] : packageInventory
		workbook.activeContent.push(
			...collectActiveContent(archive, packageInventory, sheetPathToAnchor, sheetRelsByPath, {
				hydrateOpaqueBinarySummaries: !isPartial,
				hydrateWorksheetControlDetails: !isPartial,
			}),
		)
		if (!isPartial) workbook.connectionParts.push(...collectConnectionParts(archive, capsules))
		if (!isPartial) workbook.dataModelParts.push(...collectDataModelParts(capsules))
		if (!isPartial) attachChartParts(archive, workbook, capsules)
		const report = buildReport(contentTypes, formulaFeatures, workbook, packageInventory, loadInfo)
		if (loadInfo.isPartial) {
			workbook.sourceArchiveBytes = null
			return ok({ workbook, report, capsules: [], loadInfo })
		}

		return ok({ workbook, report, capsules, loadInfo, sourceArchive: archive })
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

function definedNameOptions(dn: DefinedNameEntry): DefinedNameOptions {
	return {
		...(dn.hidden !== undefined ? { hidden: dn.hidden } : {}),
		...(dn.extraAttributes && dn.extraAttributes.length > 0
			? { extraAttributes: dn.extraAttributes }
			: {}),
	}
}

function documentPropertyXml(
	archive: ZipArchive,
	relationships: readonly Relationship[],
	type: string,
): string | undefined {
	const rel = relationships.find(
		(relationship) => relationship.type === type && relationship.targetMode !== 'External',
	)
	return rel ? readPart(archive, resolvePath('', rel.target)) : undefined
}

function readPart(archive: ZipArchive, path: string): string | undefined {
	return archive.readText(path)
}

function readPartWithPath(
	archive: ZipArchive,
	path: string,
): { readonly path: string; readonly text: string } | undefined {
	const text = archive.readText(path)
	if (text !== undefined) return { path, text }
	const resolvedPath = archive.resolvePathCaseInsensitive(path)
	if (!resolvedPath || resolvedPath === path) return undefined
	const resolvedText = archive.readText(resolvedPath)
	return resolvedText === undefined ? undefined : { path: resolvedPath, text: resolvedText }
}

function readPartBytes(archive: ZipArchive, path: string): Uint8Array | undefined {
	return archive.readBytes(path)
}

function pivotCacheRecordMaterializeLimit(options: ReadXlsxOptions): number {
	if (options.pivotCacheRecordMaterializeLimit === 'all') return Number.POSITIVE_INFINITY
	if (typeof options.pivotCacheRecordMaterializeLimit === 'number') {
		return Math.max(0, Math.floor(options.pivotCacheRecordMaterializeLimit))
	}
	return 2048
}

function recoverWorkbookRelationships(
	archive: ZipArchive,
	contentTypes: ContentTypes,
	workbookPath: string,
	sheets: readonly SheetEntry[],
	relationships: readonly Relationship[],
): Relationship[] {
	const recovered = [...relationships]
	const existingRelIds = new Set(recovered.map((rel) => rel.id))
	const existingTargets = new Set(recovered.map((rel) => resolvePath(workbookPath, rel.target)))

	const worksheetParts = availablePartsForContentType(archive, contentTypes, CT_WORKSHEET).filter(
		(path) => !existingTargets.has(path),
	)
	const chartsheetParts = availablePartsForContentType(archive, contentTypes, CT_CHARTSHEET).filter(
		(path) => !existingTargets.has(path),
	)
	const missingSheets = sheets.filter((sheet) => !existingRelIds.has(sheet.rId))
	if (missingSheets.length > 0) {
		for (const { sheet, partPath, type } of inferMissingWorkbookSheetRelationships(
			missingSheets,
			worksheetParts,
			chartsheetParts,
		)) {
			recovered.push({
				id: sheet.rId,
				type,
				target: relationshipTargetFromWorkbook(workbookPath, partPath),
			})
			existingRelIds.add(sheet.rId)
			existingTargets.add(partPath)
		}
	}

	if (relationships.length > 0 && missingSheets.length === 0) return recovered

	recoverWorkbookPartRelationship(
		recovered,
		existingRelIds,
		existingTargets,
		workbookPath,
		archive,
		contentTypes,
		REL_SHARED_STRINGS,
		CT_SHARED_STRINGS,
		'xl/sharedStrings.xml',
	)
	recoverWorkbookPartRelationship(
		recovered,
		existingRelIds,
		existingTargets,
		workbookPath,
		archive,
		contentTypes,
		REL_STYLES,
		CT_STYLES,
		'xl/styles.xml',
	)
	recoverWorkbookPartRelationship(
		recovered,
		existingRelIds,
		existingTargets,
		workbookPath,
		archive,
		contentTypes,
		REL_THEME,
		CT_THEME,
		'xl/theme/theme1.xml',
	)

	return recovered
}

function inferMissingWorkbookSheetRelationships(
	sheets: readonly SheetEntry[],
	worksheetParts: readonly string[],
	chartsheetParts: readonly string[],
): Array<{ sheet: SheetEntry; partPath: string; type: string }> {
	if (sheets.length === 0) return []
	const candidates = [
		...worksheetParts.map((partPath) => ({ partPath, type: REL_WORKSHEET })),
		...chartsheetParts.map((partPath) => ({ partPath, type: REL_CHARTSHEET })),
	].sort((a, b) => comparePartPaths(a.partPath, b.partPath))
	if (candidates.length < sheets.length) return []
	const inferred: Array<{ sheet: SheetEntry; partPath: string; type: string }> = []
	for (let index = 0; index < sheets.length; index++) {
		const sheet = sheets[index]
		const candidate = candidates[index]
		if (!sheet || !candidate) continue
		inferred.push({ sheet, partPath: candidate.partPath, type: candidate.type })
	}
	return inferred
}

function recoverWorkbookPartRelationship(
	relationships: Relationship[],
	existingRelIds: Set<string>,
	existingTargets: Set<string>,
	workbookPath: string,
	archive: ZipArchive,
	contentTypes: ContentTypes,
	type: string,
	contentType: string,
	conventionalPath: string,
): void {
	if (relationships.some((rel) => rel.type === type)) return
	const partPath = firstAvailablePartForContentType(
		archive,
		contentTypes,
		contentType,
		conventionalPath,
	)
	if (!partPath || existingTargets.has(partPath)) return
	relationships.push({
		id: nextRecoveredRelId(existingRelIds),
		type,
		target: relationshipTargetFromWorkbook(workbookPath, partPath),
	})
	existingTargets.add(partPath)
}

function firstAvailablePartForContentType(
	archive: ZipArchive,
	contentTypes: ContentTypes,
	contentType: string,
	conventionalPath: string,
): string | undefined {
	const contentTypePart = availablePartsForContentType(archive, contentTypes, contentType)[0]
	if (contentTypePart) return contentTypePart
	return archive.has(conventionalPath) ? conventionalPath : undefined
}

function availablePartsForContentType(
	archive: ZipArchive,
	contentTypes: ContentTypes,
	contentType: string,
): string[] {
	return [...contentTypes.overrides]
		.filter(
			([path, overrideContentType]) => overrideContentType === contentType && archive.has(path),
		)
		.map(([path]) => path)
		.sort(comparePartPaths)
}

function relationshipTargetFromWorkbook(workbookPath: string, partPath: string): string {
	const workbookDir = workbookPath.substring(0, workbookPath.lastIndexOf('/') + 1)
	return partPath.startsWith(workbookDir) ? partPath.slice(workbookDir.length) : `/${partPath}`
}

function nextRecoveredRelId(existingRelIds: Set<string>): string {
	let index = 1
	for (;;) {
		const relId = `ascendRecoveredRel${index}`
		if (!existingRelIds.has(relId)) {
			existingRelIds.add(relId)
			return relId
		}
		index++
	}
}

function comparePartPaths(a: string, b: string): number {
	return a.localeCompare(b, 'en', { numeric: true })
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

function chartPartPathsForSheetRelationships(
	archive: ZipArchive,
	sheetPath: string,
	relationships: readonly Relationship[],
): string[] {
	const chartPartPaths: string[] = []
	for (const rel of relationships) {
		if (rel.targetMode === 'External') continue
		if (rel.type === REL_CHART) {
			chartPartPaths.push(resolvePath(sheetPath, rel.target))
			continue
		}
		if (rel.type !== REL_DRAWING) continue
		const drawingPath = resolvePath(sheetPath, rel.target)
		const drawingRelsXml = readPart(archive, getRelsPath(drawingPath))
		if (!drawingRelsXml) continue
		for (const drawingRel of parseRelationships(drawingRelsXml)) {
			if (drawingRel.type !== REL_CHART || drawingRel.targetMode === 'External') continue
			chartPartPaths.push(resolvePath(drawingPath, drawingRel.target))
		}
	}
	return [...new Set(chartPartPaths)]
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
		let relTypeRaw: string | undefined
		let relId: string | undefined
		const directSheetAnchor = sheetPathToAnchor.get(partPath)
		if (directSheetAnchor) {
			anchor = {
				kind: 'sheet',
				sheetId: directSheetAnchor.sheetId,
				sheetName: directSheetAnchor.sheetName,
			}
		}

		const wbRef = wbRelByTarget.get(partPath)
		if (wbRef) {
			relType = wbRef.type
			relTypeRaw = wbRef.rawType
			relId = wbRef.id
		}
		const rootRef = rootRelByTarget.get(partPath)
		if (rootRef) {
			relType = rootRef.type
			relTypeRaw = rootRef.rawType
			relId = rootRef.id
		}

		const sheetRef = sheetRelByTarget.get(partPath)
		if (sheetRef) {
			anchor = { kind: 'sheet', sheetId: sheetRef.sheetId, sheetName: sheetRef.sheetName }
			relType = sheetRef.rel.type
			relTypeRaw = sheetRef.rel.rawType
			relId = sheetRef.rel.id
		}

		const capsuleRelsXml = readPart(archive, getRelsPath(partPath))
		const relationships = capsuleRelsXml
			? parseRelationships(capsuleRelsXml).map((r) => ({
					id: r.id,
					type: r.type,
					...(r.rawType ? { rawType: r.rawType } : {}),
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
		if (relTypeRaw) capsule.relTypeRaw = relTypeRaw
		if (relId) capsule.relId = relId
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

function capsuleFamily(capsule: PreservationCapsule): string {
	return classifyPackageFeatureFamily(capsule.partPath, capsule.contentType, capsule.relType)
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
		const family = capsuleFamily(capsule)
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
	sheetPathToAnchor: ReadonlyMap<string, { sheetId: string; sheetName: string }>,
	sheetRelsByPath: ReadonlyMap<string, readonly Relationship[]>,
	options: {
		readonly hydrateOpaqueBinarySummaries: boolean
		readonly hydrateWorksheetControlDetails: boolean
	},
): ActiveContentInfo[] {
	const activeContent: ActiveContentInfo[] = []
	const worksheetControls =
		options.hydrateWorksheetControlDetails && hasWorksheetControlActiveContent(capsules)
			? collectWorksheetControls(archive, sheetPathToAnchor, sheetRelsByPath)
			: new Map<string, NonNullable<ActiveContentInfo['worksheetControl']>>()
	const capsuleRelationshipByTarget = mapCapsuleRelationshipsByTarget(capsules)
	for (const capsule of capsules) {
		const kind = classifyActiveContent(capsule)
		if (!kind) continue
		const sourceRelationship = capsuleRelationshipByTarget.get(capsule.partPath)
		const worksheetControl =
			capsule.anchor.kind === 'sheet' && capsule.anchor.sheetName && capsule.relId
				? worksheetControls.get(controlKey(capsule.anchor.sheetName, capsule.relId))
				: undefined
		const entry = archive.get(capsule.partPath)
		const vbaProject =
			kind === 'vbaProject' && options.hydrateOpaqueBinarySummaries
				? summarizeVbaProject(archive.readBytes(capsule.partPath) ?? new Uint8Array())
				: undefined
		const activeX =
			kind === 'activeX'
				? parseActiveXControlInfo(
						readXmlMetadataPart(archive, capsule.partPath, capsule.contentType),
						capsule.relationships,
					)
				: undefined
		const formControl =
			kind === 'formControl'
				? parseFormControlInfo(readXmlMetadataPart(archive, capsule.partPath, capsule.contentType))
				: undefined
		const customUi =
			kind === 'customUi'
				? parseCustomUiInfo(readXmlMetadataPart(archive, capsule.partPath, capsule.contentType))
				: undefined
		const relType = capsule.relType ?? sourceRelationship?.relationship.type
		const sourceRelationshipId = capsule.relId ?? sourceRelationship?.relationship.id
		activeContent.push({
			kind,
			partPath: capsule.partPath,
			contentType: capsule.contentType,
			anchor: capsule.anchor.kind,
			...(capsule.anchor.kind === 'sheet' ? { sheetName: capsule.anchor.sheetName } : {}),
			...(sourceRelationship ? { sourcePartPath: sourceRelationship.sourcePartPath } : {}),
			...(relType ? { relType } : {}),
			...(sourceRelationshipId ? { sourceRelationshipId } : {}),
			relationshipCount: capsule.relationships.length,
			...(kind === 'vbaProject' && entry ? { byteSize: entry.uncompressedSize } : {}),
			...(kind === 'vbaProject' ? { opaque: true, executionPolicy: 'blocked' as const } : {}),
			...(kind === 'macroSheet' ? { opaque: true, executionPolicy: 'blocked' as const } : {}),
			...(kind === 'activeX' ||
			kind === 'unknownActiveContent' ||
			(kind === 'formControl' && formControl?.macro)
				? { executionPolicy: 'blocked' as const }
				: {}),
			...(kind === 'customUi' ? { executionPolicy: 'blocked' as const } : {}),
			...(kind === 'digitalSignature' || kind === 'vbaSignature'
				? {
						invalidationPolicy: 'invalidatedByPackageEdit' as const,
						resigningPolicy: 'notSupported' as const,
					}
				: {}),
			...(vbaProject ? { vbaProject } : {}),
			...(activeX ? { activeX } : {}),
			...(formControl ? { formControl } : {}),
			...(customUi ? { customUi } : {}),
			...(worksheetControl ? { worksheetControl } : {}),
		})
	}
	activeContent.push(
		...collectDrawingShapeMacros(archive, capsules, sheetPathToAnchor, sheetRelsByPath),
	)
	return activeContent
}

function collectDrawingShapeMacros(
	archive: ZipArchive,
	capsules: readonly PreservationCapsule[],
	sheetPathToAnchor: ReadonlyMap<string, { sheetId: string; sheetName: string }>,
	sheetRelsByPath: ReadonlyMap<string, readonly Relationship[]>,
): ActiveContentInfo[] {
	const capsuleByPath = new Map(capsules.map((capsule) => [capsule.partPath, capsule]))
	const activeContent: ActiveContentInfo[] = []
	for (const [sheetPath, anchor] of sheetPathToAnchor) {
		const relationships = sheetRelsByPath.get(sheetPath) ?? []
		for (const rel of relationships) {
			if (rel.type !== REL_DRAWING || rel.targetMode === 'External') continue
			const drawingPath = resolvePath(sheetPath, rel.target)
			const xml = readPart(archive, drawingPath)
			if (!xml) continue
			const capsule = capsuleByPath.get(drawingPath)
			for (const shapeMacro of parseDrawingShapeMacroInfos(xml)) {
				activeContent.push({
					kind: 'shapeMacro',
					partPath: drawingPath,
					contentType:
						capsule?.contentType ?? 'application/vnd.openxmlformats-officedocument.drawing+xml',
					anchor: 'sheet',
					sheetName: anchor.sheetName,
					sourcePartPath: sheetPath,
					relType: rel.type,
					sourceRelationshipId: rel.id,
					relationshipCount: capsule?.relationships.length ?? 0,
					executionPolicy: 'blocked',
					shapeMacro,
				})
			}
		}
	}
	return activeContent
}

function hasWorksheetControlActiveContent(capsules: readonly PreservationCapsule[]): boolean {
	for (const capsule of capsules) {
		if (capsule.anchor.kind !== 'sheet' || !capsule.relId) continue
		const kind = classifyActiveContent(capsule)
		if (kind === 'activeX' || kind === 'formControl' || kind === 'unknownActiveContent') {
			return true
		}
	}
	return false
}

function mapCapsuleRelationshipsByTarget(capsules: readonly PreservationCapsule[]): Map<
	string,
	{
		readonly sourcePartPath: string
		readonly relationship: PreservationCapsule['relationships'][number]
	}
> {
	const byTarget = new Map<
		string,
		{
			readonly sourcePartPath: string
			readonly relationship: PreservationCapsule['relationships'][number]
		}
	>()
	for (const capsule of capsules) {
		for (const relationship of capsule.relationships) {
			if (relationship.targetMode === 'External') continue
			const targetPath = resolvePath(capsule.partPath, relationship.target)
			if (!byTarget.has(targetPath)) {
				byTarget.set(targetPath, { sourcePartPath: capsule.partPath, relationship })
			}
		}
	}
	return byTarget
}

function collectWorksheetControls(
	archive: ZipArchive,
	sheetPathToAnchor: ReadonlyMap<string, { sheetId: string; sheetName: string }>,
	sheetRelsByPath: ReadonlyMap<string, readonly Relationship[]>,
): Map<string, NonNullable<ActiveContentInfo['worksheetControl']>> {
	const controlsBySheetRel = new Map<string, NonNullable<ActiveContentInfo['worksheetControl']>>()
	for (const [sheetPath, anchor] of sheetPathToAnchor) {
		const sheetXml = readPart(archive, sheetPath)
		if (!sheetXml || !sheetXml.includes('<controls')) continue
		const relationships = sheetRelsByPath.get(sheetPath) ?? []
		const vmlControls = relationships
			.filter((rel) => rel.type === REL_VML_DRAWING)
			.flatMap((rel) => {
				const vmlPath = resolvePath(sheetPath, rel.target)
				const vmlXml = readPart(archive, vmlPath)
				const vmlRelsXml = readPart(archive, getRelsPath(vmlPath))
				const vmlRelationships = vmlRelsXml ? parseRelationships(vmlRelsXml) : []
				return parseVmlControlInfos(vmlXml, vmlPath, vmlRelationships)
			})
		for (const control of parseWorksheetControlInfos(
			sheetXml,
			sheetPath,
			relationships,
			vmlControls,
		)) {
			if (!control.relationshipId) continue
			controlsBySheetRel.set(controlKey(anchor.sheetName, control.relationshipId), control)
		}
	}
	return controlsBySheetRel
}

function controlKey(sheetName: string, relId: string): string {
	return `${sheetName}\u0000${relId}`
}

function readXmlMetadataPart(
	archive: ZipArchive,
	partPath: string,
	contentType: string,
): string | undefined {
	const lowerPath = partPath.toLowerCase()
	const lowerContentType = contentType.toLowerCase()
	if (!lowerPath.endsWith('.xml') && !lowerContentType.includes('xml')) return undefined
	return readPart(archive, partPath)
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
	if (isCustomUiPart(path, contentType, relType)) return 'customUi'
	if (relType.includes('control') || relType.includes('macro')) return 'unknownActiveContent'
	return null
}

function isCustomUiPart(path: string, contentType: string, relType: string): boolean {
	return (
		path.startsWith('customui/') ||
		contentType.includes('customui') ||
		relType.endsWith('/relationships/ui/extensibility')
	)
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
	if (feature === 'preservedChartStyle') {
		return 'Chart style parts are inventoried separately from chart definitions and preserved exactly where possible.'
	}
	if (feature === 'preservedChartColor') {
		return 'Chart color style parts are inventoried separately from chart definitions and preserved exactly where possible.'
	}
	if (feature === 'preservedPivot') {
		return 'Pivot table and pivot cache parts are inventoried and preserved exactly where possible; pivot execution is not performed headlessly.'
	}
	if (feature === 'preservedSlicer') {
		return 'Slicer and slicer cache parts are inventoried with cache/item metadata and preserved exactly where possible.'
	}
	if (feature === 'preservedTimeline') {
		return 'Timeline and timeline cache parts are inventoried with date-range state and preserved exactly where possible.'
	}
	if (feature === 'preservedDrawing') {
		return 'Drawing parts are inventoried and preserved exactly where possible; drawing-object semantics are not yet editable.'
	}
	if (feature === 'preservedVml') {
		return 'VML drawing parts are inventoried and preserved, including legacy comment and form-control drawing anchors.'
	}
	if (feature === 'preservedMedia') {
		return 'Media parts are relationship-inventoried and preserved exactly where possible.'
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
	if (feature === 'preservedTable') {
		return 'Table parts are parsed where supported and preserved with package relationship identity across safe edits.'
	}
	if (feature === 'preservedExternalLink') {
		return 'External link package parts are inventoried and preserved; link target edits should use explicit external-link operations.'
	}
	if (feature === 'preservedTheme') {
		return 'Theme parts are inventoried and preserved exactly where possible; full theme editing is not yet first-class.'
	}
	if (feature === 'preservedStyles') {
		return 'Style parts are parsed into the style registry and preserved or patched according to style edit scope.'
	}
	if (feature === 'preservedMetadata') {
		return 'Workbook metadata sidecars are inventoried and preserved; unsupported rich metadata remains inspect-only.'
	}
	if (feature === 'preservedEmbedding') {
		return 'Embedded object package parts are inventoried and preserved exactly where possible; embedded payload execution is blocked.'
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
		return 'Document property parts are preserved exactly when untouched; core, app, and custom docProps are inspectable and editable through setDocumentProperties.'
	}
	if (feature === 'preservedCustomXml') {
		return 'Custom XML data and workbook XML map parts are inventoried and preserved exactly where possible; XML mappings are not semantically editable.'
	}
	if (feature === 'preservedRevision') {
		return 'Workbook revision tracking parts are inventoried and preserved exactly where possible; tracked-change semantics are inspect-only.'
	}
	if (feature === 'preservedCustomUi') {
		return 'Office RibbonX custom UI parts are inventoried with callback metadata and preserved; callback execution is blocked and semantic editing is not yet supported.'
	}
	if (feature === 'preservedVendorSecurity') {
		return 'Vendor security policy and encrypted sidecar parts are inventoried and preserved exactly; Ascend does not execute or decrypt vendor protection payloads.'
	}
	if (feature === 'preservedWorksheetSidecar') {
		return 'Worksheet-like sidecar parts without workbook relationships are preserved exactly but not modeled as active sheets.'
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

	if (
		workbook.calcSettings.calcMode === 'manual' ||
		workbook.calcSettings.fullCalcOnLoad ||
		workbook.calcSettings.calcCompleted === false ||
		workbook.calcSettings.forceFullCalc === true
	) {
		const reasons: string[] = []
		if (workbook.calcSettings.calcMode === 'manual') reasons.push('manual calculation mode')
		if (workbook.calcSettings.fullCalcOnLoad) reasons.push('full recalculation requested on load')
		if (workbook.calcSettings.calcCompleted === false) reasons.push('calculation not completed')
		if (workbook.calcSettings.forceFullCalc === true) reasons.push('forced full recalculation')
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
		const reasons =
			loadInfo.partialReasons.length > 0
				? loadInfo.partialReasons
				: buildPartialLoadReasons(loadInfo)
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

function buildPartialLoadReasons(
	loadInfo: Pick<
		ReadXlsxLoadInfo,
		| 'mode'
		| 'isPartial'
		| 'cellsHydrated'
		| 'richSheetMetadataHydrated'
		| 'hasAllSheets'
		| 'maxRows'
	>,
): readonly string[] {
	if (!loadInfo.isPartial) return []
	const reasons: string[] = []
	if (!loadInfo.hasAllSheets) reasons.push('only selected sheets are loaded')
	if (loadInfo.maxRows !== undefined) {
		reasons.push(`only the first ${loadInfo.maxRows} row(s) are hydrated per loaded sheet`)
	}
	if (!loadInfo.cellsHydrated) reasons.push('sheet cells are not hydrated')
	if (loadInfo.mode === 'values') {
		reasons.push(
			loadInfo.richSheetMetadataHydrated
				? 'formulas, styles, and preservation capsules are not hydrated'
				: 'only cell values are hydrated',
		)
	}
	if (loadInfo.mode === 'formula') reasons.push('styles and preservation capsules are not hydrated')
	if (loadInfo.mode === 'selective' && reasons.length === 0) {
		reasons.push('only selected workbook data is hydrated')
	}
	return reasons
}

function attachTables(
	archive: ZipArchive,
	contentTypes: ContentTypes,
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
		const tableContentType = resolveContentTypeInfo(tablePath, contentTypes)
		const tableRelsXml = readPart(archive, getRelsPath(tablePath))
		const tableRelationships = tableRelsXml
			? parseRelationships(tableRelsXml).filter((entry) => entry.type === REL_QUERY_TABLE)
			: []
		const table = parseTable(tableXml, sheet.id, {
			tablePath,
			contentType: tableContentType.value,
			contentTypeSource: tableContentType.source,
			sourcePartPath: sheetPath,
			sourceRelationshipPart: getRelsPath(sheetPath),
			sourceRelationship: rel,
			sourceRelationshipResolvedTarget: tablePath,
			relationships: tableRelationships,
		})
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
	const vmlRel = sheetRelationships.find((rel) => rel.type === REL_VML_DRAWING)
	if (vmlRel && sheet.comments.size > 0) {
		const vmlPath = resolvePath(sheetPath, vmlRel.target)
		const vmlXml = readPart(archive, vmlPath)
		if (vmlXml) {
			for (const [ref, legacyDrawing] of parseCommentVmlXml(vmlXml)) {
				const comment = sheet.comments.get(ref)
				if (comment) sheet.comments.set(ref, { ...comment, legacyDrawing })
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
		sheet.drawingObjectRefs.push(...parseDrawingObjectRefs(drawingXml, drawingPath, relationships))
	}
	for (const vmlRel of sheetRelationships.filter((rel) => rel.type === REL_VML_DRAWING)) {
		const vmlPath = resolvePath(sheetPath, vmlRel.target)
		const vmlXml = readPart(archive, vmlPath)
		if (!vmlXml) continue
		const vmlRelsXml = readPart(archive, getRelsPath(vmlPath))
		const relationships = vmlRelsXml ? parseRelationships(vmlRelsXml) : []
		sheet.drawingObjectRefs.push(...parseVmlDrawingObjectRefs(vmlXml, vmlPath, relationships))
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
