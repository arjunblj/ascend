import { cloneCellStyle, type SheetId, type Workbook } from '@ascend/core'
import type { AscendError, CellValue, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import type { PreservationCapsule } from '../preserve.ts'
import {
	getRelsPath,
	REL_CALC_CHAIN,
	REL_COMMENTS,
	REL_DRAWING,
	REL_IMAGE,
	REL_PIVOT_CACHE_DEFINITION,
	REL_SHARED_STRINGS,
	REL_SHEET_METADATA,
	REL_STYLES,
	REL_TABLE,
	REL_THEME,
	REL_VML_DRAWING,
	REL_WORKSHEET,
} from '../reader/relationships.ts'
import { parseSharedStrings } from '../reader/shared-strings.ts'
import { extractZip, type ZipArchive } from '../reader/zip.ts'
import { updateChartXml } from './chart.ts'
import { buildCommentsVml, buildCommentsXml } from './comments.ts'
import { buildContentTypesXml } from './content-types.ts'
import { buildAppPropsXml, buildCorePropsXml } from './doc-props.ts'
import { buildDrawingXml } from './drawing.ts'
import { buildDynamicArrayMetadataXml } from './metadata.ts'
import { updatePivotCacheDefinitionXml } from './pivot-cache.ts'
import {
	summarizeWritePlan,
	type WritePartOwner,
	WritePlanBuilder,
	type WritePlanResult,
	type WritePlanSummary,
} from './plan.ts'
import type { RelEntry } from './relationships.ts'
import { buildRelsXml } from './relationships.ts'
import { IncrementalSharedStringTable, scanWorkbookWriteFactsFast } from './shared-strings.ts'
import { buildSheetXml, buildSheetXmlStreaming } from './sheet.ts'
import { buildPreservedStylesXml, buildStylesXml } from './styles.ts'
import { buildTableXml } from './table.ts'
import { buildThemeXml } from './theme.ts'
import { buildWorkbookXml } from './workbook.ts'
import { createZip, encode, StreamingZipBuilder } from './zip.ts'

const REL_OFFICE_DOC =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const REL_CORE_PROPS =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const REL_EXT_PROPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
const REL_HYPERLINK =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const CT_SHEET_METADATA =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'
const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
const CT_VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
const CT_THEME = 'application/vnd.openxmlformats-officedocument.theme+xml'

export interface WriteXlsxOptions {
	readonly dirtySheetNames?: readonly string[]
	readonly workbookMetaDirty?: boolean
	readonly calcStateDirty?: boolean
	readonly sharedStringsDirty?: boolean
	readonly stylesDirty?: boolean
	readonly summaryOnly?: boolean
	readonly sourceArchive?: ZipArchive
	/** Use shared string table (default: true). When false, strings are written inline per cell. */
	readonly useSharedStrings?: boolean
	/** Use inline strings instead of shared string table. When true, overrides useSharedStrings. */
	readonly useInlineStrings?: boolean
	readonly streaming?: boolean
}

export function writeXlsx(
	workbook: Workbook,
	capsules?: PreservationCapsule[],
	options: WriteXlsxOptions = {},
): Result<Uint8Array, AscendError> {
	try {
		const plan = planWriteXlsx(workbook, capsules, options)
		if (!plan.ok) return plan
		return ok(createZip(plan.value.parts))
	} catch (e) {
		return err(
			ascendError(
				'EXPORT_ERROR',
				`Failed to write XLSX: ${e instanceof Error ? e.message : 'unknown'}`,
			),
		)
	}
}

export async function writeXlsxStreaming(
	workbook: Workbook,
	capsules?: PreservationCapsule[],
	options: WriteXlsxOptions = {},
): Promise<Result<Uint8Array, AscendError>> {
	try {
		const plan = planWriteXlsx(workbook, capsules, { ...options, streaming: true })
		if (!plan.ok) return plan
		const zip = await createZipStreaming(plan.value)
		return ok(zip)
	} catch (e) {
		return err(
			ascendError(
				'EXPORT_ERROR',
				`Failed to write XLSX: ${e instanceof Error ? e.message : 'unknown'}`,
			),
		)
	}
}

async function createZipStreaming(plan: import('./plan.ts').WritePlanResult): Promise<Uint8Array> {
	const builder = new StreamingZipBuilder()
	for (const descriptor of plan.descriptors) {
		if (descriptor.streamingBuild) {
			builder.addStreamingEntry(descriptor.path)
			descriptor.streamingBuild((chunk) => builder.writeChunk(encode(chunk)))
			await builder.closeEntry()
		} else {
			const part = plan.parts.get(descriptor.path)
			if (part) builder.addEntry(descriptor.path, part)
		}
	}
	return builder.finalize()
}

export function planWriteXlsx(
	workbook: Workbook,
	capsules?: PreservationCapsule[],
	options: WriteXlsxOptions = {},
): Result<WritePlanResult, AscendError> {
	try {
		const plan = new WritePlanBuilder(!options.summaryOnly)
		const recordXml = (
			path: string,
			descriptor: Omit<import('./plan.ts').WritePartDescriptor, 'path'>,
			buildXml: () => string,
		): void => {
			if (options.summaryOnly) {
				plan.recordOnly(path, descriptor)
				return
			}
			plan.putXml(path, buildXml(), descriptor)
		}
		const recordStreamingSheet = (
			path: string,
			descriptor: Omit<import('./plan.ts').WritePartDescriptor, 'path'>,
			build: (onChunk: (chunk: string) => void) => void,
		): void => {
			if (options.summaryOnly) {
				plan.recordOnly(path, descriptor)
				return
			}
			plan.putStreamingSheet(path, descriptor, build)
		}
		const recordBytes = (
			path: string,
			descriptor: Omit<import('./plan.ts').WritePartDescriptor, 'path'>,
			buildBytes: () => Uint8Array,
		): void => {
			if (options.summaryOnly) {
				plan.recordOnly(path, descriptor)
				return
			}
			plan.putBytes(path, buildBytes(), descriptor)
		}
		const sourceArchive =
			options.sourceArchive ??
			(workbook.sourceArchiveBytes ? extractZip(workbook.sourceArchiveBytes) : undefined)
		const dirtyPatchMode = (options.dirtySheetNames?.length ?? 0) > 0 && sourceArchive !== undefined
		const effectiveStylesDirty = options.stylesDirty ?? dirtyPatchMode
		const effectiveSharedStringsDirty = options.sharedStringsDirty ?? dirtyPatchMode
		const effectiveWorkbookMetaDirty = options.workbookMetaDirty ?? dirtyPatchMode
		const sheetNameById = new Map<string, string>(
			workbook.sheets.map((sheet) => [sheet.id as string, sheet.name]),
		)
		const sheetCapsuleMap = new Map<string, PreservationCapsule[]>()
		const workbookCapsules: PreservationCapsule[] = []
		let nextGeneratedTableNumber = 1
		let nextGeneratedCommentsNumber = 1
		let nextGeneratedVmlNumber = 1
		let nextGeneratedDrawingNumber = 1

		const preservedSharedStringsXml = workbook.preservedSharedStrings
			? resolvePreservedText(
					sourceArchive,
					workbook.preservedSharedStrings.xml,
					workbook.preservedSharedStrings.path,
				)
			: undefined
		const hasPreservedSheetXmlInDirtyPatchMode =
			dirtyPatchMode &&
			workbook.sheets.some(
				(sheet) =>
					!(options.dirtySheetNames ?? []).includes(sheet.name) &&
					hasPreservedPart(sourceArchive, sheet.preservedXml?.xml, sheet.preservedXml?.partPath),
			)
		const useInlineStringsRequested =
			options.useInlineStrings ?? !(options.useSharedStrings ?? true)
		const useInlineStrings = useInlineStringsRequested && !hasPreservedSheetXmlInDirtyPatchMode
		const useSharedStrings = !useInlineStrings
		const preserveSharedStrings = Boolean(
			workbook.preservedSharedStrings &&
				!effectiveSharedStringsDirty &&
				preservedSharedStringsXml !== undefined,
		)
		const preservedSharedStringEntries =
			!options.summaryOnly && preservedSharedStringsXml
				? materializeSharedStringEntries(preservedSharedStringsXml)
				: []
		const workbookWriteFacts = scanWorkbookWriteFactsFast(workbook)
		const ssTable =
			options.summaryOnly || useInlineStrings
				? {
						getIndex(): number | undefined {
							return undefined
						},
						toXml(): string {
							return ''
						},
						count: preserveSharedStrings || workbookWriteFacts.hasStringCells ? 1 : 0,
						facts: workbookWriteFacts,
					}
				: new IncrementalSharedStringTable(
						preserveSharedStrings || hasPreservedSheetXmlInDirtyPatchMode
							? preservedSharedStringEntries
							: [],
						workbookWriteFacts,
					)
		const hasSharedStrings =
			useSharedStrings && (preserveSharedStrings || workbookWriteFacts.hasStringCells)

		const preservedStyles = workbook.preservedStyles ?? undefined
		const preserveStyles = preservedStyles !== undefined && !effectiveStylesDirty
		const canReusePreservedStyles =
			preservedStyles !== undefined &&
			!effectiveStylesDirty &&
			hasCompletePreservedStyleMap(preservedStyles.xfByStyleId, workbook.styles.size)
		const preservedStylesXml =
			preserveStyles && preservedStyles && !options.summaryOnly
				? resolvePreservedText(sourceArchive, preservedStyles.xml, preservedStyles.path)
				: undefined
		const preservedStyleBytes =
			preserveStyles && preservedStyles
				? resolvePreservedBytes(sourceArchive, preservedStyles.path)
				: undefined
		const stylesResult =
			preservedStyles !== undefined &&
			preservedStylesXml !== undefined &&
			!options.summaryOnly &&
			!canReusePreservedStyles
				? buildPreservedStylesXml(preservedStylesXml, preservedStyles, workbook.styles)
				: undefined
		const needsGeneratedStyles =
			!options.summaryOnly && (!preservedStyles || canReusePreservedStyles || !preservedStylesXml)
		const cfDxfIdOverridesBySheet = needsGeneratedStyles
			? collectCfRuleDxfOverrides(workbook)
			: new Map<string, Map<string, number>>()
		const generatedStylesResult = needsGeneratedStyles
			? buildStylesXml(workbook.styles, workbook.differentialStyles)
			: undefined
		const resolvedStylesResult = stylesResult ?? generatedStylesResult
		const xfMap =
			canReusePreservedStyles && preservedStyles
				? new Map(
						Object.entries(preservedStyles.xfByStyleId).map(([styleId, xfIndex]) => [
							Number(styleId),
							xfIndex,
						]),
					)
				: options.summaryOnly || !resolvedStylesResult
					? new Map<number, number>()
					: resolvedStylesResult.xfMap
		const stylesXml = options.summaryOnly
			? ''
			: preserveStyles && (stylesResult?.xml ?? preservedStylesXml)
				? (stylesResult?.xml ?? preservedStylesXml ?? '')
				: (generatedStylesResult?.xml ?? '')
		if (preservedStyles && !options.summaryOnly) {
			workbook.preservedStyles = {
				...preservedStyles,
				xfByStyleId: Object.fromEntries(xfMap.entries()),
			}
		}

		if (capsules) {
			for (const capsule of capsules) {
				if (isCalcChainCapsule(capsule)) continue
				if (capsule.anchor.kind === 'sheet') {
					const sheetId: SheetId = capsule.anchor.sheetId as SheetId
					if (!sheetNameById.has(sheetId)) continue
					let list = sheetCapsuleMap.get(sheetId)
					if (!list) {
						list = []
						sheetCapsuleMap.set(sheetId, list)
					}
					list.push(capsule)
				} else {
					workbookCapsules.push(capsule)
				}
			}
		}

		const dynamicArrayMetadata = { entries: ssTable.facts.dynamicArrayMetadataEntries }
		const shouldWriteDynamicArrayMetadata =
			dynamicArrayMetadata.entries.length > 0 || workbook.preservedMetadata !== null
		const dynamicArrayMetadataPath = workbook.preservedMetadata?.path ?? 'xl/metadata.xml'
		const dynamicArrayMetadataTarget = dynamicArrayMetadataPath.replace(/^xl\//, '')
		const preserveDynamicArrayMetadata = workbook.preservedMetadata !== null
		const preservedDynamicArrayMetadataBytes = workbook.preservedMetadata
			? resolvePreservedBytes(sourceArchive, workbook.preservedMetadata.path)
			: undefined
		const preservedDynamicArrayMetadataText =
			workbook.preservedMetadata && !options.summaryOnly
				? resolvePreservedText(
						sourceArchive,
						workbook.preservedMetadata.xml,
						workbook.preservedMetadata.path,
					)
				: undefined

		let rIdCounter = 1
		const wbRels: RelEntry[] = []
		for (let i = 0; i < workbook.sheets.length; i++) {
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: REL_WORKSHEET,
				target: `worksheets/sheet${i + 1}.xml`,
			})
			rIdCounter++
		}

		wbRels.push({ id: `rId${rIdCounter}`, type: REL_STYLES, target: 'styles.xml' })
		rIdCounter++

		const hasPreservedTheme = workbook.preservedTheme
			? hasPreservedPart(sourceArchive, workbook.preservedTheme.xml, workbook.preservedTheme.path)
			: false
		const shouldGenerateTheme = !workbook.preservedTheme && hasThemeMetadata(workbook.themeMetadata)
		const generatedThemePath = workbook.preservedTheme?.path ?? 'xl/theme/theme1.xml'
		const generatedThemeTarget = generatedThemePath.replace(/^xl\//, '')
		const generatedThemeContentType = workbook.preservedTheme?.contentType ?? CT_THEME
		const preservedThemeXml =
			workbook.preservedTheme && !options.summaryOnly
				? resolvePreservedText(
						sourceArchive,
						workbook.preservedTheme.xml,
						workbook.preservedTheme.path,
					)
				: undefined
		if (
			(workbook.preservedTheme && (hasPreservedTheme || preservedThemeXml)) ||
			shouldGenerateTheme
		) {
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: REL_THEME,
				target: generatedThemeTarget,
			})
			rIdCounter++
		}

		if (hasSharedStrings) {
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: REL_SHARED_STRINGS,
				target: 'sharedStrings.xml',
			})
			rIdCounter++
		}

		if (shouldWriteDynamicArrayMetadata) {
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: REL_SHEET_METADATA,
				target: dynamicArrayMetadataTarget,
			})
			rIdCounter++
		}

		const pivotCachePartPaths = new Set(workbook.pivotCaches.map((c) => c.partPath))
		const pivotCacheRelIds: string[] = []
		for (const cache of workbook.pivotCaches) {
			const target = cache.partPath.replace(/^xl\//, '')
			pivotCacheRelIds.push(`rId${rIdCounter}`)
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: REL_PIVOT_CACHE_DEFINITION,
				target,
			})
			rIdCounter++
		}

		for (const capsule of workbookCapsules) {
			if (!capsule.relType) continue
			if (pivotCachePartPaths.has(capsule.partPath)) continue
			const target = capsule.partPath.replace(/^xl\//, '')
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: capsule.relType,
				target,
			})
			rIdCounter++
		}

		const externalReferenceRelIds = wbRels
			.filter(
				(rel) =>
					rel.type ===
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
			)
			.map((rel) => rel.id)

		const preservedWorkbookXml = workbook.preservedXml
		const hasPreservedWorkbookXml =
			!effectiveWorkbookMetaDirty &&
			!!preservedWorkbookXml &&
			hasPreservedPart(
				sourceArchive,
				preservedWorkbookXml.workbookXml,
				preservedWorkbookXml.workbookPath,
			)
		const hasPreservedWorkbookRels =
			!effectiveWorkbookMetaDirty &&
			!!preservedWorkbookXml &&
			hasPreservedPart(
				sourceArchive,
				preservedWorkbookXml.workbookRelsXml,
				preservedWorkbookXml.workbookRelsPath,
			)
		const preservedWorkbookXmlText =
			hasPreservedWorkbookXml && !options.summaryOnly
				? resolvePreservedText(
						sourceArchive,
						preservedWorkbookXml?.workbookXml,
						preservedWorkbookXml?.workbookPath,
					)
				: undefined
		const preservedWorkbookRelsText =
			hasPreservedWorkbookRels && !options.summaryOnly
				? resolvePreservedText(
						sourceArchive,
						preservedWorkbookXml?.workbookRelsXml,
						preservedWorkbookXml?.workbookRelsPath,
					)
				: undefined
		const preservedWorkbookXmlBytes =
			!effectiveWorkbookMetaDirty && preservedWorkbookXml
				? resolvePreservedBytes(sourceArchive, preservedWorkbookXml.workbookPath)
				: undefined
		const preservedWorkbookRelsBytes =
			!effectiveWorkbookMetaDirty && preservedWorkbookXml
				? resolvePreservedBytes(sourceArchive, preservedWorkbookXml.workbookRelsPath)
				: undefined
		const preserveWorkbookXml = options.summaryOnly
			? hasPreservedWorkbookXml && hasPreservedWorkbookRels
			: !!(preservedWorkbookXmlText && preservedWorkbookRelsText)
		const preservedRelsHasCalcChain = preservedWorkbookRelsText?.includes('calcChain') === true
		const preserveWorkbookCalcState =
			preserveWorkbookXml && !options.calcStateDirty && !preservedRelsHasCalcChain
		const preserveWorkbookRels =
			preserveWorkbookCalcState &&
			(!shouldWriteDynamicArrayMetadata ||
				preservedWorkbookRelsText?.includes(REL_SHEET_METADATA) === true)
		const workbookContentType =
			preservedWorkbookXml?.contentType ??
			'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
		if (preserveWorkbookCalcState && preservedWorkbookXmlBytes && !options.summaryOnly) {
			recordBytes(
				'xl/workbook.xml',
				{
					owner: { kind: 'workbook' },
					origin: resolvePreservedOrigin(preservedWorkbookXml?.workbookXml),
					contentType: workbookContentType,
				},
				() => preservedWorkbookXmlBytes,
			)
		} else {
			recordXml(
				'xl/workbook.xml',
				{
					owner: { kind: 'workbook' },
					origin: preserveWorkbookCalcState
						? resolvePreservedOrigin(preservedWorkbookXml?.workbookXml)
						: 'generated',
					contentType: workbookContentType,
				},
				() =>
					preserveWorkbookCalcState
						? (preservedWorkbookXmlText ?? '')
						: buildWorkbookXml(workbook, {
								externalReferenceRelIds,
								pivotCacheRelIds,
								...(options.calcStateDirty !== undefined
									? { calcStateDirty: options.calcStateDirty }
									: {}),
							}),
			)
		}
		if (canReusePreservedStyles && preservedStyleBytes && !options.summaryOnly) {
			recordBytes(
				'xl/styles.xml',
				{
					owner: { kind: 'workbook' },
					origin: resolvePreservedOrigin(workbook.preservedStyles?.xml),
					contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
				},
				() => preservedStyleBytes,
			)
		} else {
			recordXml(
				'xl/styles.xml',
				{
					owner: { kind: 'workbook' },
					origin:
						workbook.preservedStyles && preserveStyles && preservedStylesXml
							? resolvePreservedOrigin(workbook.preservedStyles.xml)
							: 'generated',
					contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
				},
				() => stylesXml,
			)
		}
		const preservedThemeBytes = workbook.preservedTheme
			? resolvePreservedBytes(sourceArchive, workbook.preservedTheme.path)
			: undefined
		if (workbook.preservedTheme && preservedThemeXml) {
			if (preservedThemeBytes && !options.summaryOnly) {
				recordBytes(
					workbook.preservedTheme.path,
					{
						owner: { kind: 'workbook' },
						origin: resolvePreservedOrigin(workbook.preservedTheme.xml),
						contentType: workbook.preservedTheme.contentType,
					},
					() => preservedThemeBytes,
				)
			} else {
				recordXml(
					workbook.preservedTheme.path,
					{
						owner: { kind: 'workbook' },
						origin: resolvePreservedOrigin(workbook.preservedTheme.xml),
						contentType: workbook.preservedTheme.contentType,
					},
					() => preservedThemeXml,
				)
			}
			plan.addOverride(workbook.preservedTheme.path, workbook.preservedTheme.contentType)
		} else if (shouldGenerateTheme) {
			recordXml(
				generatedThemePath,
				{
					owner: { kind: 'workbook' },
					origin: 'generated',
					contentType: generatedThemeContentType,
				},
				() => buildThemeXml(workbook.themeMetadata),
			)
			plan.addOverride(generatedThemePath, generatedThemeContentType)
		}

		for (let i = 0; i < workbook.sheets.length; i++) {
			const sheet = workbook.sheets[i]
			if (!sheet) continue
			const sheetCapsules = sheetCapsuleMap.get(sheet.id) ?? []
			const preservedSheetXml = sheet.preservedXml
			const sheetRels: RelEntry[] = []
			let sheetRelId = 1
			let commentsRelId: string | undefined
			let drawingRelId: string | undefined
			let legacyDrawingRelId: string | undefined
			let commentsCapsulePath: string | undefined
			const tableRelIds: string[] = []
			const commentsCapsule = sheetCapsules.find((capsule) => capsule.relType === REL_COMMENTS)
			const tableCapsules = sheetCapsules.filter((capsule) => capsule.relType === REL_TABLE)
			const hasPreservedSheetXml = hasPreservedPart(
				sourceArchive,
				preservedSheetXml?.xml,
				preservedSheetXml?.partPath,
			)
			const hasPreservedSheetRels = hasPreservedPart(
				sourceArchive,
				preservedSheetXml?.relsXml,
				preservedSheetXml?.relsPath,
			)
			const preservedSheetXmlText =
				!options.summaryOnly && hasPreservedSheetXml
					? resolvePreservedText(sourceArchive, preservedSheetXml?.xml, preservedSheetXml?.partPath)
					: undefined
			const preservedSheetXmlBytes = resolvePreservedBytes(
				sourceArchive,
				preservedSheetXml?.partPath,
			)
			const preservedSheetRelsText =
				!options.summaryOnly && hasPreservedSheetRels
					? resolvePreservedText(
							sourceArchive,
							preservedSheetXml?.relsXml,
							preservedSheetXml?.relsPath,
						)
					: undefined
			const preservedSheetRelsBytes = resolvePreservedBytes(
				sourceArchive,
				preservedSheetXml?.relsPath,
			)
			const hyperlinkEntries: Array<{
				ref: string
				relId?: string
				location?: string
				display?: string
				tooltip?: string
			}> = []
			let hasGeneratedDrawing = false
			for (const capsule of sheetCapsules) {
				if (!capsule.relType) continue
				if (capsule.relType === REL_TABLE) continue
				const relId = `rId${sheetRelId}`
				sheetRels.push({
					id: relId,
					type: capsule.relType,
					target: computeRelativePath('xl/worksheets/', capsule.partPath),
				})
				if (capsule.relType === REL_COMMENTS && !commentsRelId) {
					commentsRelId = relId
					commentsCapsulePath = capsule.partPath
				}
				if (capsule.relType === REL_DRAWING && !drawingRelId) drawingRelId = relId
				if (capsule.relType === REL_VML_DRAWING && !legacyDrawingRelId) legacyDrawingRelId = relId
				sheetRelId++
			}
			for (const [ref, hyperlink] of sheet.hyperlinks) {
				if (hyperlink.target) {
					const relId = `rId${sheetRelId}`
					sheetRels.push({
						id: relId,
						type: REL_HYPERLINK,
						target: hyperlink.target,
						targetMode: 'External',
					})
					hyperlinkEntries.push({
						ref,
						relId,
						...(hyperlink.location ? { location: hyperlink.location } : {}),
						...(hyperlink.display ? { display: hyperlink.display } : {}),
						...(hyperlink.tooltip ? { tooltip: hyperlink.tooltip } : {}),
					})
					sheetRelId++
					continue
				}
				hyperlinkEntries.push({
					ref,
					...(hyperlink.location ? { location: hyperlink.location } : {}),
					...(hyperlink.display ? { display: hyperlink.display } : {}),
					...(hyperlink.tooltip ? { tooltip: hyperlink.tooltip } : {}),
				})
			}
			const preserveSheetXml =
				!(options.dirtySheetNames ?? []).includes(sheet.name) &&
				(options.summaryOnly ? hasPreservedSheetXml : !!preservedSheetXmlText)
			if (!preserveSheetXml) {
				if (sheet.comments.size > 0) {
					const commentsPartPath =
						commentsCapsule?.partPath ?? `xl/comments${nextGeneratedCommentsNumber}.xml`
					const vmlPartPath =
						commentsCapsulePath && legacyDrawingRelId
							? sheetCapsules.find((capsule) => capsule.relType === REL_VML_DRAWING)?.partPath
							: undefined
					const resolvedVmlPartPath =
						vmlPartPath ?? `xl/drawings/vmlDrawing${nextGeneratedVmlNumber}.vml`
					recordXml(
						commentsPartPath,
						{
							owner: { kind: 'sheet', sheetName: sheet.name },
							origin: 'generated',
							contentType: CT_COMMENTS,
						},
						() => buildCommentsXml(sheet),
					)
					recordXml(
						resolvedVmlPartPath,
						{
							owner: { kind: 'sheet', sheetName: sheet.name },
							origin: 'generated',
							contentType: CT_VML,
						},
						() => buildCommentsVml(sheet),
					)
					plan.addOverride(commentsPartPath, CT_COMMENTS)
					plan.addOverride(resolvedVmlPartPath, CT_VML)
					if (commentsCapsule) plan.skipCapsulePath(commentsCapsule.partPath)
					const existingVmlCapsule = sheetCapsules.find(
						(capsule) => capsule.relType === REL_VML_DRAWING,
					)
					if (existingVmlCapsule) plan.skipCapsulePath(existingVmlCapsule.partPath)
					commentsRelId = `rId${sheetRelId}`
					sheetRels.push({
						id: commentsRelId,
						type: REL_COMMENTS,
						target: computeRelativePath('xl/worksheets/', commentsPartPath),
					})
					sheetRelId++
					legacyDrawingRelId = `rId${sheetRelId}`
					sheetRels.push({
						id: legacyDrawingRelId,
						type: REL_VML_DRAWING,
						target: computeRelativePath('xl/worksheets/', resolvedVmlPartPath),
					})
					sheetRelId++
					nextGeneratedCommentsNumber++
					nextGeneratedVmlNumber++
				}
				for (let tableIndex = 0; tableIndex < sheet.tables.length; tableIndex++) {
					const table = sheet.tables[tableIndex]
					if (!table) continue
					const tableCapsule = tableCapsules[tableIndex]
					const tablePartPath =
						tableCapsule?.partPath ?? `xl/tables/table${nextGeneratedTableNumber}.xml`
					const tableContentType = tableCapsule?.contentType ?? CT_TABLE
					recordXml(
						tablePartPath,
						{
							owner: { kind: 'sheet', sheetName: sheet.name },
							origin: 'generated',
							contentType: tableContentType,
						},
						() => buildTableXml(table, nextGeneratedTableNumber),
					)
					plan.addOverride(tablePartPath, tableContentType)
					if (tableCapsule) plan.skipCapsulePath(tableCapsule.partPath)
					const relId = `rId${sheetRelId}`
					sheetRels.push({
						id: relId,
						type: REL_TABLE,
						target: computeRelativePath('xl/worksheets/', tablePartPath),
					})
					tableRelIds.push(relId)
					sheetRelId++
					nextGeneratedTableNumber++
				}
				const generatedImages = sheet.imageRefs.filter(
					(image) => image.content && image.contentType,
				)
				if (generatedImages.length > 0) {
					for (const image of generatedImages) {
						recordBytes(
							image.targetPath,
							{
								owner: { kind: 'sheet', sheetName: sheet.name },
								origin: 'generated',
								contentType: image.contentType as string,
							},
							() => image.content as Uint8Array,
						)
						plan.addOverride(image.targetPath, image.contentType as string)
						plan.skipCapsulePath(image.targetPath)
					}
				}
				if (generatedImages.length > 0 && !drawingRelId) {
					const drawingPartPath =
						generatedImages[0]?.drawingPartPath ||
						`xl/drawings/drawing${nextGeneratedDrawingNumber}.xml`
					recordXml(
						drawingPartPath,
						{
							owner: { kind: 'sheet', sheetName: sheet.name },
							origin: 'generated',
							contentType: CT_DRAWING,
						},
						() => buildDrawingXml(generatedImages),
					)
					const drawingDir = drawingPartPath.substring(0, drawingPartPath.lastIndexOf('/') + 1)
					recordXml(
						getRelsPath(drawingPartPath),
						{
							owner: { kind: 'sheet', sheetName: sheet.name },
							origin: 'generated',
						},
						() =>
							buildRelsXml(
								generatedImages.map((image) => ({
									id: image.relId,
									type: REL_IMAGE,
									target: computeRelativePath(drawingDir, image.targetPath),
								})),
							),
					)
					plan.addOverride(drawingPartPath, CT_DRAWING)
					drawingRelId = `rId${sheetRelId}`
					sheetRels.push({
						id: drawingRelId,
						type: REL_DRAWING,
						target: computeRelativePath('xl/worksheets/', drawingPartPath),
					})
					sheetRelId++
					nextGeneratedDrawingNumber++
					hasGeneratedDrawing = true
				}
			}
			if (preserveSheetXml && preservedSheetXmlBytes && !options.summaryOnly) {
				recordBytes(
					`xl/worksheets/sheet${i + 1}.xml`,
					{
						owner: { kind: 'sheet', sheetName: sheet.name },
						origin: resolvePreservedOrigin(preservedSheetXml?.xml),
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
					},
					() => preservedSheetXmlBytes,
				)
			} else if (options.streaming && !preserveSheetXml) {
				const cfOverrides = cfDxfIdOverridesBySheet.get(sheet.name)
				const sheetOptions = {
					tableRelIds,
					...((sheet.drawingRefs.hasDrawing || hasGeneratedDrawing) && drawingRelId
						? { drawingRelId }
						: {}),
					hyperlinks: hyperlinkEntries,
					...((sheet.drawingRefs.hasLegacyDrawing || sheet.comments.size > 0) && legacyDrawingRelId
						? { legacyDrawingRelId }
						: {}),
					useInlineStrings,
					...(cfOverrides ? { cfDxfIdOverrides: cfOverrides } : {}),
				}
				recordStreamingSheet(
					`xl/worksheets/sheet${i + 1}.xml`,
					{
						owner: { kind: 'sheet', sheetName: sheet.name },
						origin: 'generated',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
					},
					(onChunk) => buildSheetXmlStreaming(sheet, ssTable, xfMap, sheetOptions, onChunk),
				)
			} else {
				recordXml(
					`xl/worksheets/sheet${i + 1}.xml`,
					{
						owner: { kind: 'sheet', sheetName: sheet.name },
						origin: preserveSheetXml ? resolvePreservedOrigin(preservedSheetXml?.xml) : 'generated',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
					},
					() => {
						const cfOverrides = cfDxfIdOverridesBySheet.get(sheet.name)
						return preserveSheetXml
							? (preservedSheetXmlText ?? '')
							: buildSheetXml(sheet, ssTable, xfMap, {
									tableRelIds,
									...((sheet.drawingRefs.hasDrawing || hasGeneratedDrawing) && drawingRelId
										? { drawingRelId }
										: {}),
									hyperlinks: hyperlinkEntries,
									...((sheet.drawingRefs.hasLegacyDrawing || sheet.comments.size > 0) &&
									legacyDrawingRelId
										? { legacyDrawingRelId }
										: {}),
									useInlineStrings,
									...(cfOverrides ? { cfDxfIdOverrides: cfOverrides } : {}),
								})
					},
				)
			}
			if (
				preserveSheetXml &&
				(options.summaryOnly ? hasPreservedSheetRels : !!preservedSheetRelsText)
			) {
				if (preservedSheetRelsBytes && !options.summaryOnly) {
					recordBytes(
						`xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
						{
							owner: { kind: 'sheet', sheetName: sheet.name },
							origin: resolvePreservedOrigin(preservedSheetXml?.relsXml),
						},
						() => preservedSheetRelsBytes,
					)
				} else {
					recordXml(
						`xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
						{
							owner: { kind: 'sheet', sheetName: sheet.name },
							origin: resolvePreservedOrigin(preservedSheetXml?.relsXml),
						},
						() => preservedSheetRelsText ?? '',
					)
				}
			} else if (sheetRels.length > 0) {
				recordXml(
					`xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
					{
						owner: { kind: 'sheet', sheetName: sheet.name },
						origin: 'generated',
					},
					() => buildRelsXml(sheetRels),
				)
			}
		}

		if (hasSharedStrings) {
			const canReuseSharedStringsXml =
				preserveSharedStrings &&
				(useInlineStrings || ssTable.count <= preservedSharedStringEntries.length)
			const preservedSharedStringBytes = canReuseSharedStringsXml
				? resolvePreservedBytes(sourceArchive, workbook.preservedSharedStrings?.path)
				: undefined
			if (canReuseSharedStringsXml && preservedSharedStringBytes && !options.summaryOnly) {
				recordBytes(
					'xl/sharedStrings.xml',
					{
						owner: { kind: 'workbook' },
						origin: resolvePreservedOrigin(workbook.preservedSharedStrings?.xml),
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
					},
					() => preservedSharedStringBytes,
				)
			} else {
				recordXml(
					'xl/sharedStrings.xml',
					{
						owner: { kind: 'workbook' },
						origin:
							canReuseSharedStringsXml && preservedSharedStringsXml !== undefined
								? resolvePreservedOrigin(workbook.preservedSharedStrings?.xml)
								: 'generated',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
					},
					() => (canReuseSharedStringsXml ? (preservedSharedStringsXml ?? '') : ssTable.toXml()),
				)
			}
		}

		recordXml(
			'docProps/core.xml',
			{
				owner: { kind: 'package' },
				origin: 'generated',
				contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
			},
			() => buildCorePropsXml(),
		)
		recordXml(
			'docProps/app.xml',
			{
				owner: { kind: 'package' },
				origin: 'generated',
				contentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
			},
			() => buildAppPropsXml(),
		)

		const rootRels: RelEntry[] = [
			{ id: 'rId1', type: REL_OFFICE_DOC, target: 'xl/workbook.xml' },
			{ id: 'rId2', type: REL_CORE_PROPS, target: 'docProps/core.xml' },
			{ id: 'rId3', type: REL_EXT_PROPS, target: 'docProps/app.xml' },
		]
		recordXml(
			'_rels/.rels',
			{
				owner: { kind: 'package' },
				origin: 'generated',
			},
			() => buildRelsXml(rootRels),
		)

		if (capsules) {
			for (const capsule of capsules) {
				if (isCalcChainCapsule(capsule)) continue
				if (plan.isCapsulePathSkipped(capsule.partPath)) continue
				const content = capsule.content ?? sourceArchive?.readBytes(capsule.partPath)
				if (!content) continue
				const owner = resolveCapsuleOwner(capsule, sheetNameById)
				if (!owner) continue
				const pivotCache = workbook.pivotCaches.find((cache) => cache.partPath === capsule.partPath)
				if (pivotCache && isPivotCacheDefinitionCapsule(capsule)) {
					recordXml(
						capsule.partPath,
						{
							owner,
							origin: 'generated',
							contentType: capsule.contentType,
						},
						() => updatePivotCacheDefinitionXml(new TextDecoder().decode(content), pivotCache),
					)
					plan.addOverride(capsule.partPath, capsule.contentType)
					if (capsule.relationships.length > 0) {
						const capsuleRelsPath = getRelsPath(capsule.partPath)
						recordXml(
							capsuleRelsPath,
							{
								owner,
								origin: 'capsule',
							},
							() => buildRelsXml(capsule.relationships),
						)
					}
					continue
				}
				const chart = workbook.chartParts.find((entry) => entry.partPath === capsule.partPath)
				if (chart && isChartCapsule(capsule)) {
					recordXml(
						capsule.partPath,
						{
							owner,
							origin: 'generated',
							contentType: capsule.contentType,
						},
						() => updateChartXml(new TextDecoder().decode(content), chart),
					)
					plan.addOverride(capsule.partPath, capsule.contentType)
					if (capsule.relationships.length > 0) {
						const capsuleRelsPath = getRelsPath(capsule.partPath)
						recordXml(
							capsuleRelsPath,
							{
								owner,
								origin: 'capsule',
							},
							() => buildRelsXml(capsule.relationships),
						)
					}
					continue
				}
				recordBytes(
					capsule.partPath,
					{
						owner,
						origin: 'capsule',
						contentType: capsule.contentType,
					},
					() => content,
				)

				if (capsule.relationships.length > 0) {
					const capsuleRelsPath = getRelsPath(capsule.partPath)
					recordXml(
						capsuleRelsPath,
						{
							owner,
							origin: 'capsule',
						},
						() => buildRelsXml(capsule.relationships),
					)
				}
			}
		}

		if (
			preserveWorkbookRels &&
			preservedWorkbookRelsText &&
			preservedWorkbookRelsBytes &&
			!options.summaryOnly
		) {
			recordBytes(
				'xl/_rels/workbook.xml.rels',
				{
					owner: { kind: 'workbook' },
					origin: resolvePreservedOrigin(preservedWorkbookXml?.workbookRelsXml),
				},
				() => preservedWorkbookRelsBytes,
			)
		} else {
			recordXml(
				'xl/_rels/workbook.xml.rels',
				{
					owner: { kind: 'workbook' },
					origin:
						preserveWorkbookRels && preservedWorkbookRelsText
							? resolvePreservedOrigin(preservedWorkbookXml?.workbookRelsXml)
							: 'generated',
				},
				() =>
					preserveWorkbookRels && preservedWorkbookRelsText
						? preservedWorkbookRelsText
						: buildRelsXml(wbRels),
			)
		}

		if (shouldWriteDynamicArrayMetadata) {
			plan.addOverride(dynamicArrayMetadataPath, CT_SHEET_METADATA)
			if (
				preserveDynamicArrayMetadata &&
				preservedDynamicArrayMetadataBytes &&
				!options.summaryOnly
			) {
				recordBytes(
					dynamicArrayMetadataPath,
					{
						owner: { kind: 'workbook' },
						origin: resolvePreservedOrigin(workbook.preservedMetadata?.xml),
						contentType: workbook.preservedMetadata?.contentType ?? CT_SHEET_METADATA,
					},
					() => preservedDynamicArrayMetadataBytes,
				)
			} else {
				recordXml(
					dynamicArrayMetadataPath,
					{
						owner: { kind: 'workbook' },
						origin:
							preserveDynamicArrayMetadata && preservedDynamicArrayMetadataText
								? resolvePreservedOrigin(workbook.preservedMetadata?.xml)
								: 'generated',
						contentType: workbook.preservedMetadata?.contentType ?? CT_SHEET_METADATA,
					},
					() =>
						preserveDynamicArrayMetadata && preservedDynamicArrayMetadataText
							? preservedDynamicArrayMetadataText
							: buildDynamicArrayMetadataXml(dynamicArrayMetadata.entries),
				)
			}
		}

		recordXml(
			'[Content_Types].xml',
			{
				owner: { kind: 'package' },
				origin: 'generated',
			},
			() => {
				const built = plan.build()
				return buildContentTypesXml(
					workbook.sheets.length,
					hasSharedStrings,
					workbookContentType,
					capsules?.filter(
						(capsule) =>
							!isCalcChainCapsule(capsule) && !built.skippedCapsulePaths.has(capsule.partPath),
					),
					built.extraOverrides.length > 0 ? built.extraOverrides : undefined,
				)
			},
		)

		return ok(plan.build())
	} catch (e) {
		return err(
			ascendError(
				'EXPORT_ERROR',
				`Failed to write XLSX: ${e instanceof Error ? e.message : 'unknown'}`,
			),
		)
	}
}

export function summarizePlannedWrite(
	workbook: Workbook,
	capsules?: PreservationCapsule[],
	options: WriteXlsxOptions = {},
): Result<WritePlanSummary, AscendError> {
	const plan = planWriteXlsx(workbook, capsules, { ...options, summaryOnly: true })
	if (!plan.ok) return plan
	return ok(summarizeWritePlan(plan.value))
}

function resolvePreservedOrigin(
	inlineText: string | undefined,
): 'preserved-inline' | 'preserved-source' {
	return inlineText !== undefined ? 'preserved-inline' : 'preserved-source'
}

function materializeSharedStringEntries(xml: string): CellValue[] {
	const resolver = parseSharedStrings(xml)
	const entries: CellValue[] = []
	for (let index = 0; index < resolver.count; index++) {
		const value = resolver.get(index)
		if (value) entries.push(value)
	}
	return entries
}

function resolveCapsuleOwner(
	capsule: PreservationCapsule,
	sheetNameById: ReadonlyMap<string, string>,
): WritePartOwner | null {
	if (capsule.anchor.kind !== 'sheet') return { kind: 'workbook' }
	const sheetName = sheetNameById.get(capsule.anchor.sheetId)
	return sheetName ? { kind: 'sheet', sheetName } : null
}

function resolvePreservedText(
	archive: ZipArchive | undefined,
	inlineText: string | undefined,
	partPath: string | undefined,
): string | undefined {
	if (inlineText !== undefined) return inlineText
	if (!archive || !partPath) return undefined
	return archive.readText(partPath)
}

function resolvePreservedBytes(
	archive: ZipArchive | undefined,
	partPath: string | undefined,
): Uint8Array | undefined {
	if (!archive || !partPath) return undefined
	return archive.readBytes(partPath)
}

function hasPreservedPart(
	archive: ZipArchive | undefined,
	inlineText: string | undefined,
	partPath: string | undefined,
): boolean {
	return inlineText !== undefined || (!!archive && !!partPath && archive.has(partPath))
}

function hasThemeMetadata(metadata: Workbook['themeMetadata']): boolean {
	return (
		metadata.colorCount > 0 ||
		(metadata.name?.trim().length ?? 0) > 0 ||
		(metadata.colorSchemeName?.trim().length ?? 0) > 0 ||
		(metadata.majorFontLatin?.trim().length ?? 0) > 0 ||
		(metadata.minorFontLatin?.trim().length ?? 0) > 0
	)
}

function collectCfRuleDxfOverrides(workbook: Workbook): Map<string, Map<string, number>> {
	const bySheet = new Map<string, Map<string, number>>()
	for (const sheet of workbook.sheets) {
		const overrides = new Map<string, number>()
		for (let cfIdx = 0; cfIdx < sheet.conditionalFormats.length; cfIdx++) {
			const cf = sheet.conditionalFormats[cfIdx]
			if (!cf) continue
			for (let ruleIdx = 0; ruleIdx < cf.rules.length; ruleIdx++) {
				const rule = cf.rules[ruleIdx]
				if (!rule || rule.dxfId !== undefined || !rule.style) continue
				const dxfId = workbook.differentialStyles.length
				workbook.differentialStyles.push(cloneCellStyle(rule.style))
				overrides.set(`${cfIdx}:${ruleIdx}`, dxfId)
			}
		}
		if (overrides.size > 0) bySheet.set(sheet.name, overrides)
	}
	return bySheet
}

function hasCompletePreservedStyleMap(
	xfByStyleId: Readonly<Record<number, number>>,
	styleCount: number,
): boolean {
	for (let index = 0; index < styleCount; index++) {
		if (xfByStyleId[index] === undefined) return false
	}
	return true
}

function isCalcChainCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.relType === REL_CALC_CHAIN ||
		capsule.contentType.includes('calcChain+xml') ||
		capsule.partPath.endsWith('/calcChain.xml')
	)
}

function isPivotCacheDefinitionCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.relType === REL_PIVOT_CACHE_DEFINITION ||
		capsule.contentType.includes('pivotCacheDefinition+xml') ||
		capsule.partPath.includes('/pivotCache/pivotCacheDefinition')
	)
}

function isChartCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.contentType.includes('chart+xml') ||
		capsule.partPath.includes('/charts/') ||
		capsule.partPath.includes('/chartEx/')
	)
}

function computeRelativePath(fromDir: string, toPath: string): string {
	const from = fromDir.replace(/^\//, '').split('/').filter(Boolean)
	const to = toPath.replace(/^\//, '').split('/')

	let common = 0
	while (common < from.length && common < to.length - 1 && from[common] === to[common]) {
		common++
	}

	const ups = from.length - common
	const rest = to.slice(common)
	return '../'.repeat(ups) + rest.join('/')
}
