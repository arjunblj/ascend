import {
	cloneCellStyle,
	type SheetDrawingObjectRef,
	type SheetId,
	type Table,
	type Workbook,
} from '@ascend/core'
import type { AscendError, CellValue, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import type { PreservationCapsule } from '../preserve.ts'
import { parseCommentsXml, parseCommentVmlXml } from '../reader/comments.ts'
import { parseConnectionPartInfos } from '../reader/connections.ts'
import { parseDrawingObjectRefs } from '../reader/drawing.ts'
import {
	parsePivotCacheDefinitionXml,
	parsePivotTableXml,
	parseSlicerCacheXml,
	parseTimelineCacheXml,
} from '../reader/pivots.ts'
import {
	getRelsPath,
	isExternalLinkPathRelationshipType,
	parseRelationships,
	REL_CALC_CHAIN,
	REL_CHARTSHEET,
	REL_COMMENTS,
	REL_DRAWING,
	REL_EXTERNAL_LINK_PATH,
	REL_IMAGE,
	REL_MACROSHEET,
	REL_PIVOT_CACHE_DEFINITION,
	REL_PIVOT_TABLE,
	REL_QUERY_TABLE,
	REL_SHARED_STRINGS,
	REL_SHEET_METADATA,
	REL_SLICER_CACHE,
	REL_STYLES,
	REL_TABLE,
	REL_THEME,
	REL_TIMELINE_CACHE,
	REL_VML_DRAWING,
	REL_WORKSHEET,
	type Relationship,
	resolvePath,
} from '../reader/relationships.ts'
import { parseSharedStrings } from '../reader/shared-strings.ts'
import { parseTable } from '../reader/table.ts'
import { extractZip, type ZipArchive } from '../reader/zip.ts'
import { updateChartXml } from './chart.ts'
import { buildCommentsVml, buildCommentsXml } from './comments.ts'
import { updateConnectionPartXml } from './connection.ts'
import { buildContentTypesXml } from './content-types.ts'
import { buildAppPropsXml, buildCorePropsXml, buildCustomPropsXml } from './doc-props.ts'
import { buildDrawingXml, type DrawingTextUpdate, updateDrawingTextXml } from './drawing.ts'
import { buildDynamicArrayMetadataXml, type DynamicArrayMetadataEntry } from './metadata.ts'
import { updatePivotCacheDefinitionXml } from './pivot-cache.ts'
import { updatePivotTableDefinitionXml } from './pivot-table.ts'
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
import { updateSlicerCacheDefinitionXml } from './slicer-cache.ts'
import { buildPreservedStylesXml, buildStylesXml } from './styles.ts'
import { buildTableXml } from './table.ts'
import { buildThemeXml, themeXmlMatches, updateThemeXml } from './theme.ts'
import {
	readThreadedCommentTextRefs,
	type ThreadedCommentTextRef,
	updateThreadedCommentsXml,
} from './threaded-comments.ts'
import { updateTimelineCacheDefinitionXml } from './timeline-cache.ts'
import { buildWorkbookXml } from './workbook.ts'
import { createZip, encode, StreamingZipBuilder } from './zip.ts'

const REL_OFFICE_DOC =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const REL_CORE_PROPS =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const REL_EXT_PROPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
const REL_CUSTOM_PROPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
const REL_DIGITAL_SIGNATURE_ORIGIN =
	'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin'
const REL_HYPERLINK =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const CT_SHEET_METADATA =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'
const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
const CT_VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
const CT_THEME = 'application/vnd.openxmlformats-officedocument.theme+xml'
const STREAMING_XML_BATCH_CHARS = 524_288

function worksheetPartPath(sheet: Workbook['sheets'][number] | undefined, index: number): string {
	return sheet?.preservedXml?.partPath ?? `xl/worksheets/sheet${index + 1}.xml`
}

function worksheetRelsPath(sheet: Workbook['sheets'][number] | undefined, index: number): string {
	return sheet?.preservedXml?.relsPath ?? getRelsPath(worksheetPartPath(sheet, index))
}

function filterPreservedContentTypeOverrides(
	overrides: NonNullable<Workbook['preservedXml']>['contentTypeOverrides'] | undefined,
	defaults: NonNullable<Workbook['preservedXml']>['contentTypeDefaults'] | undefined,
	finalPartPaths: ReadonlySet<string>,
): readonly { partPath: string; contentType: string }[] | undefined {
	if (!overrides) return undefined
	const defaultMap = new Map((defaults ?? []).map((entry) => [entry.extension, entry.contentType]))
	const filtered = overrides.filter((override) => {
		const partPath = stripLeadingSlash(override.partPath)
		if (!finalPartPaths.has(partPath)) return false
		const extension = partPath.split('.').pop()
		const coveredByDefault =
			extension !== undefined && defaultMap.get(extension) === override.contentType
		return !coveredByDefault || extension === 'rels'
	})
	return filtered.length > 0 ? filtered : undefined
}

function stripLeadingSlash(path: string): string {
	return path.startsWith('/') ? path.slice(1) : path
}

interface XmlStreamingBatchSink {
	write(chunk: string): void
	flush(): Uint8Array | undefined
	readonly length: number
}

type BunArrayBufferSink = {
	start(options?: { asUint8Array?: boolean; highWaterMark?: number; stream?: boolean }): void
	write(chunk: string): number | undefined
	flush(): number | Uint8Array | ArrayBuffer
}

type BunArrayBufferSinkConstructor = new () => BunArrayBufferSink

export interface WriteXlsxOptions {
	readonly dirtySheetNames?: readonly string[]
	readonly workbookMetaDirty?: boolean
	readonly documentPropertiesDirty?: boolean
	readonly calcStateDirty?: boolean
	readonly calcChainDirty?: boolean
	readonly sharedStringsDirty?: boolean
	readonly stylesDirty?: boolean
	readonly summaryOnly?: boolean
	readonly sourceArchive?: ZipArchive
	/** Use shared string table (default: true). When false, strings are written inline per cell. */
	readonly useSharedStrings?: boolean
	/** Use inline strings instead of shared string table. When true, overrides useSharedStrings. */
	readonly useInlineStrings?: boolean
	/** Use plain string cells (`t="str"`) for scalar strings and avoid a shared string table. */
	readonly usePlainStrings?: boolean
	/** Omit cell ref attributes on dense contiguous default-style scalar rows. */
	readonly omitDenseCellRefs?: boolean
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
			writeStreamingXmlEntry(builder, descriptor.streamingBuild)
			await builder.closeEntry()
		} else {
			const part = plan.parts.get(descriptor.path)
			if (part) builder.addEntry(descriptor.path, part)
		}
	}
	return builder.finalize()
}

function writeStreamingXmlEntry(
	builder: StreamingZipBuilder,
	build: (onChunk: (chunk: string) => void) => void,
): void {
	let batch = createXmlStreamingBatchSink()
	const flush = () => {
		const bytes = batch.flush()
		if (!bytes) return
		builder.writeChunk(bytes)
		batch = createXmlStreamingBatchSink()
	}
	build((chunk) => {
		batch.write(chunk)
		if (batch.length >= STREAMING_XML_BATCH_CHARS) flush()
	})
	flush()
}

function createXmlStreamingBatchSink(): XmlStreamingBatchSink {
	return createBunStreamingBatchSink() ?? createStringStreamingBatchSink()
}

function createBunStreamingBatchSink(): XmlStreamingBatchSink | undefined {
	const ctor = (
		globalThis as { readonly Bun?: { readonly ArrayBufferSink?: BunArrayBufferSinkConstructor } }
	).Bun?.ArrayBufferSink
	if (typeof ctor !== 'function') return undefined
	const sink = new ctor()
	sink.start({
		asUint8Array: true,
		stream: true,
		highWaterMark: STREAMING_XML_BATCH_CHARS,
	})
	let chars = 0
	return {
		write(chunk) {
			chars += chunk.length
			sink.write(chunk)
		},
		flush() {
			if (chars === 0) return undefined
			chars = 0
			const flushed = sink.flush()
			if (flushed instanceof Uint8Array) return flushed
			if (flushed instanceof ArrayBuffer) return new Uint8Array(flushed)
			return undefined
		},
		get length() {
			return chars
		},
	}
}

function createStringStreamingBatchSink(): XmlStreamingBatchSink {
	const chunks: string[] = []
	let chars = 0
	return {
		write(chunk) {
			chunks.push(chunk)
			chars += chunk.length
		},
		flush() {
			if (chars === 0) return undefined
			const bytes = encode(chunks.join(''))
			chunks.length = 0
			chars = 0
			return bytes
		},
		get length() {
			return chars
		},
	}
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
		const recordUnchangedCapsule = (
			capsule: PreservationCapsule,
			owner: WritePartOwner,
			content: Uint8Array,
		): void => {
			recordBytes(
				capsule.partPath,
				{
					owner,
					origin: 'capsule',
					contentType: capsule.contentType,
				},
				() => content,
			)
			if (capsule.relationships.length === 0) return
			const capsuleRelsPath = getRelsPath(capsule.partPath)
			const relsBytes = sourceArchive?.readBytes(capsuleRelsPath)
			if (relsBytes) {
				recordBytes(
					capsuleRelsPath,
					{
						owner,
						origin: 'capsule',
					},
					() => relsBytes,
				)
				return
			}
			recordXml(
				capsuleRelsPath,
				{
					owner,
					origin: 'capsule',
				},
				() => buildRelsXml(resolveCapsuleRelationships(workbook, capsule)),
			)
		}
		const sourceArchive =
			options.sourceArchive ??
			(workbook.sourceArchiveBytes ? extractZip(workbook.sourceArchiveBytes) : undefined)
		const preservedRootRelsText = sourceArchive?.readText('_rels/.rels')
		const preservedWorkbookRelsTextForTargets =
			sourceArchive && workbook.preservedXml
				? resolvePreservedText(
						sourceArchive,
						workbook.preservedXml.workbookRelsXml,
						workbook.preservedXml.workbookRelsPath,
					)
				: undefined
		const preservedWorkbookRels = preservedWorkbookRelsTextForTargets
			? parseRelationships(preservedWorkbookRelsTextForTargets)
			: []
		const rootRelTarget = (type: string, partPath: string, fallback: string): string =>
			preservedRelationshipTarget(preservedRootRelsText, '', type, partPath) ?? fallback
		const workbookRelTarget = (type: string, partPath: string, fallback: string): string =>
			preservedRelationshipTarget(
				preservedWorkbookRelsTextForTargets,
				'xl/workbook.xml',
				type,
				partPath,
			) ?? fallback
		const dirtyPatchMode = (options.dirtySheetNames?.length ?? 0) > 0 && sourceArchive !== undefined
		const effectiveStylesDirty = options.stylesDirty ?? dirtyPatchMode
		const effectiveSharedStringsDirty = options.sharedStringsDirty ?? dirtyPatchMode
		const effectiveWorkbookMetaDirty = options.workbookMetaDirty ?? dirtyPatchMode
		const effectiveCalcChainDirty = options.calcChainDirty ?? options.calcStateDirty ?? false
		const invalidateDigitalSignatures =
			effectiveWorkbookMetaDirty ||
			effectiveSharedStringsDirty ||
			effectiveStylesDirty ||
			effectiveCalcChainDirty ||
			(options.dirtySheetNames?.length ?? 0) > 0
		const sheetNameById = new Map<string, string>([
			...workbook.sheets.map((sheet) => [sheet.id as string, sheet.name] as const),
			...workbook.chartSheets.map((sheet) => [sheet.sheetId, sheet.name] as const),
			...workbook.macroSheets.map((sheet) => [sheet.sheetId, sheet.name] as const),
		])
		const sheetCapsuleMap = new Map<string, PreservationCapsule[]>()
		const workbookCapsules: PreservationCapsule[] = []
		const liveQueryTablePartPaths = collectLiveQueryTablePartPaths(workbook)
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
		const usePlainStringsRequested = options.usePlainStrings ?? false
		const usePlainStrings = usePlainStringsRequested && !hasPreservedSheetXmlInDirtyPatchMode
		const useInlineStringsRequested =
			options.useInlineStrings ??
			(!(options.useSharedStrings ?? true) ||
				(dirtyPatchMode && workbook.preservedSharedStrings === null))
		const useInlineStrings =
			!usePlainStrings && useInlineStringsRequested && !hasPreservedSheetXmlInDirtyPatchMode
		const useSharedStrings = !useInlineStrings && !usePlainStrings
		const omitDenseCellRefs = options.omitDenseCellRefs ?? true
		const preserveSharedStrings = Boolean(
			workbook.preservedSharedStrings &&
				!effectiveSharedStringsDirty &&
				preservedSharedStringsXml !== undefined,
		)
		const preservedSharedStringEntries =
			!options.summaryOnly && preservedSharedStringsXml
				? materializeSharedStringEntries(preservedSharedStringsXml)
				: []
		const hasSharedStringEligibleCells = workbookHasSharedStringEligibleCells(workbook)
		const needsWriteFactScan = workbookHasFormulaInfoCells(workbook)
		const workbookWriteFacts = needsWriteFactScan
			? scanWorkbookWriteFactsFast(workbook)
			: { hasStringCells: hasSharedStringEligibleCells, dynamicArrayMetadataEntries: [] }
		const ssTable =
			options.summaryOnly || useInlineStrings || usePlainStrings
				? {
						getIndex(): number | undefined {
							return undefined
						},
						toXml(): string {
							return ''
						},
						count: preserveSharedStrings || workbookWriteFacts.hasStringCells ? 1 : 0,
						entryCount: preservedSharedStringEntries.length,
						facts: workbookWriteFacts,
					}
				: new IncrementalSharedStringTable(
						preserveSharedStrings ||
							(dirtyPatchMode && preservedSharedStringsXml !== undefined) ||
							hasPreservedSheetXmlInDirtyPatchMode
							? preservedSharedStringEntries
							: [],
						workbookWriteFacts,
						preservedSharedStringsXml,
					)
		const hasSharedStrings =
			useSharedStrings &&
			(preserveSharedStrings ||
				workbookWriteFacts.hasStringCells ||
				(dirtyPatchMode && preservedSharedStringsXml !== undefined))

		const preservedStyles = workbook.preservedStyles ?? undefined
		const hasWorkbookStyleContent =
			preservedStyles !== undefined ||
			workbook.styles.size > 1 ||
			workbook.differentialStyles.length > 0
		const shouldWriteStyles = !dirtyPatchMode || hasWorkbookStyleContent
		const hasNumberFormatOnlyStylePatches =
			preservedStyles?.baseStyleIdByStyleId !== undefined &&
			Object.keys(preservedStyles.baseStyleIdByStyleId).length > 0
		const preserveStyles =
			preservedStyles !== undefined && (!effectiveStylesDirty || hasNumberFormatOnlyStylePatches)
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
			shouldWriteStyles &&
			!options.summaryOnly &&
			(!preservedStyles || canReusePreservedStyles || !preservedStylesXml || !stylesResult)
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
				if (isDeletedQueryTableCapsule(capsule, liveQueryTablePartPaths)) {
					plan.skipCapsulePath(capsule.partPath)
					continue
				}
				if (invalidateDigitalSignatures && isDigitalSignatureCapsule(capsule)) {
					plan.skipCapsulePath(capsule.partPath)
					continue
				}
				if (isCalcChainCapsule(capsule) && effectiveCalcChainDirty) {
					plan.skipCapsulePath(capsule.partPath)
					continue
				}
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
		const preserveDynamicArrayMetadata =
			workbook.preservedMetadata !== null &&
			canPreserveDynamicArrayMetadata(
				dynamicArrayMetadata.entries,
				workbook.preservedMetadata.dynamicArrayMetadata,
			)
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
		const reservedWorkbookRelIds = new Set(preservedWorkbookRels.map((rel) => rel.id))
		const usedWorkbookRelIds = new Set<string>()
		const wbRels: RelEntry[] = []
		const addWorkbookRel = (type: string, partPath: string, fallback: string): string => {
			const rawType = preservedRelationshipRawType(
				preservedWorkbookRels,
				'xl/workbook.xml',
				type,
				partPath,
			)
			const id = allocateWorkbookRelId(
				preservedWorkbookRels,
				reservedWorkbookRelIds,
				usedWorkbookRelIds,
				'xl/workbook.xml',
				type,
				partPath,
				() => `rId${rIdCounter++}`,
			)
			wbRels.push({
				id,
				type,
				...(rawType ? { rawType } : {}),
				target: workbookRelTarget(type, partPath, fallback),
			})
			return id
		}
		const worksheetRelIds: string[] = []
		for (let i = 0; i < workbook.sheets.length; i++) {
			const sheet = workbook.sheets[i]
			const partPath = worksheetPartPath(sheet, i)
			worksheetRelIds.push(addWorkbookRel(REL_WORKSHEET, partPath, partPath.replace(/^xl\//, '')))
		}

		if (shouldWriteStyles) {
			addWorkbookRel(REL_STYLES, 'xl/styles.xml', 'styles.xml')
		}

		const hasPreservedTheme = workbook.preservedTheme
			? hasPreservedPart(sourceArchive, workbook.preservedTheme.xml, workbook.preservedTheme.path)
			: false
		const shouldGenerateTheme = !workbook.preservedTheme && hasThemeContent(workbook)
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
			addWorkbookRel(REL_THEME, generatedThemePath, generatedThemeTarget)
		}

		if (hasSharedStrings) {
			addWorkbookRel(REL_SHARED_STRINGS, 'xl/sharedStrings.xml', 'sharedStrings.xml')
		}

		if (shouldWriteDynamicArrayMetadata) {
			addWorkbookRel(REL_SHEET_METADATA, dynamicArrayMetadataPath, dynamicArrayMetadataTarget)
		}

		const pivotCachePartPaths = new Set(workbook.pivotCaches.map((c) => c.partPath))
		const pivotCacheRelIds: string[] = []
		for (const cache of workbook.pivotCaches) {
			const target = cache.partPath.replace(/^xl\//, '')
			pivotCacheRelIds.push(addWorkbookRel(REL_PIVOT_CACHE_DEFINITION, cache.partPath, target))
		}

		const workbookCapsulePartPaths = new Set(workbookCapsules.map((capsule) => capsule.partPath))
		const slicerCachePartPaths = new Set(workbook.slicerCaches.map((cache) => cache.partPath))
		const slicerCacheRelIds: string[] = []
		for (const cache of workbook.slicerCaches) {
			if (!workbookCapsulePartPaths.has(cache.partPath)) continue
			const target = cache.partPath.replace(/^xl\//, '')
			slicerCacheRelIds.push(addWorkbookRel(REL_SLICER_CACHE, cache.partPath, target))
		}

		const timelineCachePartPaths = new Set(workbook.timelineCaches.map((cache) => cache.partPath))
		const timelineCacheRelIds: string[] = []
		for (const cache of workbook.timelineCaches) {
			if (!workbookCapsulePartPaths.has(cache.partPath)) continue
			const target = cache.partPath.replace(/^xl\//, '')
			timelineCacheRelIds.push(addWorkbookRel(REL_TIMELINE_CACHE, cache.partPath, target))
		}

		const chartSheetPartPaths = new Set(workbook.chartSheets.map((sheet) => sheet.partPath))
		const chartSheetRelIds: string[] = []
		for (const chartSheet of workbook.chartSheets) {
			const target = chartSheet.partPath.replace(/^xl\//, '')
			const relId = addWorkbookRel(REL_CHARTSHEET, chartSheet.partPath, target)
			chartSheetRelIds.push(relId)
		}

		const macroSheetPartPaths = new Set(workbook.macroSheets.map((sheet) => sheet.partPath))
		const macroSheetRelIds: string[] = []
		for (const macroSheet of workbook.macroSheets) {
			const target = macroSheet.partPath.replace(/^xl\//, '')
			const relId = addWorkbookRel(REL_MACROSHEET, macroSheet.partPath, target)
			macroSheetRelIds.push(relId)
		}

		for (const capsule of workbookCapsules) {
			if (!capsule.relType) continue
			if (isPackageDocPropsCapsule(capsule)) continue
			if (isPackageSignatureOriginCapsule(capsule)) continue
			if (chartSheetPartPaths.has(capsule.partPath)) continue
			if (macroSheetPartPaths.has(capsule.partPath)) continue
			if (pivotCachePartPaths.has(capsule.partPath)) continue
			if (slicerCachePartPaths.has(capsule.partPath)) continue
			if (timelineCachePartPaths.has(capsule.partPath)) continue
			const target = computeRelativePath('xl/', capsule.partPath)
			addWorkbookRel(capsule.relType, capsule.partPath, target)
		}
		const orderedWbRels = orderWorkbookRels(wbRels, preservedWorkbookRels)

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
		const preserveWorkbookCalcState = preserveWorkbookXml && !options.calcStateDirty
		const preserveWorkbookRels =
			preserveWorkbookCalcState &&
			(!shouldWriteDynamicArrayMetadata ||
				preservedWorkbookRelsText?.includes(REL_SHEET_METADATA) === true)
		const preserveCalcChainCapsules = !effectiveCalcChainDirty
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
								worksheetRelIds,
								externalReferenceRelIds,
								pivotCacheRelIds,
								slicerCacheRelIds,
								timelineCacheRelIds,
								chartSheetRelIds,
								macroSheetRelIds,
								...(options.calcStateDirty !== undefined
									? { calcStateDirty: options.calcStateDirty }
									: {}),
							}),
			)
		}
		if (shouldWriteStyles) {
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
		}
		const preservedThemeBytes = workbook.preservedTheme
			? resolvePreservedBytes(sourceArchive, workbook.preservedTheme.path)
			: undefined
		if (workbook.preservedTheme && preservedThemeXml) {
			const updatePreservedTheme = !themeXmlMatches(
				preservedThemeXml,
				workbook.themeMetadata,
				workbook.themeColors,
			)
			if (updatePreservedTheme) {
				recordXml(
					workbook.preservedTheme.path,
					{
						owner: { kind: 'workbook' },
						origin: 'generated',
						contentType: workbook.preservedTheme.contentType,
					},
					() => updateThemeXml(preservedThemeXml, workbook.themeMetadata, workbook.themeColors),
				)
			} else if (preservedThemeBytes && !options.summaryOnly) {
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
				() => buildThemeXml(workbook.themeMetadata, workbook.themeColors),
			)
			plan.addOverride(generatedThemePath, generatedThemeContentType)
		}

		for (let i = 0; i < workbook.sheets.length; i++) {
			const sheet = workbook.sheets[i]
			if (!sheet) continue
			const sheetPartPath = worksheetPartPath(sheet, i)
			const sheetRelsPath = worksheetRelsPath(sheet, i)
			const sheetCapsules = sheetCapsuleMap.get(sheet.id) ?? []
			const preservedSheetXml = sheet.preservedXml
			const sheetRels: RelEntry[] = []
			let sheetRelId = 1
			let commentsRelId: string | undefined
			let drawingRelId: string | undefined
			let legacyDrawingRelId: string | undefined
			const tableRelIds: string[] = []
			const usedSheetRelIds = new Set<string>()
			const nextSheetRelId = (preferred?: string): string => {
				if (preferred && !usedSheetRelIds.has(preferred)) {
					usedSheetRelIds.add(preferred)
					return preferred
				}
				while (usedSheetRelIds.has(`rId${sheetRelId}`)) sheetRelId++
				const relId = `rId${sheetRelId}`
				usedSheetRelIds.add(relId)
				sheetRelId++
				return relId
			}
			const commentsCapsule = sheetCapsules.find((capsule) => capsule.relType === REL_COMMENTS)
			const tableCapsules = sheetCapsules.filter((capsule) => capsule.relType === REL_TABLE)
			const usedTableCapsules = new Set<PreservationCapsule>()
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
			const sheetRelTarget = (type: string, partPath: string, fallback: string): string =>
				preservedRelationshipTarget(preservedSheetRelsText, sheetPartPath, type, partPath) ??
				fallback
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
				const relId = nextSheetRelId(capsule.relId)
				sheetRels.push({
					id: relId,
					type: capsule.relType,
					...(capsule.relTypeRaw ? { rawType: capsule.relTypeRaw } : {}),
					target: sheetRelTarget(
						capsule.relType,
						capsule.partPath,
						computeRelativePath('xl/worksheets/', capsule.partPath),
					),
				})
				if (capsule.relType === REL_COMMENTS && !commentsRelId) {
					commentsRelId = relId
				}
				if (capsule.relType === REL_DRAWING && !drawingRelId) drawingRelId = relId
				if (capsule.relType === REL_VML_DRAWING && !legacyDrawingRelId) legacyDrawingRelId = relId
			}
			for (const [ref, hyperlink] of sheet.hyperlinks) {
				if (hyperlink.target) {
					const relId = nextSheetRelId()
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
					const existingVmlCapsule = sheetCapsules.find(
						(capsule) => capsule.relType === REL_VML_DRAWING,
					)
					const commentsPartPath =
						commentsCapsule?.partPath ?? `xl/comments${nextGeneratedCommentsNumber}.xml`
					const vmlPartPath =
						existingVmlCapsule?.partPath ?? `xl/drawings/vmlDrawing${nextGeneratedVmlNumber}.vml`
					const canPreserveComments =
						commentsCapsule &&
						existingVmlCapsule &&
						commentsMatchSource(sourceArchive, commentsCapsule, sheet.comments)
					const canPreserveCommentVml =
						existingVmlCapsule &&
						commentVmlRefsMatchSource(sourceArchive, existingVmlCapsule, sheet.comments)
					if (!canPreserveComments) {
						recordXml(
							commentsPartPath,
							{
								owner: { kind: 'sheet', sheetName: sheet.name },
								origin: 'generated',
								contentType: CT_COMMENTS,
							},
							() => buildCommentsXml(sheet),
						)
						plan.addOverride(commentsPartPath, CT_COMMENTS)
						if (commentsCapsule) plan.skipCapsulePath(commentsCapsule.partPath)
					}
					if (!canPreserveCommentVml) {
						recordXml(
							vmlPartPath,
							{
								owner: { kind: 'sheet', sheetName: sheet.name },
								origin: 'generated',
								contentType: CT_VML,
							},
							() => buildCommentsVml(sheet),
						)
						plan.addOverride(vmlPartPath, CT_VML)
						if (existingVmlCapsule) plan.skipCapsulePath(existingVmlCapsule.partPath)
					}
					if (!canPreserveComments || !canPreserveCommentVml) {
						if (!commentsRelId) {
							commentsRelId = nextSheetRelId()
							sheetRels.push({
								id: commentsRelId,
								type: REL_COMMENTS,
								target: sheetRelTarget(
									REL_COMMENTS,
									commentsPartPath,
									computeRelativePath('xl/worksheets/', commentsPartPath),
								),
							})
						}
						if (!legacyDrawingRelId) {
							legacyDrawingRelId = nextSheetRelId()
							sheetRels.push({
								id: legacyDrawingRelId,
								type: REL_VML_DRAWING,
								target: sheetRelTarget(
									REL_VML_DRAWING,
									vmlPartPath,
									computeRelativePath('xl/worksheets/', vmlPartPath),
								),
							})
						}
						if (!commentsCapsule) nextGeneratedCommentsNumber++
						if (!existingVmlCapsule) nextGeneratedVmlNumber++
					}
				}
				for (let tableIndex = 0; tableIndex < sheet.tables.length; tableIndex++) {
					const table = sheet.tables[tableIndex]
					if (!table) continue
					const tableCapsule =
						findPreservableTableCapsule(
							table,
							tableCapsules,
							usedTableCapsules,
							sourceArchive,
							sheet.id,
						) ??
						findTableCapsuleByIdentity(
							table,
							tableCapsules,
							usedTableCapsules,
							sourceArchive,
							sheet.id,
						) ??
						tableCapsules.find((capsule) => !usedTableCapsules.has(capsule))
					if (tableCapsule) usedTableCapsules.add(tableCapsule)
					const tablePartPath =
						tableCapsule?.partPath ?? `xl/tables/table${nextGeneratedTableNumber}.xml`
					const tableContentType = tableCapsule?.contentType ?? CT_TABLE
					const tableCapsuleContent = readCapsuleBytes(tableCapsule, sourceArchive)
					if (
						tableCapsuleContent &&
						canPreserveTableCapsule(
							table,
							tableCapsuleContent,
							sheet.id,
							tablePartPath,
							tableCapsule?.relationships,
						)
					) {
						recordBytes(
							tablePartPath,
							{
								owner: { kind: 'sheet', sheetName: sheet.name },
								origin: 'capsule',
								contentType: tableContentType,
							},
							() => tableCapsuleContent,
						)
					} else {
						recordXml(
							tablePartPath,
							{
								owner: { kind: 'sheet', sheetName: sheet.name },
								origin: 'generated',
								contentType: tableContentType,
							},
							() => buildTableXml(table, nextGeneratedTableNumber),
						)
					}
					const tablePartRelationships = buildTablePartRelationships(
						table,
						tableCapsule,
						tablePartPath,
					)
					if (tablePartRelationships.length > 0) {
						recordXml(
							getRelsPath(tablePartPath),
							{
								owner: { kind: 'sheet', sheetName: sheet.name },
								origin: tableCapsule ? 'capsule' : 'generated',
							},
							() => buildRelsXml(tablePartRelationships),
						)
					}
					plan.addOverride(tablePartPath, tableContentType)
					if (tableCapsule) plan.skipCapsulePath(tableCapsule.partPath)
					const relId = nextSheetRelId(tableCapsule?.relId)
					sheetRels.push({
						id: relId,
						type: REL_TABLE,
						...(tableCapsule?.relTypeRaw ? { rawType: tableCapsule.relTypeRaw } : {}),
						target: sheetRelTarget(
							REL_TABLE,
							tablePartPath,
							computeRelativePath('xl/worksheets/', tablePartPath),
						),
					})
					tableRelIds.push(relId)
					nextGeneratedTableNumber++
				}
				for (const tableCapsule of tableCapsules) {
					if (!usedTableCapsules.has(tableCapsule)) plan.skipCapsulePath(tableCapsule.partPath)
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
					drawingRelId = nextSheetRelId()
					sheetRels.push({
						id: drawingRelId,
						type: REL_DRAWING,
						target: sheetRelTarget(
							REL_DRAWING,
							drawingPartPath,
							computeRelativePath('xl/worksheets/', drawingPartPath),
						),
					})
					nextGeneratedDrawingNumber++
					hasGeneratedDrawing = true
				}
			}
			if (preserveSheetXml && preservedSheetXmlBytes && !options.summaryOnly) {
				recordBytes(
					sheetPartPath,
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
					usePlainStrings,
					batchRows: true,
					omitDenseCellRefs,
					...(cfOverrides ? { cfDxfIdOverrides: cfOverrides } : {}),
				}
				recordStreamingSheet(
					sheetPartPath,
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
					sheetPartPath,
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
									usePlainStrings,
									batchRows: true,
									omitDenseCellRefs,
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
						sheetRelsPath,
						{
							owner: { kind: 'sheet', sheetName: sheet.name },
							origin: resolvePreservedOrigin(preservedSheetXml?.relsXml),
						},
						() => preservedSheetRelsBytes,
					)
				} else {
					recordXml(
						sheetRelsPath,
						{
							owner: { kind: 'sheet', sheetName: sheet.name },
							origin: resolvePreservedOrigin(preservedSheetXml?.relsXml),
						},
						() => preservedSheetRelsText ?? '',
					)
				}
			} else if (sheetRels.length > 0) {
				recordXml(
					sheetRelsPath,
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
				preservedSharedStringsXml !== undefined &&
				(preserveSharedStrings || dirtyPatchMode || hasPreservedSheetXmlInDirtyPatchMode) &&
				(useInlineStrings || ssTable.entryCount <= preservedSharedStringEntries.length)
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

		const corePropsPath = packageDocPropsPath(workbookCapsules, REL_CORE_PROPS, 'docProps/core.xml')
		const appPropsPath = packageDocPropsPath(workbookCapsules, REL_EXT_PROPS, 'docProps/app.xml')
		const customPropsPath = packageDocPropsPath(
			workbookCapsules,
			REL_CUSTOM_PROPS,
			'docProps/custom.xml',
		)
		const hasCoreOrAppDocProps =
			workbookCapsules.some((capsule) => capsule.relType === REL_CORE_PROPS) ||
			workbookCapsules.some((capsule) => capsule.relType === REL_EXT_PROPS)
		const documentPropertiesDirty = options.documentPropertiesDirty === true
		const customDocumentProperties = workbook.documentProperties.custom ?? []
		const shouldWriteDocProps = documentPropertiesDirty || !dirtyPatchMode || hasCoreOrAppDocProps
		const shouldWriteCustomDocProps = documentPropertiesDirty
			? customDocumentProperties.length > 0
			: workbookCapsules.some(
					(capsule) => capsule.partPath === customPropsPath || capsule.relType === REL_CUSTOM_PROPS,
				)

		if (shouldWriteDocProps) {
			recordDocPropsPart(
				plan,
				sourceArchive,
				capsules,
				corePropsPath,
				'application/vnd.openxmlformats-package.core-properties+xml',
				() => buildCorePropsXml(workbook.documentProperties.core),
				!documentPropertiesDirty,
			)
			recordDocPropsPart(
				plan,
				sourceArchive,
				capsules,
				appPropsPath,
				'application/vnd.openxmlformats-officedocument.extended-properties+xml',
				() => buildAppPropsXml(workbook.documentProperties.app),
				!documentPropertiesDirty,
			)
			if (shouldWriteCustomDocProps) {
				recordDocPropsPart(
					plan,
					sourceArchive,
					capsules,
					customPropsPath,
					'application/vnd.openxmlformats-officedocument.custom-properties+xml',
					() => buildCustomPropsXml(customDocumentProperties),
					!documentPropertiesDirty,
				)
				plan.addOverride(
					customPropsPath,
					'application/vnd.openxmlformats-officedocument.custom-properties+xml',
				)
			} else if (documentPropertiesDirty) {
				plan.skipCapsulePath(customPropsPath)
			}
		}

		const rootRels: RelEntry[] = []
		const preservedRootRels = preservedRootRelsText ? parseRelationships(preservedRootRelsText) : []
		const reservedRootRelIds = new Set(preservedRootRels.map((rel) => rel.id))
		const usedRootRelIds = new Set<string>()
		let rootRIdCounter = 1
		const addRootRel = (type: string, partPath: string, fallback: string): string => {
			const preserved = preservedPackageRelationship(preservedRootRels, '', type, partPath)
			const id = allocateRelationshipId(
				preserved?.id,
				reservedRootRelIds,
				usedRootRelIds,
				() => `rId${rootRIdCounter++}`,
			)
			rootRels.push({
				id,
				type: preserved?.type ?? type,
				...(preserved?.rawType ? { rawType: preserved.rawType } : {}),
				target: preserved?.target ?? rootRelTarget(type, partPath, fallback),
			})
			return id
		}
		addRootRel(REL_OFFICE_DOC, 'xl/workbook.xml', 'xl/workbook.xml')
		if (shouldWriteDocProps) {
			addRootRel(REL_CORE_PROPS, corePropsPath, corePropsPath)
			addRootRel(REL_EXT_PROPS, appPropsPath, appPropsPath)
			if (shouldWriteCustomDocProps) addRootRel(REL_CUSTOM_PROPS, customPropsPath, customPropsPath)
		}
		if (capsules) {
			for (const capsule of capsules) {
				if (plan.isCapsulePathSkipped(capsule.partPath)) continue
				const preservedRootRelationship = capsule.relType
					? preservedPackageRelationship(preservedRootRels, '', capsule.relType, capsule.partPath)
					: undefined
				if (
					!isPackageDocPropsCapsule(capsule) &&
					!isPackageSignatureOriginCapsule(capsule) &&
					!preservedRootRelationship
				) {
					continue
				}
				if (!capsule.relType) continue
				if (
					capsule.partPath === corePropsPath ||
					capsule.partPath === appPropsPath ||
					capsule.partPath === customPropsPath
				) {
					continue
				}
				addRootRel(capsule.relType, capsule.partPath, capsule.partPath)
			}
		}
		const orderedRootRels = orderWorkbookRels(rootRels, preservedRootRels)
		recordXml(
			'_rels/.rels',
			{
				owner: { kind: 'package' },
				origin: 'generated',
			},
			() => buildRelsXml(orderedRootRels),
		)

		if (capsules) {
			for (const capsule of capsules) {
				if (isCalcChainCapsule(capsule) && !preserveCalcChainCapsules) continue
				if (plan.isCapsulePathSkipped(capsule.partPath)) continue
				const content = capsule.content ?? sourceArchive?.readBytes(capsule.partPath)
				if (!content) continue
				const owner = resolveCapsuleOwner(capsule, sheetNameById)
				if (!owner) continue
				const pivotTable = workbook.pivotTables.find((pivot) => pivot.partPath === capsule.partPath)
				if (pivotTable && isPivotTableDefinitionCapsule(capsule)) {
					const sourceXml = new TextDecoder().decode(content)
					if (!shouldUpdatePivotTableDefinitionXml(sourceXml, pivotTable)) {
						recordUnchangedCapsule(capsule, owner, content)
						continue
					}
					recordXml(
						capsule.partPath,
						{
							owner,
							origin: 'generated',
							contentType: capsule.contentType,
						},
						() => updatePivotTableDefinitionXml(sourceXml, pivotTable),
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
				const pivotCache = workbook.pivotCaches.find((cache) => cache.partPath === capsule.partPath)
				if (pivotCache && isPivotCacheDefinitionCapsule(capsule)) {
					const sourceXml = new TextDecoder().decode(content)
					if (!shouldUpdatePivotCacheDefinitionXml(sourceXml, pivotCache, capsule.relationships)) {
						recordUnchangedCapsule(capsule, owner, content)
						continue
					}
					recordXml(
						capsule.partPath,
						{
							owner,
							origin: 'generated',
							contentType: capsule.contentType,
						},
						() => updatePivotCacheDefinitionXml(sourceXml, pivotCache),
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
				const slicerCache = workbook.slicerCaches.find(
					(cache) => cache.partPath === capsule.partPath,
				)
				if (slicerCache && isSlicerCacheDefinitionCapsule(capsule)) {
					const sourceXml = new TextDecoder().decode(content)
					if (!shouldUpdateSlicerCacheDefinitionXml(sourceXml, slicerCache)) {
						recordUnchangedCapsule(capsule, owner, content)
						continue
					}
					recordXml(
						capsule.partPath,
						{
							owner,
							origin: 'generated',
							contentType: capsule.contentType,
						},
						() => updateSlicerCacheDefinitionXml(sourceXml, slicerCache),
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
				const timelineCache = workbook.timelineCaches.find(
					(cache) => cache.partPath === capsule.partPath,
				)
				if (timelineCache && isTimelineCacheDefinitionCapsule(capsule)) {
					const sourceXml = new TextDecoder().decode(content)
					if (!shouldUpdateTimelineCacheDefinitionXml(sourceXml, timelineCache)) {
						recordUnchangedCapsule(capsule, owner, content)
						continue
					}
					recordXml(
						capsule.partPath,
						{
							owner,
							origin: 'generated',
							contentType: capsule.contentType,
						},
						() => updateTimelineCacheDefinitionXml(sourceXml, timelineCache),
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
				const connectionParts = workbook.connectionParts.filter(
					(part) => part.partPath === capsule.partPath,
				)
				if (connectionParts.length > 0 && isConnectionRefreshCapsule(capsule)) {
					const sourceXml = new TextDecoder().decode(content)
					if (!shouldUpdateConnectionPartXml(capsule, sourceXml, connectionParts)) {
						recordUnchangedCapsule(capsule, owner, content)
						continue
					}
					recordXml(
						capsule.partPath,
						{
							owner,
							origin: 'generated',
							contentType: capsule.contentType,
						},
						() => updateConnectionPartXml(sourceXml, connectionParts),
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
				if (isDrawingCapsule(capsule)) {
					const xml = new TextDecoder().decode(content)
					const drawingTextUpdates = collectDrawingTextUpdates(workbook, capsule.partPath, xml)
					if (drawingTextUpdates.length > 0) {
						recordXml(
							capsule.partPath,
							{
								owner,
								origin: 'generated',
								contentType: capsule.contentType,
							},
							() => updateDrawingTextXml(xml, drawingTextUpdates),
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
								() => buildRelsXml(resolveCapsuleRelationships(workbook, capsule)),
							)
						}
						continue
					}
				}
				if (isThreadedCommentCapsule(capsule)) {
					const xml = new TextDecoder().decode(content)
					const threadedCommentUpdates = collectThreadedCommentTextUpdates(
						workbook,
						capsule.partPath,
						xml,
					)
					if (threadedCommentUpdates.length > 0) {
						recordXml(
							capsule.partPath,
							{
								owner,
								origin: 'generated',
								contentType: capsule.contentType,
							},
							() => updateThreadedCommentsXml(xml, threadedCommentUpdates),
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
						() => buildRelsXml(resolveCapsuleRelationships(workbook, capsule)),
					)
				} else {
					const generatedExternalLinkRels = buildGeneratedExternalLinkRelationships(
						workbook,
						capsule.partPath,
					)
					if (generatedExternalLinkRels.length > 0) {
						const capsuleRelsPath = getRelsPath(capsule.partPath)
						recordXml(
							capsuleRelsPath,
							{
								owner,
								origin: 'generated',
							},
							() => buildRelsXml(generatedExternalLinkRels),
						)
					}
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
						: buildRelsXml(orderedWbRels),
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
				const finalPartPaths = new Set(built.descriptors.map((descriptor) => descriptor.path))
				return buildContentTypesXml(
					workbook.sheets.map((sheet, index) => worksheetPartPath(sheet, index)),
					hasSharedStrings,
					workbookContentType,
					capsules?.filter(
						(capsule) =>
							(!isCalcChainCapsule(capsule) || preserveCalcChainCapsules) &&
							!built.skippedCapsulePaths.has(capsule.partPath),
					),
					built.extraOverrides.length > 0 ? built.extraOverrides : undefined,
					workbook.preservedXml?.contentTypeDefaults,
					filterPreservedContentTypeOverrides(
						workbook.preservedXml?.contentTypeOverrides,
						workbook.preservedXml?.contentTypeDefaults,
						finalPartPaths,
					),
					{ corePropsPath, appPropsPath },
					{ includeStyles: shouldWriteStyles, includeDocProps: shouldWriteDocProps },
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

function collectLiveQueryTablePartPaths(workbook: Workbook): ReadonlySet<string> {
	const paths = new Set<string>()
	for (const sheet of workbook.sheets) {
		for (const table of sheet.tables) {
			if (table.queryTable?.partPath) paths.add(table.queryTable.partPath)
		}
	}
	for (const part of workbook.connectionParts) {
		if (part.kind === 'queryTable') paths.add(part.partPath)
	}
	return paths
}

function isDeletedQueryTableCapsule(
	capsule: PreservationCapsule,
	liveQueryTablePartPaths: ReadonlySet<string>,
): boolean {
	if (isQueryTableDefinitionCapsule(capsule)) return !liveQueryTablePartPaths.has(capsule.partPath)
	return false
}

function canPreserveTableCapsule(
	table: Table,
	content: Uint8Array,
	sheetId: SheetId,
	tablePath?: string,
	relationships: readonly RelEntry[] = [],
): boolean {
	const parsed = parseTable(new TextDecoder().decode(content), sheetId, {
		...(tablePath ? { tablePath } : {}),
		relationships,
	})
	return parsed ? tablesHaveSameWritableModel(table, parsed) : false
}

function findPreservableTableCapsule(
	table: Table,
	capsules: readonly PreservationCapsule[],
	used: ReadonlySet<PreservationCapsule>,
	sourceArchive: ZipArchive | undefined,
	sheetId: SheetId,
): PreservationCapsule | undefined {
	for (const capsule of capsules) {
		if (used.has(capsule)) continue
		const content = readCapsuleBytes(capsule, sourceArchive)
		if (
			content &&
			canPreserveTableCapsule(table, content, sheetId, capsule.partPath, capsule.relationships)
		) {
			return capsule
		}
	}
	return undefined
}

function findTableCapsuleByIdentity(
	table: Table,
	capsules: readonly PreservationCapsule[],
	used: ReadonlySet<PreservationCapsule>,
	sourceArchive: ZipArchive | undefined,
	sheetId: SheetId,
): PreservationCapsule | undefined {
	for (const capsule of capsules) {
		if (used.has(capsule)) continue
		const content = readCapsuleBytes(capsule, sourceArchive)
		const parsed = content
			? parseTable(new TextDecoder().decode(content), sheetId, {
					tablePath: capsule.partPath,
					relationships: capsule.relationships,
				})
			: null
		if (parsed?.name === table.name) return capsule
	}
	return undefined
}

function readCapsuleBytes(
	capsule: PreservationCapsule | undefined,
	sourceArchive: ZipArchive | undefined,
): Uint8Array | undefined {
	return capsule ? (capsule.content ?? sourceArchive?.readBytes(capsule.partPath)) : undefined
}

function tablesHaveSameWritableModel(left: Table, right: Table): boolean {
	return (
		left.name === right.name &&
		(left.nameAttribute === undefined || left.nameAttribute === right.nameAttribute) &&
		left.sheetId === right.sheetId &&
		(left.uid ?? null) === (right.uid ?? null) &&
		left.hasHeaders === right.hasHeaders &&
		left.hasTotals === right.hasTotals &&
		(left.tableType ?? null) === (right.tableType ?? null) &&
		(left.insertRow ?? null) === (right.insertRow ?? null) &&
		(left.insertRowShift ?? null) === (right.insertRowShift ?? null) &&
		(left.altText ?? null) === (right.altText ?? null) &&
		(left.altTextSummary ?? null) === (right.altTextSummary ?? null) &&
		(left.dxfId ?? null) === (right.dxfId ?? null) &&
		(left.dataCellStyle ?? null) === (right.dataCellStyle ?? null) &&
		(left.headerRowDxfId ?? null) === (right.headerRowDxfId ?? null) &&
		(left.headerRowCellStyle ?? null) === (right.headerRowCellStyle ?? null) &&
		(left.dataDxfId ?? null) === (right.dataDxfId ?? null) &&
		(left.totalsRowDxfId ?? null) === (right.totalsRowDxfId ?? null) &&
		(left.headerRowBorderDxfId ?? null) === (right.headerRowBorderDxfId ?? null) &&
		(left.tableBorderDxfId ?? null) === (right.tableBorderDxfId ?? null) &&
		rangesEqual(left.ref, right.ref) &&
		tableColumnsEqual(left.columns, right.columns) &&
		stableJson(left.autoFilter) === stableJson(right.autoFilter) &&
		stableJson(left.sortState) === stableJson(right.sortState) &&
		stableJson(left.tableStyleInfo) === stableJson(right.tableStyleInfo) &&
		stableJson(left.queryTable) === stableJson(right.queryTable)
	)
}

function rangesEqual(left: Table['ref'], right: Table['ref']): boolean {
	return (
		left.start.row === right.start.row &&
		left.start.col === right.start.col &&
		left.end.row === right.end.row &&
		left.end.col === right.end.col
	)
}

function tableColumnsEqual(
	left: readonly Table['columns'][number][],
	right: readonly Table['columns'][number][],
): boolean {
	if (left.length !== right.length) return false
	for (let i = 0; i < left.length; i++) {
		const leftColumn = left[i]
		const rightColumn = right[i]
		if (!leftColumn || !rightColumn) return false
		if (
			(leftColumn.id ?? null) !== (rightColumn.id ?? null) ||
			(leftColumn.uid ?? null) !== (rightColumn.uid ?? null) ||
			(leftColumn.uniqueName ?? null) !== (rightColumn.uniqueName ?? null) ||
			leftColumn.name !== rightColumn.name ||
			(leftColumn.formula ?? null) !== (rightColumn.formula ?? null) ||
			(leftColumn.formulaIsArray ?? null) !== (rightColumn.formulaIsArray ?? null) ||
			stableJson(leftColumn.xmlColumnPr) !== stableJson(rightColumn.xmlColumnPr) ||
			(leftColumn.totalsRowFunction ?? null) !== (rightColumn.totalsRowFunction ?? null) ||
			(leftColumn.totalsRowFormula ?? null) !== (rightColumn.totalsRowFormula ?? null) ||
			(leftColumn.totalsRowLabel ?? null) !== (rightColumn.totalsRowLabel ?? null) ||
			(leftColumn.queryTableFieldId ?? null) !== (rightColumn.queryTableFieldId ?? null) ||
			(leftColumn.dataCellStyle ?? null) !== (rightColumn.dataCellStyle ?? null) ||
			(leftColumn.dataDxfId ?? null) !== (rightColumn.dataDxfId ?? null) ||
			(leftColumn.headerRowDxfId ?? null) !== (rightColumn.headerRowDxfId ?? null) ||
			(leftColumn.totalsRowDxfId ?? null) !== (rightColumn.totalsRowDxfId ?? null)
		) {
			return false
		}
	}
	return true
}

function buildTablePartRelationships(
	table: Table,
	capsule: PreservationCapsule | undefined,
	tablePartPath: string,
): readonly RelEntry[] {
	const relationships = [...(capsule?.relationships ?? [])]
	if (!table.queryTable) return relationships

	const queryTableRel: RelEntry = {
		id: table.queryTable.relationshipId,
		type: table.queryTable.relationshipType || REL_QUERY_TABLE,
		...(table.queryTable.relationshipRawType
			? { rawType: table.queryTable.relationshipRawType }
			: {}),
		target: table.queryTable.targetMode
			? table.queryTable.target
			: computeRelativePath(partDir(tablePartPath), table.queryTable.partPath),
		...(table.queryTable.targetMode ? { targetMode: table.queryTable.targetMode } : {}),
	}
	const index = relationships.findIndex(
		(rel) => rel.id === queryTableRel.id || rel.type === REL_QUERY_TABLE,
	)
	if (index >= 0) {
		relationships[index] = queryTableRel
	} else {
		relationships.push(queryTableRel)
	}
	return relationships
}

function stableJson(value: unknown): string {
	if (value === undefined) return ''
	if (value === null || typeof value !== 'object') return JSON.stringify(value)
	if (Array.isArray(value)) return `[${value.map(stableJson).join(',')}]`
	const record = value as Record<string, unknown>
	return `{${Object.keys(record)
		.sort()
		.map((key) => `${JSON.stringify(key)}:${stableJson(record[key])}`)
		.join(',')}}`
}

function recordDocPropsPart(
	plan: WritePlanBuilder,
	sourceArchive: ZipArchive | undefined,
	capsules: readonly PreservationCapsule[] | undefined,
	path: string,
	contentType: string,
	buildXml: () => string,
	preserveExisting = true,
): void {
	const capsule = capsules?.find((entry) => entry.partPath === path)
	const content = capsule?.content ?? sourceArchive?.readBytes(path)
	if (preserveExisting && capsule && content) {
		plan.putBytes(path, content, {
			owner: { kind: 'package' },
			origin: 'capsule',
			contentType: capsule.contentType || contentType,
		})
		plan.skipCapsulePath(path)
		return
	}
	if (capsule) plan.skipCapsulePath(path)
	plan.putXml(path, buildXml(), {
		owner: { kind: 'package' },
		origin: 'generated',
		contentType,
	})
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

function hasThemeContent(workbook: Workbook): boolean {
	const metadata = workbook.themeMetadata
	return (
		workbook.themeColors.length > 0 ||
		metadata.colorCount > 0 ||
		(metadata.name?.trim().length ?? 0) > 0 ||
		(metadata.colorSchemeName?.trim().length ?? 0) > 0 ||
		(metadata.majorFontLatin?.trim().length ?? 0) > 0 ||
		(metadata.minorFontLatin?.trim().length ?? 0) > 0
	)
}

function canPreserveDynamicArrayMetadata(
	entries: readonly DynamicArrayMetadataEntry[],
	preserved:
		| readonly { readonly metadataIndex: number; readonly collapsed?: boolean }[]
		| undefined,
): boolean {
	if (entries.length === 0) return true
	if (!preserved) return false
	const preservedByIndex = new Map<number, boolean | undefined>()
	for (const entry of preserved) {
		preservedByIndex.set(entry.metadataIndex, entry.collapsed)
	}
	for (const entry of entries) {
		if (!preservedByIndex.has(entry.metadataIndex)) return false
		if ((preservedByIndex.get(entry.metadataIndex) ?? false) !== entry.collapsed) return false
	}
	return true
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

function workbookHasSharedStringEligibleCells(workbook: Workbook): boolean {
	return workbook.sheets.some(
		(sheet) => sheet.cells.stringCellCount() > 0 || sheet.cells.richTextCellCount() > 0,
	)
}

function workbookHasFormulaInfoCells(workbook: Workbook): boolean {
	return workbook.sheets.some((sheet) => sheet.cells.formulaInfoCellCount() > 0)
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

function commentsMatchSource(
	sourceArchive: ZipArchive | undefined,
	commentsCapsule: PreservationCapsule,
	comments: ReadonlyMap<string, { readonly text: string; readonly author?: string }>,
): boolean {
	const xml = sourceArchive?.readText(commentsCapsule.partPath)
	if (!xml) return false
	try {
		return commentsEqual(parseCommentsXml(xml), comments)
	} catch {
		return false
	}
}

function commentVmlRefsMatchSource(
	sourceArchive: ZipArchive | undefined,
	vmlCapsule: PreservationCapsule,
	comments: ReadonlyMap<string, { readonly text: string; readonly author?: string }>,
): boolean {
	const xml = sourceArchive?.readText(vmlCapsule.partPath)
	if (!xml) return false
	try {
		const layouts = parseCommentVmlXml(xml)
		if (layouts.size !== comments.size) return false
		for (const ref of comments.keys()) {
			if (!layouts.has(ref)) return false
		}
		return true
	} catch {
		return false
	}
}

function commentsEqual(
	left: ReadonlyMap<string, { readonly text: string; readonly author?: string }>,
	right: ReadonlyMap<string, { readonly text: string; readonly author?: string }>,
): boolean {
	if (left.size !== right.size) return false
	for (const [ref, comment] of left) {
		const other = right.get(ref)
		if (!other || other.text !== comment.text || other.author !== comment.author) return false
	}
	return true
}

function isPackageDocPropsCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.partPath.startsWith('docProps/') ||
		capsule.relType === REL_CORE_PROPS ||
		capsule.relType === REL_EXT_PROPS
	)
}

function isPackageSignatureOriginCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.relType === REL_DIGITAL_SIGNATURE_ORIGIN ||
		capsule.contentType.includes('digital-signature-origin') ||
		capsule.partPath === '_xmlsignatures/origin.sigs'
	)
}

function isDigitalSignatureCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.partPath.startsWith('_xmlsignatures/') ||
		capsule.contentType.includes('digital-signature') ||
		capsule.relType?.includes('digital-signature') === true
	)
}

function packageDocPropsPath(
	capsules: readonly PreservationCapsule[],
	relType: string,
	fallback: string,
): string {
	return capsules.find((capsule) => capsule.relType === relType)?.partPath ?? fallback
}

function allocateWorkbookRelId(
	preservedRels: readonly Relationship[],
	reservedRelIds: ReadonlySet<string>,
	usedRelIds: Set<string>,
	sourcePart: string,
	type: string,
	resolvedPartPath: string,
	nextGenerated: () => string,
): string {
	const preserved = preservedRelationshipId(preservedRels, sourcePart, type, resolvedPartPath)
	return allocateRelationshipId(preserved, reservedRelIds, usedRelIds, nextGenerated)
}

function allocateRelationshipId(
	preservedId: string | undefined,
	reservedRelIds: ReadonlySet<string>,
	usedRelIds: Set<string>,
	nextGenerated: () => string,
): string {
	if (preservedId && !usedRelIds.has(preservedId)) {
		usedRelIds.add(preservedId)
		return preservedId
	}
	let generated = nextGenerated()
	while (usedRelIds.has(generated) || reservedRelIds.has(generated)) generated = nextGenerated()
	usedRelIds.add(generated)
	return generated
}

function preservedRelationshipId(
	rels: readonly Relationship[],
	sourcePart: string,
	type: string,
	resolvedPartPath: string,
): string | undefined {
	for (const rel of rels) {
		if (rel.type !== type) continue
		if (resolvePath(sourcePart, rel.target) === resolvedPartPath) return rel.id
	}
	return undefined
}

function preservedRelationshipRawType(
	rels: readonly Relationship[],
	sourcePart: string,
	type: string,
	resolvedPartPath: string,
): string | undefined {
	return preservedPackageRelationship(rels, sourcePart, type, resolvedPartPath)?.rawType
}

function preservedPackageRelationship(
	rels: readonly Relationship[],
	sourcePart: string,
	type: string,
	resolvedPartPath: string,
): Relationship | undefined {
	for (const rel of rels) {
		if (rel.type === type && resolvePath(sourcePart, rel.target) === resolvedPartPath) return rel
	}
	if (type === REL_CORE_PROPS) {
		return rels.find(
			(rel) =>
				rel.type.endsWith('/metadata/core-properties') &&
				resolvePath(sourcePart, rel.target) === resolvedPartPath,
		)
	}
	return undefined
}

function orderWorkbookRels(
	entries: readonly RelEntry[],
	preservedRels: readonly Relationship[],
): readonly RelEntry[] {
	if (preservedRels.length === 0) return entries
	const pending = new Map(entries.map((entry) => [entry.id, entry] as const))
	const ordered: RelEntry[] = []
	for (const rel of preservedRels) {
		const entry = pending.get(rel.id)
		if (!entry) continue
		ordered.push(entry)
		pending.delete(rel.id)
	}
	for (const entry of entries) {
		if (pending.delete(entry.id)) ordered.push(entry)
	}
	return ordered
}

function preservedRelationshipTarget(
	relsXml: string | undefined,
	sourcePart: string,
	type: string,
	resolvedPartPath: string,
): string | undefined {
	if (!relsXml) return undefined
	for (const rel of parseRelationships(relsXml)) {
		if (rel.type !== type) continue
		if (resolvePath(sourcePart, rel.target) === resolvedPartPath) return rel.target
	}
	return undefined
}

type PivotTableInfo = Workbook['pivotTables'][number]
type PivotCacheInfo = Workbook['pivotCaches'][number]
type SlicerCacheInfo = Workbook['slicerCaches'][number]
type TimelineCacheInfo = Workbook['timelineCaches'][number]
type ConnectionPartInfo = Workbook['connectionParts'][number]

function shouldUpdatePivotTableDefinitionXml(xml: string, pivot: PivotTableInfo): boolean {
	const parsed = parsePivotTableXml(xml, pivot.partPath, pivot.sheetName)
	if (!parsed) return true
	return stableJson(pivotTableWritableState(parsed)) !== stableJson(pivotTableWritableState(pivot))
}

function pivotTableWritableState(pivot: PivotTableInfo): unknown {
	return {
		fields: pivot.fields.map((field) => ({
			index: field.index,
			items: (field.items ?? []).map((item) => ({
				index: item.index,
				cacheIndex: item.cacheIndex ?? null,
				itemType: item.itemType ?? null,
				caption: item.caption ?? null,
				hidden: item.hidden ?? null,
				manualFilter: item.manualFilter ?? null,
				showDetails: item.showDetails ?? null,
				calculated: item.calculated ?? null,
				missing: item.missing ?? null,
				childItems: item.childItems ?? null,
				expanded: item.expanded ?? null,
				drillAcrossAttributes: item.drillAcrossAttributes ?? null,
			})),
		})),
		pageFields: pivot.pageFields.map((field) => ({
			index: field.index,
			item: field.item ?? null,
		})),
	}
}

function shouldUpdatePivotCacheDefinitionXml(
	xml: string,
	cache: PivotCacheInfo,
	relationships: readonly RelEntry[],
): boolean {
	const parsed = parsePivotCacheDefinitionXml(
		xml,
		cache.partPath,
		cache.cacheId,
		cache.relId,
		relationships,
	)
	if (!parsed) return true
	return stableJson(pivotCacheWritableState(parsed)) !== stableJson(pivotCacheWritableState(cache))
}

function pivotCacheWritableState(cache: PivotCacheInfo): unknown {
	return {
		sourceSheet: cache.sourceSheet ?? null,
		sourceRef: cache.sourceRef ?? null,
		refreshOnLoad: cache.refreshOnLoad ?? null,
		enableRefresh: cache.enableRefresh ?? null,
		invalid: cache.invalid ?? null,
		saveData: cache.saveData ?? null,
	}
}

function shouldUpdateSlicerCacheDefinitionXml(xml: string, cache: SlicerCacheInfo): boolean {
	const parsed = parseSlicerCacheXml(xml, cache.partPath)
	if (!parsed) return true
	return (
		stableJson(slicerCacheWritableState(parsed)) !== stableJson(slicerCacheWritableState(cache))
	)
}

function shouldUpdateTimelineCacheDefinitionXml(xml: string, cache: TimelineCacheInfo): boolean {
	const parsed = parseTimelineCacheXml(xml, cache.partPath)
	if (!parsed) return true
	return (
		stableJson(timelineCacheWritableState(parsed)) !== stableJson(timelineCacheWritableState(cache))
	)
}

function shouldUpdateConnectionPartXml(
	capsule: PreservationCapsule,
	xml: string,
	parts: readonly ConnectionPartInfo[],
): boolean {
	const parsed = parseConnectionPartInfos(capsule, xml)
	return (
		stableJson(parts.map(connectionRefreshWritableState)) !==
		stableJson(parsed.map(connectionRefreshWritableState))
	)
}

function slicerCacheWritableState(cache: SlicerCacheInfo): unknown {
	return {
		items: (cache.items ?? []).map((item) => ({
			index: item.index,
			selected: item.selected ?? null,
			noData: item.noData ?? null,
		})),
	}
}

function timelineCacheWritableState(cache: TimelineCacheInfo): unknown {
	return {
		selection: cache.state?.selection
			? {
					startDate: cache.state.selection.startDate,
					endDate: cache.state.selection.endDate,
				}
			: null,
		singleRangeFilterState: cache.state?.singleRangeFilterState ?? null,
	}
}

function connectionRefreshWritableState(part: ConnectionPartInfo): unknown {
	return {
		kind: part.kind,
		partPath: part.partPath,
		name: part.name ?? null,
		connectionId: part.connectionId ?? null,
		sheetName: part.sheetName ?? null,
		refreshOnLoad: part.refreshOnLoad ?? null,
		saveData: part.saveData ?? null,
		refreshedVersion: part.refreshedVersion ?? null,
	}
}

function isPivotCacheDefinitionCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.relType === REL_PIVOT_CACHE_DEFINITION ||
		capsule.contentType.includes('pivotCacheDefinition+xml') ||
		capsule.partPath.includes('/pivotCache/pivotCacheDefinition')
	)
}

function isPivotTableDefinitionCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.relType === REL_PIVOT_TABLE ||
		capsule.contentType.includes('pivotTable+xml') ||
		capsule.partPath.includes('/pivotTables/')
	)
}

function isQueryTableDefinitionCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.relType === REL_QUERY_TABLE ||
		capsule.contentType.includes('queryTable+xml') ||
		capsule.partPath.includes('/queryTables/')
	)
}

function isSlicerCacheDefinitionCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.contentType.includes('slicerCache+xml') ||
		capsule.partPath.includes('/slicerCaches/slicerCache')
	)
}

function isTimelineCacheDefinitionCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.contentType.includes('timelineCache+xml') ||
		capsule.partPath.includes('/timelineCaches/timelineCache')
	)
}

function isConnectionRefreshCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.contentType.includes('connections+xml') ||
		capsule.contentType.includes('queryTable+xml') ||
		capsule.partPath.endsWith('/connections.xml') ||
		capsule.partPath.includes('/queryTables/')
	)
}

function isChartCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.contentType.includes('chart+xml') ||
		capsule.partPath.includes('/charts/') ||
		capsule.partPath.includes('/chartEx/')
	)
}

function isDrawingCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.relType === REL_DRAWING ||
		capsule.contentType.includes('drawing+xml') ||
		capsule.partPath.includes('/drawings/drawing')
	)
}

function isThreadedCommentCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.contentType.includes('threadedcomments+xml') ||
		capsule.partPath.includes('/threadedComments/')
	)
}

function collectDrawingTextUpdates(
	workbook: Workbook,
	partPath: string,
	sourceXml: string,
): readonly DrawingTextUpdate[] {
	const sourceRefs = parseDrawingObjectRefs(sourceXml, partPath)
	const updates: DrawingTextUpdate[] = []
	const seen = new Set<string>()
	for (const sheet of workbook.sheets) {
		for (const ref of sheet.drawingObjectRefs) {
			if (ref.drawingPartPath !== partPath || ref.text === undefined) continue
			const source = findMatchingDrawingObject(sourceRefs, ref)
			if (!source || source.text === ref.text) continue
			const key = ref.id !== undefined ? `id:${ref.id}` : ref.name ? `name:${ref.name}` : null
			if (!key || seen.has(key)) continue
			updates.push({
				...(ref.id !== undefined ? { id: ref.id } : {}),
				...(ref.name ? { name: ref.name } : {}),
				text: ref.text,
			})
			seen.add(key)
		}
	}
	return updates
}

function collectThreadedCommentTextUpdates(
	workbook: Workbook,
	partPath: string,
	sourceXml: string,
): readonly ThreadedCommentTextRef[] {
	const sourceRefs = readThreadedCommentTextRefs(sourceXml)
	const updates: ThreadedCommentTextRef[] = []
	const seen = new Set<string>()
	for (const sheet of workbook.sheets) {
		for (const comment of sheet.threadedComments) {
			if (comment.partPath !== partPath) continue
			const source = findMatchingThreadedComment(sourceRefs, comment)
			if (!source || source.text === comment.text) continue
			const key = comment.id !== undefined ? `id:${comment.id}` : `ref:${comment.ref}`
			if (seen.has(key)) continue
			updates.push({
				...(comment.id !== undefined ? { id: comment.id } : { ref: comment.ref }),
				text: comment.text,
			})
			seen.add(key)
		}
	}
	return updates
}

function findMatchingDrawingObject(
	sourceRefs: readonly SheetDrawingObjectRef[],
	ref: SheetDrawingObjectRef,
): SheetDrawingObjectRef | undefined {
	const matches = sourceRefs.filter((sourceRef) => {
		if (ref.id !== undefined) return sourceRef.id === ref.id
		if (ref.name !== undefined) return sourceRef.name === ref.name
		return false
	})
	return matches.length === 1 ? matches[0] : undefined
}

function findMatchingThreadedComment(
	sourceRefs: readonly ThreadedCommentTextRef[],
	comment: { readonly id?: string; readonly ref: string },
): ThreadedCommentTextRef | undefined {
	const matches = sourceRefs.filter((sourceRef) => {
		if (comment.id !== undefined) return sourceRef.id === comment.id
		return sourceRef.ref?.toUpperCase() === comment.ref.toUpperCase()
	})
	return matches.length === 1 ? matches[0] : undefined
}

function resolveCapsuleRelationships(
	workbook: Workbook,
	capsule: PreservationCapsule,
): readonly RelEntry[] {
	const updatedExternalLinkRels = buildGeneratedExternalLinkRelationships(
		workbook,
		capsule.partPath,
		capsule.relationships,
	)
	return updatedExternalLinkRels.length > 0 ? updatedExternalLinkRels : capsule.relationships
}

function buildGeneratedExternalLinkRelationships(
	workbook: Workbook,
	partPath: string,
	relationships: readonly RelEntry[] = [],
): readonly RelEntry[] {
	const detail = workbook.externalReferenceDetails.find((entry) => entry.partPath === partPath)
	if (!detail?.target) return []
	const target = detail.target
	if (relationships.length === 0) {
		return [
			{
				id: detail.linkRelId ?? detail.externalBookRelId ?? 'rId1',
				type: detail.linkRelationshipType ?? REL_EXTERNAL_LINK_PATH,
				...(detail.linkRelationshipRawType ? { rawType: detail.linkRelationshipRawType } : {}),
				target,
				targetMode: detail.targetMode ?? 'External',
			},
		]
	}
	let changed = false
	const updated = relationships.map((rel) => {
		const matchesLinkRelId = detail.linkRelId !== undefined && rel.id === detail.linkRelId
		const matchesDefaultLink =
			detail.linkRelId === undefined && isExternalLinkPathRelationshipType(rel.type)
		if (!matchesLinkRelId && !matchesDefaultLink) return rel
		changed = true
		return {
			...rel,
			target,
			...(detail.targetMode !== undefined ? { targetMode: detail.targetMode } : {}),
		}
	})
	if (changed) return updated
	if (
		detail.externalBookRelId !== undefined &&
		!relationships.some((rel) => rel.id === detail.externalBookRelId)
	) {
		return [
			...relationships,
			{
				id: detail.externalBookRelId,
				type: detail.linkRelationshipType ?? REL_EXTERNAL_LINK_PATH,
				...(detail.linkRelationshipRawType ? { rawType: detail.linkRelationshipRawType } : {}),
				target,
				targetMode: detail.targetMode ?? 'External',
			},
		]
	}
	return []
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

function partDir(partPath: string): string {
	const index = partPath.lastIndexOf('/')
	return index >= 0 ? partPath.slice(0, index + 1) : ''
}
