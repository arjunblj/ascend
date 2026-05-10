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
import { parseCommentsXml } from '../reader/comments.ts'
import { parseDrawingObjectRefs } from '../reader/drawing.ts'
import {
	getRelsPath,
	parseRelationships,
	REL_CALC_CHAIN,
	REL_CHARTSHEET,
	REL_COMMENTS,
	REL_DRAWING,
	REL_IMAGE,
	REL_MACROSHEET,
	REL_PIVOT_CACHE_DEFINITION,
	REL_PIVOT_TABLE,
	REL_SHARED_STRINGS,
	REL_SHEET_METADATA,
	REL_STYLES,
	REL_TABLE,
	REL_THEME,
	REL_VML_DRAWING,
	REL_WORKSHEET,
	resolvePath,
} from '../reader/relationships.ts'
import { parseSharedStrings } from '../reader/shared-strings.ts'
import { parseTable } from '../reader/table.ts'
import { extractZip, type ZipArchive } from '../reader/zip.ts'
import { updateChartXml } from './chart.ts'
import { buildCommentsVml, buildCommentsXml } from './comments.ts'
import { buildContentTypesXml } from './content-types.ts'
import { buildAppPropsXml, buildCorePropsXml } from './doc-props.ts'
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
import { buildThemeXml } from './theme.ts'
import { buildWorkbookXml } from './workbook.ts'
import { createZip, encode, StreamingZipBuilder } from './zip.ts'

const REL_OFFICE_DOC =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const REL_CORE_PROPS =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const REL_EXT_PROPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
const REL_DIGITAL_SIGNATURE_ORIGIN =
	'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin'
const REL_HYPERLINK =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const REL_EXTERNAL_LINK_PATH =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath'
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
			options.useInlineStrings ?? !(options.useSharedStrings ?? true)
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
		const needsWriteFactScan = useSharedStrings || workbookHasFormulaCells(workbook)
		const workbookWriteFacts = needsWriteFactScan
			? scanWorkbookWriteFactsFast(workbook)
			: { hasStringCells: false, dynamicArrayMetadataEntries: [] }
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
						preserveSharedStrings || hasPreservedSheetXmlInDirtyPatchMode
							? preservedSharedStringEntries
							: [],
						workbookWriteFacts,
					)
		const hasSharedStrings =
			useSharedStrings && (preserveSharedStrings || workbookWriteFacts.hasStringCells)

		const preservedStyles = workbook.preservedStyles ?? undefined
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
				if (invalidateDigitalSignatures && isDigitalSignatureCapsule(capsule)) {
					plan.skipCapsulePath(capsule.partPath)
					continue
				}
				if (isCalcChainCapsule(capsule) && effectiveCalcChainDirty) continue
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
		const wbRels: RelEntry[] = []
		for (let i = 0; i < workbook.sheets.length; i++) {
			const sheet = workbook.sheets[i]
			const partPath = worksheetPartPath(sheet, i)
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: REL_WORKSHEET,
				target: workbookRelTarget(REL_WORKSHEET, partPath, partPath.replace(/^xl\//, '')),
			})
			rIdCounter++
		}

		wbRels.push({
			id: `rId${rIdCounter}`,
			type: REL_STYLES,
			target: workbookRelTarget(REL_STYLES, 'xl/styles.xml', 'styles.xml'),
		})
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
				target: workbookRelTarget(REL_THEME, generatedThemePath, generatedThemeTarget),
			})
			rIdCounter++
		}

		if (hasSharedStrings) {
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: REL_SHARED_STRINGS,
				target: workbookRelTarget(REL_SHARED_STRINGS, 'xl/sharedStrings.xml', 'sharedStrings.xml'),
			})
			rIdCounter++
		}

		if (shouldWriteDynamicArrayMetadata) {
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: REL_SHEET_METADATA,
				target: workbookRelTarget(
					REL_SHEET_METADATA,
					dynamicArrayMetadataPath,
					dynamicArrayMetadataTarget,
				),
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
				target: workbookRelTarget(REL_PIVOT_CACHE_DEFINITION, cache.partPath, target),
			})
			rIdCounter++
		}

		const chartSheetPartPaths = new Set(workbook.chartSheets.map((sheet) => sheet.partPath))
		const chartSheetRelIds: string[] = []
		for (const chartSheet of workbook.chartSheets) {
			const target = chartSheet.partPath.replace(/^xl\//, '')
			const relId = `rId${rIdCounter}`
			chartSheetRelIds.push(relId)
			wbRels.push({
				id: relId,
				type: REL_CHARTSHEET,
				target: workbookRelTarget(REL_CHARTSHEET, chartSheet.partPath, target),
			})
			rIdCounter++
		}

		const macroSheetPartPaths = new Set(workbook.macroSheets.map((sheet) => sheet.partPath))
		const macroSheetRelIds: string[] = []
		for (const macroSheet of workbook.macroSheets) {
			const target = macroSheet.partPath.replace(/^xl\//, '')
			const relId = `rId${rIdCounter}`
			macroSheetRelIds.push(relId)
			wbRels.push({
				id: relId,
				type: REL_MACROSHEET,
				target: workbookRelTarget(REL_MACROSHEET, macroSheet.partPath, target),
			})
			rIdCounter++
		}

		for (const capsule of workbookCapsules) {
			if (!capsule.relType) continue
			if (isPackageDocPropsCapsule(capsule)) continue
			if (isPackageSignatureOriginCapsule(capsule)) continue
			if (chartSheetPartPaths.has(capsule.partPath)) continue
			if (macroSheetPartPaths.has(capsule.partPath)) continue
			if (pivotCachePartPaths.has(capsule.partPath)) continue
			const target = computeRelativePath('xl/', capsule.partPath)
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: capsule.relType,
				target: workbookRelTarget(capsule.relType, capsule.partPath, target),
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
								externalReferenceRelIds,
								pivotCacheRelIds,
								chartSheetRelIds,
								macroSheetRelIds,
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
						recordXml(
							vmlPartPath,
							{
								owner: { kind: 'sheet', sheetName: sheet.name },
								origin: 'generated',
								contentType: CT_VML,
							},
							() => buildCommentsVml(sheet),
						)
						plan.addOverride(commentsPartPath, CT_COMMENTS)
						plan.addOverride(vmlPartPath, CT_VML)
						if (commentsCapsule) plan.skipCapsulePath(commentsCapsule.partPath)
						if (existingVmlCapsule) plan.skipCapsulePath(existingVmlCapsule.partPath)
						if (!commentsRelId) {
							commentsRelId = `rId${sheetRelId}`
							sheetRels.push({
								id: commentsRelId,
								type: REL_COMMENTS,
								target: computeRelativePath('xl/worksheets/', commentsPartPath),
							})
							sheetRelId++
						}
						if (!legacyDrawingRelId) {
							legacyDrawingRelId = `rId${sheetRelId}`
							sheetRels.push({
								id: legacyDrawingRelId,
								type: REL_VML_DRAWING,
								target: computeRelativePath('xl/worksheets/', vmlPartPath),
							})
							sheetRelId++
						}
						if (!commentsCapsule) nextGeneratedCommentsNumber++
						if (!existingVmlCapsule) nextGeneratedVmlNumber++
					}
				}
				for (let tableIndex = 0; tableIndex < sheet.tables.length; tableIndex++) {
					const table = sheet.tables[tableIndex]
					if (!table) continue
					const tableCapsule = tableCapsules[tableIndex]
					const tablePartPath =
						tableCapsule?.partPath ?? `xl/tables/table${nextGeneratedTableNumber}.xml`
					const tableContentType = tableCapsule?.contentType ?? CT_TABLE
					const tableCapsuleContent = tableCapsule
						? (tableCapsule.content ?? sourceArchive?.readBytes(tableCapsule.partPath))
						: undefined
					if (
						tableCapsuleContent &&
						canPreserveTableCapsule(table, tableCapsuleContent, sheet.id)
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
				preserveSharedStrings &&
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

		recordDocPropsPart(
			plan,
			sourceArchive,
			capsules,
			corePropsPath,
			'application/vnd.openxmlformats-package.core-properties+xml',
			() => buildCorePropsXml(),
		)
		recordDocPropsPart(
			plan,
			sourceArchive,
			capsules,
			appPropsPath,
			'application/vnd.openxmlformats-officedocument.extended-properties+xml',
			() => buildAppPropsXml(),
		)

		const rootRels: RelEntry[] = [
			{
				id: 'rId1',
				type: REL_OFFICE_DOC,
				target: rootRelTarget(REL_OFFICE_DOC, 'xl/workbook.xml', 'xl/workbook.xml'),
			},
			{
				id: 'rId2',
				type: REL_CORE_PROPS,
				target: rootRelTarget(REL_CORE_PROPS, corePropsPath, corePropsPath),
			},
			{
				id: 'rId3',
				type: REL_EXT_PROPS,
				target: rootRelTarget(REL_EXT_PROPS, appPropsPath, appPropsPath),
			},
		]
		let rootRIdCounter = 4
		if (capsules) {
			for (const capsule of capsules) {
				if (plan.isCapsulePathSkipped(capsule.partPath)) continue
				if (!isPackageDocPropsCapsule(capsule) && !isPackageSignatureOriginCapsule(capsule)) {
					continue
				}
				if (!capsule.relType) continue
				if (capsule.partPath === corePropsPath || capsule.partPath === appPropsPath) {
					continue
				}
				rootRels.push({
					id: `rId${rootRIdCounter}`,
					type: capsule.relType,
					target: capsule.partPath,
				})
				rootRIdCounter++
			}
		}
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
				if (isCalcChainCapsule(capsule) && !preserveCalcChainCapsules) continue
				if (plan.isCapsulePathSkipped(capsule.partPath)) continue
				const content = capsule.content ?? sourceArchive?.readBytes(capsule.partPath)
				if (!content) continue
				const owner = resolveCapsuleOwner(capsule, sheetNameById)
				if (!owner) continue
				const pivotTable = workbook.pivotTables.find((pivot) => pivot.partPath === capsule.partPath)
				if (pivotTable && isPivotTableDefinitionCapsule(capsule)) {
					recordXml(
						capsule.partPath,
						{
							owner,
							origin: 'generated',
							contentType: capsule.contentType,
						},
						() => updatePivotTableDefinitionXml(new TextDecoder().decode(content), pivotTable),
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
				const slicerCache = workbook.slicerCaches.find(
					(cache) => cache.partPath === capsule.partPath,
				)
				if (slicerCache && isSlicerCacheDefinitionCapsule(capsule)) {
					recordXml(
						capsule.partPath,
						{
							owner,
							origin: 'generated',
							contentType: capsule.contentType,
						},
						() => updateSlicerCacheDefinitionXml(new TextDecoder().decode(content), slicerCache),
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
							(!isCalcChainCapsule(capsule) || preserveCalcChainCapsules) &&
							!built.skippedCapsulePaths.has(capsule.partPath),
					),
					built.extraOverrides.length > 0 ? built.extraOverrides : undefined,
					workbook.preservedXml?.contentTypeDefaults,
					{ corePropsPath, appPropsPath },
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

function canPreserveTableCapsule(table: Table, content: Uint8Array, sheetId: SheetId): boolean {
	const parsed = parseTable(new TextDecoder().decode(content), sheetId)
	return parsed ? tablesHaveSameWritableModel(table, parsed) : false
}

function tablesHaveSameWritableModel(left: Table, right: Table): boolean {
	return (
		left.name === right.name &&
		left.sheetId === right.sheetId &&
		left.hasHeaders === right.hasHeaders &&
		left.hasTotals === right.hasTotals &&
		(left.dxfId ?? null) === (right.dxfId ?? null) &&
		(left.headerRowDxfId ?? null) === (right.headerRowDxfId ?? null) &&
		(left.dataDxfId ?? null) === (right.dataDxfId ?? null) &&
		(left.totalsRowDxfId ?? null) === (right.totalsRowDxfId ?? null) &&
		(left.headerRowBorderDxfId ?? null) === (right.headerRowBorderDxfId ?? null) &&
		rangesEqual(left.ref, right.ref) &&
		tableColumnsEqual(left.columns, right.columns) &&
		stableJson(left.autoFilter) === stableJson(right.autoFilter) &&
		stableJson(left.sortState) === stableJson(right.sortState) &&
		stableJson(left.tableStyleInfo) === stableJson(right.tableStyleInfo)
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
			leftColumn.name !== rightColumn.name ||
			(leftColumn.formula ?? null) !== (rightColumn.formula ?? null) ||
			(leftColumn.totalsRowFunction ?? null) !== (rightColumn.totalsRowFunction ?? null) ||
			(leftColumn.totalsRowFormula ?? null) !== (rightColumn.totalsRowFormula ?? null) ||
			(leftColumn.totalsRowLabel ?? null) !== (rightColumn.totalsRowLabel ?? null) ||
			(leftColumn.dataDxfId ?? null) !== (rightColumn.dataDxfId ?? null) ||
			(leftColumn.headerRowDxfId ?? null) !== (rightColumn.headerRowDxfId ?? null) ||
			(leftColumn.totalsRowDxfId ?? null) !== (rightColumn.totalsRowDxfId ?? null)
		) {
			return false
		}
	}
	return true
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
): void {
	const capsule = capsules?.find((entry) => entry.partPath === path)
	const content = capsule?.content ?? sourceArchive?.readBytes(path)
	if (capsule && content) {
		plan.putBytes(path, content, {
			owner: { kind: 'package' },
			origin: 'capsule',
			contentType: capsule.contentType || contentType,
		})
		plan.skipCapsulePath(path)
		return
	}
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

function hasThemeMetadata(metadata: Workbook['themeMetadata']): boolean {
	return (
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

function workbookHasFormulaCells(workbook: Workbook): boolean {
	return workbook.sheets.some((sheet) => sheet.cells.formulaCellCount() > 0)
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

function isSlicerCacheDefinitionCapsule(capsule: PreservationCapsule): boolean {
	return (
		capsule.contentType.includes('slicerCache+xml') ||
		capsule.partPath.includes('/slicerCaches/slicerCache')
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
				id: detail.linkRelId ?? 'rId1',
				type: REL_EXTERNAL_LINK_PATH,
				target,
				targetMode: detail.targetMode ?? 'External',
			},
		]
	}
	let changed = false
	const updated = relationships.map((rel, index) => {
		const matchesLinkRelId = detail.linkRelId !== undefined && rel.id === detail.linkRelId
		const matchesDefaultLink =
			detail.linkRelId === undefined && (rel.type === REL_EXTERNAL_LINK_PATH || index === 0)
		if (!matchesLinkRelId && !matchesDefaultLink) return rel
		changed = true
		return {
			...rel,
			target,
			...(detail.targetMode !== undefined ? { targetMode: detail.targetMode } : {}),
		}
	})
	return changed ? updated : []
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
