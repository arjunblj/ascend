import type { Workbook } from '@ascend/core'
import type { AscendError, CellValue, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import type { PreservationCapsule } from '../preserve.ts'
import {
	getRelsPath,
	REL_COMMENTS,
	REL_DRAWING,
	REL_SHARED_STRINGS,
	REL_STYLES,
	REL_TABLE,
	REL_THEME,
	REL_VML_DRAWING,
	REL_WORKSHEET,
} from '../reader/relationships.ts'
import { parseSharedStrings } from '../reader/shared-strings.ts'
import { extractZip, type ZipArchive } from '../reader/zip.ts'
import { buildCommentsVml, buildCommentsXml } from './comments.ts'
import { buildContentTypesXml } from './content-types.ts'
import { buildAppPropsXml, buildCorePropsXml } from './doc-props.ts'
import {
	summarizeWritePlan,
	WritePlanBuilder,
	type WritePlanResult,
	type WritePlanSummary,
} from './plan.ts'
import type { RelEntry } from './relationships.ts'
import { buildRelsXml } from './relationships.ts'
import { buildSharedStrings } from './shared-strings.ts'
import { buildSheetXml } from './sheet.ts'
import { buildPreservedStylesXml, buildStylesXml } from './styles.ts'
import { buildTableXml } from './table.ts'
import { buildWorkbookXml } from './workbook.ts'
import { createZip } from './zip.ts'

const REL_OFFICE_DOC =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const REL_CORE_PROPS =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const REL_EXT_PROPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
const REL_HYPERLINK =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const CT_COMMENTS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
const CT_VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing'

export interface WriteXlsxOptions {
	readonly dirtySheetNames?: readonly string[]
	readonly workbookMetaDirty?: boolean
	readonly sharedStringsDirty?: boolean
	readonly stylesDirty?: boolean
	readonly summaryOnly?: boolean
	readonly sourceArchive?: ZipArchive
}

export function writeXlsx(
	workbook: Workbook,
	capsules?: PreservationCapsule[],
	options: WriteXlsxOptions = {},
): Result<Uint8Array, AscendError> {
	try {
		const plan = planWriteXlsx(workbook, capsules, options)
		if (!plan.ok) return plan
		return ok(createZip(new Map(plan.value.parts)))
	} catch (e) {
		return err(
			ascendError(
				'EXPORT_ERROR',
				`Failed to write XLSX: ${e instanceof Error ? e.message : 'unknown'}`,
			),
		)
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
		const sheetCapsuleMap = new Map<string, PreservationCapsule[]>()
		const workbookCapsules: PreservationCapsule[] = []
		let nextGeneratedTableNumber = 1
		let nextGeneratedCommentsNumber = 1
		let nextGeneratedVmlNumber = 1

		const preserveSharedStrings = Boolean(
			workbook.preservedSharedStrings && !options.sharedStringsDirty,
		)
		const preservedSharedStringsXml = preserveSharedStrings
			? resolvePreservedText(
					sourceArchive,
					workbook.preservedSharedStrings?.xml,
					workbook.preservedSharedStrings?.path,
				)
			: undefined
		const preservedSharedStringEntries =
			preserveSharedStrings && !options.summaryOnly && preservedSharedStringsXml
				? materializeSharedStringEntries(preservedSharedStringsXml)
				: []
		const ssTable = options.summaryOnly
			? {
					getIndex(): number | undefined {
						return undefined
					},
					toXml(): string {
						return ''
					},
					count: preserveSharedStrings || workbookHasStringCells(workbook) ? 1 : 0,
				}
			: buildSharedStrings(workbook, preservedSharedStringEntries)
		const hasSharedStrings =
			preserveSharedStrings ||
			(!options.summaryOnly ? ssTable.count > 0 : workbookHasStringCells(workbook))

		const preservedStyles = workbook.preservedStyles ?? undefined
		const preserveStyles = preservedStyles !== undefined && !options.stylesDirty
		const canReusePreservedStyles =
			preservedStyles !== undefined &&
			!options.stylesDirty &&
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
		const generatedStylesResult = options.summaryOnly
			? undefined
			: preservedStyles !== undefined && !canReusePreservedStyles
				? undefined
				: buildStylesXml(workbook.styles, workbook.differentialStyles)
		const xfMap =
			canReusePreservedStyles && preservedStyles
				? new Map(
						Object.entries(preservedStyles.xfByStyleId).map(([styleId, xfIndex]) => [
							Number(styleId),
							xfIndex,
						]),
					)
				: options.summaryOnly
					? new Map<number, number>()
					: (
							stylesResult ??
							generatedStylesResult ??
							buildStylesXml(workbook.styles, workbook.differentialStyles)
						).xfMap
		const stylesXml = options.summaryOnly
			? ''
			: preservedStyles !== undefined
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
				if (capsule.anchor.kind === 'sheet') {
					const sheetName = capsule.anchor.sheetName
					let list = sheetCapsuleMap.get(sheetName)
					if (!list) {
						list = []
						sheetCapsuleMap.set(sheetName, list)
					}
					list.push(capsule)
				} else {
					workbookCapsules.push(capsule)
				}
			}
		}

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
		const preservedThemeXml =
			workbook.preservedTheme && !options.summaryOnly
				? resolvePreservedText(
						sourceArchive,
						workbook.preservedTheme.xml,
						workbook.preservedTheme.path,
					)
				: undefined
		if (workbook.preservedTheme && (hasPreservedTheme || preservedThemeXml)) {
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: REL_THEME,
				target: workbook.preservedTheme.path.replace(/^xl\//, ''),
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

		for (const capsule of workbookCapsules) {
			if (!capsule.relType) continue
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
			!options.workbookMetaDirty &&
			!!preservedWorkbookXml &&
			hasPreservedPart(
				sourceArchive,
				preservedWorkbookXml.workbookXml,
				preservedWorkbookXml.workbookPath,
			)
		const hasPreservedWorkbookRels =
			!options.workbookMetaDirty &&
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
			!options.workbookMetaDirty && preservedWorkbookXml
				? resolvePreservedBytes(sourceArchive, preservedWorkbookXml.workbookPath)
				: undefined
		const preservedWorkbookRelsBytes =
			!options.workbookMetaDirty && preservedWorkbookXml
				? resolvePreservedBytes(sourceArchive, preservedWorkbookXml.workbookRelsPath)
				: undefined
		const preserveWorkbookXml = options.summaryOnly
			? hasPreservedWorkbookXml && hasPreservedWorkbookRels
			: !!(preservedWorkbookXmlText && preservedWorkbookRelsText)
		const workbookContentType =
			preservedWorkbookXml?.contentType ??
			'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
		if (preserveWorkbookXml && preservedWorkbookXmlBytes && !options.summaryOnly) {
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
					origin: preserveWorkbookXml
						? resolvePreservedOrigin(preservedWorkbookXml?.workbookXml)
						: 'generated',
					contentType: workbookContentType,
				},
				() =>
					preserveWorkbookXml
						? (preservedWorkbookXmlText ?? '')
						: buildWorkbookXml(workbook, { externalReferenceRelIds }),
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
		}

		if (hasSharedStrings) {
			const preservedSharedStringBytes = preserveSharedStrings
				? resolvePreservedBytes(sourceArchive, workbook.preservedSharedStrings?.path)
				: undefined
			if (preserveSharedStrings && preservedSharedStringBytes && !options.summaryOnly) {
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
							preservedSharedStringsXml !== undefined
								? resolvePreservedOrigin(workbook.preservedSharedStrings?.xml)
								: 'generated',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
					},
					() => preservedSharedStringsXml ?? ssTable.toXml(),
				)
			}
		}

		for (let i = 0; i < workbook.sheets.length; i++) {
			const sheet = workbook.sheets[i]
			if (!sheet) continue
			const sheetCapsules = sheetCapsuleMap.get(sheet.name) ?? []
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
			} else {
				recordXml(
					`xl/worksheets/sheet${i + 1}.xml`,
					{
						owner: { kind: 'sheet', sheetName: sheet.name },
						origin: preserveSheetXml ? resolvePreservedOrigin(preservedSheetXml?.xml) : 'generated',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
					},
					() =>
						preserveSheetXml
							? (preservedSheetXmlText ?? '')
							: buildSheetXml(sheet, ssTable, xfMap, {
									tableRelIds,
									...(sheet.drawingRefs.hasDrawing && drawingRelId ? { drawingRelId } : {}),
									hyperlinks: hyperlinkEntries,
									...((sheet.drawingRefs.hasLegacyDrawing || sheet.comments.size > 0) &&
									legacyDrawingRelId
										? { legacyDrawingRelId }
										: {}),
								}),
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
				if (plan.isCapsulePathSkipped(capsule.partPath)) continue
				const content = capsule.content ?? sourceArchive?.readBytes(capsule.partPath)
				if (!content) continue
				recordBytes(
					capsule.partPath,
					{
						owner:
							capsule.anchor.kind === 'sheet'
								? { kind: 'sheet', sheetName: capsule.anchor.sheetName }
								: { kind: 'workbook' },
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
							owner:
								capsule.anchor.kind === 'sheet'
									? { kind: 'sheet', sheetName: capsule.anchor.sheetName }
									: { kind: 'workbook' },
							origin: 'capsule',
						},
						() => buildRelsXml(capsule.relationships),
					)
				}
			}
		}

		if (
			preserveWorkbookXml &&
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
						preserveWorkbookXml && preservedWorkbookRelsText
							? resolvePreservedOrigin(preservedWorkbookXml?.workbookRelsXml)
							: 'generated',
				},
				() =>
					preserveWorkbookXml && preservedWorkbookRelsText
						? preservedWorkbookRelsText
						: buildRelsXml(wbRels),
			)
		}

		recordXml(
			'[Content_Types].xml',
			{
				owner: { kind: 'package' },
				origin: 'generated',
			},
			() =>
				buildContentTypesXml(
					workbook.sheets.length,
					hasSharedStrings,
					workbookContentType,
					capsules,
					plan.build().extraOverrides.length > 0 ? plan.build().extraOverrides : undefined,
				),
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

function hasCompletePreservedStyleMap(
	xfByStyleId: Readonly<Record<number, number>>,
	styleCount: number,
): boolean {
	for (let index = 0; index < styleCount; index++) {
		if (xfByStyleId[index] === undefined) return false
	}
	return true
}

function workbookHasStringCells(workbook: Workbook): boolean {
	for (const sheet of workbook.sheets) {
		for (const [, , cell] of sheet.cells.iterate()) {
			if (cell.value.kind === 'string' || cell.value.kind === 'richText') return true
		}
	}
	return false
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
