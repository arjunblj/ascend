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
		const plan = new WritePlanBuilder()
		const sourceArchive = workbook.sourceArchiveBytes
			? extractZip(workbook.sourceArchiveBytes)
			: undefined
		const sheetCapsuleMap = new Map<string, PreservationCapsule[]>()
		const workbookCapsules: PreservationCapsule[] = []
		let nextGeneratedTableNumber = 1
		let nextGeneratedCommentsNumber = 1
		let nextGeneratedVmlNumber = 1

		const preservedSharedStringsXml =
			workbook.preservedSharedStrings && !options.sharedStringsDirty
				? resolvePreservedText(
						sourceArchive,
						workbook.preservedSharedStrings.xml,
						workbook.preservedSharedStrings.path,
					)
				: undefined
		const preservedSharedStringEntries = preservedSharedStringsXml
			? materializeSharedStringEntries(preservedSharedStringsXml)
			: []
		const ssTable = buildSharedStrings(workbook, preservedSharedStringEntries)
		const hasSharedStrings = ssTable.count > 0 || preservedSharedStringsXml !== undefined

		const preservedStylesXml = workbook.preservedStyles
			? resolvePreservedText(
					sourceArchive,
					workbook.preservedStyles.xml,
					workbook.preservedStyles.path,
				)
			: undefined
		const stylesResult =
			workbook.preservedStyles && preservedStylesXml
				? buildPreservedStylesXml(preservedStylesXml, workbook.preservedStyles, workbook.styles)
				: undefined
		const { xml: stylesXml, xfMap } =
			stylesResult ?? buildStylesXml(workbook.styles, workbook.differentialStyles)
		if (workbook.preservedStyles) {
			workbook.preservedStyles = {
				...workbook.preservedStyles,
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

		const preservedThemeXml = workbook.preservedTheme
			? resolvePreservedText(
					sourceArchive,
					workbook.preservedTheme.xml,
					workbook.preservedTheme.path,
				)
			: undefined
		if (workbook.preservedTheme && preservedThemeXml) {
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
		const preservedWorkbookXmlText =
			!options.workbookMetaDirty && preservedWorkbookXml
				? resolvePreservedText(
						sourceArchive,
						preservedWorkbookXml.workbookXml,
						preservedWorkbookXml.workbookPath,
					)
				: undefined
		const preservedWorkbookRelsText =
			!options.workbookMetaDirty && preservedWorkbookXml
				? resolvePreservedText(
						sourceArchive,
						preservedWorkbookXml.workbookRelsXml,
						preservedWorkbookXml.workbookRelsPath,
					)
				: undefined
		const preserveWorkbookXml = preservedWorkbookXmlText && preservedWorkbookRelsText
		const workbookContentType =
			preservedWorkbookXml?.contentType ??
			'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
		plan.putXml(
			'xl/workbook.xml',
			preserveWorkbookXml
				? preservedWorkbookXmlText
				: buildWorkbookXml(workbook, { externalReferenceRelIds }),
			{
				owner: { kind: 'workbook' },
				origin: preserveWorkbookXml
					? resolvePreservedOrigin(preservedWorkbookXml?.workbookXml)
					: 'generated',
				contentType: workbookContentType,
			},
		)
		plan.putXml('xl/styles.xml', stylesXml, {
			owner: { kind: 'workbook' },
			origin:
				workbook.preservedStyles && preservedStylesXml
					? resolvePreservedOrigin(workbook.preservedStyles.xml)
					: 'generated',
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
		})
		if (workbook.preservedTheme && preservedThemeXml) {
			plan.putXml(workbook.preservedTheme.path, preservedThemeXml, {
				owner: { kind: 'workbook' },
				origin: resolvePreservedOrigin(workbook.preservedTheme.xml),
				contentType: workbook.preservedTheme.contentType,
			})
			plan.addOverride(workbook.preservedTheme.path, workbook.preservedTheme.contentType)
		}

		if (hasSharedStrings) {
			plan.putXml('xl/sharedStrings.xml', preservedSharedStringsXml ?? ssTable.toXml(), {
				owner: { kind: 'workbook' },
				origin:
					preservedSharedStringsXml !== undefined
						? resolvePreservedOrigin(workbook.preservedSharedStrings?.xml)
						: 'generated',
				contentType:
					'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
			})
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
			const preservedSheetXmlText = resolvePreservedText(
				sourceArchive,
				preservedSheetXml?.xml,
				preservedSheetXml?.partPath,
			)
			const preservedSheetRelsText = resolvePreservedText(
				sourceArchive,
				preservedSheetXml?.relsXml,
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
				!(options.dirtySheetNames ?? []).includes(sheet.name) && preservedSheetXmlText
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
					plan.putXml(commentsPartPath, buildCommentsXml(sheet), {
						owner: { kind: 'sheet', sheetName: sheet.name },
						origin: 'generated',
						contentType: CT_COMMENTS,
					})
					plan.putXml(resolvedVmlPartPath, buildCommentsVml(sheet), {
						owner: { kind: 'sheet', sheetName: sheet.name },
						origin: 'generated',
						contentType: CT_VML,
					})
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
					plan.putXml(tablePartPath, buildTableXml(table, nextGeneratedTableNumber), {
						owner: { kind: 'sheet', sheetName: sheet.name },
						origin: 'generated',
						contentType: tableContentType,
					})
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
			plan.putXml(
				`xl/worksheets/sheet${i + 1}.xml`,
				preserveSheetXml
					? preservedSheetXmlText
					: buildSheetXml(sheet, ssTable, xfMap, {
							tableRelIds,
							...(sheet.drawingRefs.hasDrawing && drawingRelId ? { drawingRelId } : {}),
							hyperlinks: hyperlinkEntries,
							...((sheet.drawingRefs.hasLegacyDrawing || sheet.comments.size > 0) &&
							legacyDrawingRelId
								? { legacyDrawingRelId }
								: {}),
						}),
				{
					owner: { kind: 'sheet', sheetName: sheet.name },
					origin: preserveSheetXml ? resolvePreservedOrigin(preservedSheetXml?.xml) : 'generated',
					contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
				},
			)
			if (preserveSheetXml && preservedSheetRelsText) {
				plan.putXml(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, preservedSheetRelsText, {
					owner: { kind: 'sheet', sheetName: sheet.name },
					origin: resolvePreservedOrigin(preservedSheetXml?.relsXml),
				})
			} else if (sheetRels.length > 0) {
				plan.putXml(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, buildRelsXml(sheetRels), {
					owner: { kind: 'sheet', sheetName: sheet.name },
					origin: 'generated',
				})
			}
		}

		plan.putXml('docProps/core.xml', buildCorePropsXml(), {
			owner: { kind: 'package' },
			origin: 'generated',
			contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
		})
		plan.putXml('docProps/app.xml', buildAppPropsXml(), {
			owner: { kind: 'package' },
			origin: 'generated',
			contentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
		})

		const rootRels: RelEntry[] = [
			{ id: 'rId1', type: REL_OFFICE_DOC, target: 'xl/workbook.xml' },
			{ id: 'rId2', type: REL_CORE_PROPS, target: 'docProps/core.xml' },
			{ id: 'rId3', type: REL_EXT_PROPS, target: 'docProps/app.xml' },
		]
		plan.putXml('_rels/.rels', buildRelsXml(rootRels), {
			owner: { kind: 'package' },
			origin: 'generated',
		})

		if (capsules) {
			for (const capsule of capsules) {
				if (plan.isCapsulePathSkipped(capsule.partPath)) continue
				const content = capsule.content ?? sourceArchive?.readBytes(capsule.partPath)
				if (!content) continue
				plan.putBytes(capsule.partPath, content, {
					owner:
						capsule.anchor.kind === 'sheet'
							? { kind: 'sheet', sheetName: capsule.anchor.sheetName }
							: { kind: 'workbook' },
					origin: 'capsule',
					contentType: capsule.contentType,
				})

				if (capsule.relationships.length > 0) {
					const capsuleRelsPath = getRelsPath(capsule.partPath)
					plan.putXml(capsuleRelsPath, buildRelsXml(capsule.relationships), {
						owner:
							capsule.anchor.kind === 'sheet'
								? { kind: 'sheet', sheetName: capsule.anchor.sheetName }
								: { kind: 'workbook' },
						origin: 'capsule',
					})
				}
			}
		}

		plan.putXml(
			'xl/_rels/workbook.xml.rels',
			preserveWorkbookXml && preservedWorkbookRelsText
				? preservedWorkbookRelsText
				: buildRelsXml(wbRels),
			{
				owner: { kind: 'workbook' },
				origin:
					preserveWorkbookXml && preservedWorkbookRelsText
						? resolvePreservedOrigin(preservedWorkbookXml?.workbookRelsXml)
						: 'generated',
			},
		)

		plan.putXml(
			'[Content_Types].xml',
			buildContentTypesXml(
				workbook.sheets.length,
				hasSharedStrings,
				workbookContentType,
				capsules,
				plan.build().extraOverrides.length > 0 ? plan.build().extraOverrides : undefined,
			),
			{
				owner: { kind: 'package' },
				origin: 'generated',
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
	const plan = planWriteXlsx(workbook, capsules, options)
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
