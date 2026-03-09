import type { Workbook } from '@ascend/core'
import type { AscendError, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import type { PreservationCapsule } from '../preserve.ts'
import {
	getRelsPath,
	REL_SHARED_STRINGS,
	REL_STYLES,
	REL_TABLE,
	REL_THEME,
	REL_WORKSHEET,
} from '../reader/relationships.ts'
import { buildContentTypesXml } from './content-types.ts'
import { buildAppPropsXml, buildCorePropsXml } from './doc-props.ts'
import type { RelEntry } from './relationships.ts'
import { buildRelsXml } from './relationships.ts'
import { buildSharedStrings } from './shared-strings.ts'
import { buildSheetXml } from './sheet.ts'
import { buildPreservedStylesXml, buildStylesXml } from './styles.ts'
import { buildTableXml } from './table.ts'
import { buildWorkbookXml } from './workbook.ts'
import { createZip, encode } from './zip.ts'

const REL_OFFICE_DOC =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const REL_CORE_PROPS =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const REL_EXT_PROPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
const REL_HYPERLINK =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'

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
		const parts = new Map<string, Uint8Array>()
		const sheetCapsuleMap = new Map<string, PreservationCapsule[]>()
		const workbookCapsules: PreservationCapsule[] = []
		const skippedCapsulePaths = new Set<string>()
		const extraOverrides: Array<{ partPath: string; contentType: string }> = []
		let nextGeneratedTableNumber = 1

		const ssTable = buildSharedStrings(workbook)
		const hasSharedStrings = ssTable.count > 0

		const stylesResult =
			workbook.preservedStyles && !stylesNeedRebuild(workbook)
				? buildPreservedStylesXml(workbook.preservedStyles, workbook.styles)
				: undefined
		const { xml: stylesXml, xfMap } = stylesResult ?? buildStylesXml(workbook.styles)

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

		if (workbook.preservedTheme) {
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
		const preserveWorkbookXml =
			!options.workbookMetaDirty &&
			preservedWorkbookXml?.workbookXml &&
			preservedWorkbookXml.workbookRelsXml
		parts.set(
			'xl/workbook.xml',
			encode(
				preserveWorkbookXml
					? preservedWorkbookXml.workbookXml
					: buildWorkbookXml(workbook, { externalReferenceRelIds }),
			),
		)
		parts.set('xl/styles.xml', encode(stylesXml))
		if (workbook.preservedTheme) {
			parts.set(workbook.preservedTheme.path, encode(workbook.preservedTheme.xml))
			extraOverrides.push({
				partPath: workbook.preservedTheme.path,
				contentType: workbook.preservedTheme.contentType,
			})
		}

		if (hasSharedStrings) {
			parts.set('xl/sharedStrings.xml', encode(ssTable.toXml()))
		}

		for (let i = 0; i < workbook.sheets.length; i++) {
			const sheet = workbook.sheets[i]
			if (!sheet) continue
			const sheetCapsules = sheetCapsuleMap.get(sheet.name) ?? []
			const preservedSheetXml = sheet.preservedXml
			const sheetRels: RelEntry[] = []
			let sheetRelId = 1
			const tableRelIds: string[] = []
			const tableCapsules = sheetCapsules.filter((capsule) => capsule.relType === REL_TABLE)
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
				!options.sharedStringsDirty &&
				!(options.dirtySheetNames ?? []).includes(sheet.name) &&
				preservedSheetXml?.xml
			if (!preserveSheetXml) {
				for (let tableIndex = 0; tableIndex < sheet.tables.length; tableIndex++) {
					const table = sheet.tables[tableIndex]
					if (!table) continue
					const tableCapsule = tableCapsules[tableIndex]
					const tablePartPath =
						tableCapsule?.partPath ?? `xl/tables/table${nextGeneratedTableNumber}.xml`
					const tableContentType = tableCapsule?.contentType ?? CT_TABLE
					parts.set(tablePartPath, encode(buildTableXml(table, nextGeneratedTableNumber)))
					extraOverrides.push({ partPath: tablePartPath, contentType: tableContentType })
					if (tableCapsule) skippedCapsulePaths.add(tableCapsule.partPath)
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
			parts.set(
				`xl/worksheets/sheet${i + 1}.xml`,
				encode(
					preserveSheetXml
						? preservedSheetXml.xml
						: buildSheetXml(sheet, ssTable, xfMap, { tableRelIds, hyperlinks: hyperlinkEntries }),
				),
			)
			if (preserveSheetXml && preservedSheetXml?.relsXml) {
				parts.set(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, encode(preservedSheetXml.relsXml))
			} else if (sheetRels.length > 0) {
				parts.set(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, encode(buildRelsXml(sheetRels)))
			}
		}

		parts.set('docProps/core.xml', encode(buildCorePropsXml()))
		parts.set('docProps/app.xml', encode(buildAppPropsXml()))

		const rootRels: RelEntry[] = [
			{ id: 'rId1', type: REL_OFFICE_DOC, target: 'xl/workbook.xml' },
			{ id: 'rId2', type: REL_CORE_PROPS, target: 'docProps/core.xml' },
			{ id: 'rId3', type: REL_EXT_PROPS, target: 'docProps/app.xml' },
		]
		parts.set('_rels/.rels', encode(buildRelsXml(rootRels)))

		if (capsules) {
			for (const capsule of capsules) {
				if (skippedCapsulePaths.has(capsule.partPath)) continue
				parts.set(capsule.partPath, capsule.content)

				if (capsule.relationships.length > 0) {
					const capsuleRelsPath = getRelsPath(capsule.partPath)
					parts.set(capsuleRelsPath, encode(buildRelsXml(capsule.relationships)))
				}
			}
		}

		parts.set(
			'xl/_rels/workbook.xml.rels',
			encode(
				preserveWorkbookXml && workbook.preservedXml?.workbookRelsXml
					? workbook.preservedXml.workbookRelsXml
					: buildRelsXml(wbRels),
			),
		)

		parts.set(
			'[Content_Types].xml',
			encode(
				buildContentTypesXml(
					workbook.sheets.length,
					hasSharedStrings,
					capsules,
					extraOverrides.length > 0 ? extraOverrides : undefined,
				),
			),
		)

		return ok(createZip(parts))
	} catch (e) {
		return err(
			ascendError(
				'EXPORT_ERROR',
				`Failed to write XLSX: ${e instanceof Error ? e.message : 'unknown'}`,
			),
		)
	}
}

function stylesNeedRebuild(workbook: Workbook): boolean {
	if (!workbook.preservedStyles) return true
	for (let i = 0; i < workbook.styles.size; i++) {
		if (workbook.preservedStyles.xfByStyleId[i] === undefined) return true
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
