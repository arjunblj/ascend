import type { Workbook } from '@ascend/core'
import type { AscendError, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import type { PreservationCapsule } from '../preserve.ts'
import {
	getRelsPath,
	REL_SHARED_STRINGS,
	REL_STYLES,
	REL_WORKSHEET,
} from '../reader/relationships.ts'
import { buildContentTypesXml } from './content-types.ts'
import { buildAppPropsXml, buildCorePropsXml } from './doc-props.ts'
import type { RelEntry } from './relationships.ts'
import { buildRelsXml } from './relationships.ts'
import { buildSharedStrings } from './shared-strings.ts'
import { buildSheetXml } from './sheet.ts'
import { buildStylesXml } from './styles.ts'
import { buildWorkbookXml } from './workbook.ts'
import { createZip, encode } from './zip.ts'

const REL_OFFICE_DOC =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const REL_CORE_PROPS =
	'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const REL_EXT_PROPS =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'

export function writeXlsx(
	workbook: Workbook,
	capsules?: PreservationCapsule[],
): Result<Uint8Array, AscendError> {
	try {
		const parts = new Map<string, Uint8Array>()

		const ssTable = buildSharedStrings(workbook)
		const hasSharedStrings = ssTable.count > 0

		const { xml: stylesXml, xfMap } = buildStylesXml(workbook.styles)

		parts.set('xl/workbook.xml', encode(buildWorkbookXml(workbook)))
		parts.set('xl/styles.xml', encode(stylesXml))

		if (hasSharedStrings) {
			parts.set('xl/sharedStrings.xml', encode(ssTable.toXml()))
		}

		for (let i = 0; i < workbook.sheets.length; i++) {
			const sheet = workbook.sheets[i]
			if (!sheet) continue
			parts.set(`xl/worksheets/sheet${i + 1}.xml`, encode(buildSheetXml(sheet, ssTable, xfMap)))
		}

		parts.set('docProps/core.xml', encode(buildCorePropsXml()))
		parts.set('docProps/app.xml', encode(buildAppPropsXml()))

		const rootRels: RelEntry[] = [
			{ id: 'rId1', type: REL_OFFICE_DOC, target: 'xl/workbook.xml' },
			{ id: 'rId2', type: REL_CORE_PROPS, target: 'docProps/core.xml' },
			{ id: 'rId3', type: REL_EXT_PROPS, target: 'docProps/app.xml' },
		]
		parts.set('_rels/.rels', encode(buildRelsXml(rootRels)))

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

		if (hasSharedStrings) {
			wbRels.push({
				id: `rId${rIdCounter}`,
				type: REL_SHARED_STRINGS,
				target: 'sharedStrings.xml',
			})
			rIdCounter++
		}

		if (capsules) {
			const sheetCapsuleMap = new Map<string, PreservationCapsule[]>()

			for (const capsule of capsules) {
				parts.set(capsule.partPath, capsule.content)

				if (capsule.relationships.length > 0) {
					const capsuleRelsPath = getRelsPath(capsule.partPath)
					parts.set(capsuleRelsPath, encode(buildRelsXml(capsule.relationships)))
				}

				if (capsule.anchor.kind === 'workbook' && capsule.relType) {
					const target = capsule.partPath.replace(/^xl\//, '')
					wbRels.push({
						id: `rId${rIdCounter}`,
						type: capsule.relType,
						target,
					})
					rIdCounter++
				} else if (capsule.anchor.kind === 'sheet') {
					const sheetName = capsule.anchor.sheetName
					let list = sheetCapsuleMap.get(sheetName)
					if (!list) {
						list = []
						sheetCapsuleMap.set(sheetName, list)
					}
					list.push(capsule)
				}
			}

			for (const [sheetName, sheetCapsules] of sheetCapsuleMap) {
				const sheetIndex = workbook.sheets.findIndex((s) => s.name === sheetName)
				if (sheetIndex === -1) continue

				const sheetRels: RelEntry[] = []
				let sheetRelId = 1

				for (const capsule of sheetCapsules) {
					if (capsule.relType) {
						sheetRels.push({
							id: `rId${sheetRelId}`,
							type: capsule.relType,
							target: computeRelativePath('xl/worksheets/', capsule.partPath),
						})
						sheetRelId++
					}
				}

				if (sheetRels.length > 0) {
					parts.set(
						`xl/worksheets/_rels/sheet${sheetIndex + 1}.xml.rels`,
						encode(buildRelsXml(sheetRels)),
					)
				}
			}
		}

		parts.set('xl/_rels/workbook.xml.rels', encode(buildRelsXml(wbRels)))

		parts.set(
			'[Content_Types].xml',
			encode(buildContentTypesXml(workbook.sheets.length, hasSharedStrings, capsules)),
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
