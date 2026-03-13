import type { PreservationCapsule } from '../preserve.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

const CT_RELS = 'application/vnd.openxmlformats-package.relationships+xml'
const CT_XML = 'application/xml'
const CT_WORKBOOK = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
const CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
const CT_SHARED_STRINGS =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
const CT_CORE_PROPS = 'application/vnd.openxmlformats-package.core-properties+xml'
const CT_APP_PROPS = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'

export function buildContentTypesXml(
	sheetCount: number,
	hasSharedStrings: boolean,
	workbookContentType = CT_WORKBOOK,
	capsules?: PreservationCapsule[],
	extraOverrides?: readonly { partPath: string; contentType: string }[],
): string {
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(`<Types xmlns="${NS}">`)
	out.push(`<Default Extension="rels" ContentType="${CT_RELS}"/>`)
	out.push(`<Default Extension="xml" ContentType="${CT_XML}"/>`)
	out.push(`<Override PartName="/xl/workbook.xml" ContentType="${workbookContentType}"/>`)

	for (let i = 1; i <= sheetCount; i++) {
		out.push(`<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="${CT_WORKSHEET}"/>`)
	}

	if (hasSharedStrings) {
		out.push(`<Override PartName="/xl/sharedStrings.xml" ContentType="${CT_SHARED_STRINGS}"/>`)
	}

	out.push(`<Override PartName="/xl/styles.xml" ContentType="${CT_STYLES}"/>`)
	out.push(`<Override PartName="/docProps/core.xml" ContentType="${CT_CORE_PROPS}"/>`)
	out.push(`<Override PartName="/docProps/app.xml" ContentType="${CT_APP_PROPS}"/>`)

	if (capsules) {
		for (const capsule of capsules) {
			if (capsule.contentType) {
				const pn = capsule.partPath.startsWith('/') ? capsule.partPath : `/${capsule.partPath}`
				out.push(`<Override PartName="${pn}" ContentType="${capsule.contentType}"/>`)
			}
		}
	}

	if (extraOverrides) {
		for (const override of extraOverrides) {
			const pn = override.partPath.startsWith('/') ? override.partPath : `/${override.partPath}`
			out.push(`<Override PartName="${pn}" ContentType="${override.contentType}"/>`)
		}
	}

	out.push('</Types>')
	return out.toString()
}
