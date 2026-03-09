import type { PreservationCapsule } from '../preserve.ts'

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
	capsules?: PreservationCapsule[],
	extraOverrides?: readonly { partPath: string; contentType: string }[],
): string {
	const parts: string[] = [
		XML_HEADER,
		`<Types xmlns="${NS}">`,
		`<Default Extension="rels" ContentType="${CT_RELS}"/>`,
		`<Default Extension="xml" ContentType="${CT_XML}"/>`,
		`<Override PartName="/xl/workbook.xml" ContentType="${CT_WORKBOOK}"/>`,
	]

	for (let i = 1; i <= sheetCount; i++) {
		parts.push(`<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="${CT_WORKSHEET}"/>`)
	}

	if (hasSharedStrings) {
		parts.push(`<Override PartName="/xl/sharedStrings.xml" ContentType="${CT_SHARED_STRINGS}"/>`)
	}

	parts.push(`<Override PartName="/xl/styles.xml" ContentType="${CT_STYLES}"/>`)
	parts.push(`<Override PartName="/docProps/core.xml" ContentType="${CT_CORE_PROPS}"/>`)
	parts.push(`<Override PartName="/docProps/app.xml" ContentType="${CT_APP_PROPS}"/>`)

	if (capsules) {
		for (const capsule of capsules) {
			if (capsule.contentType) {
				const pn = capsule.partPath.startsWith('/') ? capsule.partPath : `/${capsule.partPath}`
				parts.push(`<Override PartName="${pn}" ContentType="${capsule.contentType}"/>`)
			}
		}
	}

	if (extraOverrides) {
		for (const override of extraOverrides) {
			const pn = override.partPath.startsWith('/') ? override.partPath : `/${override.partPath}`
			parts.push(`<Override PartName="${pn}" ContentType="${override.contentType}"/>`)
		}
	}

	parts.push('</Types>')
	return parts.join('')
}
