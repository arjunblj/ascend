const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_DYNAMIC_ARRAY = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray'
const DYNAMIC_ARRAY_URI = '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}'

export interface DynamicArrayMetadataEntry {
	readonly metadataIndex: number
	readonly collapsed: boolean
}

export function buildDynamicArrayMetadataXml(
	entries: readonly DynamicArrayMetadataEntry[],
): string {
	const parts: string[] = [
		XML_HEADER,
		`<metadata xmlns="${NS_MAIN}" xmlns:xda="${NS_DYNAMIC_ARRAY}">`,
		'<metadataTypes count="1">',
		'<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>',
		'</metadataTypes>',
		`<futureMetadata name="XLDAPR" count="${entries.length}">`,
	]

	for (const entry of entries) {
		parts.push(
			`<bk><extLst><ext uri="${DYNAMIC_ARRAY_URI}"><xda:dynamicArrayProperties fDynamic="1"`,
		)
		if (entry.collapsed) parts.push(' fCollapsed="1"')
		else parts.push(' fCollapsed="0"')
		parts.push('/></ext></extLst></bk>')
	}

	parts.push(`</futureMetadata><cellMetadata count="${entries.length}">`)
	for (const entry of entries) {
		parts.push(`<bk><rc t="1" v="${entry.metadataIndex - 1}"/></bk>`)
	}
	parts.push('</cellMetadata></metadata>')
	return parts.join('')
}
