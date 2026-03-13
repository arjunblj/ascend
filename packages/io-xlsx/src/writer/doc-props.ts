import { ChunkedStringBuilder } from './chunked-string-builder.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

export function buildCorePropsXml(): string {
	const now = new Date().toISOString()
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push('<cp:coreProperties')
	out.push(' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"')
	out.push(' xmlns:dc="http://purl.org/dc/elements/1.1/"')
	out.push(' xmlns:dcterms="http://purl.org/dc/terms/"')
	out.push(' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">')
	out.push(`<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`)
	out.push(`<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`)
	out.push('<dc:creator>Ascend</dc:creator>')
	out.push('</cp:coreProperties>')
	return out.toString()
}

export function buildAppPropsXml(): string {
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(
		'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">',
	)
	out.push('<Application>Ascend</Application>')
	out.push('</Properties>')
	return out.toString()
}
