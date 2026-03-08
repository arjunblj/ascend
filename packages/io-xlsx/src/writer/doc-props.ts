const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

export function buildCorePropsXml(): string {
	const now = new Date().toISOString()
	return [
		XML_HEADER,
		'<cp:coreProperties',
		' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"',
		' xmlns:dc="http://purl.org/dc/elements/1.1/"',
		' xmlns:dcterms="http://purl.org/dc/terms/"',
		' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
		`<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`,
		`<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`,
		'<dc:creator>Ascend</dc:creator>',
		'</cp:coreProperties>',
	].join('')
}

export function buildAppPropsXml(): string {
	return [
		XML_HEADER,
		'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">',
		'<Application>Ascend</Application>',
		'</Properties>',
	].join('')
}
