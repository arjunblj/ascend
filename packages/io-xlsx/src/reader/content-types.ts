import { asArray, attr, parseXml, type XmlNode } from '../xml.ts'

export interface ContentTypes {
	readonly defaults: ReadonlyMap<string, string>
	readonly overrides: ReadonlyMap<string, string>
}

export function parseContentTypes(xml: string): ContentTypes {
	const doc = parseXml(xml)
	const types = doc.Types as XmlNode | undefined
	const defaults = new Map<string, string>()
	const overrides = new Map<string, string>()

	if (!types) return { defaults, overrides }

	for (const entry of asArray<XmlNode>(types.Default as XmlNode | XmlNode[])) {
		const ext = attr(entry, 'Extension')
		const ct = attr(entry, 'ContentType')
		if (ext && ct) defaults.set(ext, ct)
	}

	for (const entry of asArray<XmlNode>(types.Override as XmlNode | XmlNode[])) {
		const pn = attr(entry, 'PartName')
		const ct = attr(entry, 'ContentType')
		if (pn && ct) overrides.set(pn.replace(/^\//, ''), ct)
	}

	return { defaults, overrides }
}
