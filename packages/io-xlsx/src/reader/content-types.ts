export interface ContentTypes {
	readonly defaults: ReadonlyMap<string, string>
	readonly overrides: ReadonlyMap<string, string>
}

const DEFAULT_RE = /<Default\b([^>]*)\/>/g
const OVERRIDE_RE = /<Override\b([^>]*)\/>/g
const ATTR_RE = /([A-Za-z_][\w:.-]*)="([^"]*)"/g

export function parseContentTypes(xml: string): ContentTypes {
	const defaults = new Map<string, string>()
	const overrides = new Map<string, string>()
	collectAttributes(xml, DEFAULT_RE, (attrs) => {
		const ext = attrs.get('Extension')
		const ct = attrs.get('ContentType')
		if (ext && ct) defaults.set(ext, ct)
	})
	collectAttributes(xml, OVERRIDE_RE, (attrs) => {
		const partName = attrs.get('PartName')
		const ct = attrs.get('ContentType')
		if (partName && ct) overrides.set(partName.replace(/^\//, ''), ct)
	})

	return { defaults, overrides }
}

function collectAttributes(
	xml: string,
	pattern: RegExp,
	visit: (attrs: Map<string, string>) => void,
): void {
	for (const match of xml.matchAll(pattern)) {
		const rawAttrs = match[1]
		if (!rawAttrs) continue
		const attrs = new Map<string, string>()
		for (const attrMatch of rawAttrs.matchAll(ATTR_RE)) {
			const key = attrMatch[1]
			const value = attrMatch[2]
			if (!key || value === undefined) continue
			attrs.set(key, value)
		}
		visit(attrs)
	}
}
