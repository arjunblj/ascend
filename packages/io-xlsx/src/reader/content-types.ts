import { parseXmlAttributes } from './xml-utils.ts'

export interface ContentTypes {
	readonly defaults: ReadonlyMap<string, string>
	readonly overrides: ReadonlyMap<string, string>
	readonly defaultEntries: readonly ContentTypeDefaultEntry[]
	readonly overrideEntries: readonly ContentTypeOverrideEntry[]
}

export interface ContentTypeDefaultEntry {
	readonly extension: string
	readonly contentType: string
	readonly extraAttributes?: readonly ContentTypeEntryAttribute[]
}

export interface ContentTypeOverrideEntry {
	readonly partPath: string
	readonly contentType: string
	readonly extraAttributes?: readonly ContentTypeEntryAttribute[]
}

export interface ContentTypeEntryAttribute {
	readonly name: string
	readonly value: string
}

const DEFAULT_RE = /<Default\b([^>]*)\/>/g
const OVERRIDE_RE = /<Override\b([^>]*)\/>/g

export function parseContentTypes(xml: string): ContentTypes {
	const defaults = new Map<string, string>()
	const overrides = new Map<string, string>()
	const defaultEntries: ContentTypeDefaultEntry[] = []
	const overrideEntries: ContentTypeOverrideEntry[] = []
	collectAttributes(xml, DEFAULT_RE, (attrs) => {
		const ext = attrs.get('Extension')
		const ct = attrs.get('ContentType')
		if (ext && ct) {
			defaults.set(ext, ct)
			defaultEntries.push({
				extension: ext,
				contentType: ct,
				...contentTypeEntryExtraAttributes(attrs),
			})
		}
	})
	collectAttributes(xml, OVERRIDE_RE, (attrs) => {
		const partName = attrs.get('PartName')
		const ct = attrs.get('ContentType')
		if (partName && ct) {
			const partPath = partName.replace(/^\//, '')
			overrides.set(partPath, ct)
			overrideEntries.push({
				partPath,
				contentType: ct,
				...contentTypeEntryExtraAttributes(attrs),
			})
		}
	})

	return { defaults, overrides, defaultEntries, overrideEntries }
}

function contentTypeEntryExtraAttributes(
	attrs: ReadonlyMap<string, string>,
): { readonly extraAttributes: readonly ContentTypeEntryAttribute[] } | Record<string, never> {
	const extraAttributes: ContentTypeEntryAttribute[] = []
	for (const [name, value] of attrs) {
		if (isKnownContentTypeEntryAttribute(name)) continue
		extraAttributes.push({ name, value })
	}
	return extraAttributes.length > 0 ? { extraAttributes } : {}
}

function isKnownContentTypeEntryAttribute(name: string): boolean {
	return name === 'Extension' || name === 'PartName' || name === 'ContentType'
}

function collectAttributes(
	xml: string,
	pattern: RegExp,
	visit: (attrs: Map<string, string>) => void,
): void {
	for (const match of xml.matchAll(pattern)) {
		const rawAttrs = match[1]
		if (!rawAttrs) continue
		visit(parseXmlAttributes(rawAttrs))
	}
}
