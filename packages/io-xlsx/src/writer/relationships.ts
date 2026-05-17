import { ChunkedStringBuilder } from './chunked-string-builder.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const XML_ATTR_RE = /([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

export interface RelEntry {
	readonly id: string
	readonly type: string
	readonly rawType?: string
	readonly target: string
	readonly targetMode?: string
}

export interface RelsXmlOptions {
	readonly preservedRelationshipsXml?: string
}

export function buildRelsXml(entries: readonly RelEntry[], options: RelsXmlOptions = {}): string {
	if (entries.length === 0) return ''
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(buildRelationshipsOpenTag(options.preservedRelationshipsXml))
	for (const e of entries) {
		const attrs = [
			`Id="${escapeXmlAttr(e.id)}"`,
			`Type="${escapeXmlAttr(e.rawType ?? e.type)}"`,
			`Target="${escapeXmlAttr(e.target)}"`,
		]
		if (e.targetMode) attrs.push(`TargetMode="${escapeXmlAttr(e.targetMode)}"`)
		out.push(`<Relationship ${attrs.join(' ')}/>`)
	}
	out.push('</Relationships>')
	return out.toString()
}

function buildRelationshipsOpenTag(sourceXml: string | undefined): string {
	const attrs = new Map<string, string>([['xmlns', NS]])
	for (const [name, value] of extractSourceRelationshipsRootAttrs(sourceXml)) {
		if (!attrs.has(name)) attrs.set(name, value)
	}
	return `<Relationships ${Array.from(attrs, ([name, value]) => `${name}="${escapeXmlAttr(value)}"`).join(' ')}>`
}

function extractSourceRelationshipsRootAttrs(
	sourceXml: string | undefined,
): readonly [string, string][] {
	if (!sourceXml) return []
	const rootMatch = /<Relationships\b([^>]*)>/i.exec(sourceXml)
	if (!rootMatch) return []
	const rawAttrs = rootMatch[1]
	if (!rawAttrs) return []
	const attrs: [string, string][] = []
	XML_ATTR_RE.lastIndex = 0
	for (const match of rawAttrs.matchAll(XML_ATTR_RE)) {
		const name = match[1]
		const value = match[2] ?? match[3]
		if (!name || value === undefined || name === 'xmlns' || !isXmlAttributeName(name)) continue
		attrs.push([name, value])
	}
	return attrs
}

function isXmlAttributeName(name: string): boolean {
	return /^[A-Za-z_][\w:.-]*$/.test(name)
}

function escapeXmlAttr(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/"/g, '&quot;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
}
