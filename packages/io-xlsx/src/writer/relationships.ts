import { ChunkedStringBuilder } from './chunked-string-builder.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

export interface RelEntry {
	readonly id: string
	readonly type: string
	readonly rawType?: string
	readonly target: string
	readonly targetMode?: string
}

export function buildRelsXml(entries: readonly RelEntry[]): string {
	if (entries.length === 0) return ''
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(`<Relationships xmlns="${NS}">`)
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

function escapeXmlAttr(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/"/g, '&quot;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
}
