const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

export interface RelEntry {
	readonly id: string
	readonly type: string
	readonly target: string
	readonly targetMode?: string
}

export function buildRelsXml(entries: readonly RelEntry[]): string {
	if (entries.length === 0) return ''
	const parts: string[] = [XML_HEADER, `<Relationships xmlns="${NS}">`]
	for (const e of entries) {
		const attrs = [`Id="${e.id}"`, `Type="${e.type}"`, `Target="${e.target}"`]
		if (e.targetMode) attrs.push(`TargetMode="${e.targetMode}"`)
		parts.push(`<Relationship ${attrs.join(' ')}/>`)
	}
	parts.push('</Relationships>')
	return parts.join('')
}
