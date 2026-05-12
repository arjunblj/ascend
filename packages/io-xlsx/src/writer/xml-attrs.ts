import { escapeXml } from '../xml.ts'

export function readXmlAttr(attrs: string, name: string): string | undefined {
	const match = xmlAttrPattern(name).exec(attrs)
	return match?.[1] ?? match?.[2]
}

export function readNumberXmlAttr(attrs: string, name: string): number | undefined {
	const value = readXmlAttr(attrs, name)
	if (value === undefined) return undefined
	const parsed = Number(value)
	return Number.isFinite(parsed) ? parsed : undefined
}

export function setXmlAttr(
	attrs: string,
	name: string,
	value: string | number | boolean | undefined,
): string {
	if (value === undefined) return attrs
	const serialized = typeof value === 'boolean' ? (value ? '1' : '0') : escapeXml(String(value))
	const attrText = `${name}="${serialized}"`
	const pattern = xmlAttrPattern(name)
	if (pattern.test(attrs)) return attrs.replace(pattern, ` ${attrText}`)
	return `${attrs} ${attrText}`
}

export function removeXmlAttr(attrs: string, name: string): string {
	return attrs.replace(xmlAttrPattern(name, 'g'), '')
}

export function xmlAttrPattern(name: string, flags = ''): RegExp {
	return new RegExp(String.raw`\s${escapeRegExp(name)}\s*=\s*(?:"([^"]*)"|'([^']*)')`, flags)
}

function escapeRegExp(value: string): string {
	return value.replace(/[\\^$.*+?()[\]{}|]/g, '\\$&')
}
