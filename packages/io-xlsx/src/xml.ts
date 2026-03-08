import { XMLParser } from 'fast-xml-parser'

export type XmlNode = Record<string, unknown>

const parser = new XMLParser({
	attributeNamePrefix: '@_',
	ignoreAttributes: false,
	parseTagValue: true,
	trimValues: true,
	processEntities: true,
})

export function parseXml(content: string): XmlNode {
	return parser.parse(content) as XmlNode
}

export function asArray<T>(val: T | T[] | undefined | null): T[] {
	if (val === undefined || val === null) return []
	return Array.isArray(val) ? val : [val]
}

export function attr(node: XmlNode, name: string): string | undefined {
	const val = node[`@_${name}`]
	if (val === undefined || val === null) return undefined
	return String(val)
}

export function numAttr(node: XmlNode, name: string): number | undefined {
	const val = node[`@_${name}`]
	if (val === undefined || val === null) return undefined
	const n = Number(val)
	return Number.isNaN(n) ? undefined : n
}

export function boolAttr(node: XmlNode, name: string): boolean | undefined {
	const val = attr(node, name)
	if (val === undefined) return undefined
	return val === '1' || val === 'true'
}
