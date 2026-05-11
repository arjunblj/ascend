import type {
	WorkbookCoreDocumentProperties,
	WorkbookCustomDocumentProperty,
	WorkbookDocumentProperties,
	WorkbookDocumentPropertyAppValue,
} from '@ascend/core'
import { asArray, attr, numAttr, parseXml, type XmlNode } from '../xml.ts'

export interface DocumentPropertiesXmlParts {
	readonly coreXml?: string | undefined
	readonly appXml?: string | undefined
	readonly customXml?: string | undefined
}

const CORE_PROPERTY_FIELDS = [
	['title', 'title'],
	['subject', 'subject'],
	['creator', 'creator'],
	['keywords', 'keywords'],
	['description', 'description'],
	['lastModifiedBy', 'lastModifiedBy'],
	['revision', 'revision'],
	['created', 'created'],
	['modified', 'modified'],
	['category', 'category'],
	['contentStatus', 'contentStatus'],
	['language', 'language'],
	['identifier', 'identifier'],
	['version', 'version'],
] as const satisfies readonly (readonly [keyof WorkbookCoreDocumentProperties, string])[]

export function parseDocumentProperties(
	parts: DocumentPropertiesXmlParts,
): WorkbookDocumentProperties {
	const core = parts.coreXml ? parseCoreDocumentProperties(parts.coreXml) : undefined
	const app = parts.appXml ? parseAppDocumentProperties(parts.appXml) : undefined
	const custom = parts.customXml ? parseCustomDocumentProperties(parts.customXml) : undefined
	return {
		...(core && Object.keys(core).length > 0 ? { core } : {}),
		...(app && Object.keys(app).length > 0 ? { app } : {}),
		...(custom && custom.length > 0 ? { custom } : {}),
	}
}

function parseCoreDocumentProperties(xml: string): WorkbookCoreDocumentProperties {
	const root = findRoot(parseXml(xml))
	if (!root) return {}
	const properties: Writable<WorkbookCoreDocumentProperties> = {}
	for (const [field, localName] of CORE_PROPERTY_FIELDS) {
		const value = textChild(root, localName)
		if (value !== undefined) properties[field] = value
	}
	return properties
}

function parseAppDocumentProperties(
	xml: string,
): Readonly<Record<string, WorkbookDocumentPropertyAppValue>> {
	const root = findRoot(parseXml(xml))
	if (!root) return {}
	const properties: Record<string, WorkbookDocumentPropertyAppValue> = {}
	for (const [key, value] of Object.entries(root)) {
		if (key.startsWith('@_')) continue
		const localName = localPart(key)
		const primitive = primitiveValue(value)
		if (primitive !== undefined) {
			properties[localName] = primitive
			continue
		}
		const values = primitiveDescendants(value)
		if (values.length > 0) properties[localName] = values
	}
	return properties
}

function parseCustomDocumentProperties(xml: string): WorkbookCustomDocumentProperty[] {
	const root = findRoot(parseXml(xml))
	if (!root) return []
	const properties: WorkbookCustomDocumentProperty[] = []
	for (const property of asArray(root.property as XmlNode | XmlNode[] | undefined)) {
		if (!isXmlNode(property)) continue
		const name = attr(property, 'name')
		if (!name) continue
		const valueEntry = firstElementEntry(property)
		const value = valueEntry
			? coerceCustomValue(localPart(valueEntry.key), valueEntry.value)
			: undefined
		if (value === undefined) continue
		const pid = numAttr(property, 'pid')
		const fmtid = attr(property, 'fmtid')
		properties.push({
			name,
			value,
			...(valueEntry ? { type: localPart(valueEntry.key) } : {}),
			...(pid !== undefined ? { pid } : {}),
			...(fmtid ? { fmtid } : {}),
		})
	}
	return properties
}

function findRoot(parsed: XmlNode): XmlNode | undefined {
	const rootEntry = Object.entries(parsed).find(
		([key, value]) => !key.startsWith('@_') && !key.startsWith('?') && isXmlNode(value),
	)
	return rootEntry ? (rootEntry[1] as XmlNode) : undefined
}

function textChild(node: XmlNode, localName: string): string | undefined {
	const value = childValue(node, localName)
	if (value === undefined) return undefined
	const primitive = primitiveValue(value)
	return primitive === undefined ? undefined : String(primitive)
}

function childValue(node: XmlNode, localName: string): unknown {
	for (const [key, value] of Object.entries(node)) {
		if (key.startsWith('@_')) continue
		if (localPart(key) === localName) return value
	}
	return undefined
}

function primitiveValue(value: unknown): string | number | boolean | undefined {
	if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
		return value
	}
	if (!isXmlNode(value)) return undefined
	const text = value['#text']
	return typeof text === 'string' || typeof text === 'number' || typeof text === 'boolean'
		? text
		: undefined
}

function primitiveDescendants(value: unknown): Array<string | number | boolean> {
	const primitive = primitiveValue(value)
	if (primitive !== undefined) return [primitive]
	if (Array.isArray(value)) return value.flatMap((entry) => primitiveDescendants(entry))
	if (!isXmlNode(value)) return []
	return Object.entries(value).flatMap(([key, child]) =>
		key.startsWith('@_') ? [] : primitiveDescendants(child),
	)
}

function firstElementEntry(
	node: XmlNode,
): { readonly key: string; readonly value: unknown } | undefined {
	for (const [key, value] of Object.entries(node)) {
		if (!key.startsWith('@_')) return { key, value }
	}
	return undefined
}

function coerceCustomValue(type: string, value: unknown): string | number | boolean | undefined {
	const primitive = primitiveValue(value)
	if (primitive === undefined) return undefined
	switch (type) {
		case 'bool':
			if (typeof primitive === 'boolean') return primitive
			return primitive === 1 || primitive === '1' || String(primitive).toLowerCase() === 'true'
		case 'i1':
		case 'i2':
		case 'i4':
		case 'i8':
		case 'int':
		case 'uint':
		case 'ui1':
		case 'ui2':
		case 'ui4':
		case 'ui8':
		case 'r4':
		case 'r8': {
			const number = Number(primitive)
			return Number.isNaN(number) ? undefined : number
		}
		default:
			return typeof primitive === 'string' ? primitive : String(primitive)
	}
}

function localPart(name: string): string {
	return name.includes(':') ? (name.split(':').pop() ?? name) : name
}

function isXmlNode(value: unknown): value is XmlNode {
	return value !== null && typeof value === 'object' && !Array.isArray(value)
}

type Writable<T> = { -readonly [K in keyof T]?: T[K] }
