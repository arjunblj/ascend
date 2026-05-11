import type {
	WorkbookCoreDocumentProperties,
	WorkbookCustomDocumentProperty,
	WorkbookDocumentPropertyAppValue,
	WorkbookDocumentPropertyScalar,
} from '@ascend/core'
import { escapeXml } from '../xml.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

const CORE_PROPS_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
const DC_NS = 'http://purl.org/dc/elements/1.1/'
const DCTERMS_NS = 'http://purl.org/dc/terms/'
const XSI_NS = 'http://www.w3.org/2001/XMLSchema-instance'
const APP_PROPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
const CUSTOM_PROPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
const VT_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'

const CORE_TEXT_TAGS = [
	['title', 'dc:title'],
	['subject', 'dc:subject'],
	['creator', 'dc:creator'],
	['keywords', 'cp:keywords'],
	['description', 'dc:description'],
	['lastModifiedBy', 'cp:lastModifiedBy'],
	['revision', 'cp:revision'],
	['category', 'cp:category'],
	['contentStatus', 'cp:contentStatus'],
	['language', 'dc:language'],
	['identifier', 'dc:identifier'],
	['version', 'cp:version'],
] as const satisfies readonly (readonly [keyof WorkbookCoreDocumentProperties, string])[]

export function buildCorePropsXml(properties?: WorkbookCoreDocumentProperties): string {
	const now = new Date().toISOString()
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push('<cp:coreProperties')
	out.push(` xmlns:cp="${CORE_PROPS_NS}"`)
	out.push(` xmlns:dc="${DC_NS}"`)
	out.push(` xmlns:dcterms="${DCTERMS_NS}"`)
	out.push(` xmlns:xsi="${XSI_NS}">`)
	for (const [field, tag] of CORE_TEXT_TAGS) {
		const value = properties?.[field]
		if (value !== undefined) out.push(`<${tag}>${escapeXml(value)}</${tag}>`)
	}
	out.push(
		`<dcterms:created xsi:type="dcterms:W3CDTF">${escapeXml(properties?.created ?? now)}</dcterms:created>`,
	)
	out.push(
		`<dcterms:modified xsi:type="dcterms:W3CDTF">${escapeXml(properties?.modified ?? now)}</dcterms:modified>`,
	)
	if (properties?.creator === undefined) out.push('<dc:creator>Ascend</dc:creator>')
	out.push('</cp:coreProperties>')
	return out.toString()
}

export function buildAppPropsXml(
	properties?: Readonly<Record<string, WorkbookDocumentPropertyAppValue>>,
): string {
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(`<Properties xmlns="${APP_PROPS_NS}" xmlns:vt="${VT_NS}">`)
	if (properties && Object.keys(properties).length > 0) {
		for (const [key, value] of Object.entries(properties)) {
			if (Array.isArray(value)) {
				out.push(`<${key}>${buildAppVectorXml(key, value)}</${key}>`)
			} else {
				out.push(`<${key}>${escapeXml(String(value))}</${key}>`)
			}
		}
	} else {
		out.push('<Application>Ascend</Application>')
	}
	out.push('</Properties>')
	return out.toString()
}

function buildAppVectorXml(key: string, values: readonly WorkbookDocumentPropertyScalar[]): string {
	const useVariant = key === 'HeadingPairs' || new Set(values.map((value) => typeof value)).size > 1
	if (useVariant) {
		return `<vt:vector size="${values.length}" baseType="variant">${values
			.map((value) => `<vt:variant>${buildTypedValueXml(value)}</vt:variant>`)
			.join('')}</vt:vector>`
	}
	const baseType = values.length === 0 ? 'lpstr' : appVectorBaseType(values[0])
	return `<vt:vector size="${values.length}" baseType="${baseType}">${values
		.map((value) => buildTypedValueXml(value, baseType))
		.join('')}</vt:vector>`
}

function appVectorBaseType(value: WorkbookDocumentPropertyScalar | undefined): string {
	if (typeof value === 'boolean') return 'bool'
	if (typeof value === 'number') return Number.isInteger(value) ? 'i4' : 'r8'
	return 'lpstr'
}

function buildTypedValueXml(value: WorkbookDocumentPropertyScalar, forcedType?: string): string {
	const type = forcedType ?? appVectorBaseType(value)
	const text = typeof value === 'boolean' ? (value ? 'true' : 'false') : String(value)
	return `<vt:${type}>${escapeXml(text)}</vt:${type}>`
}

export function buildCustomPropsXml(properties: readonly WorkbookCustomDocumentProperty[]): string {
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(`<Properties xmlns="${CUSTOM_PROPS_NS}" xmlns:vt="${VT_NS}">`)
	for (let index = 0; index < properties.length; index++) {
		const property = properties[index]
		if (!property) continue
		const pid = property.pid ?? index + 2
		const fmtid = property.fmtid ?? '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'
		const type = property.type ?? customPropertyType(property.value)
		out.push(
			`<property fmtid="${escapeXml(fmtid)}" pid="${pid}" name="${escapeXml(property.name)}">`,
		)
		out.push(`<vt:${type}>${escapeXml(customPropertyValueText(property.value))}</vt:${type}>`)
		out.push('</property>')
	}
	out.push('</Properties>')
	return out.toString()
}

function customPropertyType(value: string | number | boolean): string {
	if (typeof value === 'boolean') return 'bool'
	if (typeof value === 'number') return Number.isInteger(value) ? 'i4' : 'r8'
	return 'lpwstr'
}

function customPropertyValueText(value: string | number | boolean): string {
	if (typeof value === 'boolean') return value ? 'true' : 'false'
	return String(value)
}
