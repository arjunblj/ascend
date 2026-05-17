import type { PreservationCapsule } from '../preserve.ts'
import { escapeXml } from '../xml.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

const CT_RELS = 'application/vnd.openxmlformats-package.relationships+xml'
const CT_XML = 'application/xml'
const CT_WORKBOOK = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
const CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
const CT_SHARED_STRINGS =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
const CT_CORE_PROPS = 'application/vnd.openxmlformats-package.core-properties+xml'
const CT_APP_PROPS = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
const XML_ATTR_RE = /([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

export interface ContentTypeDefault {
	readonly extension: string
	readonly contentType: string
	readonly extraAttributes?: readonly ContentTypeEntryAttribute[]
}

export interface ContentTypeOverride {
	readonly partPath: string
	readonly contentType: string
	readonly extraAttributes?: readonly ContentTypeEntryAttribute[]
}

export interface ContentTypeEntryAttribute {
	readonly name: string
	readonly value: string
}

export function buildContentTypesXml(
	sheetPartPaths: readonly string[],
	hasSharedStrings: boolean,
	workbookContentType = CT_WORKBOOK,
	capsules?: PreservationCapsule[],
	extraOverrides?: readonly { partPath: string; contentType: string }[],
	preservedDefaults?: readonly ContentTypeDefault[],
	preservedOverrides?: readonly ContentTypeOverride[],
	docPropsPaths: { readonly corePropsPath: string; readonly appPropsPath: string } = {
		corePropsPath: 'docProps/core.xml',
		appPropsPath: 'docProps/app.xml',
	},
	options: {
		readonly includeStyles?: boolean
		readonly includeDocProps?: boolean
		readonly preservedContentTypesXml?: string
	} = {},
): string {
	const out = new ChunkedStringBuilder()
	const defaults = new Map<string, string>()
	const overrides = new Set<string>()
	const preservedOverrideByPartName = new Map(
		(preservedOverrides ?? []).map((entry) => [normalizePartName(entry.partPath), entry]),
	)
	const pushDefault = (
		extension: string,
		contentType: string,
		extraAttributes?: readonly ContentTypeEntryAttribute[],
	) => {
		if (defaults.has(extension)) return
		defaults.set(extension, contentType)
		const attrs = [
			`Extension="${escapeXml(extension)}"`,
			`ContentType="${escapeXml(contentType)}"`,
			...formatContentTypeExtraAttributes(extraAttributes),
		]
		out.push(`<Default ${attrs.join(' ')}/>`)
	}
	const isDefaultCovered = (partPath: string, contentType: string) => {
		const ext = partPath.split('.').pop()
		return ext !== undefined && defaults.get(ext) === contentType
	}
	const pushOverride = (
		partPath: string,
		contentType: string,
		skipDefaultCovered = false,
		extraAttributes?: readonly ContentTypeEntryAttribute[],
	) => {
		if (skipDefaultCovered && isDefaultCovered(partPath, contentType)) return
		const pn = normalizePartName(partPath)
		if (overrides.has(pn)) return
		overrides.add(pn)
		const preservedOverride = preservedOverrideByPartName.get(pn)
		const resolvedExtraAttributes =
			extraAttributes ??
			(preservedOverride?.contentType === contentType
				? preservedOverride.extraAttributes
				: undefined)
		const attrs = [
			`PartName="${escapeXml(pn)}"`,
			`ContentType="${escapeXml(contentType)}"`,
			...formatContentTypeExtraAttributes(resolvedExtraAttributes),
		]
		out.push(`<Override ${attrs.join(' ')}/>`)
	}
	out.push(XML_HEADER)
	out.push(buildTypesOpenTag(options.preservedContentTypesXml))
	if (preservedDefaults) {
		for (const entry of preservedDefaults) {
			pushDefault(entry.extension, entry.contentType, entry.extraAttributes)
		}
	}
	const preserveSourceDefaultSet =
		preservedDefaults !== undefined && preservedOverrides !== undefined
	if (!preserveSourceDefaultSet) {
		pushDefault('rels', CT_RELS)
		pushDefault('xml', CT_XML)
	}
	pushOverride('xl/workbook.xml', workbookContentType, true)

	for (const sheetPartPath of sheetPartPaths) {
		pushOverride(sheetPartPath, CT_WORKSHEET)
	}

	if (hasSharedStrings) {
		pushOverride('xl/sharedStrings.xml', CT_SHARED_STRINGS)
	}

	if (options.includeStyles ?? true) {
		pushOverride('xl/styles.xml', CT_STYLES)
	}
	if (options.includeDocProps ?? true) {
		pushOverride(docPropsPaths.corePropsPath, CT_CORE_PROPS, true)
		pushOverride(docPropsPaths.appPropsPath, CT_APP_PROPS, true)
	}

	if (capsules) {
		for (const capsule of capsules) {
			if (capsule.contentType) {
				pushOverride(capsule.partPath, capsule.contentType, capsule.contentTypeSource === 'default')
			}
		}
	}

	if (extraOverrides) {
		for (const override of extraOverrides) {
			pushOverride(override.partPath, override.contentType, true)
		}
	}

	if (preservedOverrides) {
		for (const override of preservedOverrides) {
			pushOverride(override.partPath, override.contentType, false, override.extraAttributes)
		}
	}

	out.push('</Types>')
	return out.toString()
}

function buildTypesOpenTag(sourceXml: string | undefined): string {
	const attrs = new Map<string, string>([['xmlns', NS]])
	for (const [name, value] of extractSourceTypesRootAttrs(sourceXml)) {
		if (!attrs.has(name)) attrs.set(name, value)
	}
	return `<Types ${Array.from(attrs, ([name, value]) => `${name}="${escapeXml(value)}"`).join(' ')}>`
}

function extractSourceTypesRootAttrs(sourceXml: string | undefined): readonly [string, string][] {
	if (!sourceXml) return []
	const rootMatch = /<Types\b([^>]*)>/i.exec(sourceXml)
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

function normalizePartName(partPath: string): string {
	return partPath.startsWith('/') ? partPath : `/${partPath}`
}

function formatContentTypeExtraAttributes(
	extraAttributes: readonly ContentTypeEntryAttribute[] | undefined,
): string[] {
	const attrs: string[] = []
	for (const extra of extraAttributes ?? []) {
		if (isKnownContentTypeEntryAttribute(extra.name) || !isXmlAttributeName(extra.name)) continue
		attrs.push(`${extra.name}="${escapeXml(extra.value)}"`)
	}
	return attrs
}

function isKnownContentTypeEntryAttribute(name: string): boolean {
	return name === 'Extension' || name === 'PartName' || name === 'ContentType'
}
