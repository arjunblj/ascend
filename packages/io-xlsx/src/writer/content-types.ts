import type { PreservationCapsule } from '../preserve.ts'
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

export interface ContentTypeDefault {
	readonly extension: string
	readonly contentType: string
}

export function buildContentTypesXml(
	sheetCount: number,
	hasSharedStrings: boolean,
	workbookContentType = CT_WORKBOOK,
	capsules?: PreservationCapsule[],
	extraOverrides?: readonly { partPath: string; contentType: string }[],
	preservedDefaults?: readonly ContentTypeDefault[],
	docPropsPaths: { readonly corePropsPath: string; readonly appPropsPath: string } = {
		corePropsPath: 'docProps/core.xml',
		appPropsPath: 'docProps/app.xml',
	},
): string {
	const out = new ChunkedStringBuilder()
	const defaults = new Map<string, string>()
	const overrides = new Set<string>()
	const pushDefault = (extension: string, contentType: string) => {
		if (defaults.has(extension)) return
		defaults.set(extension, contentType)
		out.push(`<Default Extension="${extension}" ContentType="${contentType}"/>`)
	}
	const isDefaultCovered = (partPath: string, contentType: string) => {
		const ext = partPath.split('.').pop()
		return ext !== undefined && defaults.get(ext) === contentType
	}
	const pushOverride = (partPath: string, contentType: string, skipDefaultCovered = false) => {
		if (skipDefaultCovered && isDefaultCovered(partPath, contentType)) return
		const pn = partPath.startsWith('/') ? partPath : `/${partPath}`
		if (overrides.has(pn)) return
		overrides.add(pn)
		out.push(`<Override PartName="${pn}" ContentType="${contentType}"/>`)
	}
	out.push(XML_HEADER)
	out.push(`<Types xmlns="${NS}">`)
	if (preservedDefaults) {
		for (const entry of preservedDefaults) {
			pushDefault(entry.extension, entry.contentType)
		}
	}
	pushDefault('rels', CT_RELS)
	pushDefault('xml', CT_XML)
	pushOverride('xl/workbook.xml', workbookContentType, true)

	for (let i = 1; i <= sheetCount; i++) {
		pushOverride(`xl/worksheets/sheet${i}.xml`, CT_WORKSHEET)
	}

	if (hasSharedStrings) {
		pushOverride('xl/sharedStrings.xml', CT_SHARED_STRINGS)
	}

	pushOverride('xl/styles.xml', CT_STYLES)
	pushOverride(docPropsPaths.corePropsPath, CT_CORE_PROPS, true)
	pushOverride(docPropsPaths.appPropsPath, CT_APP_PROPS, true)

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

	out.push('</Types>')
	return out.toString()
}
