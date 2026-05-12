import type { PivotCacheInfo } from '@ascend/core'
import { escapeXml } from '../xml.ts'
import { setXmlAttr } from './xml-attrs.ts'

export function updatePivotCacheDefinitionXml(xml: string, cache: PivotCacheInfo): string {
	let out = updateRootAttributes(xml, cache)
	out = updateWorksheetSource(out, cache)
	return out
}

function updateRootAttributes(xml: string, cache: PivotCacheInfo): string {
	return xml.replace(pivotCacheDefinitionOpenPattern, (_match, tagName: string, attrs: string) => {
		let updated = attrs as string
		updated = setXmlAttr(updated, 'refreshOnLoad', cache.refreshOnLoad)
		updated = setXmlAttr(updated, 'enableRefresh', cache.enableRefresh)
		updated = setXmlAttr(updated, 'invalid', cache.invalid)
		updated = setXmlAttr(updated, 'saveData', cache.saveData)
		return `<${tagName}${updated}>`
	})
}

function updateWorksheetSource(xml: string, cache: PivotCacheInfo): string {
	if (cache.sourceSheet === undefined && cache.sourceRef === undefined) return xml
	const worksheetSourceMatch = worksheetSourceSelfClosingPattern.exec(xml)
	if (worksheetSourceMatch) {
		return xml.replace(
			worksheetSourceSelfClosingPattern,
			(_match, tagName: string, attrs: string) => {
				let updated = attrs as string
				updated = setXmlAttr(updated, 'sheet', cache.sourceSheet)
				updated = setXmlAttr(updated, 'ref', cache.sourceRef)
				return `<${tagName}${updated}/>`
			},
		)
	}
	const cacheSourceSelfClosingMatch = cacheSourceSelfClosingPattern.exec(xml)
	if (cacheSourceSelfClosingMatch) {
		return xml.replace(
			cacheSourceSelfClosingPattern,
			(_match, tagName: string, attrs: string) =>
				`<${tagName}${attrs}>${buildWorksheetSource(cache, tagPrefix(tagName))}</${tagName}>`,
		)
	}
	const cacheSourceMatch = cacheSourceOpenPattern.exec(xml)
	if (cacheSourceMatch) {
		return xml.replace(
			cacheSourceOpenPattern,
			(match, tagName: string) => `${match}${buildWorksheetSource(cache, tagPrefix(tagName))}`,
		)
	}
	return xml.replace(pivotCacheDefinitionOpenPattern, (match, tagName: string) => {
		const prefix = tagPrefix(tagName)
		return `${match}<${prefix}cacheSource type="worksheet">${buildWorksheetSource(cache, prefix)}</${prefix}cacheSource>`
	})
}

const pivotCacheDefinitionOpenPattern = /<((?:[A-Za-z_][\w.-]*:)?pivotCacheDefinition)\b([^>]*)>/
const cacheSourceSelfClosingPattern = /<((?:[A-Za-z_][\w.-]*:)?cacheSource)\b([^>]*)\/>/
const cacheSourceOpenPattern = /<((?:[A-Za-z_][\w.-]*:)?cacheSource)\b[^>]*>/
const worksheetSourceSelfClosingPattern = /<((?:[A-Za-z_][\w.-]*:)?worksheetSource)\b([^>]*)\/>/

function tagPrefix(tagName: string): string {
	const separatorIndex = tagName.indexOf(':')
	return separatorIndex === -1 ? '' : tagName.slice(0, separatorIndex + 1)
}

function buildWorksheetSource(cache: PivotCacheInfo, prefix = ''): string {
	const attrs: string[] = []
	if (cache.sourceRef !== undefined) attrs.push(`ref="${escapeXml(cache.sourceRef)}"`)
	if (cache.sourceSheet !== undefined) attrs.push(`sheet="${escapeXml(cache.sourceSheet)}"`)
	return `<${prefix}worksheetSource ${attrs.join(' ')}/>`
}
