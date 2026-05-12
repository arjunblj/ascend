import type { PivotCacheInfo } from '@ascend/core'
import { escapeXml } from '../xml.ts'
import { setXmlAttr } from './xml-attrs.ts'

export function updatePivotCacheDefinitionXml(xml: string, cache: PivotCacheInfo): string {
	let out = updateRootAttributes(xml, cache)
	out = updateWorksheetSource(out, cache)
	return out
}

function updateRootAttributes(xml: string, cache: PivotCacheInfo): string {
	return xml.replace(/<pivotCacheDefinition\b([^>]*)>/, (_match, attrs: string) => {
		let updated = attrs as string
		updated = setXmlAttr(updated, 'refreshOnLoad', cache.refreshOnLoad)
		updated = setXmlAttr(updated, 'enableRefresh', cache.enableRefresh)
		updated = setXmlAttr(updated, 'invalid', cache.invalid)
		updated = setXmlAttr(updated, 'saveData', cache.saveData)
		return `<pivotCacheDefinition${updated}>`
	})
}

function updateWorksheetSource(xml: string, cache: PivotCacheInfo): string {
	if (cache.sourceSheet === undefined && cache.sourceRef === undefined) return xml
	const nextSource = buildWorksheetSource(cache)
	if (/<worksheetSource\b[^>]*\/>/.test(xml)) {
		return xml.replace(/<worksheetSource\b([^>]*)\/>/, (_match, attrs: string) => {
			let updated = attrs as string
			updated = setXmlAttr(updated, 'sheet', cache.sourceSheet)
			updated = setXmlAttr(updated, 'ref', cache.sourceRef)
			return `<worksheetSource${updated}/>`
		})
	}
	if (/<cacheSource\b[^>]*>/.test(xml)) {
		return xml.replace(/(<cacheSource\b[^>]*>)/, `$1${nextSource}`)
	}
	return xml.replace(
		/(<pivotCacheDefinition\b[^>]*>)/,
		`$1<cacheSource type="worksheet">${nextSource}</cacheSource>`,
	)
}

function buildWorksheetSource(cache: PivotCacheInfo): string {
	const attrs: string[] = []
	if (cache.sourceRef !== undefined) attrs.push(`ref="${escapeXml(cache.sourceRef)}"`)
	if (cache.sourceSheet !== undefined) attrs.push(`sheet="${escapeXml(cache.sourceSheet)}"`)
	return `<worksheetSource ${attrs.join(' ')}/>`
}
