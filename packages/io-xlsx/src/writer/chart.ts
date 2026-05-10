import type { ChartPartInfo } from '@ascend/core'
import { decodeXmlText } from '../reader/xml-utils.ts'
import { escapeXml } from '../xml.ts'

const SERIES_RE = /<(?:c:)?ser\b[\s\S]*?<\/(?:c:)?ser>/g

export function updateChartXml(xml: string, chart: ChartPartInfo): string {
	let index = 0
	return xml.replace(SERIES_RE, (seriesXml) => {
		const series = chart.series[index]
		index += 1
		if (!series) return seriesXml
		let next = seriesXml
		if (series.nameRef !== undefined) next = replaceFormula(next, 'tx', series.nameRef).xml
		if (series.categoryRef !== undefined) {
			next = replaceFirstFormula(next, ['cat', 'xVal'], series.categoryRef)
		}
		if (series.valueRef !== undefined) {
			next = replaceFirstFormula(next, ['val', 'yVal'], series.valueRef)
		}
		return next
	})
}

function replaceFirstFormula(xml: string, tags: readonly string[], formula: string): string {
	for (const tag of tags) {
		const result = replaceFormula(xml, tag, formula)
		if (result.matched) return result.xml
	}
	return xml
}

function replaceFormula(
	xml: string,
	tag: string,
	formula: string,
): { xml: string; matched: boolean } {
	const pattern = new RegExp(`(<(?:c:)?${tag}\\b[\\s\\S]*?<(?:c:)?f>)([\\s\\S]*?)(</(?:c:)?f>)`)
	let matched = false
	const updated = xml.replace(pattern, (match, open: string, oldFormula: string, close: string) => {
		matched = true
		if (decodeXmlText(oldFormula) === formula) return match
		return `${open}${escapeXml(formula)}${close}`
	})
	return { xml: updated, matched }
}
