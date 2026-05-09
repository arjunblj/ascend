import type { ChartPartInfo, ChartSeriesInfo } from '@ascend/core'
import { decodeXmlText } from './xml-utils.ts'

const CHART_TYPE_RE = /<(?:c:)?([A-Za-z]+Chart)\b/
const TITLE_RE = /<(?:c:)?title\b[\s\S]*?<a:t>([\s\S]*?)<\/a:t>[\s\S]*?<\/(?:c:)?title>/
const SERIES_RE = /<(?:c:)?ser\b[\s\S]*?<\/(?:c:)?ser>/g
const TX_REF_RE = /<(?:c:)?tx\b[\s\S]*?<(?:c:)?strRef\b[\s\S]*?<(?:c:)?f>([\s\S]*?)<\/(?:c:)?f>/
const TX_TEXT_RE = /<(?:c:)?tx\b[\s\S]*?<(?:c:)?v>([\s\S]*?)<\/(?:c:)?v>[\s\S]*?<\/(?:c:)?tx>/
const CAT_REF_RE =
	/<(?:c:)?cat\b[\s\S]*?<(?:c:)?(?:strRef|numRef|multiLvlStrRef)\b[\s\S]*?<(?:c:)?f>([\s\S]*?)<\/(?:c:)?f>/
const VAL_REF_RE = /<(?:c:)?val\b[\s\S]*?<(?:c:)?numRef\b[\s\S]*?<(?:c:)?f>([\s\S]*?)<\/(?:c:)?f>/
const X_VAL_REF_RE =
	/<(?:c:)?xVal\b[\s\S]*?<(?:c:)?(?:strRef|numRef)\b[\s\S]*?<(?:c:)?f>([\s\S]*?)<\/(?:c:)?f>/
const Y_VAL_REF_RE =
	/<(?:c:)?yVal\b[\s\S]*?<(?:c:)?numRef\b[\s\S]*?<(?:c:)?f>([\s\S]*?)<\/(?:c:)?f>/

export function parseChartXml(xml: string, partPath: string, sheetName?: string): ChartPartInfo {
	const chartType = CHART_TYPE_RE.exec(xml)?.[1]
	const title = TITLE_RE.exec(xml)?.[1]
	const series: ChartSeriesInfo[] = []
	for (const match of xml.matchAll(SERIES_RE)) {
		const serXml = match[0]
		const nameRef = readFormula(serXml, TX_REF_RE)
		const nameText = readText(serXml, TX_TEXT_RE)
		const categoryRef = readFormula(serXml, CAT_REF_RE) ?? readFormula(serXml, X_VAL_REF_RE)
		const valueRef = readFormula(serXml, VAL_REF_RE) ?? readFormula(serXml, Y_VAL_REF_RE)
		series.push({
			...(nameRef ? { nameRef } : {}),
			...(nameText ? { nameText } : {}),
			...(categoryRef ? { categoryRef } : {}),
			...(valueRef ? { valueRef } : {}),
		})
	}
	return {
		partPath,
		...(sheetName ? { sheetName } : {}),
		...(chartType ? { chartType } : {}),
		...(title ? { title: decodeXmlText(title) } : {}),
		series,
	}
}

function readFormula(xml: string, pattern: RegExp): string | undefined {
	const value = pattern.exec(xml)?.[1]
	return value ? decodeXmlText(value.trim()) : undefined
}

function readText(xml: string, pattern: RegExp): string | undefined {
	const value = pattern.exec(xml)?.[1]
	return value ? decodeXmlText(value.trim()) : undefined
}
