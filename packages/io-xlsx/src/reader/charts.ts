import type { ChartPartInfo, ChartSeriesInfo } from '@ascend/core'
import { decodeXmlText } from './xml-utils.ts'

const NS_PREFIX = String.raw`(?:[A-Za-z_][\w.-]*:)?`
const CHART_TYPE_RE = new RegExp(`<${NS_PREFIX}([A-Za-z]+Chart)\\b`)
const TITLE_RE = new RegExp(
	`<${NS_PREFIX}title\\b[\\s\\S]*?<${NS_PREFIX}t>([\\s\\S]*?)<\\/${NS_PREFIX}t>[\\s\\S]*?<\\/${NS_PREFIX}title>`,
)
const SERIES_RE = new RegExp(`<${NS_PREFIX}ser\\b[\\s\\S]*?<\\/${NS_PREFIX}ser>`, 'g')
const TX_REF_RE = new RegExp(
	`<${NS_PREFIX}tx\\b[\\s\\S]*?<${NS_PREFIX}strRef\\b[\\s\\S]*?<${NS_PREFIX}f>([\\s\\S]*?)<\\/${NS_PREFIX}f>`,
)
const TX_TEXT_RE = new RegExp(
	`<${NS_PREFIX}tx\\b[\\s\\S]*?<${NS_PREFIX}v>([\\s\\S]*?)<\\/${NS_PREFIX}v>[\\s\\S]*?<\\/${NS_PREFIX}tx>`,
)
const CAT_REF_RE = new RegExp(
	`<${NS_PREFIX}cat\\b[\\s\\S]*?<${NS_PREFIX}(?:strRef|numRef|multiLvlStrRef)\\b[\\s\\S]*?<${NS_PREFIX}f>([\\s\\S]*?)<\\/${NS_PREFIX}f>`,
)
const VAL_REF_RE = new RegExp(
	`<${NS_PREFIX}val\\b[\\s\\S]*?<${NS_PREFIX}numRef\\b[\\s\\S]*?<${NS_PREFIX}f>([\\s\\S]*?)<\\/${NS_PREFIX}f>`,
)
const X_VAL_REF_RE = new RegExp(
	`<${NS_PREFIX}xVal\\b[\\s\\S]*?<${NS_PREFIX}(?:strRef|numRef)\\b[\\s\\S]*?<${NS_PREFIX}f>([\\s\\S]*?)<\\/${NS_PREFIX}f>`,
)
const Y_VAL_REF_RE = new RegExp(
	`<${NS_PREFIX}yVal\\b[\\s\\S]*?<${NS_PREFIX}numRef\\b[\\s\\S]*?<${NS_PREFIX}f>([\\s\\S]*?)<\\/${NS_PREFIX}f>`,
)

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
