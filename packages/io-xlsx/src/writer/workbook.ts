import type { Workbook } from '@ascend/core'
import { escapeXml } from '../xml.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export function buildWorkbookXml(workbook: Workbook): string {
	const parts: string[] = [XML_HEADER, `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`]

	if (workbook.calcSettings.dateSystem === '1904') {
		parts.push('<workbookPr date1904="1"/>')
	}

	parts.push('<sheets>')
	for (let i = 0; i < workbook.sheets.length; i++) {
		const sheet = workbook.sheets[i]
		if (!sheet) continue
		const attrs = [`name="${escapeXml(sheet.name)}"`, `sheetId="${i + 1}"`, `r:id="rId${i + 1}"`]
		if (sheet.state !== 'visible') {
			attrs.push(`state="${sheet.state}"`)
		}
		parts.push(`<sheet ${attrs.join(' ')}/>`)
	}
	parts.push('</sheets>')

	if (workbook.definedNames.size > 0) {
		parts.push('<definedNames>')
		for (const [name, formula] of workbook.definedNames) {
			parts.push(`<definedName name="${escapeXml(name)}">${escapeXml(formula)}</definedName>`)
		}
		parts.push('</definedNames>')
	}

	const cs = workbook.calcSettings
	const calcAttrs: string[] = []
	if (cs.calcMode !== 'auto') calcAttrs.push(`calcMode="${cs.calcMode}"`)
	if (cs.fullCalcOnLoad) calcAttrs.push('fullCalcOnLoad="1"')
	if (cs.iterativeCalc.enabled) {
		calcAttrs.push('iterate="1"')
		calcAttrs.push(`iterateCount="${cs.iterativeCalc.maxIterations}"`)
		calcAttrs.push(`iterateDelta="${cs.iterativeCalc.maxChange}"`)
	}
	if (calcAttrs.length > 0) {
		parts.push(`<calcPr ${calcAttrs.join(' ')}/>`)
	}

	parts.push('</workbook>')
	return parts.join('')
}
