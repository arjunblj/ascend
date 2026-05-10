import type { SheetSparklineGroupInfo } from '@ascend/core'
import { escapeXml } from '../xml.ts'

const XML_NAME = String.raw`[A-Za-z_][\w.-]*`
const PREFIXED_TAG = `(?:${XML_NAME}:)?`
const SPARKLINE_GROUP_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}sparklineGroup)\b([^>]*)>([\s\S]*?)<\/\1>`,
	'g',
)

export function updateSparklineExtLstXml(
	xml: string,
	groups: readonly SheetSparklineGroupInfo[],
): string {
	if (groups.length === 0) return xml
	let groupIndex = 0
	return xml.replace(SPARKLINE_GROUP_RE, (match, tag: string, attrs: string, body: string) => {
		const group = groups.find((entry) => entry.groupIndex === groupIndex++)
		if (!group) return match
		let updatedAttrs = attrs
		updatedAttrs = setOptionalAttr(updatedAttrs, 'type', group.type)
		updatedAttrs = setOptionalAttr(updatedAttrs, 'displayEmptyCellsAs', group.displayEmptyCellsAs)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'dateAxis', group.dateAxis)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'markers', group.markers)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'high', group.highPoint)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'low', group.lowPoint)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'first', group.firstPoint)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'last', group.lastPoint)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'negative', group.negative)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'displayXAxis', group.displayXAxis)
		let updatedBody = body
		updatedBody = setFirstSparklineText(updatedBody, 'f', group.range)
		updatedBody = setFirstSparklineText(updatedBody, 'sqref', group.locationRange)
		return `<${tag}${updatedAttrs}>${updatedBody}</${tag}>`
	})
}

function setFirstSparklineText(xml: string, localName: string, value: string | undefined): string {
	if (value === undefined) return xml
	const tagRe = new RegExp(String.raw`<(${PREFIXED_TAG}${localName})\b([^>]*)>[\s\S]*?<\/\1>`)
	if (tagRe.test(xml)) {
		return xml.replace(tagRe, (_match, tag: string, attrs: string) => {
			return `<${tag}${attrs}>${escapeXml(value)}</${tag}>`
		})
	}
	const attrRe = new RegExp(String.raw`\s${localName}="[^"]*"`)
	return attrRe.test(xml) ? xml.replace(attrRe, ` ${localName}="${escapeXml(value)}"`) : xml
}

function setOptionalAttr(attrs: string, name: string, value: string | undefined): string {
	if (value === undefined) return attrs
	return setXmlAttr(attrs, name, value)
}

function setOptionalBoolAttr(attrs: string, name: string, value: boolean | undefined): string {
	if (value === undefined) return attrs
	return setXmlAttr(attrs, name, value ? '1' : '0')
}

function setXmlAttr(attrs: string, name: string, value: string): string {
	const attrText = `${name}="${escapeXml(value)}"`
	const attrPattern = new RegExp(String.raw`\s${name}="[^"]*"`)
	if (attrPattern.test(attrs)) return attrs.replace(attrPattern, ` ${attrText}`)
	return `${attrs} ${attrText}`
}
