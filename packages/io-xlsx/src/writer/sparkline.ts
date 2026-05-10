import type {
	SheetSparklineGroupInfo,
	SheetSparklineInfo,
	SheetX14ConditionalFormatInfo,
	SheetX14DataValidationInfo,
} from '@ascend/core'
import { escapeXml } from '../xml.ts'

const XML_NAME = String.raw`[A-Za-z_][\w.-]*`
const PREFIXED_TAG = `(?:${XML_NAME}:)?`
const SPARKLINE_GROUP_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}sparklineGroup)\b([^>]*)>([\s\S]*?)<\/\1>`,
	'g',
)
const CONDITIONAL_FORMATTING_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}conditionalFormatting)\b([^>]*)>([\s\S]*?)<\/\1>`,
	'g',
)
const DATA_VALIDATION_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}dataValidation)\b([^>]*)>([\s\S]*?)<\/\1>`,
	'g',
)

export function updateWorksheetExtLstXml(
	xml: string,
	updates: {
		readonly sparklineGroups: readonly SheetSparklineGroupInfo[]
		readonly x14ConditionalFormats: readonly SheetX14ConditionalFormatInfo[]
		readonly x14DataValidations: readonly SheetX14DataValidationInfo[]
	},
): string {
	let next = updateSparklineExtLstXml(xml, updates.sparklineGroups)
	next = updateX14ConditionalFormattingExtLstXml(next, updates.x14ConditionalFormats)
	next = updateX14DataValidationExtLstXml(next, updates.x14DataValidations)
	return next
}

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
		updatedAttrs = setOptionalNumberAttr(updatedAttrs, 'manualMax', group.manualMax)
		updatedAttrs = setOptionalNumberAttr(updatedAttrs, 'manualMin', group.manualMin)
		updatedAttrs = setOptionalNumberAttr(updatedAttrs, 'lineWeight', group.lineWeight)
		updatedAttrs = setOptionalAttr(updatedAttrs, 'xr2:uid', group.uid)
		updatedAttrs = setOptionalAttr(updatedAttrs, 'displayEmptyCellsAs', group.displayEmptyCellsAs)
		updatedAttrs = setOptionalAttr(updatedAttrs, 'minAxisType', group.minAxisType)
		updatedAttrs = setOptionalAttr(updatedAttrs, 'maxAxisType', group.maxAxisType)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'dateAxis', group.dateAxis)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'markers', group.markers)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'high', group.highPoint)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'low', group.lowPoint)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'first', group.firstPoint)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'last', group.lastPoint)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'negative', group.negative)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'displayXAxis', group.displayXAxis)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'displayHidden', group.displayHidden)
		updatedAttrs = setOptionalBoolAttr(updatedAttrs, 'rightToLeft', group.rightToLeft)
		let updatedBody = body
		updatedBody = setDirectChildTextBefore(updatedBody, 'sparklines', 'f', group.dateAxisRange)
		updatedBody = setSparklineRefs(updatedBody, group)
		return `<${tag}${updatedAttrs}>${updatedBody}</${tag}>`
	})
}

function setSparklineRefs(xml: string, group: SheetSparklineGroupInfo): string {
	const refs = sparklineRefsForGroup(group)
	if (refs.length === 0) return xml
	let index = 0
	const sparklineRe = new RegExp(
		String.raw`<(${PREFIXED_TAG}sparkline)\b([^>]*)>([\s\S]*?)<\/\1>`,
		'g',
	)
	return xml.replace(sparklineRe, (match, tag: string, attrs: string, body: string) => {
		const ref = refs[index++]
		if (!ref) return match
		let updatedBody = body
		updatedBody = setFirstSparklineText(updatedBody, 'f', ref.range)
		updatedBody = setFirstSparklineText(updatedBody, 'sqref', ref.locationRange)
		return `<${tag}${attrs}>${updatedBody}</${tag}>`
	})
}

function sparklineRefsForGroup(group: SheetSparklineGroupInfo): readonly SheetSparklineInfo[] {
	const refs: SheetSparklineInfo[] =
		group.sparklines && group.sparklines.length > 0
			? group.sparklines.map((ref) => ({ ...ref }))
			: [sparklineRefFromGroup(group)].filter((ref) => Object.keys(ref).length > 0)
	if (refs.length === 0) return refs
	const first = refs[0]
	if (!first) return refs
	refs[0] = {
		...first,
		...(group.range !== undefined ? { range: group.range } : {}),
		...(group.locationRange !== undefined ? { locationRange: group.locationRange } : {}),
	}
	return refs
}

function sparklineRefFromGroup(group: SheetSparklineGroupInfo): SheetSparklineInfo {
	return {
		...(group.range !== undefined ? { range: group.range } : {}),
		...(group.locationRange !== undefined ? { locationRange: group.locationRange } : {}),
	}
}

function setDirectChildTextBefore(
	xml: string,
	stopLocalName: string,
	localName: string,
	value: string | undefined,
): string {
	if (value === undefined) return xml
	const stopRe = new RegExp(String.raw`<${PREFIXED_TAG}${stopLocalName}\b`)
	const stop = stopRe.exec(xml)
	const prefix = stop ? xml.slice(0, stop.index) : xml
	const suffix = stop ? xml.slice(stop.index) : ''
	const updatedPrefix = setFirstSparklineText(prefix, localName, value)
	return `${updatedPrefix}${suffix}`
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

function updateX14ConditionalFormattingExtLstXml(
	xml: string,
	formats: readonly SheetX14ConditionalFormatInfo[],
): string {
	if (formats.length === 0) return xml
	let entryIndex = 0
	return xml.replace(
		CONDITIONAL_FORMATTING_RE,
		(match, tag: string, attrs: string, body: string) => {
			const format = formats.find((entry) => entry.index === entryIndex)
			entryIndex += 1
			if (!format) return match
			if (format.deleted) return ''
			let updatedBody = setFirstSparklineText(body, 'sqref', format.sqref)
			updatedBody = setOrderedElementTexts(updatedBody, 'f', format.formulas)
			return `<${tag}${attrs}>${updatedBody}</${tag}>`
		},
	)
}

function updateX14DataValidationExtLstXml(
	xml: string,
	validations: readonly SheetX14DataValidationInfo[],
): string {
	if (validations.length === 0) return xml
	let entryIndex = 0
	return xml.replace(DATA_VALIDATION_RE, (match, tag: string, attrs: string, body: string) => {
		const validation = validations.find((entry) => entry.index === entryIndex)
		entryIndex += 1
		if (!validation) return match
		if (validation.deleted) return ''
		let updatedAttrs = attrs
		if (attrRe('sqref').test(updatedAttrs)) {
			updatedAttrs = updatedAttrs.replace(
				attrRe('sqref'),
				` sqref="${escapeXml(validation.sqref)}"`,
			)
		}
		let updatedBody = setFirstSparklineText(body, 'sqref', validation.sqref)
		updatedBody = setNestedFormulaText(updatedBody, 'formula1', validation.formula1)
		updatedBody = setNestedFormulaText(updatedBody, 'formula2', validation.formula2)
		return `<${tag}${updatedAttrs}>${updatedBody}</${tag}>`
	})
}

function setNestedFormulaText(
	xml: string,
	localName: 'formula1' | 'formula2',
	value: string | undefined,
): string {
	if (value === undefined) return xml
	const tagRe = new RegExp(String.raw`<(${PREFIXED_TAG}${localName})\b([^>]*)>([\s\S]*?)<\/\1>`)
	return xml.replace(tagRe, (_match, tag: string, attrs: string, body: string) => {
		const updatedBody = setFirstSparklineText(body, 'f', value)
		return `<${tag}${attrs}>${updatedBody}</${tag}>`
	})
}

function setOrderedElementTexts(xml: string, localName: string, values: readonly string[]): string {
	if (values.length === 0) return xml
	let index = 0
	const tagRe = new RegExp(String.raw`<(${PREFIXED_TAG}${localName})\b([^>]*)>[\s\S]*?<\/\1>`, 'g')
	return xml.replace(tagRe, (match, tag: string, attrs: string) => {
		const value = values[index++]
		return value === undefined ? match : `<${tag}${attrs}>${escapeXml(value)}</${tag}>`
	})
}

function setOptionalAttr(attrs: string, name: string, value: string | undefined): string {
	if (value === undefined) return attrs
	return setXmlAttr(attrs, name, value)
}

function setOptionalNumberAttr(attrs: string, name: string, value: number | undefined): string {
	if (value === undefined) return attrs
	return setXmlAttr(attrs, name, String(value))
}

function setOptionalBoolAttr(attrs: string, name: string, value: boolean | undefined): string {
	if (value === undefined) return attrs
	return setXmlAttr(attrs, name, value ? '1' : '0')
}

function setXmlAttr(attrs: string, name: string, value: string): string {
	const attrText = `${name}="${escapeXml(value)}"`
	const attrPattern = attrRe(name)
	if (attrPattern.test(attrs)) return attrs.replace(attrPattern, ` ${attrText}`)
	return `${attrs} ${attrText}`
}

function attrRe(name: string): RegExp {
	return new RegExp(String.raw`\s${name}="[^"]*"`)
}
