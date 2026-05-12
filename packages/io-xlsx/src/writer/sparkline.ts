import type {
	SheetConditionalFormatColor,
	SheetConditionalFormatValueObject,
	SheetSparklineGroupInfo,
	SheetSparklineInfo,
	SheetX14ConditionalFormatInfo,
	SheetX14DataValidationInfo,
} from '@ascend/core'
import { escapeXml } from '../xml.ts'
import { readXmlAttr, setXmlAttr as setXmlAttrBase, xmlAttrPattern } from './xml-attrs.ts'

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
const DATA_VALIDATION_ENTRY_SOURCE = String.raw`<(${PREFIXED_TAG}dataValidation)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/\1>)`
const DATA_VALIDATION_RE = new RegExp(DATA_VALIDATION_ENTRY_SOURCE, 'g')
const DATA_VALIDATIONS_CONTAINER_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}dataValidations)\b([^>]*)>([\s\S]*?)<\/\1>`,
	'g',
)
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main'
const X14_WORKSHEET_EXT_URI = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}'

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

export function buildWorksheetExtLstXml(updates: {
	readonly x14ConditionalFormats: readonly SheetX14ConditionalFormatInfo[]
	readonly x14DataValidations: readonly SheetX14DataValidationInfo[]
}): string {
	const conditionalFormats = updates.x14ConditionalFormats.filter((entry) => !entry.deleted)
	const dataValidations = updates.x14DataValidations.filter((entry) => !entry.deleted)
	if (conditionalFormats.length === 0 && dataValidations.length === 0) return ''
	const body: string[] = []
	if (conditionalFormats.length > 0) {
		body.push('<x14:conditionalFormattings>')
		for (const format of conditionalFormats) body.push(buildX14ConditionalFormattingXml(format))
		body.push('</x14:conditionalFormattings>')
	}
	if (dataValidations.length > 0) {
		body.push(`<x14:dataValidations count="${dataValidations.length}">`)
		for (const validation of dataValidations) body.push(buildX14DataValidationXml(validation))
		body.push('</x14:dataValidations>')
	}
	return `<extLst xmlns:x14="${X14_NS}" xmlns:xm="${XM_NS}"><ext uri="${X14_WORKSHEET_EXT_URI}">${body.join('')}</ext></extLst>`
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
	const attrRe = xmlAttrPattern(localName)
	return attrRe.test(xml) ? xml.replace(attrRe, ` ${localName}="${escapeXml(value)}"`) : xml
}

function updateX14ConditionalFormattingExtLstXml(
	xml: string,
	formats: readonly SheetX14ConditionalFormatInfo[],
): string {
	if (formats.length === 0) return xml
	let entryIndex = 0
	const next = xml.replace(
		CONDITIONAL_FORMATTING_RE,
		(match, tag: string, attrs: string, body: string) => {
			const format = formats.find((entry) => entry.index === entryIndex)
			entryIndex += 1
			if (!format) return match
			if (format.deleted) return ''
			let updatedBody = setFirstSparklineText(body, 'sqref', format.sqref)
			updatedBody = setOrderedElementTexts(
				updatedBody,
				'f',
				x14ConditionalFormatFormulaTexts(format),
			)
			return `<${tag}${attrs}>${updatedBody}</${tag}>`
		},
	)
	const missing = formats.filter((entry) => !entry.deleted && entry.index >= entryIndex)
	return appendX14ConditionalFormattings(next, missing)
}

function updateX14DataValidationExtLstXml(
	xml: string,
	validations: readonly SheetX14DataValidationInfo[],
): string {
	if (validations.length === 0) return xml
	let entryIndex = 0
	const next = xml.replace(
		DATA_VALIDATION_RE,
		(match, tag: string, attrs: string, body: string | undefined) => {
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
			let updatedBody = setFirstSparklineText(body ?? '', 'sqref', validation.sqref)
			updatedBody = setNestedFormulaText(updatedBody, 'formula1', validation.formula1)
			updatedBody = setNestedFormulaText(updatedBody, 'formula2', validation.formula2)
			if (updatedBody.length === 0) return `<${tag}${updatedAttrs}/>`
			return `<${tag}${updatedAttrs}>${updatedBody}</${tag}>`
		},
	)
	const missing = validations.filter((entry) => !entry.deleted && entry.index >= entryIndex)
	return normalizeX14DataValidationCounts(appendX14DataValidations(next, missing))
}

function normalizeX14DataValidationCounts(xml: string): string {
	return xml.replace(
		DATA_VALIDATIONS_CONTAINER_RE,
		(_match, tag: string, attrs: string, body: string) => {
			const count = [...body.matchAll(new RegExp(DATA_VALIDATION_ENTRY_SOURCE, 'g'))].length
			return `<${tag}${setXmlAttr(attrs, 'count', String(count))}>${body}</${tag}>`
		},
	)
}

function appendX14ConditionalFormattings(
	xml: string,
	formats: readonly SheetX14ConditionalFormatInfo[],
): string {
	if (formats.length === 0) return xml
	const additions = formats.map(buildX14ConditionalFormattingXml).join('')
	const closing = new RegExp(String.raw`<\/(${PREFIXED_TAG}conditionalFormattings)>`)
	if (closing.test(xml)) return xml.replace(closing, `${additions}</$1>`)
	return insertIntoFirstExt(
		xml,
		`<x14:conditionalFormattings>${additions}</x14:conditionalFormattings>`,
	)
}

function appendX14DataValidations(
	xml: string,
	validations: readonly SheetX14DataValidationInfo[],
): string {
	if (validations.length === 0) return xml
	const additions = validations.map(buildX14DataValidationXml).join('')
	const opening = new RegExp(String.raw`<(${PREFIXED_TAG}dataValidations)\b([^>]*)>`)
	const closing = new RegExp(String.raw`<\/(${PREFIXED_TAG}dataValidations)>`)
	if (opening.test(xml) && closing.test(xml)) {
		const withCount = xml.replace(opening, (_match, tag: string, attrs: string) => {
			const existingCount = Number(attr(attrs, 'count') ?? 0)
			const count = Number.isFinite(existingCount)
				? existingCount + validations.length
				: validations.length
			return `<${tag}${setXmlAttr(attrs, 'count', String(count))}>`
		})
		return withCount.replace(closing, `${additions}</$1>`)
	}
	return insertIntoFirstExt(
		xml,
		`<x14:dataValidations count="${validations.length}">${additions}</x14:dataValidations>`,
	)
}

function insertIntoFirstExt(xml: string, body: string): string {
	const closing = new RegExp(String.raw`<\/(${PREFIXED_TAG}ext)>`)
	if (closing.test(xml)) return xml.replace(closing, `${body}</$1>`)
	return xml
}

function buildX14ConditionalFormattingXml(format: SheetX14ConditionalFormatInfo): string {
	const ruleAttrs: string[] = []
	if (format.type) ruleAttrs.push(`type="${escapeXml(format.type)}"`)
	if (format.priority !== undefined) ruleAttrs.push(`priority="${format.priority}"`)
	if (format.id) ruleAttrs.push(`id="${escapeXml(format.id)}"`)
	const ruleBody: string[] = []
	for (const formula of format.formulas) ruleBody.push(`<xm:f>${escapeXml(formula)}</xm:f>`)
	if (format.dataBar) ruleBody.push(buildX14DataBarXml(format.dataBar))
	if (format.iconSet) ruleBody.push(buildX14IconSetXml(format.iconSet))
	return `<x14:conditionalFormatting><x14:cfRule${ruleAttrs.length > 0 ? ` ${ruleAttrs.join(' ')}` : ''}>${ruleBody.join('')}</x14:cfRule><xm:sqref>${escapeXml(format.sqref)}</xm:sqref></x14:conditionalFormatting>`
}

function buildX14DataValidationXml(validation: SheetX14DataValidationInfo): string {
	const attrs: string[] = []
	if (validation.type) attrs.push(`type="${escapeXml(validation.type)}"`)
	if (validation.operator) attrs.push(`operator="${escapeXml(validation.operator)}"`)
	if (validation.allowBlank !== undefined)
		attrs.push(`allowBlank="${validation.allowBlank ? '1' : '0'}"`)
	if (validation.showInputMessage !== undefined) {
		attrs.push(`showInputMessage="${validation.showInputMessage ? '1' : '0'}"`)
	}
	if (validation.showErrorMessage !== undefined) {
		attrs.push(`showErrorMessage="${validation.showErrorMessage ? '1' : '0'}"`)
	}
	if (validation.showDropDown !== undefined) {
		attrs.push(`showDropDown="${validation.showDropDown ? '1' : '0'}"`)
	}
	if (validation.promptTitle) attrs.push(`promptTitle="${escapeXml(validation.promptTitle)}"`)
	if (validation.prompt) attrs.push(`prompt="${escapeXml(validation.prompt)}"`)
	if (validation.errorTitle) attrs.push(`errorTitle="${escapeXml(validation.errorTitle)}"`)
	if (validation.error) attrs.push(`error="${escapeXml(validation.error)}"`)
	if (validation.errorStyle) attrs.push(`errorStyle="${escapeXml(validation.errorStyle)}"`)
	if (validation.imeMode) attrs.push(`imeMode="${escapeXml(validation.imeMode)}"`)
	const body: string[] = []
	if (validation.formula1)
		body.push(`<x14:formula1><xm:f>${escapeXml(validation.formula1)}</xm:f></x14:formula1>`)
	if (validation.formula2)
		body.push(`<x14:formula2><xm:f>${escapeXml(validation.formula2)}</xm:f></x14:formula2>`)
	body.push(`<xm:sqref>${escapeXml(validation.sqref)}</xm:sqref>`)
	return `<x14:dataValidation${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}>${body.join('')}</x14:dataValidation>`
}

function buildX14DataBarXml(
	dataBar: NonNullable<SheetX14ConditionalFormatInfo['dataBar']>,
): string {
	const attrs: string[] = []
	if (dataBar.minLength !== undefined) attrs.push(`minLength="${dataBar.minLength}"`)
	if (dataBar.maxLength !== undefined) attrs.push(`maxLength="${dataBar.maxLength}"`)
	if (dataBar.border !== undefined) attrs.push(`border="${dataBar.border ? '1' : '0'}"`)
	if (dataBar.negativeBarBorderColorSameAsPositive !== undefined) {
		attrs.push(
			`negativeBarBorderColorSameAsPositive="${dataBar.negativeBarBorderColorSameAsPositive ? '1' : '0'}"`,
		)
	}
	const body = [
		...dataBar.cfvo.map(buildX14CfvoXml),
		colorXml('x14:borderColor', dataBar.borderColor),
		colorXml('x14:negativeFillColor', dataBar.negativeFillColor),
		colorXml('x14:negativeBorderColor', dataBar.negativeBorderColor),
		colorXml('x14:axisColor', dataBar.axisColor),
	].filter((entry) => entry.length > 0)
	return `<x14:dataBar${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}>${body.join('')}</x14:dataBar>`
}

function buildX14IconSetXml(
	iconSet: NonNullable<SheetX14ConditionalFormatInfo['iconSet']>,
): string {
	const attrs: string[] = []
	if (iconSet.iconSet) attrs.push(`iconSet="${escapeXml(iconSet.iconSet)}"`)
	if (iconSet.custom !== undefined) attrs.push(`custom="${iconSet.custom ? '1' : '0'}"`)
	if (iconSet.showValue !== undefined) attrs.push(`showValue="${iconSet.showValue ? '1' : '0'}"`)
	if (iconSet.percent !== undefined) attrs.push(`percent="${iconSet.percent ? '1' : '0'}"`)
	if (iconSet.reverse !== undefined) attrs.push(`reverse="${iconSet.reverse ? '1' : '0'}"`)
	const body = [
		...iconSet.cfvo.map(buildX14CfvoXml),
		...(iconSet.icons ?? []).map((icon) => {
			const iconAttrs: string[] = []
			if (icon.iconSet) iconAttrs.push(`iconSet="${escapeXml(icon.iconSet)}"`)
			if (icon.iconId !== undefined) iconAttrs.push(`iconId="${icon.iconId}"`)
			return `<x14:cfIcon${iconAttrs.length > 0 ? ` ${iconAttrs.join(' ')}` : ''}/>`
		}),
	]
	return `<x14:iconSet${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}>${body.join('')}</x14:iconSet>`
}

function buildX14CfvoXml(cfvo: SheetConditionalFormatValueObject): string {
	const attrs: string[] = []
	if (cfvo.type) attrs.push(`type="${escapeXml(cfvo.type)}"`)
	if (cfvo.gte !== undefined) attrs.push(`gte="${cfvo.gte ? '1' : '0'}"`)
	if (cfvo.type === 'formula' && cfvo.value !== undefined) {
		return `<x14:cfvo${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}><xm:f>${escapeXml(cfvo.value)}</xm:f></x14:cfvo>`
	}
	if (cfvo.value !== undefined) attrs.push(`val="${escapeXml(cfvo.value)}"`)
	return `<x14:cfvo${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}/>`
}

function colorXml(tag: string, color: SheetConditionalFormatColor | undefined): string {
	if (!color) return ''
	const attrs: string[] = []
	if (color.rgb) attrs.push(`rgb="${escapeXml(color.rgb)}"`)
	if (color.theme !== undefined) attrs.push(`theme="${color.theme}"`)
	if (color.tint !== undefined) attrs.push(`tint="${color.tint}"`)
	if (color.indexed !== undefined) attrs.push(`indexed="${color.indexed}"`)
	if (color.auto !== undefined) attrs.push(`auto="${color.auto ? '1' : '0'}"`)
	return attrs.length > 0 ? `<${tag} ${attrs.join(' ')}/>` : ''
}

function x14ConditionalFormatFormulaTexts(format: SheetX14ConditionalFormatInfo): string[] {
	return [
		...format.formulas,
		...(format.dataBar?.cfvo ?? [])
			.filter((entry) => entry.type === 'formula' && entry.value !== undefined)
			.map((entry) => entry.value as string),
		...(format.iconSet?.cfvo ?? [])
			.filter((entry) => entry.type === 'formula' && entry.value !== undefined)
			.map((entry) => entry.value as string),
	]
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
	return setXmlAttrBase(attrs, name, value)
}

function attr(attrs: string, name: string): string | undefined {
	return readXmlAttr(attrs, name)
}

function attrRe(name: string): RegExp {
	return xmlAttrPattern(name)
}
