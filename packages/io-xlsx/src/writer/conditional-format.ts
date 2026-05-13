import type { SheetConditionalFormatRule } from '@ascend/core'
import { escapeXml } from '../xml.ts'

type ColorScale = NonNullable<SheetConditionalFormatRule['colorScale']>
type DataBar = NonNullable<SheetConditionalFormatRule['dataBar']>
type IconSet = NonNullable<SheetConditionalFormatRule['iconSet']>
type Cfvo = ColorScale['cfvo'][number]
type CfColor = ColorScale['colors'][number]

export function buildColorScaleXml(colorScale: ColorScale): string {
	const attrs = preservedAttrsXml(colorScale.preservedAttributes)
	const parts = [attrs.length > 0 ? `<colorScale ${attrs}>` : '<colorScale>']
	for (const cfvo of colorScale.cfvo) parts.push(cfvoXml(cfvo))
	for (const color of colorScale.colors) parts.push(cfColorXml(color))
	parts.push(...(colorScale.preservedChildXml ?? []))
	parts.push('</colorScale>')
	return parts.join('')
}

export function buildDataBarXml(dataBar: DataBar): string {
	const attrs: string[] = []
	if (dataBar.minLength !== undefined) attrs.push(`minLength="${dataBar.minLength}"`)
	if (dataBar.maxLength !== undefined) attrs.push(`maxLength="${dataBar.maxLength}"`)
	if (dataBar.showValue !== undefined) attrs.push(`showValue="${dataBar.showValue ? '1' : '0'}"`)
	const parts = [attrs.length > 0 ? `<dataBar ${attrs.join(' ')}>` : '<dataBar>']
	for (const cfvo of dataBar.cfvo) parts.push(cfvoXml(cfvo))
	if (dataBar.color) parts.push(cfColorXml(dataBar.color))
	parts.push('</dataBar>')
	return parts.join('')
}

export function buildIconSetXml(iconSet: IconSet): string {
	const attrs: string[] = []
	if (iconSet.iconSet) attrs.push(`iconSet="${escapeXml(iconSet.iconSet)}"`)
	if (iconSet.showValue !== undefined) attrs.push(`showValue="${iconSet.showValue ? '1' : '0'}"`)
	if (iconSet.percent !== undefined) attrs.push(`percent="${iconSet.percent ? '1' : '0'}"`)
	if (iconSet.reverse !== undefined) attrs.push(`reverse="${iconSet.reverse ? '1' : '0'}"`)
	const parts = [attrs.length > 0 ? `<iconSet ${attrs.join(' ')}>` : '<iconSet>']
	for (const cfvo of iconSet.cfvo) parts.push(cfvoXml(cfvo))
	parts.push('</iconSet>')
	return parts.join('')
}

function cfvoXml(cfvo: Cfvo): string {
	const attrs: string[] = []
	if (cfvo.type) attrs.push(`type="${escapeXml(cfvo.type)}"`)
	if (cfvo.value !== undefined) attrs.push(`val="${escapeXml(cfvo.value)}"`)
	if (cfvo.gte !== undefined) attrs.push(`gte="${cfvo.gte ? '1' : '0'}"`)
	return `<cfvo ${attrs.join(' ')}/>`
}

function cfColorXml(color: CfColor): string {
	const attrs: string[] = []
	if (color.rgb) attrs.push(`rgb="${escapeXml(color.rgb)}"`)
	if (color.theme !== undefined) attrs.push(`theme="${color.theme}"`)
	if (color.tint !== undefined) attrs.push(`tint="${color.tint}"`)
	if (color.indexed !== undefined) attrs.push(`indexed="${color.indexed}"`)
	if (color.auto !== undefined) attrs.push(`auto="${color.auto ? '1' : '0'}"`)
	return `<color ${attrs.join(' ')}/>`
}

function preservedAttrsXml(attrs: Readonly<Record<string, string>> | undefined): string {
	return Object.entries(attrs ?? {})
		.filter(([name]) => name !== 'xmlns' && !name.startsWith('xmlns:'))
		.map(([name, value]) => `${name}="${escapeXml(value)}"`)
		.join(' ')
}
