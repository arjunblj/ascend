import type { WorkbookThemeColor, WorkbookThemeMetadata } from '@ascend/core'
import { attr, parseXml, type XmlNode } from '../xml.ts'

export function parseThemeXml(xml: string): WorkbookThemeMetadata {
	const doc = parseXml(xml)
	const theme = (doc['a:theme'] ?? doc.theme) as XmlNode | undefined
	if (!theme) return { colorCount: 0 }

	const themeElements = (theme['a:themeElements'] ?? theme.themeElements) as XmlNode | undefined
	const colorScheme = (themeElements?.['a:clrScheme'] ?? themeElements?.clrScheme) as
		| XmlNode
		| undefined
	const fontScheme = (themeElements?.['a:fontScheme'] ?? themeElements?.fontScheme) as
		| XmlNode
		| undefined
	const majorFont = (fontScheme?.['a:majorFont'] ?? fontScheme?.majorFont) as XmlNode | undefined
	const minorFont = (fontScheme?.['a:minorFont'] ?? fontScheme?.minorFont) as XmlNode | undefined

	const metadata: WorkbookThemeMetadata = {
		colorCount: countChildElements(colorScheme),
	}
	const themeName = attr(theme, 'name')
	if (themeName) Object.assign(metadata, { name: themeName })
	const colorSchemeName = colorScheme ? attr(colorScheme, 'name') : undefined
	if (colorSchemeName) Object.assign(metadata, { colorSchemeName })
	const majorFontLatin = readLatinTypeface(majorFont)
	if (majorFontLatin) Object.assign(metadata, { majorFontLatin })
	const minorFontLatin = readLatinTypeface(minorFont)
	if (minorFontLatin) Object.assign(metadata, { minorFontLatin })
	return metadata
}

export function parseThemeColorsXml(xml: string): WorkbookThemeColor[] {
	const colorScheme = readColorScheme(parseXml(xml))
	if (!colorScheme) return []
	const colors: WorkbookThemeColor[] = []
	for (const [key, value] of Object.entries(colorScheme)) {
		if (key.startsWith('@_')) continue
		const node = Array.isArray(value) ? value[0] : value
		if (!node || typeof node !== 'object') continue
		const color = readThemeColor(stripNamespace(key), node as XmlNode)
		if (color) colors.push(color)
	}
	return colors
}

function readLatinTypeface(node: XmlNode | undefined): string | undefined {
	if (!node) return undefined
	const latin = (node['a:latin'] ?? node.latin) as XmlNode | undefined
	return latin ? attr(latin, 'typeface') : undefined
}

function readColorScheme(doc: XmlNode): XmlNode | undefined {
	const theme = (doc['a:theme'] ?? doc.theme) as XmlNode | undefined
	const themeElements = (theme?.['a:themeElements'] ?? theme?.themeElements) as XmlNode | undefined
	return (themeElements?.['a:clrScheme'] ?? themeElements?.clrScheme) as XmlNode | undefined
}

function readThemeColor(slot: string, node: XmlNode): WorkbookThemeColor | null {
	const srgb = (node['a:srgbClr'] ?? node.srgbClr) as XmlNode | undefined
	if (srgb) {
		const rgb = attr(srgb, 'val')
		return rgb ? { slot, rgb } : { slot }
	}
	const system = (node['a:sysClr'] ?? node.sysClr) as XmlNode | undefined
	if (system) {
		const systemColor = attr(system, 'val')
		const lastColor = attr(system, 'lastClr')
		return {
			slot,
			...(systemColor ? { systemColor } : {}),
			...(lastColor ? { lastColor } : {}),
		}
	}
	return { slot }
}

function stripNamespace(name: string): string {
	const colon = name.indexOf(':')
	return colon >= 0 ? name.slice(colon + 1) : name
}

function countChildElements(node: XmlNode | undefined): number {
	if (!node) return 0
	let count = 0
	for (const [key, value] of Object.entries(node)) {
		if (key.startsWith('@_')) continue
		if (Array.isArray(value)) count += value.length
		else if (value !== undefined && value !== null) count += 1
	}
	return count
}
