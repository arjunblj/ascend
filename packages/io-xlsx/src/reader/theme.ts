import type { WorkbookThemeColor, WorkbookThemeMetadata } from '@ascend/core'
import { attr, parseXml, type XmlNode } from '../xml.ts'

export function parseThemeXml(xml: string): WorkbookThemeMetadata {
	const doc = parseXml(xml)
	const theme = firstElement(doc, 'theme')
	if (!theme) return { colorCount: 0 }

	const themeElements = childNode(theme, 'themeElements')
	const colorScheme = childNode(themeElements, 'clrScheme')
	const fontScheme = childNode(themeElements, 'fontScheme')
	const majorFont = childNode(fontScheme, 'majorFont')
	const minorFont = childNode(fontScheme, 'minorFont')

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
	const latin = childNode(node, 'latin')
	return latin ? attr(latin, 'typeface') : undefined
}

function readColorScheme(doc: XmlNode): XmlNode | undefined {
	const theme = firstElement(doc, 'theme')
	const themeElements = childNode(theme, 'themeElements')
	return childNode(themeElements, 'clrScheme')
}

function readThemeColor(slot: string, node: XmlNode): WorkbookThemeColor | null {
	const srgb = childNode(node, 'srgbClr')
	if (srgb) {
		const rgb = attr(srgb, 'val')
		return rgb ? { slot, rgb } : { slot }
	}
	const system = childNode(node, 'sysClr')
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

function firstElement(doc: XmlNode, localName: string): XmlNode | undefined {
	for (const [key, value] of Object.entries(doc)) {
		if (key.startsWith('@_') || stripNamespace(key) !== localName || !isXmlNode(value)) continue
		return value
	}
	return undefined
}

function childNode(node: XmlNode | undefined, localName: string): XmlNode | undefined {
	if (!node) return undefined
	for (const [key, value] of Object.entries(node)) {
		if (key.startsWith('@_') || stripNamespace(key) !== localName) continue
		const child = Array.isArray(value) ? value[0] : value
		return isXmlNode(child) ? child : undefined
	}
	return undefined
}

function isXmlNode(value: unknown): value is XmlNode {
	return typeof value === 'object' && value !== null && !Array.isArray(value)
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
