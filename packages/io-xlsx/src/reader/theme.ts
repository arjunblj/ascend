import type { WorkbookThemeMetadata } from '@ascend/core'
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

function readLatinTypeface(node: XmlNode | undefined): string | undefined {
	if (!node) return undefined
	const latin = (node['a:latin'] ?? node.latin) as XmlNode | undefined
	return latin ? attr(latin, 'typeface') : undefined
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
