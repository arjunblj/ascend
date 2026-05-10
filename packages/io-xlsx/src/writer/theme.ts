import type { WorkbookThemeColor, WorkbookThemeMetadata } from '@ascend/core'
import { parseThemeColorsXml, parseThemeXml } from '../reader/theme.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const XML_NAME = String.raw`[A-Za-z_][\w.-]*`
const PREFIXED_TAG = `(?:${XML_NAME}:)?`
const THEME_RE = new RegExp(String.raw`<(${PREFIXED_TAG}theme)\b([^>]*)>`)
const CLR_SCHEME_RE = new RegExp(String.raw`<(${PREFIXED_TAG}clrScheme)\b([^>]*)>`)
const LATIN_RE = (fontTag: 'majorFont' | 'minorFont') =>
	new RegExp(
		String.raw`<(${PREFIXED_TAG}${fontTag})\b[^>]*>[\s\S]*?<(${PREFIXED_TAG}latin)\b([^>]*?)(?:\/>|>(?:[\s\S]*?)<\/\2>)[\s\S]*?<\/\1>`,
	)
const COLOR_SLOT_RE = (slot: string) =>
	new RegExp(String.raw`<(${PREFIXED_TAG}${slot})\b[^>]*>[\s\S]*?<\/\1>`)
const CLR_SCHEME_BLOCK_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}clrScheme)\b([^>]*)>([\s\S]*?)<\/\1>`,
)

const DEFAULT_THEME_COLORS: readonly WorkbookThemeColor[] = [
	{ slot: 'dk1', systemColor: 'windowText', lastColor: '000000' },
	{ slot: 'lt1', systemColor: 'window', lastColor: 'FFFFFF' },
	{ slot: 'dk2', rgb: '1F497D' },
	{ slot: 'lt2', rgb: 'EEECE1' },
	{ slot: 'accent1', rgb: '4F81BD' },
	{ slot: 'accent2', rgb: 'C0504D' },
	{ slot: 'accent3', rgb: '9BBB59' },
	{ slot: 'accent4', rgb: '8064A2' },
	{ slot: 'accent5', rgb: '4BACC6' },
	{ slot: 'accent6', rgb: 'F79646' },
	{ slot: 'hlink', rgb: '0000FF' },
	{ slot: 'folHlink', rgb: '800080' },
]

export function buildThemeXml(
	metadata: WorkbookThemeMetadata,
	colors: readonly WorkbookThemeColor[] = DEFAULT_THEME_COLORS,
): string {
	const themeName = escapeXml(metadata.name ?? 'Ascend Theme')
	const colorSchemeName = escapeXml(metadata.colorSchemeName ?? 'Ascend Colors')
	const majorFontLatin = escapeXml(metadata.majorFontLatin ?? 'Cambria')
	const minorFontLatin = escapeXml(metadata.minorFontLatin ?? 'Calibri')
	const themeColors = colors.length > 0 ? colors : DEFAULT_THEME_COLORS

	return `${XML_HEADER}<a:theme xmlns:a="${NS_A}" name="${themeName}">
  <a:themeElements>
    <a:clrScheme name="${colorSchemeName}">
${themeColors.map((color) => `      ${themeColorXml(color, 'a')}`).join('\n')}
    </a:clrScheme>
    <a:fontScheme name="Ascend Fonts">
      <a:majorFont>
        <a:latin typeface="${majorFontLatin}"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="${minorFontLatin}"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Ascend Formats">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`
}

export function updateThemeXml(
	xml: string,
	metadata: WorkbookThemeMetadata,
	colors: readonly WorkbookThemeColor[],
): string {
	let next = updateAttr(xml, THEME_RE, 'name', metadata.name)
	next = updateAttr(next, CLR_SCHEME_RE, 'name', metadata.colorSchemeName)
	next = updateLatinTypeface(next, 'majorFont', metadata.majorFontLatin)
	next = updateLatinTypeface(next, 'minorFont', metadata.minorFontLatin)
	for (const color of colors) next = updateThemeColorSlot(next, color)
	return next
}

export function themeXmlMatches(
	xml: string,
	metadata: WorkbookThemeMetadata,
	colors: readonly WorkbookThemeColor[],
): boolean {
	const sourceMetadata = parseThemeXml(xml)
	if (!themeMetadataMatches(sourceMetadata, metadata)) return false
	if (colors.length === 0) return true
	return themeColorsMatch(parseThemeColorsXml(xml), colors)
}

function updateLatinTypeface(
	xml: string,
	fontTag: 'majorFont' | 'minorFont',
	typeface: string | undefined,
): string {
	if (typeface === undefined) return xml
	return xml.replace(
		LATIN_RE(fontTag),
		(match, _fontTag: string, latinTag: string, attrs: string) =>
			updateLatinOpenTag(match, latinTag, attrs, typeface),
	)
}

function updateThemeColorSlot(xml: string, color: WorkbookThemeColor): string {
	const slotRe = COLOR_SLOT_RE(color.slot)
	if (slotRe.test(xml)) {
		return xml.replace(slotRe, (_match, tag: string) => themeColorXml(color, tagPrefix(tag)))
	}
	return xml.replace(CLR_SCHEME_BLOCK_RE, (_match, tag: string, attrs: string, body: string) => {
		return `<${tag}${attrs}>${body}\n      ${themeColorXml(color, tagPrefix(tag))}</${tag}>`
	})
}

function updateAttr(xml: string, tagRe: RegExp, name: string, value: string | undefined): string {
	if (value === undefined) return xml
	return xml.replace(tagRe, (_match, tag: string, attrs: string) => {
		return `<${tag}${setAttr(attrs, name, value)}>`
	})
}

function setAttr(attrs: string, name: string, value: string): string {
	const escaped = escapeXml(value)
	const attrRe = new RegExp(String.raw`\s${name}="[^"]*"`)
	if (attrRe.test(attrs)) return attrs.replace(attrRe, ` ${name}="${escaped}"`)
	return `${attrs} ${name}="${escaped}"`
}

function updateLatinOpenTag(
	match: string,
	latinTag: string,
	attrs: string,
	typeface: string,
): string {
	const tagRe = new RegExp(
		String.raw`<${latinTag}\b${escapeRegExp(attrs)}(?:\/>|>[\s\S]*?<\/${latinTag}>)`,
	)
	return match.replace(tagRe, `<${latinTag}${setAttr(attrs, 'typeface', typeface)}/>`)
}

function themeMetadataMatches(
	source: WorkbookThemeMetadata,
	target: WorkbookThemeMetadata,
): boolean {
	return (
		(source.name ?? undefined) === (target.name ?? undefined) &&
		(source.colorSchemeName ?? undefined) === (target.colorSchemeName ?? undefined) &&
		(source.majorFontLatin ?? undefined) === (target.majorFontLatin ?? undefined) &&
		(source.minorFontLatin ?? undefined) === (target.minorFontLatin ?? undefined)
	)
}

function themeColorsMatch(
	source: readonly WorkbookThemeColor[],
	target: readonly WorkbookThemeColor[],
): boolean {
	if (source.length !== target.length) return false
	const sourceBySlot = new Map(source.map((color) => [color.slot, color]))
	for (const color of target) {
		const other = sourceBySlot.get(color.slot)
		if (
			!other ||
			(other.rgb ?? undefined) !== (color.rgb ?? undefined) ||
			(other.systemColor ?? undefined) !== (color.systemColor ?? undefined) ||
			(other.lastColor ?? undefined) !== (color.lastColor ?? undefined)
		) {
			return false
		}
	}
	return true
}

function themeColorXml(color: WorkbookThemeColor, prefix: string): string {
	const slotTag = themeTagName(prefix, color.slot)
	const colorTag = color.systemColor
		? themeTagName(prefix, 'sysClr')
		: themeTagName(prefix, 'srgbClr')
	if (color.systemColor) {
		return `<${slotTag}><${colorTag} val="${escapeXml(color.systemColor)}"${color.lastColor ? ` lastClr="${escapeXml(color.lastColor)}"` : ''}/></${slotTag}>`
	}
	return `<${slotTag}><${colorTag} val="${escapeXml(color.rgb ?? '000000')}"/></${slotTag}>`
}

function themeTagName(prefix: string, localName: string): string {
	return prefix ? `${prefix}:${localName}` : localName
}

function tagPrefix(tag: string): string {
	const colon = tag.indexOf(':')
	return colon >= 0 ? tag.slice(0, colon) : ''
}

function escapeRegExp(value: string): string {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function escapeXml(value: string): string {
	return value
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&apos;')
}
