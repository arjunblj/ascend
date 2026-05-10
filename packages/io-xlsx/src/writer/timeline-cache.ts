import type { TimelineCacheInfo, TimelineRangeInfo } from '@ascend/core'
import { escapeXml } from '../xml.ts'

const XML_NAME = String.raw`[A-Za-z_][\w.-]*`
const PREFIXED_TAG = `(?:${XML_NAME}:)?`
const ROOT_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}timelineCacheDefinition)\b([^>]*?)(\/>|>([\s\S]*?)<\/\1>)`,
)
const STATE_RE = new RegExp(String.raw`<(${PREFIXED_TAG}state)\b([^>]*?)(\/>|>([\s\S]*?)<\/\1>)`)
const SELECTION_RE = new RegExp(
	String.raw`<(${PREFIXED_TAG}selection)\b([^>]*?)(\/>|>([\s\S]*?)<\/\1>)`,
)

export function updateTimelineCacheDefinitionXml(xml: string, cache: TimelineCacheInfo): string {
	if (!cache.state?.selection) return xml
	return xml.replace(ROOT_RE, (_match, tag: string, attrs: string, tail: string, body = '') => {
		if (tail === '/>') {
			const stateTag = deriveChildTag(tag, 'state')
			return `<${tag}${attrs}>${buildStateXml(stateTag, cache)}</${tag}>`
		}
		const updatedBody = STATE_RE.test(body)
			? body.replace(
					STATE_RE,
					(
						_stateMatch: string,
						stateTag: string,
						stateAttrs: string,
						stateTail: string,
						stateBody = '',
					) => updateStateXml(stateTag, stateAttrs, stateTail, stateBody, cache),
				)
			: insertStateXml(body, deriveChildTag(tag, 'state'), cache)
		return `<${tag}${attrs}>${updatedBody}</${tag}>`
	})
}

function updateStateXml(
	tag: string,
	attrs: string,
	tail: string,
	body: string,
	cache: TimelineCacheInfo,
): string {
	const selection = cache.state?.selection
	if (!selection) return tail === '/>' ? `<${tag}${attrs}/>` : `<${tag}${attrs}>${body}</${tag}>`
	let updatedAttrs = attrs
	if (cache.state?.singleRangeFilterState !== undefined) {
		updatedAttrs = setXmlAttr(
			updatedAttrs,
			'singleRangeFilterState',
			cache.state.singleRangeFilterState ? '1' : '0',
		)
	}
	const selectionTag = deriveChildTag(tag, 'selection')
	const originalBody = tail === '/>' ? '' : body
	const updatedBody = SELECTION_RE.test(originalBody)
		? originalBody.replace(
				SELECTION_RE,
				(selectionNode, childTag: string, selectionAttrs: string) =>
					updateRangeXml(selectionNode, childTag, selectionAttrs, selection),
			)
		: `${buildRangeXml(selectionTag, selection)}${originalBody}`
	return `<${tag}${updatedAttrs}>${updatedBody}</${tag}>`
}

function insertStateXml(body: string, stateTag: string, cache: TimelineCacheInfo): string {
	const extIndex = body.search(new RegExp(String.raw`<${PREFIXED_TAG}extLst\b`))
	const stateXml = buildStateXml(stateTag, cache)
	return extIndex === -1
		? `${body}${stateXml}`
		: `${body.slice(0, extIndex)}${stateXml}${body.slice(extIndex)}`
}

function buildStateXml(tag: string, cache: TimelineCacheInfo): string {
	let attrs = ''
	if (cache.state?.singleRangeFilterState !== undefined) {
		attrs = setXmlAttr(
			attrs,
			'singleRangeFilterState',
			cache.state.singleRangeFilterState ? '1' : '0',
		)
	}
	const selection = cache.state?.selection
	const body = selection ? buildRangeXml(deriveChildTag(tag, 'selection'), selection) : ''
	return `<${tag}${attrs}>${body}</${tag}>`
}

function updateRangeXml(
	_node: string,
	tag: string,
	attrs: string,
	range: TimelineRangeInfo,
): string {
	return `<${tag}${rangeAttrs(attrs, range)}/>`
}

function buildRangeXml(tag: string, range: TimelineRangeInfo): string {
	return `<${tag}${rangeAttrs('', range)}/>`
}

function rangeAttrs(attrs: string, range: TimelineRangeInfo): string {
	return setXmlAttr(setXmlAttr(attrs, 'startDate', range.startDate), 'endDate', range.endDate)
}

function deriveChildTag(parentTag: string, localName: string): string {
	const colon = parentTag.indexOf(':')
	return colon === -1 ? localName : `${parentTag.slice(0, colon + 1)}${localName}`
}

function setXmlAttr(attrs: string, name: string, value: string): string {
	const attrText = `${name}="${escapeXml(value)}"`
	const attrPattern = new RegExp(String.raw`\s${name}="[^"]*"`)
	if (attrPattern.test(attrs)) return attrs.replace(attrPattern, ` ${attrText}`)
	return `${attrs} ${attrText}`
}
