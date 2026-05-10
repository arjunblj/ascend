import type { ActiveXControlInfo, FormControlInfo } from '@ascend/core'
import { attr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import type { Relationship } from './relationships.ts'

type ActiveXControlBuilder = {
	-readonly [K in keyof ActiveXControlInfo]?: ActiveXControlInfo[K]
}

type FormControlBuilder = {
	-readonly [K in keyof FormControlInfo]?: FormControlInfo[K]
}

export function parseActiveXControlInfo(
	xml: string | undefined,
	relationships: readonly Relationship[],
): ActiveXControlInfo | undefined {
	const info: ActiveXControlBuilder = {}
	assignActiveXBinaryInfo(info, relationships)
	if (!xml) return nonEmpty(info)

	const root = firstElement(parseXml(xml), 'ocx')
	if (!root) return nonEmpty(info)

	assignText(info, 'classId', attr(root, 'ax:classid') ?? attr(root, 'classid'))
	assignText(info, 'persistence', attr(root, 'ax:persistence') ?? attr(root, 'persistence'))
	assignText(info, 'relationshipId', attr(root, 'r:id') ?? attr(root, 'id'))
	return nonEmpty(info)
}

export function parseFormControlInfo(xml: string | undefined): FormControlInfo | undefined {
	if (!xml) return undefined
	const root = firstElement(parseXml(xml), 'formControlPr')
	if (!root) return undefined
	const linkedCell = attr(root, 'linkedCell') ?? attr(root, 'fmlaLink')
	const listFillRange = attr(root, 'listFillRange') ?? attr(root, 'fmlaRange')
	const info: FormControlBuilder = {}
	assignText(info, 'objectType', attr(root, 'objectType'))
	assignText(info, 'macro', attr(root, 'macro'))
	assignText(info, 'linkedCell', linkedCell)
	assignText(info, 'listFillRange', listFillRange)
	assignText(info, 'checked', attr(root, 'checked'))
	const dropLines = numAttr(root, 'dropLines')
	if (dropLines !== undefined) info.dropLines = dropLines
	return nonEmpty(info)
}

function assignActiveXBinaryInfo(
	info: ActiveXControlBuilder,
	relationships: readonly Relationship[],
): void {
	const binary = relationships.find((rel) =>
		rel.type.toLowerCase().includes('activexcontrolbinary'),
	)
	if (!binary) return
	info.binaryRelationshipId = binary.id
	info.binaryTarget = binary.target
}

function firstElement(doc: XmlNode, localName: string): XmlNode | undefined {
	for (const [key, value] of Object.entries(doc)) {
		if (localPart(key) !== localName || typeof value !== 'object' || value === null) continue
		return value as XmlNode
	}
	return undefined
}

function localPart(qualifiedName: string): string {
	const colon = qualifiedName.indexOf(':')
	return colon === -1 ? qualifiedName : qualifiedName.slice(colon + 1)
}

function assignText<T extends ActiveXControlBuilder | FormControlBuilder, K extends keyof T>(
	info: T,
	key: K,
	value: string | undefined,
): void {
	if (value !== undefined && value.length > 0) info[key] = value as T[K]
}

function nonEmpty<T extends ActiveXControlBuilder | FormControlBuilder>(info: T): T | undefined {
	return Object.keys(info).length > 0 ? info : undefined
}
