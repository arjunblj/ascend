import type {
	ActiveXControlInfo,
	CustomUiInfo,
	FormControlInfo,
	ShapeMacroInfo,
	SheetAnchorMarker,
	SheetImageAnchor,
	WorksheetControlInfo,
} from '@ascend/core'
import { attr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import type { Relationship } from './relationships.ts'
import { resolvePath } from './relationships.ts'

type ActiveXControlBuilder = {
	-readonly [K in keyof ActiveXControlInfo]?: ActiveXControlInfo[K]
}

type FormControlBuilder = {
	-readonly [K in keyof FormControlInfo]?: FormControlInfo[K]
}

type CustomUiBuilder = {
	-readonly [K in keyof CustomUiInfo]?: CustomUiInfo[K]
}

type WorksheetControlBuilder = {
	-readonly [K in keyof WorksheetControlInfo]?: WorksheetControlInfo[K]
}

export interface VmlControlInfo {
	readonly shapeId?: number
	readonly shapeName?: string
	readonly shapeSpid?: string
	readonly objectType?: string
	readonly mapOcx?: boolean
	readonly imageRelationshipId?: string
	readonly imageRelationshipType?: string
	readonly imageTarget?: string
}

type VmlControlBuilder = {
	-readonly [K in keyof VmlControlInfo]?: VmlControlInfo[K]
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
	assignText(info, 'relationshipId', relationshipIdFromNode(root))
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

export function parseCustomUiInfo(xml: string | undefined): CustomUiInfo | undefined {
	if (!xml) return undefined
	const root = firstElement(parseXml(xml), 'customUI')
	const callbacks = collectCustomUiCallbacks(xml)
	const namespaceUri = root ? attr(root, 'xmlns') : undefined
	return {
		callbackCount: callbacks.length,
		callbacks,
		...(namespaceUri ? { namespaceUri } : {}),
	}
}

export function parseDrawingShapeMacroInfos(xml: string | undefined): readonly ShapeMacroInfo[] {
	if (!xml || !xml.includes(' macro=')) return []
	const infos: ShapeMacroInfo[] = []
	const seen = new Set<string>()
	for (const match of xml.matchAll(DRAWING_MACRO_SHAPE_RE)) {
		const attrs = parseAttrs(match[2] ?? '')
		const macro = attrs.get('macro')
		if (!macro) continue
		const body = match[3] ?? ''
		const cNvPrAttrs = firstTagAttrs(body, 'cNvPr')
		const shapeId = parseOptionalNumber(cNvPrAttrs.get('id'))
		const shapeName = cNvPrAttrs.get('name')
		const parsed: ShapeMacroInfo = {
			macro,
			...(shapeId !== undefined ? { shapeId } : {}),
			...(shapeName ? { shapeName } : {}),
		}
		const key = `${macro}\u0000${parsed.shapeId ?? ''}\u0000${parsed.shapeName ?? ''}`
		if (seen.has(key)) continue
		seen.add(key)
		infos.push(parsed)
	}
	return infos
}

export function parseWorksheetControlInfos(
	xml: string | undefined,
	sheetPath: string,
	relationships: readonly Relationship[],
	vmlControls: readonly VmlControlInfo[] = [],
): readonly WorksheetControlInfo[] {
	if (!xml || !CONTROLS_CONTAINER_RE.test(xml)) return []
	const controls: WorksheetControlInfo[] = []
	const relationshipsById = new Map(relationships.map((rel) => [rel.id, rel]))
	const vmlByShapeId = new Map(
		vmlControls.flatMap((info) => (info.shapeId !== undefined ? [[info.shapeId, info]] : [])),
	)
	const vmlByName = new Map(
		vmlControls.flatMap((info) => (info.shapeName ? [[info.shapeName, info]] : [])),
	)

	for (const { attrs, body } of extractControlNodes(xml)) {
		const relationshipId = relationshipIdFromAttributes(attrs)
		const shapeId = parseOptionalNumber(attrs.get('shapeId'))
		const name = attrs.get('name')
		const controlPrAttrs = firstTagAttrs(body, 'controlPr')
		const controlPrRelationshipId = relationshipIdFromAttributes(controlPrAttrs)
		const controlPrRelationship = controlPrRelationshipId
			? relationshipsById.get(controlPrRelationshipId)
			: undefined
		const vml =
			(shapeId !== undefined ? vmlByShapeId.get(shapeId) : undefined) ??
			(name ? vmlByName.get(name) : undefined)
		const info: WorksheetControlBuilder = {}
		if (shapeId !== undefined) info.shapeId = shapeId
		assignText(info, 'name', name)
		assignText(info, 'relationshipId', relationshipId)
		assignText(info, 'controlPrRelationshipId', controlPrRelationshipId)
		assignText(info, 'controlPrRelationshipType', controlPrRelationship?.type)
		if (controlPrRelationship) {
			info.controlPrTarget =
				controlPrRelationship.targetMode === 'External'
					? controlPrRelationship.target
					: resolvePath(sheetPath, controlPrRelationship.target)
		}
		const anchor = parseWorksheetControlAnchor(body)
		if (anchor) info.anchor = anchor
		if (vml) assignVmlControlInfo(info, vml)
		const parsed = nonEmpty(info)
		if (parsed) controls.push(parsed)
	}

	return controls
}

export function parseVmlControlInfos(
	xml: string | undefined,
	vmlPath: string,
	relationships: readonly Relationship[],
): readonly VmlControlInfo[] {
	if (!xml) return []
	const relationshipsById = new Map(relationships.map((rel) => [rel.id, rel]))
	const controls: VmlControlInfo[] = []
	for (const match of xml.matchAll(VML_SHAPE_RE)) {
		const attrs = parseAttrs(match[1] ?? '')
		const body = match[2] ?? ''
		const shapeName = attrs.get('id')
		const shapeSpid = attrs.get('o:spid')
		const imageAttrs = firstTagAttrs(body, 'imagedata')
		const imageRelationshipId = relationshipRelIdFromAttributes(imageAttrs)
		const imageRelationship = imageRelationshipId
			? relationshipsById.get(imageRelationshipId)
			: undefined
		const control: VmlControlBuilder = {}
		assignText(control, 'shapeName', shapeName)
		if (shapeSpid) {
			const shapeId = parseShapeSpid(shapeSpid)
			control.shapeSpid = shapeSpid
			if (shapeId !== undefined) control.shapeId = shapeId
		}
		const objectType = firstTagAttrs(body, 'ClientData').get('ObjectType')
		assignText(control, 'objectType', objectType)
		if (hasLocalTag(body, 'MapOCX')) control.mapOcx = true
		assignText(control, 'imageRelationshipId', imageRelationshipId)
		if (imageRelationship) {
			control.imageRelationshipType = imageRelationship.type
			control.imageTarget =
				imageRelationship.targetMode === 'External'
					? imageRelationship.target
					: resolvePath(vmlPath, imageRelationship.target)
		}
		controls.push(control)
	}
	return controls
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

function assignVmlControlInfo(info: WorksheetControlBuilder, vml: VmlControlInfo): void {
	assignText(info, 'vmlShapeId', vml.shapeName)
	assignText(info, 'vmlShapeSpid', vml.shapeSpid)
	assignText(info, 'vmlObjectType', vml.objectType)
	if (vml.mapOcx !== undefined) info.vmlMapOcx = vml.mapOcx
	assignText(info, 'vmlImageRelationshipId', vml.imageRelationshipId)
	assignText(info, 'vmlImageRelationshipType', vml.imageRelationshipType)
	assignText(info, 'vmlImageTarget', vml.imageTarget)
}

const CONTROL_RE =
	/<(?:[A-Za-z_][\w.-]*:)?control\b([^/>]*?)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?control>|<(?:[A-Za-z_][\w.-]*:)?control\b([^/>]*?)\/>/g
const CONTROLS_CONTAINER_RE = /<(?:[A-Za-z_][\w.-]*:)?controls\b/
const VML_SHAPE_RE =
	/<(?:[A-Za-z_][\w.-]*:)?shape\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?shape>/g
const DRAWING_MACRO_SHAPE_RE =
	/<((?:[A-Za-z_][\w.-]*:)?(?:sp|graphicFrame|cxnSp|pic))\b([^>]*)>([\s\S]*?)<\/\1>/g
const ATTR_RE = /([A-Za-z_][\w:.-]*)=(?:"([^"]*)"|'([^']*)')/g
const CUSTOM_UI_TAG_RE = /<(?:[A-Za-z_][\w.-]*:)?[A-Za-z_][\w.-]*\b([^<>]*?)(?:\/>|>)/g
const CUSTOM_UI_CALLBACK_ATTRS = new Set([
	'getContent',
	'getDescription',
	'getEnabled',
	'getImage',
	'getItemCount',
	'getItemID',
	'getItemImage',
	'getItemLabel',
	'getItemScreentip',
	'getItemSupertip',
	'getItemWidth',
	'getKeytip',
	'getLabel',
	'getPressed',
	'getScreentip',
	'getSelectedItemID',
	'getSelectedItemIndex',
	'getShowImage',
	'getShowLabel',
	'getSize',
	'getSupertip',
	'getText',
	'getTitle',
	'getVisible',
	'loadImage',
	'onAction',
	'onChange',
	'onLoad',
])

function extractControlNodes(
	xml: string,
): Array<{ attrs: ReadonlyMap<string, string>; body: string }> {
	const controls: Array<{ attrs: ReadonlyMap<string, string>; body: string }> = []
	const seen = new Set<string>()
	for (const match of xml.matchAll(CONTROL_RE)) {
		const attrs = parseAttrs(match[1] ?? match[3] ?? '')
		const body = match[2] ?? ''
		const key = `${relationshipIdFromAttributes(attrs) ?? ''}:${attrs.get('shapeId') ?? ''}`
		const existing = controls.findIndex((control) => {
			const existingKey = `${relationshipIdFromAttributes(control.attrs) ?? ''}:${control.attrs.get('shapeId') ?? ''}`
			return existingKey === key
		})
		if (existing >= 0) {
			if (body.includes('<controlPr')) controls[existing] = { attrs, body }
			continue
		}
		if (seen.has(key)) continue
		seen.add(key)
		controls.push({ attrs, body })
	}
	return controls
}

function parseAttrs(raw: string): ReadonlyMap<string, string> {
	const attrs = new Map<string, string>()
	for (const match of raw.matchAll(ATTR_RE)) {
		const name = match[1]
		const value = match[2] ?? match[3]
		if (name && value !== undefined) attrs.set(name, value)
	}
	return attrs
}

function relationshipIdFromNode(node: XmlNode): string | undefined {
	const direct = attr(node, 'r:id') ?? attr(node, 'id')
	if (direct !== undefined) return direct
	for (const [name, value] of Object.entries(node)) {
		if (name.startsWith('@_') && name.endsWith(':id')) return String(value)
	}
	return undefined
}

function relationshipIdFromAttributes(attrs: ReadonlyMap<string, string>): string | undefined {
	const direct = attrs.get('r:id') ?? attrs.get('id')
	if (direct !== undefined) return direct
	for (const [name, value] of attrs) {
		if (name.endsWith(':id')) return value
	}
	return undefined
}

function relationshipRelIdFromAttributes(attrs: ReadonlyMap<string, string>): string | undefined {
	const direct = attrs.get('o:relid') ?? attrs.get('relid')
	if (direct !== undefined) return direct
	for (const [name, value] of attrs) {
		if (name.endsWith(':relid')) return value
	}
	return undefined
}

function firstTagAttrs(xml: string, localName: string): ReadonlyMap<string, string> {
	const re = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b([^>]*)>`, 'i')
	return parseAttrs(xml.match(re)?.[1] ?? '')
}

function collectCustomUiCallbacks(xml: string): CustomUiInfo['callbacks'] {
	const callbacks: Array<{ attribute: string; macro: string }> = []
	const seen = new Set<string>()
	for (const match of xml.matchAll(CUSTOM_UI_TAG_RE)) {
		const attrs = parseAttrs(match[1] ?? '')
		for (const [attribute, macro] of attrs) {
			if (!CUSTOM_UI_CALLBACK_ATTRS.has(attribute) || macro.length === 0) continue
			const key = `${attribute}\u0000${macro}`
			if (seen.has(key)) continue
			seen.add(key)
			callbacks.push({ attribute, macro })
		}
	}
	return callbacks
}

function parseWorksheetControlAnchor(xml: string): SheetImageAnchor | undefined {
	const body = localTagBody(xml, 'anchor')
	if (!body) return undefined
	const from = parseMarker(localTagBody(body, 'from'))
	const to = parseMarker(localTagBody(body, 'to'))
	if (!from || !to) return undefined
	return { kind: 'twoCell', from, to }
}

function parseMarker(xml: string | undefined): SheetAnchorMarker | null {
	if (!xml) return null
	const col = readChildNumber(xml, 'col')
	const row = readChildNumber(xml, 'row')
	if (col === undefined || row === undefined) return null
	const marker: { col: number; row: number; colOff?: number; rowOff?: number } = { col, row }
	const colOff = readChildNumber(xml, 'colOff')
	if (colOff !== undefined) marker.colOff = colOff
	const rowOff = readChildNumber(xml, 'rowOff')
	if (rowOff !== undefined) marker.rowOff = rowOff
	return marker
}

function readChildNumber(xml: string, localName: string): number | undefined {
	const text = localTagBody(xml, localName)?.trim()
	if (!text) return undefined
	const parsed = Number(text)
	return Number.isNaN(parsed) ? undefined : parsed
}

function localTagBody(xml: string, localName: string): string | undefined {
	const re = new RegExp(
		`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`,
		'i',
	)
	return xml.match(re)?.[1]
}

function hasLocalTag(xml: string, localName: string): boolean {
	return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b`, 'i').test(xml)
}

function parseShapeSpid(value: string): number | undefined {
	const match = value.match(/_s(\d+)$/)
	return parseOptionalNumber(match?.[1])
}

function parseOptionalNumber(value: string | undefined): number | undefined {
	if (value === undefined || value.length === 0) return undefined
	const parsed = Number(value)
	return Number.isNaN(parsed) ? undefined : parsed
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

function assignText<
	T extends
		| ActiveXControlBuilder
		| FormControlBuilder
		| CustomUiBuilder
		| WorksheetControlBuilder
		| VmlControlBuilder,
	K extends keyof T,
>(info: T, key: K, value: string | undefined): void {
	if (value !== undefined && value.length > 0) info[key] = value as T[K]
}

function nonEmpty<
	T extends ActiveXControlBuilder | FormControlBuilder | CustomUiBuilder | WorksheetControlBuilder,
>(info: T): T | undefined {
	return Object.keys(info).length > 0 ? info : undefined
}
