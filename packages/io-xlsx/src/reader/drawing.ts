import type {
	SheetAnchorMarker,
	SheetDrawingObjectKind,
	SheetDrawingObjectRef,
	SheetImageAnchor,
	SheetImageRef,
} from '@ascend/core'
import { asArray, attr, numAttr, parseXml, type XmlNode } from '../xml.ts'
import type { Relationship } from './relationships.ts'
import { resolvePath } from './relationships.ts'

export function parseDrawingImageRefs(
	drawingXml: string,
	drawingPath: string,
	relationships: readonly Relationship[],
): readonly SheetImageRef[] {
	const doc = parseXml(drawingXml)
	const wsDr = (doc['xdr:wsDr'] ?? doc.wsDr) as XmlNode | undefined
	if (!wsDr) return []

	const imageRels = new Map(
		relationships.map(
			(rel) => [rel.id, { ...rel, target: resolveRelationshipTarget(drawingPath, rel) }] as const,
		),
	)
	const refs: SheetImageRef[] = []

	for (const anchor of asArray<XmlNode>(
		wsDr['xdr:oneCellAnchor'] as XmlNode | XmlNode[] | undefined,
	)) {
		const parsed = parseAnchoredImage(anchor, drawingPath, imageRels, 'oneCell')
		if (parsed) refs.push(parsed)
	}
	for (const anchor of asArray<XmlNode>(
		wsDr['xdr:twoCellAnchor'] as XmlNode | XmlNode[] | undefined,
	)) {
		const parsed = parseAnchoredImage(anchor, drawingPath, imageRels, 'twoCell')
		if (parsed) refs.push(parsed)
	}
	for (const anchor of asArray<XmlNode>(
		wsDr['xdr:absoluteAnchor'] as XmlNode | XmlNode[] | undefined,
	)) {
		const parsed = parseAnchoredImage(anchor, drawingPath, imageRels, 'absolute')
		if (parsed) refs.push(parsed)
	}

	return refs
}

export function parseDrawingObjectRefs(
	drawingXml: string,
	drawingPath: string,
	relationships: readonly Relationship[] = [],
): readonly SheetDrawingObjectRef[] {
	const doc = parseXml(drawingXml)
	const wsDr = (doc['xdr:wsDr'] ?? doc.wsDr) as XmlNode | undefined
	if (!wsDr) return []

	const relsById = new Map(
		relationships.map(
			(rel) => [rel.id, { ...rel, target: resolveRelationshipTarget(drawingPath, rel) }] as const,
		),
	)
	const refs: SheetDrawingObjectRef[] = []
	for (const { node, kind } of iterAnchors(wsDr)) {
		for (const parsed of parseAnchoredDrawingObjects(node, drawingPath, kind, relsById)) {
			refs.push(parsed)
		}
	}
	return refs
}

export function parseVmlDrawingObjectRefs(
	vmlXml: string,
	vmlPath: string,
	relationships: readonly Relationship[] = [],
): readonly SheetDrawingObjectRef[] {
	if (!vmlXml.includes('<v:shape')) return []
	const relsById = new Map(
		relationships.map(
			(rel) => [rel.id, { ...rel, target: resolveRelationshipTarget(vmlPath, rel) }] as const,
		),
	)
	const refs: SheetDrawingObjectRef[] = []
	for (const match of vmlXml.matchAll(VML_SHAPE_RE)) {
		const rawAttrs = match[1] ?? ''
		const body = match[2] ?? ''
		const objectType = firstLocalTagAttrs(body, 'ClientData').get('ObjectType')
		if (objectType === 'Note') continue
		const ref = parseVmlDrawingObject(rawAttrs, body, vmlPath, objectType, relsById)
		if (ref) refs.push(ref)
	}
	return refs
}

function iterAnchors(
	wsDr: XmlNode,
): Array<{ node: XmlNode; kind: 'oneCell' | 'twoCell' | 'absolute' }> {
	return [
		...asArray<XmlNode>(wsDr['xdr:oneCellAnchor'] as XmlNode | XmlNode[] | undefined).map(
			(node) => ({ node, kind: 'oneCell' as const }),
		),
		...asArray<XmlNode>(wsDr['xdr:twoCellAnchor'] as XmlNode | XmlNode[] | undefined).map(
			(node) => ({ node, kind: 'twoCell' as const }),
		),
		...asArray<XmlNode>(wsDr['xdr:absoluteAnchor'] as XmlNode | XmlNode[] | undefined).map(
			(node) => ({ node, kind: 'absolute' as const }),
		),
	]
}

function parseAnchoredDrawingObjects(
	anchorNode: XmlNode,
	drawingPartPath: string,
	anchorKind: 'oneCell' | 'twoCell' | 'absolute',
	relationships: ReadonlyMap<string, Relationship & { target: string }>,
): SheetDrawingObjectRef[] {
	const anchor =
		anchorKind === 'oneCell'
			? parseOneCellAnchor(anchorNode)
			: anchorKind === 'twoCell'
				? parseTwoCellAnchor(anchorNode)
				: parseAbsoluteAnchor(anchorNode)
	const refs: SheetDrawingObjectRef[] = []
	for (const { node, kind, nonVisualNames } of drawingObjectCandidates(anchorNode)) {
		const ref = parseDrawingObject(
			node,
			drawingPartPath,
			kind,
			nonVisualNames,
			anchor,
			relationships,
		)
		if (ref) refs.push(ref)
	}
	return refs
}

function drawingObjectCandidates(anchorNode: XmlNode): Array<{
	node: XmlNode
	kind: SheetDrawingObjectKind
	nonVisualNames: readonly string[]
}> {
	return [
		...asArray<XmlNode>(anchorNode['xdr:sp'] as XmlNode | XmlNode[] | undefined).map((node) => ({
			node,
			kind: hasTextBoxContent(node) ? ('textBox' as const) : ('shape' as const),
			nonVisualNames: ['xdr:nvSpPr', 'nvSpPr'],
		})),
		...asArray<XmlNode>(anchorNode['xdr:graphicFrame'] as XmlNode | XmlNode[] | undefined).map(
			(node) => ({
				node,
				kind: 'graphicFrame' as const,
				nonVisualNames: ['xdr:nvGraphicFramePr', 'nvGraphicFramePr'],
			}),
		),
		...asArray<XmlNode>(anchorNode['xdr:cxnSp'] as XmlNode | XmlNode[] | undefined).map((node) => ({
			node,
			kind: 'connector' as const,
			nonVisualNames: ['xdr:nvCxnSpPr', 'nvCxnSpPr'],
		})),
		...asArray<XmlNode>(anchorNode['xdr:grpSp'] as XmlNode | XmlNode[] | undefined).map((node) => ({
			node,
			kind: 'groupShape' as const,
			nonVisualNames: ['xdr:nvGrpSpPr', 'nvGrpSpPr'],
		})),
	]
}

function parseDrawingObject(
	node: XmlNode,
	drawingPartPath: string,
	kind: SheetDrawingObjectKind,
	nonVisualNames: readonly string[],
	anchor: SheetImageAnchor | null,
	relationships: ReadonlyMap<string, Relationship & { target: string }>,
): SheetDrawingObjectRef | null {
	const cNvPr = findNonVisualProps(node, nonVisualNames)
	const id = cNvPr ? numAttr(cNvPr, 'id') : undefined
	const name = cNvPr ? attr(cNvPr, 'name') : undefined
	const description = cNvPr ? attr(cNvPr, 'descr') : undefined
	const text = kind === 'textBox' ? extractDrawingText(node) : undefined
	const relIds = collectRelationshipIds(node)
	const relationshipRefs = relIds.flatMap((relId) => {
		const rel = relationships.get(relId)
		return rel
			? [
					{
						id: rel.id,
						type: rel.type,
						target: rel.target,
						...(rel.targetMode ? { targetMode: rel.targetMode } : {}),
					},
				]
			: []
	})
	return {
		drawingPartPath,
		source: 'drawingml',
		kind,
		...(anchor ? { anchor } : {}),
		...(id !== undefined ? { id } : {}),
		...(name ? { name } : {}),
		...(description ? { description } : {}),
		...(text ? { text } : {}),
		...(relIds.length > 0 ? { relIds } : {}),
		...(relationshipRefs.length > 0 ? { relationshipRefs } : {}),
	}
}

function parseVmlDrawingObject(
	rawAttrs: string,
	body: string,
	vmlPath: string,
	objectType: string | undefined,
	relationships: ReadonlyMap<string, Relationship & { target: string }>,
): SheetDrawingObjectRef | null {
	const attrs = parseRawAttrs(rawAttrs)
	const vmlShapeId = attrs.get('id')
	const shapeSpid = attrs.get('o:spid') ?? attrs.get('spid')
	const id = parseVmlShapeId(shapeSpid ?? vmlShapeId)
	const style = attrs.get('style')
	const text = extractVmlTextBoxText(body)
	const relIds = collectVmlRelationshipIds(body)
	const relationshipRefs = relIds.flatMap((relId) => {
		const rel = relationships.get(relId)
		return rel
			? [
					{
						id: rel.id,
						type: rel.type,
						target: rel.target,
						...(rel.targetMode ? { targetMode: rel.targetMode } : {}),
					},
				]
			: []
	})
	const anchor = parseVmlClientAnchor(body)
	const visible = parseVmlVisibility(body, style)
	const kind: SheetDrawingObjectKind = text ? 'textBox' : 'shape'
	if (!vmlShapeId && id === undefined && !text && !objectType && relIds.length === 0) return null
	return {
		drawingPartPath: vmlPath,
		source: 'vml',
		kind,
		...(anchor ? { anchor } : {}),
		...(id !== undefined ? { id } : {}),
		...(vmlShapeId ? { name: vmlShapeId, vmlShapeId } : {}),
		...(text ? { text } : {}),
		...(style ? { style } : {}),
		...(objectType ? { vmlObjectType: objectType } : {}),
		...(visible !== undefined ? { visible } : {}),
		...(relIds.length > 0 ? { relIds } : {}),
		...(relationshipRefs.length > 0 ? { relationshipRefs } : {}),
	}
}

function findNonVisualProps(node: XmlNode, names: readonly string[]): XmlNode | undefined {
	for (const name of names) {
		const nonVisual = node[name] as XmlNode | undefined
		const cNvPr = (nonVisual?.['xdr:cNvPr'] ?? nonVisual?.cNvPr) as XmlNode | undefined
		if (cNvPr) return cNvPr
	}
	return undefined
}

function hasTextBoxContent(node: XmlNode): boolean {
	return Boolean(node['xdr:txBody'] ?? node.txBody ?? node['xdr:txbx'] ?? node.txbx)
}

function extractDrawingText(node: XmlNode): string | undefined {
	const chunks: string[] = []
	collectTextRuns(node, chunks)
	const text = chunks.join('')
	return text.length > 0 ? text : undefined
}

function collectTextRuns(value: unknown, chunks: string[]): void {
	if (typeof value === 'string' || typeof value === 'number') return
	if (Array.isArray(value)) {
		for (const item of value) collectTextRuns(item, chunks)
		return
	}
	if (!value || typeof value !== 'object') return
	const node = value as XmlNode
	for (const [key, child] of Object.entries(node)) {
		if (key === 'a:t' || key === 't') {
			if (typeof child === 'string' || typeof child === 'number') chunks.push(String(child))
			else collectTextRuns(child, chunks)
			continue
		}
		collectTextRuns(child, chunks)
	}
}

function collectRelationshipIds(node: XmlNode): readonly string[] {
	const relIds = new Set<string>()
	visitXmlNodes(node, (current) => {
		for (const [key, value] of Object.entries(current)) {
			if (!key.startsWith('@_')) continue
			const local = key.slice(2)
			if (
				(local.startsWith('r:') || local === 'embed' || local === 'link') &&
				typeof value === 'string'
			) {
				relIds.add(value)
			}
		}
	})
	return [...relIds]
}

const VML_SHAPE_RE = /<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/gi
const RAW_ATTR_RE = /([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

function parseRawAttrs(rawAttrs: string): ReadonlyMap<string, string> {
	const attrs = new Map<string, string>()
	RAW_ATTR_RE.lastIndex = 0
	for (const match of rawAttrs.matchAll(RAW_ATTR_RE)) {
		const name = match[1]
		const value = match[2] ?? match[3]
		if (name && value !== undefined) attrs.set(name, decodeXmlEntities(value))
	}
	return attrs
}

function firstLocalTagAttrs(xml: string, localName: string): ReadonlyMap<string, string> {
	const re = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b([^>]*)>`, 'i')
	return parseRawAttrs(re.exec(xml)?.[1] ?? '')
}

function localTagBody(xml: string, localName: string): string | undefined {
	const re = new RegExp(
		`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`,
		'i',
	)
	return re.exec(xml)?.[1]
}

function extractVmlTextBoxText(xml: string): string | undefined {
	const body = localTagBody(xml, 'textbox')
	if (!body) return undefined
	const text = decodeXmlEntities(
		body
			.replace(/<br\s*\/?>/gi, ' ')
			.replace(/<[^>]+>/g, ' ')
			.replace(/\s+/g, ' ')
			.trim(),
	)
	return text.length > 0 ? text : undefined
}

function collectVmlRelationshipIds(xml: string): readonly string[] {
	const relIds = new Set<string>()
	RAW_ATTR_RE.lastIndex = 0
	for (const match of xml.matchAll(RAW_ATTR_RE)) {
		const name = match[1]
		const value = match[2] ?? match[3]
		if (!name || !value) continue
		if (name === 'r:id' || name.endsWith(':relid') || name === 'relid') relIds.add(value)
	}
	return [...relIds]
}

function parseVmlClientAnchor(xml: string): SheetImageAnchor | undefined {
	const anchorText = localTagBody(xml, 'Anchor')
	if (!anchorText) return undefined
	const values = anchorText
		.split(',')
		.map((part) => Number(part.trim()))
		.filter((value) => Number.isInteger(value))
	if (values.length !== 8) return undefined
	return {
		kind: 'twoCell',
		from: {
			col: values[0] ?? 0,
			colOff: values[1] ?? 0,
			row: values[2] ?? 0,
			rowOff: values[3] ?? 0,
		},
		to: {
			col: values[4] ?? 0,
			colOff: values[5] ?? 0,
			row: values[6] ?? 0,
			rowOff: values[7] ?? 0,
		},
	}
}

function parseVmlVisibility(xml: string, style: string | undefined): boolean | undefined {
	if (/<(?:[A-Za-z_][\w.-]*:)?Visible\b/i.test(xml)) return true
	if (!style) return undefined
	const match = /(?:^|;)\s*visibility\s*:\s*(visible|hidden)\b/i.exec(style)
	if (!match) return undefined
	return match[1]?.toLowerCase() === 'visible'
}

function parseVmlShapeId(value: string | undefined): number | undefined {
	if (!value) return undefined
	const match = /(?:^|_s)(\d+)$/.exec(value)
	if (!match) return undefined
	const parsed = Number(match[1])
	return Number.isInteger(parsed) ? parsed : undefined
}

function decodeXmlEntities(value: string): string {
	return value.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (entity, body: string) => {
		switch (body) {
			case 'amp':
				return '&'
			case 'lt':
				return '<'
			case 'gt':
				return '>'
			case 'quot':
				return '"'
			case 'apos':
				return "'"
			default: {
				const code = body.toLowerCase().startsWith('#x')
					? Number.parseInt(body.slice(2), 16)
					: Number.parseInt(body.slice(1), 10)
				return Number.isFinite(code) ? String.fromCodePoint(code) : entity
			}
		}
	})
}

function visitXmlNodes(value: unknown, fn: (node: XmlNode) => void): void {
	if (Array.isArray(value)) {
		for (const item of value) visitXmlNodes(item, fn)
		return
	}
	if (!value || typeof value !== 'object') return
	const node = value as XmlNode
	fn(node)
	for (const child of Object.values(node)) visitXmlNodes(child, fn)
}

function parseAnchoredImage(
	anchorNode: XmlNode,
	drawingPartPath: string,
	relationships: ReadonlyMap<string, Relationship & { target: string }>,
	kind: 'oneCell' | 'twoCell' | 'absolute',
): SheetImageRef | null {
	const pic = (anchorNode['xdr:pic'] ?? anchorNode.pic) as XmlNode | undefined
	if (!pic) return null
	const blipFill = (pic['xdr:blipFill'] ?? pic.blipFill) as XmlNode | undefined
	const blip = (blipFill?.['a:blip'] ?? blipFill?.blip) as XmlNode | undefined
	const relId = blip ? (attr(blip, 'r:embed') ?? attr(blip, 'embed')) : undefined
	if (!relId) return null
	const relationship = relationships.get(relId)
	if (!relationship) return null

	const imageRef: {
		drawingPartPath: string
		relId: string
		targetPath: string
		anchor?: SheetImageAnchor
		name?: string
		description?: string
	} = {
		drawingPartPath,
		relId,
		targetPath: relationship.target,
	}

	const nonVisual = (pic['xdr:nvPicPr'] ?? pic.nvPicPr) as XmlNode | undefined
	const cNvPr = (nonVisual?.['xdr:cNvPr'] ?? nonVisual?.cNvPr) as XmlNode | undefined
	const name = cNvPr ? attr(cNvPr, 'name') : undefined
	if (name) imageRef.name = name
	const description = cNvPr ? attr(cNvPr, 'descr') : undefined
	if (description) imageRef.description = description

	const parsedAnchor =
		kind === 'oneCell'
			? parseOneCellAnchor(anchorNode)
			: kind === 'twoCell'
				? parseTwoCellAnchor(anchorNode)
				: parseAbsoluteAnchor(anchorNode)
	if (parsedAnchor) imageRef.anchor = parsedAnchor

	return imageRef as SheetImageRef
}

function resolveRelationshipTarget(drawingPath: string, relationship: Relationship): string {
	return relationship.targetMode === 'External'
		? relationship.target
		: resolvePath(drawingPath, relationship.target)
}

function parseOneCellAnchor(node: XmlNode): SheetImageAnchor | null {
	const from = parseMarker((node['xdr:from'] ?? node.from) as XmlNode | undefined)
	if (!from) return null
	const ext = (node['xdr:ext'] ?? node.ext) as XmlNode | undefined
	const cx = ext ? numAttr(ext, 'cx') : undefined
	const cy = ext ? numAttr(ext, 'cy') : undefined
	return {
		kind: 'oneCell',
		from,
		...(cx !== undefined ? { cx } : {}),
		...(cy !== undefined ? { cy } : {}),
	}
}

function parseTwoCellAnchor(node: XmlNode): SheetImageAnchor | null {
	const from = parseMarker((node['xdr:from'] ?? node.from) as XmlNode | undefined)
	const to = parseMarker((node['xdr:to'] ?? node.to) as XmlNode | undefined)
	if (!from || !to) return null
	const editAs = attr(node, 'editAs')
	return editAs ? { kind: 'twoCell', from, to, editAs } : { kind: 'twoCell', from, to }
}

function parseAbsoluteAnchor(node: XmlNode): SheetImageAnchor | null {
	const pos = (node['xdr:pos'] ?? node.pos) as XmlNode | undefined
	const ext = (node['xdr:ext'] ?? node.ext) as XmlNode | undefined
	const x = pos ? numAttr(pos, 'x') : undefined
	const y = pos ? numAttr(pos, 'y') : undefined
	if (x === undefined || y === undefined) return null
	const cx = ext ? numAttr(ext, 'cx') : undefined
	const cy = ext ? numAttr(ext, 'cy') : undefined
	return {
		kind: 'absolute',
		x,
		y,
		...(cx !== undefined ? { cx } : {}),
		...(cy !== undefined ? { cy } : {}),
	}
}

function parseMarker(node: XmlNode | undefined): SheetAnchorMarker | null {
	if (!node) return null
	const col = readChildNumber(node, 'xdr:col', 'col')
	const row = readChildNumber(node, 'xdr:row', 'row')
	if (col === undefined || row === undefined) return null
	const marker: {
		col: number
		row: number
		colOff?: number
		rowOff?: number
	} = { col, row }
	const colOff = readChildNumber(node, 'xdr:colOff', 'colOff')
	if (colOff !== undefined) marker.colOff = colOff
	const rowOff = readChildNumber(node, 'xdr:rowOff', 'rowOff')
	if (rowOff !== undefined) marker.rowOff = rowOff
	return marker as SheetAnchorMarker
}

function readChildNumber(node: XmlNode, namespaced: string, local: string): number | undefined {
	const child = node[namespaced] ?? node[local]
	if (typeof child === 'number') return child
	if (typeof child === 'string') {
		const parsed = Number(child)
		return Number.isNaN(parsed) ? undefined : parsed
	}
	return undefined
}
