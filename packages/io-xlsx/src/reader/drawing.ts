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
