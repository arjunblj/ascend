import type { SheetAnchorMarker, SheetImageAnchor, SheetImageRef } from '@ascend/core'
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
			(rel) => [rel.id, { ...rel, target: resolvePath(drawingPath, rel.target) }] as const,
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
