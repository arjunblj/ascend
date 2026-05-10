import type { SheetAnchorMarker, SheetImageRef } from '@ascend/core'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const XML_NAME = String.raw`[A-Za-z_][\w.-]*`
const PREFIXED_TAG = `(?:${XML_NAME}:)?`
const SHAPE_RE = new RegExp(String.raw`<(${PREFIXED_TAG}sp)\b[\s\S]*?<\/\1>`, 'g')
const CNV_PR_RE = new RegExp(String.raw`<${PREFIXED_TAG}cNvPr\b([^>]*?)(?:\/>|>)`)
const TEXT_RUN_RE = new RegExp(String.raw`<(${PREFIXED_TAG}t)\b([^>]*)>([\s\S]*?)<\/\1>`, 'g')

export interface DrawingTextUpdate {
	readonly id?: number
	readonly name?: string
	readonly text: string
}

export function buildDrawingXml(images: readonly SheetImageRef[]): string {
	const out = new ChunkedStringBuilder()
	out.push(XML_HEADER)
	out.push(`<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">`)
	for (let i = 0; i < images.length; i++) {
		const image = images[i]
		if (!image) continue
		out.push(anchorXml(image, i))
	}
	out.push('</xdr:wsDr>')
	return out.toString()
}

export function updateDrawingTextXml(xml: string, updates: readonly DrawingTextUpdate[]): string {
	if (updates.length === 0) return xml
	return xml.replace(SHAPE_RE, (shapeXml) => {
		const update = findDrawingTextUpdate(shapeXml, updates)
		return update ? replaceShapeTextRuns(shapeXml, update.text) : shapeXml
	})
}

function anchorXml(image: SheetImageRef, index: number): string {
	const picXml = pictureXml(image, index)
	const clientData = '<xdr:clientData/>'
	const anchor = image.anchor ?? {
		kind: 'oneCell',
		from: { row: 0, col: 0 },
		cx: 320000,
		cy: 240000,
	}
	switch (anchor.kind) {
		case 'oneCell':
			return `<xdr:oneCellAnchor>${markerXml('from', anchor.from)}<xdr:ext cx="${anchor.cx ?? 320000}" cy="${anchor.cy ?? 240000}"/>${picXml}${clientData}</xdr:oneCellAnchor>`
		case 'twoCell':
			return `<xdr:twoCellAnchor${anchor.editAs ? ` editAs="${escapeXml(anchor.editAs)}"` : ''}>${markerXml('from', anchor.from)}${markerXml('to', anchor.to)}${picXml}${clientData}</xdr:twoCellAnchor>`
		case 'absolute':
			return `<xdr:absoluteAnchor><xdr:pos x="${anchor.x}" y="${anchor.y}"/><xdr:ext cx="${anchor.cx ?? 320000}" cy="${anchor.cy ?? 240000}"/>${picXml}${clientData}</xdr:absoluteAnchor>`
	}
}

function findDrawingTextUpdate(
	shapeXml: string,
	updates: readonly DrawingTextUpdate[],
): DrawingTextUpdate | undefined {
	const cNvPr = shapeXml.match(CNV_PR_RE)
	if (!cNvPr) return undefined
	const attrs = cNvPr[1] ?? ''
	const id = readNumberAttr(attrs, 'id')
	const name = readXmlAttr(attrs, 'name')
	return updates.find((update) => {
		if (update.id !== undefined && id !== update.id) return false
		if (update.name !== undefined && name !== update.name) return false
		return update.id !== undefined || update.name !== undefined
	})
}

function replaceShapeTextRuns(shapeXml: string, text: string): string {
	let seenTextRun = false
	return shapeXml.replace(TEXT_RUN_RE, (_match, tag: string, attrs: string) => {
		const nextText = seenTextRun ? '' : escapeXml(text)
		seenTextRun = true
		return `<${tag}${attrs}>${nextText}</${tag}>`
	})
}

function readNumberAttr(attrs: string, name: string): number | undefined {
	const value = readXmlAttr(attrs, name)
	if (value === undefined) return undefined
	const parsed = Number(value)
	return Number.isFinite(parsed) ? parsed : undefined
}

function readXmlAttr(attrs: string, name: string): string | undefined {
	const match = attrs.match(new RegExp(String.raw`\s${name}="([^"]*)"`))
	return match?.[1]
}

function pictureXml(image: SheetImageRef, index: number): string {
	const name = escapeXml(image.name ?? `Image ${index + 1}`)
	const descr = image.description ? ` descr="${escapeXml(image.description)}"` : ''
	return `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${index + 1}" name="${name}"${descr}/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${escapeXml(image.relId)}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`
}

function markerXml(tag: 'from' | 'to', marker: SheetAnchorMarker): string {
	return `<xdr:${tag}><xdr:col>${marker.col}</xdr:col><xdr:colOff>${marker.colOff ?? 0}</xdr:colOff><xdr:row>${marker.row}</xdr:row><xdr:rowOff>${marker.rowOff ?? 0}</xdr:rowOff></xdr:${tag}>`
}

function escapeXml(value: string): string {
	return value
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&apos;')
}
