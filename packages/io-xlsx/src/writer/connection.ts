import type { WorkbookConnectionPartInfo } from '@ascend/core'
import { readNumberXmlAttr, readXmlAttr, setXmlAttr } from './xml-attrs.ts'

const XML_NAME = String.raw`[A-Za-z_][\w.-]*`
const PREFIXED_TAG = `(?:${XML_NAME}:)?`
const CONNECTION_RE = new RegExp(String.raw`<(${PREFIXED_TAG}connection)\b([^>]*?)(\/>|>)`, 'g')
const QUERY_TABLE_RE = new RegExp(String.raw`<(${PREFIXED_TAG}queryTable)\b([^>]*?)(\/>|>)`)

export function updateConnectionPartXml(
	xml: string,
	parts: readonly WorkbookConnectionPartInfo[],
): string {
	const editable = parts.filter((part) => part.kind === 'connection' || part.kind === 'queryTable')
	if (editable.length === 0) return xml
	const queryTable = editable.find((part) => part.kind === 'queryTable')
	if (queryTable) {
		return xml.replace(QUERY_TABLE_RE, (_node, tag: string, attrs: string, tail: string) => {
			const updated = updateConnectionAttrs(attrs, queryTable)
			return `<${tag}${updated}${tail}`
		})
	}
	return xml.replace(CONNECTION_RE, (node, tag: string, attrs: string, tail: string) => {
		const part = matchConnectionNode(attrs, editable)
		if (!part) return node
		const updated = updateConnectionAttrs(attrs, part)
		return `<${tag}${updated}${tail}`
	})
}

function matchConnectionNode(
	attrs: string,
	parts: readonly WorkbookConnectionPartInfo[],
): WorkbookConnectionPartInfo | undefined {
	if (parts.length === 1) return parts[0]
	const id = readNumberAttr(attrs, 'id')
	const name = readXmlAttr(attrs, 'name')
	return parts.find((part) => {
		if (part.connectionId !== undefined && id !== undefined) return part.connectionId === id
		if (part.name !== undefined && name !== undefined) return part.name === name
		return false
	})
}

function updateConnectionAttrs(attrs: string, part: WorkbookConnectionPartInfo): string {
	let updated = attrs
	updated = setOptionalBoolAttr(updated, 'refreshOnLoad', part.refreshOnLoad)
	if (part.kind === 'queryTable') {
		updated = setOptionalBoolAttr(
			updated,
			'removeDataOnSave',
			part.saveData === undefined ? undefined : !part.saveData,
		)
	} else {
		updated = setOptionalBoolAttr(updated, 'saveData', part.saveData)
	}
	updated = setOptionalNumberAttr(updated, 'refreshedVersion', part.refreshedVersion)
	return updated
}

function readNumberAttr(attrs: string, name: string): number | undefined {
	return readNumberXmlAttr(attrs, name)
}

function setOptionalBoolAttr(attrs: string, name: string, value: boolean | undefined): string {
	if (value === undefined) return attrs
	return setXmlAttr(attrs, name, value ? '1' : '0')
}

function setOptionalNumberAttr(attrs: string, name: string, value: number | undefined): string {
	if (value === undefined) return attrs
	return setXmlAttr(attrs, name, String(value))
}
