import { unzipSync } from 'fflate'

export interface XmlFingerprint {
	readonly normalized: string
	readonly tagCounts: Readonly<Record<string, number>>
}

export interface PartFingerprint {
	readonly path: string
	readonly xml: XmlFingerprint
}

export interface XlsxPackageFingerprint {
	readonly partPaths: readonly string[]
	readonly workbook?: XmlFingerprint
	readonly workbookRels?: XmlFingerprint
	readonly contentTypes?: XmlFingerprint
	readonly styles?: XmlFingerprint
	readonly sheets: readonly PartFingerprint[]
	readonly sheetRels: readonly PartFingerprint[]
}

export function fingerprintXlsx(bytes: Uint8Array): XlsxPackageFingerprint {
	const parts = unzipSync(bytes)
	const partPaths = Object.keys(parts).sort()

	return {
		partPaths,
		workbook: fingerprintPart(parts['xl/workbook.xml']),
		workbookRels: fingerprintPart(parts['xl/_rels/workbook.xml.rels']),
		contentTypes: fingerprintPart(parts['[Content_Types].xml']),
		styles: fingerprintPart(parts['xl/styles.xml']),
		sheets: collectFingerprints(parts, /^xl\/worksheets\/sheet\d+\.xml$/),
		sheetRels: collectFingerprints(parts, /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/),
	}
}

export function fingerprintXlsxPart(bytes: Uint8Array, path: string): PartFingerprint | undefined {
	const parts = unzipSync(bytes)
	const content = parts[path]
	return content ? { path, xml: fingerprintXml(decode(content)) } : undefined
}

function collectFingerprints(
	parts: Record<string, Uint8Array>,
	pattern: RegExp,
): readonly PartFingerprint[] {
	return Object.keys(parts)
		.filter((path) => pattern.test(path))
		.sort()
		.map((path) => {
			const bytes = parts[path]
			if (!bytes) {
				throw new Error(`Missing ZIP part bytes for ${path}`)
			}
			return { path, xml: fingerprintXml(decode(bytes)) }
		})
}

function fingerprintPart(bytes: Uint8Array | undefined): XmlFingerprint | undefined {
	return bytes ? fingerprintXml(decode(bytes)) : undefined
}

function decode(bytes: Uint8Array): string {
	return new TextDecoder().decode(bytes)
}

export function fingerprintXml(xml: string): XmlFingerprint {
	return {
		normalized: normalizeXml(xml),
		tagCounts: countTags(xml),
	}
}

function normalizeXml(xml: string): string {
	return xml.replace(/>\s+</g, '><').replace(/\s+/g, ' ').trim()
}

function countTags(xml: string): Readonly<Record<string, number>> {
	const counts = new Map<string, number>()
	const re = /<([A-Za-z_][\w:.-]*)(?=[\s/>])/g
	for (const match of xml.matchAll(re)) {
		const tag = match[1]
		if (!tag || tag.startsWith('?') || tag.startsWith('!')) continue
		counts.set(tag, (counts.get(tag) ?? 0) + 1)
	}
	return Object.fromEntries([...counts.entries()].sort(([a], [b]) => a.localeCompare(b)))
}
