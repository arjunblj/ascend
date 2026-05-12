import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'

export interface PackageFamilyCounts {
	readonly charts: number
	readonly structuredCharts: number
	readonly drawings: number
	readonly vml: number
	readonly media: number
	readonly tables: number
	readonly comments: number
	readonly threadedComments: number
	readonly pivotTables: number
	readonly pivotCaches: number
	readonly slicers: number
	readonly slicerCaches: number
	readonly timelines: number
	readonly timelineCaches: number
	readonly macros: number
	readonly customXml: number
	readonly externalLinks: number
	readonly connections: number
	readonly calcChain: number
}

export interface OoxmlPackageSummary {
	readonly workbookContentType?: string
	readonly partCount: number
	readonly families: PackageFamilyCounts
}

export function summarizeOoxmlPackage(bytes: Uint8Array): OoxmlPackageSummary {
	const archive = extractZip(bytes)
	const paths = [...archive.entries()]
		.map((entry) => entry.path)
		.filter((path) => !path.endsWith('/'))
	const contentTypes = archive.readText('[Content_Types].xml') ?? ''
	return {
		workbookContentType: readWorkbookContentType(contentTypes),
		partCount: paths.length,
		families: summarizePackageFamilies(paths),
	}
}

export function summarizePackageFamilies(paths: readonly string[]): PackageFamilyCounts {
	return {
		charts: countPaths(paths, /^xl\/(charts|chartEx)\//),
		structuredCharts: countPaths(paths, /^xl\/charts\/chart\d+\.xml$/i),
		drawings: countPaths(paths, /^xl\/drawings\//),
		vml: countPaths(paths, /^xl\/drawings\/.*\.vml$/),
		media: countPaths(paths, /^xl\/media\//),
		tables: countPaths(paths, /^xl\/tables\/(?!_rels\/)[^/]+\.xml$/i),
		comments: countPaths(paths, /^xl\/comments\d+\.xml$/),
		threadedComments: countPaths(paths, /^xl\/threadedComments\//),
		pivotTables: countPaths(paths, /^xl\/pivotTables\/pivotTable\d+\.xml$/i),
		pivotCaches: countPaths(paths, /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/i),
		slicers: countPaths(paths, /^xl\/slicers\/(?!_rels\/)[^/]+\.xml$/i),
		slicerCaches: countPaths(paths, /^xl\/slicerCaches\/(?!_rels\/)[^/]+\.xml$/i),
		timelines: countPaths(paths, /^xl\/timelines\/(?!_rels\/)[^/]+\.xml$/i),
		timelineCaches: countPaths(paths, /^xl\/timelineCaches\/(?!_rels\/)[^/]+\.xml$/i),
		macros: countPaths(paths, /^xl\/vbaProject/i),
		customXml: countPaths(paths, /^customXml\//),
		externalLinks: countPaths(paths, /^xl\/externalLinks\//),
		connections: countPaths(paths, /^xl\/connections\.xml$/),
		calcChain: countPaths(paths, /^xl\/calcChain\.xml$/),
	}
}

function countPaths(paths: readonly string[], pattern: RegExp): number {
	return paths.filter((path) => pattern.test(path)).length
}

function readWorkbookContentType(xml: string): string | undefined {
	const match =
		/<Override\s+PartName="\/xl\/workbook\.xml"\s+ContentType="([^"]+)"/.exec(xml) ??
		/<Override\s+ContentType="([^"]+)"\s+PartName="\/xl\/workbook\.xml"/.exec(xml)
	return match?.[1]
}
