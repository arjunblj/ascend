import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CorpusManifestEntry } from '../../corpus/manifest.ts'
import { inspectOoxmlPackageFeatures } from '../../corpus/ooxml-feature-probe.ts'

const EXCELFORGE_SOURCE = 'ExcelForge test workbook'
const EXCELFORGE_BASE_URL =
	'https://raw.githubusercontent.com/node-projects/excelForge/master/src/test'

const FIXTURES = [
	{
		file: 'Book_1_unknown_part.xlsx',
		sourcePath: 'Book%201.xlsx',
		sourceSha256: '9c5426fa71ff68cc7e40e19e02b5992daf91da5754ef643d2db2f89bd70bb122',
		packageManifestSha256: 'cae1feec581eed864255cff45fa23a7e2c085cb0f2c2628d1a0187fc39de3ef7',
	},
] as const

export async function loadManifest(): Promise<CorpusManifestEntry[]> {
	const root = dirname(fileURLToPath(import.meta.url))
	const entries: CorpusManifestEntry[] = []
	for (const fixture of FIXTURES) entries.push(await buildEntry(root, fixture))
	return entries
}

async function buildEntry(
	root: string,
	fixture: (typeof FIXTURES)[number],
): Promise<CorpusManifestEntry> {
	const bytes = new Uint8Array(await readFile(join(root, fixture.file)))
	const probe = inspectOoxmlPackageFeatures(bytes)
	const features = { ...probe.features }
	return {
		file: basename(fixture.file),
		size_bytes: bytes.byteLength,
		features,
		counts: {
			worksheets: probe.counts.worksheets,
			formulas: probe.counts.formulas,
			charts: probe.counts.charts,
			tables: probe.counts.tables,
			drawings: probe.counts.drawings,
			pivot_tables: probe.counts.pivot_tables,
			pivot_caches: probe.counts.pivot_caches,
			comments: probe.counts.comments,
		},
		source: EXCELFORGE_SOURCE,
		sourceUrl: `${EXCELFORGE_BASE_URL}/${fixture.sourcePath}`,
		license: 'MIT',
		sha256: createHash('sha256').update(bytes).digest('hex'),
		redistributionAllowed: true,
		citation: 'ExcelForge src/test Book 1.xlsx fixture, MIT per upstream package manifest.',
		vendorable: true,
		benchmarkTier: 'core',
		assertionClass: 'preservation-only',
		riskClass: 'high',
		featureTags: ['excelforge', 'metadata-only', 'unknown-part'],
		notes: `Upstream package manifest SHA-256 ${fixture.packageManifestSha256}; workbook source SHA-256 ${fixture.sourceSha256}.`,
	}
}
