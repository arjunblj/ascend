#!/usr/bin/env bun
import { indexToColumn } from '../../packages/core/src/index.ts'
import { createZip, encode, StreamingZipBuilder } from '../../packages/io-xlsx/src/writer/zip.ts'
import { escapeXml } from '../../packages/io-xlsx/src/xml.ts'
import {
	buildWorkloadValues,
	denseWriteAssertions,
	expectedWorkloadValuesHash,
	type WorkloadName,
} from './competitive-io.ts'
import { UPSTREAM_PROFILES } from './upstream-profiles.ts'

interface Args {
	readonly profile?: string
	readonly rows: number
	readonly cols: number
	readonly workload: WorkloadName
	readonly xmlMode: XmlMode
	readonly zipMode: ZipMode
	readonly repeat: number
	readonly warmup: number
	readonly validate: boolean
	readonly json: boolean
}

type XmlMode = 'join' | 'sink'
type ZipMode = 'buffer' | 'stream'
type WorkloadCellValue = string | number | boolean | null
type PrimitiveAssertion = string | number | boolean | null
const XML_BYTE_BATCH_TARGET = positiveEnvInt('ASCEND_DENSE_XML_BYTE_BATCH', 2 * 1024 * 1024)
const XML_ROW_BATCH_TARGET = positiveEnvInt('ASCEND_DENSE_XML_ROW_BATCH', 1024)

interface PhaseSample {
	readonly writeMs: number
	readonly xmlBuildMs?: number
	readonly zipBuildMs?: number
	readonly validateMs?: number
	readonly assertions?: Record<string, PrimitiveAssertion>
	readonly cellsPerSecond: number
	readonly writeNsPerCell: number
	readonly bytes: number
	readonly rssAfterBytes: number
	readonly heapUsedBytes: number
}
interface BuildResult {
	readonly bytes: Uint8Array
	readonly xmlBuildMs?: number
	readonly zipBuildMs?: number
}

const WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'mixed-closedxml-10text-5number',
	'plain-text',
	'string-heavy',
])

function readOption(args: readonly string[], name: string): string | undefined {
	const index = args.indexOf(name)
	return index >= 0 ? args[index + 1] : undefined
}

function hasFlag(args: readonly string[], name: string): boolean {
	return args.includes(name)
}

function positiveInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

function positiveEnvInt(name: string, fallback: number): number {
	const raw = process.env[name]
	if (!raw) return fallback
	const value = Number.parseInt(raw, 10)
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function parseArgs(): Args {
	const argv = process.argv.slice(2)
	const profileName = readOption(argv, '--profile')
	const profile = profileName
		? UPSTREAM_PROFILES.find((entry) => entry.name === profileName)
		: undefined
	if (profileName && !profile) {
		throw new Error(
			`Unsupported --profile "${profileName}". Expected one of: ${UPSTREAM_PROFILES.map((entry) => entry.name).join(', ')}`,
		)
	}
	if (profile && profile.category !== 'write') {
		throw new Error(
			`--profile "${profile.name}" is a read profile; xlsx-row-writer-phase only supports write profiles`,
		)
	}
	const workload = readOption(argv, '--workload') ?? profile?.workload ?? 'dense-values'
	if (!WORKLOADS.has(workload)) {
		throw new Error(`Unsupported --workload "${workload}" for row writer prototype`)
	}
	const xmlMode = readOption(argv, '--xml-mode') ?? 'join'
	if (xmlMode !== 'join' && xmlMode !== 'sink') {
		throw new Error('Unsupported --xml-mode. Expected "join" or "sink"')
	}
	const zipMode = readOption(argv, '--zip-mode') ?? 'buffer'
	if (zipMode !== 'buffer' && zipMode !== 'stream') {
		throw new Error('Unsupported --zip-mode. Expected "buffer" or "stream"')
	}
	return {
		...(profile ? { profile: profile.name } : {}),
		rows: positiveInt(readOption(argv, '--rows'), profile?.rows ?? 2000),
		cols: positiveInt(readOption(argv, '--cols'), profile?.cols ?? 20),
		workload: workload as WorkloadName,
		xmlMode,
		zipMode,
		repeat: positiveInt(readOption(argv, '--repeat'), 5),
		warmup: nonNegativeInt(readOption(argv, '--warmup'), 1),
		validate: hasFlag(argv, '--validate'),
		json: hasFlag(argv, '--json'),
	}
}

function workloadValue(
	workloadName: WorkloadName,
	row: number,
	col: number,
	cols: number,
): WorkloadCellValue {
	if (workloadName === 'dense-values') return row * cols + col
	if (workloadName === 'mixed-10pct-text') {
		const key = row * cols + col
		return key % 10 === 0 ? `text-${String(key).padStart(8, '0')}` : key
	}
	if (workloadName === 'mixed-50pct-text') {
		const key = row * cols + col
		return key % 2 === 0 ? `text-${String(key).padStart(8, '0')}` : key
	}
	if (workloadName === 'mixed-closedxml-10text-5number') {
		return col < 10 ? 'Hello world' : col - 10
	}
	if (workloadName === 'plain-text') return `text-${String(row * cols + col).padStart(8, '0')}`

	const key = row * cols + col
	switch (col % 5) {
		case 0:
			return `sku-${String(key).padStart(8, '0')}`
		case 1:
			return `region-${(row % 17) + 1}`
		case 2:
			return `customer-${row % 997}-segment-${col % 13}`
		case 3:
			return `note row ${row} col ${col} token ${key % 104729}`
		default:
			return key % 2 === 0 ? `status-open-${key % 31}` : `status-closed-${key % 29}`
	}
}

function pushCellXml(
	out: string[],
	workloadName: WorkloadName,
	row: number,
	col: number,
	columnName: string,
	cols: number,
): void {
	const value = workloadValue(workloadName, row, col, cols)
	if (value === null) return
	const ref = `${columnName}${row + 1}`
	if (typeof value === 'number') {
		out.push(`<c r="${ref}"><v>${value}</v></c>`)
		return
	}
	if (typeof value === 'boolean') {
		out.push(`<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`)
		return
	}
	out.push(`<c r="${ref}" t="str"><v>${escapeXml(value)}</v></c>`)
}

function buildColumnNames(cols: number): string[] {
	return Array.from({ length: cols }, (_, index) => indexToColumn(index))
}

function sheetDimension(columnNames: readonly string[], rows: number): string {
	return `A1:${columnNames.at(-1) ?? 'A'}${Math.max(1, rows)}`
}

function buildSheetXml(args: Args): Uint8Array {
	return args.xmlMode === 'sink' ? buildSheetXmlWithSink(args) : buildSheetXmlWithJoin(args)
}

function buildSheetXmlWithJoin(args: Args): Uint8Array {
	const columnNames = buildColumnNames(args.cols)
	const out: string[] = [
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
		'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
		`<dimension ref="${sheetDimension(columnNames, args.rows)}"/>`,
		'<sheetData>',
	]
	for (let row = 0; row < args.rows; row++) {
		const cells: string[] = []
		for (let col = 0; col < args.cols; col++) {
			pushCellXml(cells, args.workload, row, col, columnNames[col] as string, args.cols)
		}
		if (cells.length > 0) out.push(`<row r="${row + 1}">${cells.join('')}</row>`)
	}
	out.push('</sheetData></worksheet>')
	return encode(out.join(''))
}

function buildSheetXmlWithSink(args: Args): Uint8Array {
	const columnNames = buildColumnNames(args.cols)
	const sink = new Bun.ArrayBufferSink()
	sink.start({
		asUint8Array: true,
		highWaterMark: sinkHighWaterMark(args.rows, args.cols),
	})
	sink.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
	sink.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
	sink.write(`<dimension ref="${sheetDimension(columnNames, args.rows)}"/>`)
	sink.write('<sheetData>')
	for (let row = 0; row < args.rows; row++) {
		const cells: string[] = []
		for (let col = 0; col < args.cols; col++) {
			pushCellXml(cells, args.workload, row, col, columnNames[col] as string, args.cols)
		}
		if (cells.length > 0) {
			sink.write(`<row r="${row + 1}">`)
			sink.write(cells.join(''))
			sink.write('</row>')
		}
	}
	sink.write('</sheetData></worksheet>')
	return sink.end() as Uint8Array
}

function sinkHighWaterMark(rows: number, cols: number): number {
	return Math.min(64 * 1024 * 1024, Math.max(1024 * 1024, rows * cols * 32))
}

function writeSheetXmlChunks(args: Args, onChunk: (chunk: string) => void): void {
	const columnNames = buildColumnNames(args.cols)
	onChunk('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
	onChunk('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
	onChunk(`<dimension ref="${sheetDimension(columnNames, args.rows)}"/><sheetData>`)
	let batch = ''
	for (let row = 0; row < args.rows; row++) {
		const cells: string[] = []
		for (let col = 0; col < args.cols; col++) {
			pushCellXml(cells, args.workload, row, col, columnNames[col] as string, args.cols)
		}
		if (cells.length === 0) continue
		batch += `<row r="${row + 1}">${cells.join('')}</row>`
		if ((row + 1) % XML_ROW_BATCH_TARGET === 0 || batch.length >= XML_BYTE_BATCH_TARGET) {
			onChunk(batch)
			batch = ''
		}
	}
	if (batch.length > 0) onChunk(batch)
	onChunk('</sheetData></worksheet>')
}

function buildBaseParts(): Map<string, Uint8Array> {
	const parts = new Map<string, Uint8Array>()
	parts.set(
		'[Content_Types].xml',
		encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`),
	)
	parts.set(
		'_rels/.rels',
		encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`),
	)
	parts.set(
		'xl/_rels/workbook.xml.rels',
		encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
	)
	parts.set(
		'xl/workbook.xml',
		encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
	)
	parts.set(
		'xl/styles.xml',
		encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`),
	)
	parts.set(
		'docProps/core.xml',
		encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:creator>Ascend</dc:creator>
</cp:coreProperties>`),
	)
	parts.set(
		'docProps/app.xml',
		encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>Ascend</Application>
</Properties>`),
	)
	return parts
}

function buildXlsxBuffered(args: Args): BuildResult {
	const parts = buildBaseParts()
	const xmlStart = performance.now()
	parts.set('xl/worksheets/sheet1.xml', buildSheetXml(args))
	const xmlBuildMs = performance.now() - xmlStart
	const zipStart = performance.now()
	const bytes = createZip(parts)
	const zipBuildMs = performance.now() - zipStart
	return { bytes, xmlBuildMs, zipBuildMs }
}

async function buildXlsxStreamed(args: Args): Promise<BuildResult> {
	const builder = new StreamingZipBuilder()
	for (const [path, data] of buildBaseParts()) {
		builder.addEntry(path, data)
	}
	const xmlStart = performance.now()
	builder.addStreamingEntry('xl/worksheets/sheet1.xml')
	writeSheetXmlChunks(args, (chunk) => builder.writeChunk(encode(chunk)))
	const xmlBuildMs = performance.now() - xmlStart
	const zipStart = performance.now()
	await builder.closeEntry()
	const bytes = builder.finalize()
	const zipBuildMs = performance.now() - zipStart
	return { bytes, xmlBuildMs, zipBuildMs }
}

function buildXlsx(args: Args): BuildResult | Promise<BuildResult> {
	return args.zipMode === 'stream' ? buildXlsxStreamed(args) : buildXlsxBuffered(args)
}

function memory() {
	const current = process.memoryUsage()
	const rss = typeof current.rss === 'function' ? current.rss() : current.rss
	return {
		rssAfterBytes: rss,
		heapUsedBytes: current.heapUsed,
	}
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const middle = Math.floor(sorted.length / 2)
	const upper = sorted[middle] ?? 0
	return sorted.length % 2 === 1 ? upper : ((sorted[middle - 1] ?? upper) + upper) / 2
}

function summarize(samples: readonly PhaseSample[]) {
	return {
		writeMedianMs: median(samples.map((sample) => sample.writeMs)),
		xmlBuildMedianMs: median(samples.map((sample) => sample.xmlBuildMs ?? 0)),
		zipBuildMedianMs: median(samples.map((sample) => sample.zipBuildMs ?? 0)),
		...(samples.some((sample) => sample.validateMs !== undefined)
			? { validateMedianMs: median(samples.map((sample) => sample.validateMs ?? 0)) }
			: {}),
		cellsPerSecondMedian: median(samples.map((sample) => sample.cellsPerSecond)),
		writeNsPerCellMedian: median(samples.map((sample) => sample.writeNsPerCell)),
		bytesMedian: median(samples.map((sample) => sample.bytes)),
		peakRssBytes: Math.max(...samples.map((sample) => sample.rssAfterBytes)),
	}
}

async function runSample(args: Args): Promise<PhaseSample> {
	const start = performance.now()
	const build = await buildXlsx(args)
	const writeMs = performance.now() - start
	const bytes = build.bytes
	let validateMs: number | undefined
	let assertions: Record<string, PrimitiveAssertion> | undefined
	if (args.validate) {
		const validateStart = performance.now()
		const materializeValues = args.rows * args.cols <= 500_000
		const values = materializeValues ? buildWorkloadValues(args.workload, args.rows, args.cols) : []
		assertions = denseWriteAssertions(bytes, {
			workloadName: args.workload,
			readSource: 'ascend-writer',
			rows: args.rows,
			cols: args.cols,
			cells: args.rows * args.cols,
			values,
			semanticCellValuesHash: expectedWorkloadValuesHash(args.workload, args.rows, args.cols),
			xlsxPath: '',
			xlsxBytes: bytes,
		})
		validateMs = performance.now() - validateStart
	}
	const cells = args.rows * args.cols
	return {
		writeMs,
		...(build.xmlBuildMs === undefined ? {} : { xmlBuildMs: build.xmlBuildMs }),
		...(build.zipBuildMs === undefined ? {} : { zipBuildMs: build.zipBuildMs }),
		...(validateMs === undefined ? {} : { validateMs }),
		...(assertions === undefined ? {} : { assertions }),
		cellsPerSecond: cells / (writeMs / 1000),
		writeNsPerCell: (writeMs * 1_000_000) / cells,
		bytes: bytes.byteLength,
		...memory(),
	}
}

const args = parseArgs()
for (let i = 0; i < args.warmup; i++) await runSample(args)
const samples: PhaseSample[] = []
for (let i = 0; i < args.repeat; i++) samples.push(await runSample(args))
const payload = {
	tool: 'xlsx-row-writer-phase',
	args,
	summary: summarize(samples),
	samples,
}
console.log(args.json ? JSON.stringify(payload, null, 2) : payload.summary)
