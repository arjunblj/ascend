#!/usr/bin/env bun
import type { StyleId } from '@ascend/core'
import {
	emptySharedStrings,
	parseSharedStrings,
} from '../../packages/io-xlsx/src/reader/shared-strings.ts'
import {
	type SheetParseContext,
	streamSheetRowsXml,
	ValueInternPool,
} from '../../packages/io-xlsx/src/reader/sheet.ts'
import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'
import { buildWorkloadDataSet, type ReadSource, type WorkloadName } from './competitive-io.ts'
import { UPSTREAM_PROFILES } from './upstream-profiles.ts'

interface Args {
	readonly profile?: string
	readonly rows: number
	readonly cols: number
	readonly workload: WorkloadName
	readonly readSource: ReadSource
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface TokenizerSample {
	readonly streamRowsMs: number
	readonly structuralScanMs: number
	readonly decodeAndStructuralScanMs: number
	readonly byteScanMs: number
}

export interface WorksheetXmlScanSummary {
	readonly rows: number
	readonly cells: number
	readonly values: number
	readonly formulas: number
	readonly inlineStrings: number
	readonly sharedStringCells: number
	readonly numericCells: number
	readonly maxCol: number
	readonly lastRow: number
	readonly checksum: number
}

const WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'mixed-closedxml-10text-5number',
	'plain-text',
	'string-heavy',
	'sparse-wide',
	'styles-heavy',
	'formula-heavy',
	'table-heavy',
	'feature-rich',
	'selected-sheet',
	'metadata-only',
	'warm-workflow',
])

export function scanWorksheetXmlStructure(xml: string): WorksheetXmlScanSummary {
	const sheetData = locateSheetData(xml)
	if (!sheetData) return emptyScan()
	let rowCursor = sheetData.contentStart
	let currentRow = -1
	let rows = 0
	let cells = 0
	let values = 0
	let formulas = 0
	let inlineStrings = 0
	let sharedStringCells = 0
	let numericCells = 0
	let maxCol = 0
	let checksum = 0

	while (true) {
		const rowOpen = xml.indexOf('<row', rowCursor)
		if (rowOpen === -1 || rowOpen >= sheetData.contentEnd) break
		const rowTagEnd = findTagEnd(xml, rowOpen)
		if (rowTagEnd === -1 || rowTagEnd >= sheetData.contentEnd) break
		const explicitRow = positiveIntAttrInRange(xml, rowOpen + 4, rowTagEnd, 'r')
		const row = explicitRow === undefined ? currentRow + 1 : explicitRow - 1
		currentRow = row
		rows += 1
		checksum = mixChecksum(checksum, row + 1)
		if (isSelfClosingTag(xml, rowOpen, rowTagEnd)) {
			rowCursor = rowTagEnd + 1
			continue
		}
		const rowClose = xml.indexOf('</row>', rowTagEnd + 1)
		if (rowClose === -1 || rowClose > sheetData.contentEnd) break
		let cellCursor = rowTagEnd + 1
		let nextCol = 0
		while (true) {
			const cellOpen = xml.indexOf('<c', cellCursor)
			if (cellOpen === -1 || cellOpen >= rowClose) break
			const cellTagEnd = findTagEnd(xml, cellOpen)
			if (cellTagEnd === -1 || cellTagEnd > rowClose) break
			const col = cellColInRange(xml, cellOpen + 2, cellTagEnd) ?? nextCol
			nextCol = col + 1
			cells += 1
			if (col + 1 > maxCol) maxCol = col + 1
			checksum = mixChecksum(checksum, (row + 1) * 131 + col + 1)
			const type = attrValueInRange(xml, cellOpen + 2, cellTagEnd, 't')
			if (type === 's') sharedStringCells += 1
			else if (type === 'inlineStr') inlineStrings += 1
			else numericCells += 1
			const selfClosing = isSelfClosingTag(xml, cellOpen, cellTagEnd)
			const cellClose = selfClosing ? -1 : xml.indexOf('</c>', cellTagEnd + 1)
			const bodyEnd = cellClose === -1 || cellClose > rowClose ? cellTagEnd + 1 : cellClose
			if (!selfClosing && hasTagInRange(xml, cellTagEnd + 1, bodyEnd, 'v')) values += 1
			if (!selfClosing && hasTagInRange(xml, cellTagEnd + 1, bodyEnd, 'f')) formulas += 1
			cellCursor = cellClose === -1 || cellClose > rowClose ? cellTagEnd + 1 : cellClose + 4
		}
		rowCursor = rowClose + 6
	}

	return {
		rows,
		cells,
		values,
		formulas,
		inlineStrings,
		sharedStringCells,
		numericCells,
		maxCol,
		lastRow: currentRow + 1,
		checksum,
	}
}

export function scanWorksheetXmlBytes(bytes: Uint8Array): WorksheetXmlScanSummary {
	const sheetData = locateSheetDataBytes(bytes)
	if (!sheetData) return emptyScan()
	let rowCursor = sheetData.contentStart
	let currentRow = -1
	let rows = 0
	let cells = 0
	let values = 0
	let formulas = 0
	let inlineStrings = 0
	let sharedStringCells = 0
	let numericCells = 0
	let maxCol = 0
	let checksum = 0

	while (true) {
		const rowOpen = indexOfTag(bytes, TAG_ROW_OPEN, rowCursor, sheetData.contentEnd, 4)
		if (rowOpen === -1) break
		const rowTagEnd = findTagEndBytes(bytes, rowOpen)
		if (rowTagEnd === -1 || rowTagEnd >= sheetData.contentEnd) break
		const explicitRow = positiveIntAttrInRangeBytes(bytes, rowOpen + 4, rowTagEnd, ATTR_R)
		const row = explicitRow === undefined ? currentRow + 1 : explicitRow - 1
		currentRow = row
		rows += 1
		checksum = mixChecksum(checksum, row + 1)
		if (isSelfClosingTagBytes(bytes, rowOpen, rowTagEnd)) {
			rowCursor = rowTagEnd + 1
			continue
		}
		const rowClose = indexOfBytes(bytes, TAG_ROW_CLOSE, rowTagEnd + 1, sheetData.contentEnd)
		if (rowClose === -1) break
		let cellCursor = rowTagEnd + 1
		let nextCol = 0
		while (true) {
			const cellOpen = indexOfTag(bytes, TAG_CELL_OPEN, cellCursor, rowClose, 2)
			if (cellOpen === -1) break
			const cellTagEnd = findTagEndBytes(bytes, cellOpen)
			if (cellTagEnd === -1 || cellTagEnd > rowClose) break
			const col = cellColInRangeBytes(bytes, cellOpen + 2, cellTagEnd) ?? nextCol
			nextCol = col + 1
			cells += 1
			if (col + 1 > maxCol) maxCol = col + 1
			checksum = mixChecksum(checksum, (row + 1) * 131 + col + 1)
			const cellType = cellTypeInRangeBytes(bytes, cellOpen + 2, cellTagEnd)
			if (cellType === BYTE_CELL_TYPE_SHARED_STRING) sharedStringCells += 1
			else if (cellType === BYTE_CELL_TYPE_INLINE_STRING) inlineStrings += 1
			else numericCells += 1
			const selfClosing = isSelfClosingTagBytes(bytes, cellOpen, cellTagEnd)
			const cellClose = selfClosing
				? -1
				: indexOfBytes(bytes, TAG_CELL_CLOSE, cellTagEnd + 1, rowClose)
			const bodyEnd = cellClose === -1 ? cellTagEnd + 1 : cellClose
			if (!selfClosing && hasTagInRangeBytes(bytes, cellTagEnd + 1, bodyEnd, TAG_VALUE_OPEN)) {
				values += 1
			}
			if (!selfClosing && hasTagInRangeBytes(bytes, cellTagEnd + 1, bodyEnd, TAG_FORMULA_OPEN)) {
				formulas += 1
			}
			cellCursor = cellClose === -1 ? cellTagEnd + 1 : cellClose + TAG_CELL_CLOSE.length
		}
		rowCursor = rowClose + TAG_ROW_CLOSE.length
	}

	return {
		rows,
		cells,
		values,
		formulas,
		inlineStrings,
		sharedStringCells,
		numericCells,
		maxCol,
		lastRow: currentRow + 1,
		checksum,
	}
}

async function main(): Promise<void> {
	const args = parseArgs(process.argv.slice(2))
	const input = await buildWorkloadDataSet(args.workload, args.rows, args.cols, args.readSource)
	const archive = extractZip(input.xlsxBytes)
	const sheetPath = [...archive.entries()].find((entry) =>
		/^xl\/worksheets\/sheet\d+\.xml$/i.test(entry.path),
	)?.path
	if (!sheetPath) throw new Error('No worksheet XML part found')
	const sheetBytes = archive.readBytes(sheetPath)
	if (!sheetBytes) throw new Error(`Missing worksheet XML part ${sheetPath}`)
	const sheetXml = new TextDecoder('utf-8').decode(sheetBytes)
	const sharedStrings = archive.readText('xl/sharedStrings.xml')
	const valuePool = new ValueInternPool()
	const ctx: SheetParseContext = {
		sharedStrings: sharedStrings
			? parseSharedStrings(sharedStrings, {
					normalize: (value) => valuePool.internValue(value),
					lazy: true,
				})
			: emptySharedStrings(),
		styleIds: [0 as StyleId],
		isDateFormat: [false],
		valuePool,
		valuesOnly: true,
	}

	const runSample = (): TokenizerSample => ({
		streamRowsMs: time(() => consumeStreamRows(sheetXml, ctx)),
		structuralScanMs: time(() => consumeScan(scanWorksheetXmlStructure(sheetXml))),
		decodeAndStructuralScanMs: time(() =>
			consumeScan(scanWorksheetXmlStructure(new TextDecoder('utf-8').decode(sheetBytes))),
		),
		byteScanMs: time(() => consumeScan(scanWorksheetXmlBytes(sheetBytes))),
	})
	for (let i = 0; i < args.warmup; i++) runSample()
	const samples = Array.from({ length: args.repeat }, runSample)
	const scan = scanWorksheetXmlStructure(sheetXml)
	const byteScan = scanWorksheetXmlBytes(sheetBytes)
	assertMatchingScans(scan, byteScan)
	const streamCells = consumeStreamRows(sheetXml, ctx)
	const payload = {
		tool: 'xlsx-xml-tokenizer',
		args,
		fixture: {
			path: input.xlsxPath,
			sheetPath,
			sheetXmlBytes: sheetBytes.byteLength,
		},
		summary: summarize(samples, scan, streamCells),
		samples,
	}
	if (args.json) console.log(JSON.stringify(payload, null, 2))
	else console.log(payload.summary)
}

function summarize(
	samples: readonly TokenizerSample[],
	scan: WorksheetXmlScanSummary,
	streamCells: number,
) {
	const streamRowsMedianMs = median(samples.map((sample) => sample.streamRowsMs))
	const structuralScanMedianMs = median(samples.map((sample) => sample.structuralScanMs))
	const decodeAndStructuralScanMedianMs = median(
		samples.map((sample) => sample.decodeAndStructuralScanMs),
	)
	const byteScanMedianMs = median(samples.map((sample) => sample.byteScanMs))
	return {
		streamRowsMedianMs,
		structuralScanMedianMs,
		decodeAndStructuralScanMedianMs,
		byteScanMedianMs,
		tokenizerHeadroom: streamRowsMedianMs / structuralScanMedianMs,
		byteTokenizerHeadroom: streamRowsMedianMs / byteScanMedianMs,
		byteVsStringScanHeadroom: structuralScanMedianMs / byteScanMedianMs,
		byteVsDecodeStringScanHeadroom: decodeAndStructuralScanMedianMs / byteScanMedianMs,
		streamCells,
		...scan,
	}
}

function assertMatchingScans(
	stringScan: WorksheetXmlScanSummary,
	byteScan: WorksheetXmlScanSummary,
): void {
	const keys = [
		'rows',
		'cells',
		'values',
		'formulas',
		'inlineStrings',
		'sharedStringCells',
		'numericCells',
		'maxCol',
		'lastRow',
		'checksum',
	] as const
	for (const key of keys) {
		if (stringScan[key] !== byteScan[key]) {
			throw new Error(
				`byte worksheet scan mismatch for ${key}: ${byteScan[key]} !== ${stringScan[key]}`,
			)
		}
	}
}

function consumeStreamRows(xml: string, ctx: SheetParseContext): number {
	let cells = 0
	for (const row of streamSheetRowsXml('Sheet1', xml, ctx)) cells += row.cells.length
	return cells
}

function consumeScan(scan: WorksheetXmlScanSummary): number {
	return scan.cells ^ scan.rows ^ scan.values ^ scan.checksum
}

function time(fn: () => number): number {
	const start = performance.now()
	const result = fn()
	if (Number.isNaN(result)) throw new Error('unreachable')
	return performance.now() - start
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const middle = Math.floor(sorted.length / 2)
	const upper = sorted[middle] ?? 0
	return sorted.length % 2 === 1 ? upper : ((sorted[middle - 1] ?? upper) + upper) / 2
}

function parseArgs(argv: readonly string[]): Args {
	const profileName = option(argv, '--profile')
	const profile = profileName
		? UPSTREAM_PROFILES.find((entry) => entry.name === profileName)
		: undefined
	if (profileName && !profile) throw new Error(`Unsupported --profile "${profileName}"`)
	if (profile && profile.category !== 'read') {
		throw new Error(`--profile "${profile.name}" is a write profile`)
	}
	const workload = option(argv, '--workload') ?? profile?.workload ?? 'dense-values'
	if (!WORKLOADS.has(workload)) throw new Error(`Unsupported --workload "${workload}"`)
	const readSource = option(argv, '--read-source') ?? profile?.readSource ?? 'raw-ooxml'
	if (readSource !== 'ascend-writer' && readSource !== 'raw-ooxml') {
		throw new Error('--read-source must be ascend-writer or raw-ooxml')
	}
	return {
		...(profile ? { profile: profile.name } : {}),
		rows: positiveInt(option(argv, '--rows'), profile?.rows ?? 5000),
		cols: positiveInt(option(argv, '--cols'), profile?.cols ?? 20),
		workload: workload as WorkloadName,
		readSource,
		repeat: positiveInt(option(argv, '--repeat'), 7),
		warmup: nonNegativeInt(option(argv, '--warmup'), 2),
		json: argv.includes('--json'),
	}
}

function option(args: readonly string[], name: string): string | undefined {
	const index = args.indexOf(name)
	return index >= 0 ? args[index + 1] : undefined
}

function positiveInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

function emptyScan(): WorksheetXmlScanSummary {
	return {
		rows: 0,
		cells: 0,
		values: 0,
		formulas: 0,
		inlineStrings: 0,
		sharedStringCells: 0,
		numericCells: 0,
		maxCol: 0,
		lastRow: 0,
		checksum: 0,
	}
}

interface ByteSheetDataLocation {
	readonly contentStart: number
	readonly contentEnd: number
}

const BYTE_CELL_TYPE_NUMERIC = 0
const BYTE_CELL_TYPE_SHARED_STRING = 1
const BYTE_CELL_TYPE_INLINE_STRING = 2
const ATTR_R = Uint8Array.of(114)
const ATTR_T = Uint8Array.of(116)
const TYPE_SHARED_STRING = Uint8Array.of(115)
const TYPE_INLINE_STRING = new TextEncoder().encode('inlineStr')
const TAG_SHEET_DATA_OPEN = new TextEncoder().encode('<sheetData')
const TAG_SHEET_DATA_CLOSE = new TextEncoder().encode('</sheetData>')
const TAG_ROW_OPEN = new TextEncoder().encode('<row')
const TAG_ROW_CLOSE = new TextEncoder().encode('</row>')
const TAG_CELL_OPEN = new TextEncoder().encode('<c')
const TAG_CELL_CLOSE = new TextEncoder().encode('</c>')
const TAG_VALUE_OPEN = new TextEncoder().encode('<v')
const TAG_FORMULA_OPEN = new TextEncoder().encode('<f')

function locateSheetDataBytes(bytes: Uint8Array): ByteSheetDataLocation | null {
	const open = indexOfBytes(bytes, TAG_SHEET_DATA_OPEN, 0, bytes.byteLength)
	if (open === -1) return null
	const tagEnd = findTagEndBytes(bytes, open)
	if (tagEnd === -1 || isSelfClosingTagBytes(bytes, open, tagEnd)) return null
	const close = indexOfBytes(bytes, TAG_SHEET_DATA_CLOSE, tagEnd + 1, bytes.byteLength)
	return close === -1 ? null : { contentStart: tagEnd + 1, contentEnd: close }
}

function indexOfTag(
	bytes: Uint8Array,
	pattern: Uint8Array,
	start: number,
	end: number,
	boundaryIndex: number,
): number {
	let cursor = start
	while (true) {
		const index = indexOfBytes(bytes, pattern, cursor, end)
		if (index === -1) return -1
		const boundary = index + boundaryIndex
		if (boundary >= end || isXmlTagBoundaryByte(bytes[boundary] ?? -1)) return index
		cursor = index + 1
	}
}

function indexOfBytes(bytes: Uint8Array, pattern: Uint8Array, start: number, end: number): number {
	const first = pattern[0]
	const last = end - pattern.length
	for (let index = start; index <= last; index++) {
		if (bytes[index] !== first) continue
		let matched = true
		for (let offset = 1; offset < pattern.length; offset++) {
			if (bytes[index + offset] === pattern[offset]) continue
			matched = false
			break
		}
		if (matched) return index
	}
	return -1
}

function findTagEndBytes(bytes: Uint8Array, start: number): number {
	let quote = 0
	for (let i = start; i < bytes.byteLength; i++) {
		const code = bytes[i] ?? -1
		if (quote !== 0) {
			if (code === quote) quote = 0
			continue
		}
		if (code === 34 || code === 39) quote = code
		else if (code === 62) return i
	}
	return -1
}

function isSelfClosingTagBytes(bytes: Uint8Array, start: number, tagEnd: number): boolean {
	let cursor = tagEnd - 1
	while (cursor > start) {
		const code = bytes[cursor] ?? -1
		if (!isXmlWhitespaceByte(code)) break
		cursor -= 1
	}
	return bytes[cursor] === 47
}

function positiveIntAttrInRangeBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	name: Uint8Array,
): number | undefined {
	const valueStart = attrValueStartInRangeBytes(bytes, start, end, name)
	if (valueStart === -1) return undefined
	const valueEnd = indexOfByte(bytes, 34, valueStart, end)
	if (valueEnd === -1) return undefined
	let value = 0
	for (let i = valueStart; i < valueEnd; i++) {
		const code = bytes[i] ?? -1
		if (code < 48 || code > 57) return undefined
		value = value * 10 + (code - 48)
	}
	return value > 0 ? value : undefined
}

function cellTypeInRangeBytes(bytes: Uint8Array, start: number, end: number): number {
	const valueStart = attrValueStartInRangeBytes(bytes, start, end, ATTR_T)
	if (valueStart === -1) return BYTE_CELL_TYPE_NUMERIC
	const valueEnd = indexOfByte(bytes, 34, valueStart, end)
	if (valueEnd === -1) return BYTE_CELL_TYPE_NUMERIC
	if (bytesEqualRange(bytes, valueStart, valueEnd, TYPE_SHARED_STRING)) {
		return BYTE_CELL_TYPE_SHARED_STRING
	}
	if (bytesEqualRange(bytes, valueStart, valueEnd, TYPE_INLINE_STRING)) {
		return BYTE_CELL_TYPE_INLINE_STRING
	}
	return BYTE_CELL_TYPE_NUMERIC
}

function attrValueStartInRangeBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	name: Uint8Array,
): number {
	let cursor = start
	while (cursor < end) {
		cursor = skipXmlWhitespaceBytes(bytes, cursor, end)
		if (cursor >= end || bytes[cursor] === 47) return -1
		const attrNameStart = cursor
		while (cursor < end) {
			const code = bytes[cursor] ?? -1
			if (code === 61 || isXmlWhitespaceByte(code)) break
			cursor += 1
		}
		const attrNameEnd = cursor
		cursor = skipXmlWhitespaceBytes(bytes, cursor, end)
		if (cursor >= end || bytes[cursor] !== 61) return -1
		cursor = skipXmlWhitespaceBytes(bytes, cursor + 1, end)
		if (cursor >= end || bytes[cursor] !== 34) return -1
		const valueStart = cursor + 1
		const valueEnd = indexOfByte(bytes, 34, valueStart, end)
		if (valueEnd === -1) return -1
		if (bytesEqualRange(bytes, attrNameStart, attrNameEnd, name)) return valueStart
		cursor = valueEnd + 1
	}
	return -1
}

function cellColInRangeBytes(bytes: Uint8Array, start: number, end: number): number | undefined {
	const valueStart = attrValueStartInRangeBytes(bytes, start, end, ATTR_R)
	if (valueStart === -1) return undefined
	const valueEnd = indexOfByte(bytes, 34, valueStart, end)
	if (valueEnd === -1) return undefined
	let index = valueStart
	let col = 0
	while (index < valueEnd) {
		const code = bytes[index] ?? -1
		if (code >= 48 && code <= 57) break
		if (code >= 65 && code <= 90) col = col * 26 + (code - 64)
		else if (code >= 97 && code <= 122) col = col * 26 + (code - 96)
		else return undefined
		index += 1
	}
	return col > 0 ? col - 1 : undefined
}

function hasTagInRangeBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	pattern: Uint8Array,
): boolean {
	const index = indexOfTag(bytes, pattern, start, end, pattern.length)
	return index !== -1
}

function indexOfByte(bytes: Uint8Array, code: number, start: number, end: number): number {
	for (let index = start; index < end; index++) {
		if (bytes[index] === code) return index
	}
	return -1
}

function bytesEqualRange(
	bytes: Uint8Array,
	start: number,
	end: number,
	expected: Uint8Array,
): boolean {
	if (end - start !== expected.length) return false
	for (let index = 0; index < expected.length; index++) {
		if (bytes[start + index] !== expected[index]) return false
	}
	return true
}

function skipXmlWhitespaceBytes(bytes: Uint8Array, cursor: number, end: number): number {
	while (cursor < end && isXmlWhitespaceByte(bytes[cursor] ?? -1)) cursor += 1
	return cursor
}

function isXmlWhitespaceByte(code: number): boolean {
	return code === 9 || code === 10 || code === 13 || code === 32
}

function isXmlTagBoundaryByte(code: number): boolean {
	return isXmlWhitespaceByte(code) || code === 47 || code === 62
}

interface SheetDataLocation {
	readonly contentStart: number
	readonly contentEnd: number
}

function locateSheetData(xml: string): SheetDataLocation | null {
	const open = xml.indexOf('<sheetData')
	if (open === -1) return null
	const tagEnd = findTagEnd(xml, open)
	if (tagEnd === -1 || isSelfClosingTag(xml, open, tagEnd)) return null
	const close = xml.indexOf('</sheetData>', tagEnd + 1)
	return close === -1 ? null : { contentStart: tagEnd + 1, contentEnd: close }
}

function findTagEnd(xml: string, start: number): number {
	let quote = 0
	for (let i = start; i < xml.length; i++) {
		const code = xml.charCodeAt(i)
		if (quote !== 0) {
			if (code === quote) quote = 0
			continue
		}
		if (code === 34 || code === 39) quote = code
		else if (code === 62) return i
	}
	return -1
}

function isSelfClosingTag(xml: string, start: number, tagEnd: number): boolean {
	let cursor = tagEnd - 1
	while (cursor > start) {
		const code = xml.charCodeAt(cursor)
		if (code !== 9 && code !== 10 && code !== 13 && code !== 32) break
		cursor -= 1
	}
	return xml.charCodeAt(cursor) === 47
}

function positiveIntAttrInRange(
	xml: string,
	start: number,
	end: number,
	name: string,
): number | undefined {
	const valueStart = attrValueStartInRange(xml, start, end, name)
	if (valueStart === -1) return undefined
	const valueEnd = xml.indexOf('"', valueStart)
	if (valueEnd === -1 || valueEnd > end) return undefined
	let value = 0
	for (let i = valueStart; i < valueEnd; i++) {
		const code = xml.charCodeAt(i)
		if (code < 48 || code > 57) return undefined
		value = value * 10 + (code - 48)
	}
	return value > 0 ? value : undefined
}

function attrValueInRange(
	xml: string,
	start: number,
	end: number,
	name: string,
): string | undefined {
	const valueStart = attrValueStartInRange(xml, start, end, name)
	if (valueStart === -1) return undefined
	const valueEnd = xml.indexOf('"', valueStart)
	return valueEnd === -1 || valueEnd > end ? undefined : xml.slice(valueStart, valueEnd)
}

function attrValueStartInRange(xml: string, start: number, end: number, name: string): number {
	let cursor = start
	const nameLength = name.length
	while (cursor < end) {
		cursor = skipXmlWhitespace(xml, cursor, end)
		if (cursor >= end || xml.charCodeAt(cursor) === 47) return -1
		const attrNameStart = cursor
		while (cursor < end) {
			const code = xml.charCodeAt(cursor)
			if (code === 61 || code === 9 || code === 10 || code === 13 || code === 32) break
			cursor += 1
		}
		const attrNameEnd = cursor
		cursor = skipXmlWhitespace(xml, cursor, end)
		if (cursor >= end || xml.charCodeAt(cursor) !== 61) return -1
		cursor = skipXmlWhitespace(xml, cursor + 1, end)
		if (cursor >= end || xml.charCodeAt(cursor) !== 34) return -1
		const valueStart = cursor + 1
		const valueEnd = xml.indexOf('"', valueStart)
		if (valueEnd === -1 || valueEnd > end) return -1
		if (attrNameEnd - attrNameStart === nameLength && xml.startsWith(name, attrNameStart)) {
			return valueStart
		}
		cursor = valueEnd + 1
	}
	return -1
}

function cellColInRange(xml: string, start: number, end: number): number | undefined {
	const valueStart = attrValueStartInRange(xml, start, end, 'r')
	if (valueStart === -1) return undefined
	const valueEnd = xml.indexOf('"', valueStart)
	if (valueEnd === -1 || valueEnd > end) return undefined
	let index = valueStart
	let col = 0
	while (index < valueEnd) {
		const code = xml.charCodeAt(index)
		if (code >= 48 && code <= 57) break
		if (code >= 65 && code <= 90) col = col * 26 + (code - 64)
		else if (code >= 97 && code <= 122) col = col * 26 + (code - 96)
		else return undefined
		index += 1
	}
	return col > 0 ? col - 1 : undefined
}

function hasTagInRange(xml: string, start: number, end: number, tagName: string): boolean {
	const first = tagName.charCodeAt(0)
	for (let cursor = start; cursor < end - tagName.length; cursor++) {
		if (xml.charCodeAt(cursor) !== 60 || xml.charCodeAt(cursor + 1) !== first) continue
		let matched = true
		for (let index = 1; index < tagName.length; index++) {
			if (xml.charCodeAt(cursor + index + 1) !== tagName.charCodeAt(index)) {
				matched = false
				break
			}
		}
		if (!matched) continue
		const next = xml.charCodeAt(cursor + tagName.length + 1)
		return next === 9 || next === 10 || next === 13 || next === 32 || next === 47 || next === 62
	}
	return false
}

function skipXmlWhitespace(xml: string, cursor: number, end: number): number {
	while (cursor < end) {
		const code = xml.charCodeAt(cursor)
		if (code !== 9 && code !== 10 && code !== 13 && code !== 32) break
		cursor += 1
	}
	return cursor
}

function mixChecksum(current: number, value: number): number {
	return Math.imul(current ^ value, 16_777_619) >>> 0 || 2_166_136_261
}

if (import.meta.main) await main()
