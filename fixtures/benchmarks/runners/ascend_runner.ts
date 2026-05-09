#!/usr/bin/env bun
import { indexToColumn } from '../../../packages/core/src/index.ts'
import {
	parseRelationships,
	REL_OFFICE_DOC,
	REL_SHARED_STRINGS,
	resolvePath,
} from '../../../packages/io-xlsx/src/reader/relationships.ts'
import {
	emptySharedStrings,
	parseSharedStrings,
	type SharedStringResolver,
} from '../../../packages/io-xlsx/src/reader/shared-strings.ts'
import { readXlsxRowsStream } from '../../../packages/io-xlsx/src/reader/stream.ts'
import { parseWorkbookXml } from '../../../packages/io-xlsx/src/reader/workbook.ts'
import {
	decodeXmlText,
	findTagEnd,
	isSelfClosingTag,
} from '../../../packages/io-xlsx/src/reader/xml-utils.ts'
import { extractZip } from '../../../packages/io-xlsx/src/reader/zip.ts'
import type { CellValue } from '../../../packages/schema/src/index.ts'
import { Ascend } from '../../../packages/sdk/src/index.ts'
import { summarizeAscendWorkbook, workbookShapeAssertions } from '../competitive-real-workbook.ts'

type Operation = 'read' | 'roundtrip'
type Mode = 'formula' | 'values' | 'full' | 'metadata-only'
type Source = 'path' | 'bytes'
const ORDERED_HASH_FLUSH_BYTES = positiveInt(process.env.ASCEND_ORDERED_HASH_FLUSH_BYTES, 1_048_576)

export interface Args {
	readonly operation: Operation
	readonly file: string
	readonly mode: Mode
	readonly source: Source
	readonly richMetadata: boolean
	readonly orderedHashes: boolean
	readonly streamOrderedHashes: boolean
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

function readOption(args: readonly string[], name: string): string | undefined {
	const index = args.indexOf(name)
	return index >= 0 ? args[index + 1] : undefined
}

function hasFlag(args: readonly string[], name: string): boolean {
	return args.includes(name)
}

function parseArgs(): Args {
	const args = process.argv.slice(2)
	const operation = readOption(args, '--operation')
	const file = readOption(args, '--file')
	const mode = readOption(args, '--mode') ?? 'formula'
	const source = readOption(args, '--source') ?? 'path'
	if (operation !== 'read' && operation !== 'roundtrip') {
		throw new Error('--operation must be read or roundtrip')
	}
	if (!file) throw new Error('--file is required')
	if (mode !== 'formula' && mode !== 'values' && mode !== 'full' && mode !== 'metadata-only') {
		throw new Error('--mode must be formula, values, full, or metadata-only')
	}
	if (source !== 'path' && source !== 'bytes') {
		throw new Error('--source must be path or bytes')
	}
	return {
		operation,
		file,
		mode,
		source,
		richMetadata: hasFlag(args, '--rich-metadata'),
		orderedHashes: hasFlag(args, '--ordered-hashes'),
		streamOrderedHashes: hasFlag(args, '--stream-ordered-hashes'),
		repeat: positiveInt(readOption(args, '--repeat'), 1),
		warmup: nonNegativeInt(readOption(args, '--warmup'), 0),
		json: hasFlag(args, '--json'),
	}
}

function positiveInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

async function openWorkbookFromSource(
	args: Args,
	bytes: Uint8Array | undefined,
): Promise<Awaited<ReturnType<typeof Ascend.open>>> {
	if (args.operation !== 'read') throw new Error('ascend runner currently supports read only')
	const openOptions =
		args.mode === 'full' && !args.richMetadata
			? undefined
			: { mode: args.mode, ...(args.richMetadata ? { richMetadata: true } : {}) }
	return Ascend.open(bytes ?? args.file, openOptions)
}

function readAssertions(
	workbook: Awaited<ReturnType<typeof Ascend.open>>,
	args: Args,
): Record<string, string | number | boolean | null> {
	const info = workbook.inspect()
	if (args.mode === 'metadata-only') {
		return {
			metadataOnlyRead: true,
			sourceSheetCount: info.load.sourceSheets.length,
			loadedSheetCount: info.load.loadedSheets.length,
			loadedSheetNames: info.load.loadedSheets.join(','),
			hasAllSheets: info.load.hasAllSheets,
			cellsHydrated: info.load.cellsHydrated,
			cellCount: info.cellCount,
			runnerVersion: 'workspace',
			runnerSource: args.source,
			runnerLoadMode: args.mode,
		}
	}
	if (args.orderedHashes) return orderedReadAssertions(workbook, args)
	return {
		...workbookShapeAssertions(summarizeAscendWorkbook(workbook)),
		...readFeatureAssertions(workbook.getWorkbookModel()),
		runnerVersion: 'workspace',
		runnerSource: args.source,
		runnerLoadMode: args.mode,
		runnerRichMetadata: args.richMetadata,
	}
}

class OrderedLineHasher {
	private readonly hash = new Bun.CryptoHasher('sha256')
	private readonly chunks: string[] = []
	private bufferedLength = 0

	update(line: string): void {
		this.chunks.push(String(line.length), ':', line, '\n')
		this.bufferedLength += line.length + String(line.length).length + 2
		if (this.bufferedLength >= ORDERED_HASH_FLUSH_BYTES) this.flush()
	}

	digest(): string {
		this.flush()
		return this.hash.digest('hex')
	}

	private flush(): void {
		if (this.bufferedLength === 0) return
		this.hash.update(this.chunks.join(''))
		this.chunks.length = 0
		this.bufferedLength = 0
	}
}

function canonicalNumber(value: number): string {
	return Object.is(value, -0) ? '0' : String(value)
}

function serializeCellValue(value: CellValue): string {
	switch (value.kind) {
		case 'empty':
			return 'empty'
		case 'number':
			return `n:${canonicalNumber(value.value)}`
		case 'date':
			return `n:${canonicalNumber(value.serial)}`
		case 'string':
			return `s:${value.value}`
		case 'boolean':
			return `b:${value.value ? 'true' : 'false'}`
		case 'error':
			return `e:${value.value}`
		case 'richText':
			return `s:${value.runs.map((run) => run.text).join('')}`
	}
}

function hashOrdered(lines: readonly string[]): string {
	const hasher = new OrderedLineHasher()
	for (const line of lines) hasher.update(line)
	return hasher.digest()
}

function orderedReadAssertions(
	workbook: Awaited<ReturnType<typeof Ascend.open>>,
	args: Args,
): Record<string, string | number | boolean | null> {
	const model = workbook.getWorkbookModel()
	const sheetNames = model.sheets.map((sheet) => sheet.name)
	let cellCount = 0
	let formulaCount = 0
	const usedRanges: string[] = []
	const physicalUsedRanges: string[] = []
	const orderedRefs = new OrderedLineHasher()
	const orderedValues = new OrderedLineHasher()
	const orderedFormulas = new OrderedLineHasher()
	for (const sheet of model.sheets) {
		formulaCount += sheet.cells.formulaCellCount()
		const usedRange = sheet.cells.usedRange()
		const usedRangeText = usedRange
			? `${sheet.name}!${indexToColumn(usedRange.start.col)}${usedRange.start.row + 1}:${indexToColumn(
					usedRange.end.col,
				)}${usedRange.end.row + 1}`
			: `${sheet.name}!empty`
		usedRanges.push(usedRangeText)
		physicalUsedRanges.push(usedRangeText)
		for (const [row, entries] of sheet.cells.iterateRows()) {
			for (const [col, cell] of entries) {
				cellCount++
				const ref = `${sheet.name}!${indexToColumn(col)}${row + 1}`
				orderedRefs.update(ref)
				orderedValues.update(`${ref}\t${serializeCellValue(cell.value)}`)
				if (cell.formula) orderedFormulas.update(`${ref}=${cell.formula}`)
			}
		}
	}
	return {
		sheetCount: model.sheets.length,
		sheetNamesHash: hashOrdered(sheetNames.map((name, index) => `${index}:${name}`).sort()),
		cellCount,
		physicalCellCount: cellCount,
		formulaCount,
		usedRangeCount: usedRanges.length,
		firstUsedRange: usedRanges[0] ?? null,
		firstPhysicalUsedRange: physicalUsedRanges[0] ?? null,
		usedRangesHash: hashOrdered([...usedRanges].sort()),
		physicalUsedRangesHash: hashOrdered([...physicalUsedRanges].sort()),
		orderedSemanticCellRefsHash: orderedRefs.digest(),
		orderedSemanticCellValuesHash: orderedValues.digest(),
		orderedFormulaTextHash: orderedFormulas.digest(),
		runnerVersion: 'workspace',
		runnerSource: args.source,
		runnerLoadMode: args.mode,
		runnerRichMetadata: args.richMetadata,
		runnerAssertionMode: 'ordered-hashes',
	}
}

async function streamedOrderedReadAssertions(
	args: Args,
	bytes: Uint8Array | undefined,
): Promise<Record<string, string | number | boolean | null> | undefined> {
	if (args.mode !== 'values' || args.richMetadata || args.source !== 'path') return undefined
	const direct = await directOrderedReadAssertions(args, bytes)
	if (direct) return direct
	const sourceBytes = bytes ?? (await Bun.file(args.file).bytes())
	const metadataWorkbook = await Ascend.open(sourceBytes, { mode: 'metadata-only' })
	const info = metadataWorkbook.inspect()
	if (info.load.sourceSheets.length !== 1) return undefined
	const sheetName = info.load.sourceSheets[0]
	if (!sheetName) return undefined
	const result = await readXlsxRowsStream(sourceBytes, {
		sheet: sheetName,
		mode: 'values',
		chunkedSheetXml: args.streamOrderedHashes,
	})
	if (!result.ok) throw new Error(result.error.message)

	let cellCount = 0
	let formulaCount = 0
	let minRow = Number.POSITIVE_INFINITY
	let minCol = Number.POSITIVE_INFINITY
	let maxRow = -1
	let maxCol = -1
	const orderedRefs = new OrderedLineHasher()
	const orderedValues = new OrderedLineHasher()
	const orderedFormulas = new OrderedLineHasher()
	for await (const row of result.value) {
		for (const [col, cell] of row.cells) {
			cellCount++
			minRow = Math.min(minRow, row.row)
			minCol = Math.min(minCol, col)
			maxRow = Math.max(maxRow, row.row)
			maxCol = Math.max(maxCol, col)
			const ref = `${sheetName}!${indexToColumn(col)}${row.row + 1}`
			orderedRefs.update(ref)
			orderedValues.update(`${ref}\t${serializeCellValue(cell.value)}`)
			if (cell.formula) {
				formulaCount++
				orderedFormulas.update(`${ref}=${cell.formula}`)
			}
		}
	}
	const usedRange =
		cellCount === 0
			? `${sheetName}!empty`
			: `${sheetName}!${indexToColumn(minCol)}${minRow + 1}:${indexToColumn(maxCol)}${maxRow + 1}`
	const usedRanges = [usedRange]
	return {
		sheetCount: 1,
		sheetNamesHash: hashOrdered([`0:${sheetName}`]),
		cellCount,
		physicalCellCount: cellCount,
		formulaCount,
		usedRangeCount: usedRanges.length,
		firstUsedRange: usedRanges[0] ?? null,
		firstPhysicalUsedRange: usedRanges[0] ?? null,
		usedRangesHash: hashOrdered([...usedRanges].sort()),
		physicalUsedRangesHash: hashOrdered([...usedRanges].sort()),
		orderedSemanticCellRefsHash: orderedRefs.digest(),
		orderedSemanticCellValuesHash: orderedValues.digest(),
		orderedFormulaTextHash: orderedFormulas.digest(),
		runnerVersion: 'workspace',
		runnerSource: args.source,
		runnerLoadMode: args.mode,
		runnerRichMetadata: args.richMetadata,
		runnerAssertionMode: 'ordered-hashes-stream',
	}
}

export async function directOrderedReadAssertions(
	args: Args,
	bytes: Uint8Array | undefined,
): Promise<Record<string, string | number | boolean | null> | undefined> {
	const sourceBytes = bytes ?? (await Bun.file(args.file).bytes())
	const archive = extractZip(sourceBytes)
	const rootRels = parseRelationships(archive.readText('_rels/.rels') ?? '')
	const docRel = rootRels.find((rel) => rel.type === REL_OFFICE_DOC)
	if (!docRel) return undefined
	const workbookPath = docRel.target.replace(/^\//, '')
	const workbookXml = archive.readText(workbookPath)
	if (!workbookXml) return undefined
	const workbookInfo = parseWorkbookXml(workbookXml)
	if (workbookInfo.sheets.length !== 1) return undefined
	const workbookRelsPath = workbookPath.replace(/(^|\/)([^/]+)$/, '$1_rels/$2.rels')
	const workbookRels = parseRelationships(archive.readText(workbookRelsPath) ?? '')
	const relMap = new Map(workbookRels.map((rel) => [rel.id, rel]))
	const sheet = workbookInfo.sheets[0]
	if (!sheet) return undefined
	const sheetRel = relMap.get(sheet.rId)
	if (!sheetRel) return undefined
	const sheetPath = resolvePath(workbookPath, sheetRel.target)
	if (!archive.has(sheetPath)) return undefined
	const sharedStringsRel = workbookRels.find((rel) => rel.type === REL_SHARED_STRINGS)
	const sharedStringsPath = sharedStringsRel
		? resolvePath(workbookPath, sharedStringsRel.target)
		: null
	const sharedStringsXml = sharedStringsPath ? archive.readText(sharedStringsPath) : undefined
	const sharedStrings = sharedStringsXml
		? parseSharedStrings(sharedStringsXml, { lazy: true })
		: emptySharedStrings()
	if (sharedStringsXml === undefined && process.env.ASCEND_ORDERED_BYTE_SCAN === '1') {
		return scanOrderedValuesSheetByteChunks(
			sheet.name,
			archive.readByteChunksAsync(sheetPath, 4 * 1024 * 1024),
			args,
		)
	}
	return scanOrderedValuesSheetTextChunks(
		sheet.name,
		process.env.ASCEND_ORDERED_SYNC_TEXT_SCAN === '1'
			? archive.readTextChunks(sheetPath, 4 * 1024 * 1024)
			: archive.readTextChunksAsync(sheetPath, 4 * 1024 * 1024),
		sharedStrings,
		args,
	)
}

async function scanOrderedValuesSheetTextChunks(
	sheetName: string,
	chunks: AsyncIterable<string> | Iterable<string>,
	sharedStrings: SharedStringResolver,
	args: Args,
): Promise<Record<string, string | number | boolean | null>> {
	const state = createOrderedSheetScanState(sheetName)
	let buffer = ''
	let started = false
	let scanOffset = 0
	let rowFallback = 0
	for await (const chunk of chunks) {
		buffer = scanOffset > 0 ? buffer.slice(scanOffset) + chunk : buffer + chunk
		scanOffset = 0
		if (!started) {
			const sheetDataOpen = buffer.indexOf('<sheetData', scanOffset)
			if (sheetDataOpen === -1) {
				buffer = trimXmlScanBuffer(buffer)
				continue
			}
			const sheetDataTagEnd = findTagEnd(buffer, sheetDataOpen)
			if (sheetDataTagEnd === -1) continue
			scanOffset = sheetDataTagEnd + 1
			started = true
		}
		while (true) {
			const sheetDataClose = buffer.indexOf('</sheetData>', scanOffset)
			const rowStart = buffer.indexOf('<row', scanOffset)
			if (rowStart === -1 || (sheetDataClose !== -1 && sheetDataClose < rowStart)) {
				if (sheetDataClose !== -1) return orderedSheetScanAssertions(state, args)
				if (scanOffset > 0) {
					buffer = buffer.slice(scanOffset)
					scanOffset = 0
				} else buffer = trimXmlScanBuffer(buffer)
				break
			}
			const rowTagEnd = findTagEnd(buffer, rowStart)
			if (rowTagEnd === -1) {
				buffer = buffer.slice(rowStart)
				scanOffset = 0
				break
			}
			if (isSelfClosingTag(buffer, rowStart, rowTagEnd)) {
				scanOffset = rowTagEnd + 1
				rowFallback += 1
				continue
			}
			const rowEnd = buffer.indexOf('</row>', rowTagEnd + 1)
			if (rowEnd === -1) {
				buffer = buffer.slice(rowStart)
				scanOffset = 0
				break
			}
			const parsedRow = parseNonNegativeIntAttr(buffer, rowStart, rowTagEnd, 'r')
			const row = parsedRow === undefined ? rowFallback : parsedRow - 1
			scanOrderedValuesRow(buffer, rowTagEnd + 1, rowEnd, row, state, sharedStrings)
			rowFallback = row + 1
			scanOffset = rowEnd + 6
			if (scanOffset > 8 * 1024 * 1024) {
				buffer = buffer.slice(scanOffset)
				scanOffset = 0
			}
		}
	}
	return orderedSheetScanAssertions(state, args)
}

const BYTE_DECODER = new TextDecoder('utf-8')
const SHEET_DATA_OPEN_BYTES = asciiBytes('<sheetData')
const SHEET_DATA_CLOSE_BYTES = asciiBytes('</sheetData>')
const ROW_OPEN_BYTES = asciiBytes('<row')
const ROW_CLOSE_BYTES = asciiBytes('</row>')
const CELL_OPEN_BYTES = asciiBytes('<c')
const CELL_CLOSE_BYTES = asciiBytes('</c>')
const VALUE_OPEN_BYTES = asciiBytes('<v>')
const VALUE_CLOSE_BYTES = asciiBytes('</v>')
const FORMULA_OPEN_BYTES = asciiBytes('<f')
const FORMULA_CLOSE_BYTES = asciiBytes('</f>')
const TEXT_OPEN_BYTES = asciiBytes('<t')
const TEXT_CLOSE_BYTES = asciiBytes('</t>')
const INLINE_STRING_OPEN_BYTES = asciiBytes('<is')
const CELL_REF_ATTR_BYTES = asciiBytes('r="')
const CELL_TYPE_ATTR_BYTES = asciiBytes('t="')
const TYPE_NONE = 0
const TYPE_SHARED_STRING = 1
const TYPE_INLINE_STRING = 2
const TYPE_BOOLEAN = 3
const TYPE_ERROR = 4
const TYPE_STRING = 5
const TYPE_NUMBER = 6

async function scanOrderedValuesSheetByteChunks(
	sheetName: string,
	chunks: AsyncIterable<Uint8Array>,
	args: Args,
): Promise<Record<string, string | number | boolean | null>> {
	const state = createOrderedSheetScanState(sheetName)
	let buffer = new Uint8Array()
	let started = false
	let scanOffset = 0
	let rowFallback = 0
	for await (const chunk of chunks) {
		buffer =
			scanOffset > 0 ? concatBytes(buffer.subarray(scanOffset), chunk) : concatBytes(buffer, chunk)
		scanOffset = 0
		if (!started) {
			const sheetDataOpen = indexOfBytes(buffer, SHEET_DATA_OPEN_BYTES, scanOffset)
			if (sheetDataOpen === -1) {
				buffer = trimByteScanBuffer(buffer)
				continue
			}
			const sheetDataTagEnd = findByteTagEnd(buffer, sheetDataOpen)
			if (sheetDataTagEnd === -1) continue
			scanOffset = sheetDataTagEnd + 1
			started = true
		}
		while (true) {
			const sheetDataClose = indexOfBytes(buffer, SHEET_DATA_CLOSE_BYTES, scanOffset)
			const rowStart = indexOfBytes(buffer, ROW_OPEN_BYTES, scanOffset)
			if (rowStart === -1 || (sheetDataClose !== -1 && sheetDataClose < rowStart)) {
				if (sheetDataClose !== -1) return orderedSheetScanAssertions(state, args)
				if (scanOffset > 0) {
					buffer = buffer.subarray(scanOffset)
					scanOffset = 0
				} else buffer = trimByteScanBuffer(buffer)
				break
			}
			const rowTagEnd = findByteTagEnd(buffer, rowStart)
			if (rowTagEnd === -1) {
				buffer = buffer.subarray(rowStart)
				scanOffset = 0
				break
			}
			if (isSelfClosingByteTag(buffer, rowStart, rowTagEnd)) {
				scanOffset = rowTagEnd + 1
				rowFallback += 1
				continue
			}
			const rowEnd = indexOfBytes(buffer, ROW_CLOSE_BYTES, rowTagEnd + 1)
			if (rowEnd === -1) {
				buffer = buffer.subarray(rowStart)
				scanOffset = 0
				break
			}
			const parsedRow = parseNonNegativeIntAttrBytes(buffer, rowStart, rowTagEnd, 114)
			const row = parsedRow === undefined ? rowFallback : parsedRow - 1
			scanOrderedValuesRowBytes(buffer, rowTagEnd + 1, rowEnd, row, state)
			rowFallback = row + 1
			scanOffset = rowEnd + ROW_CLOSE_BYTES.length
			if (scanOffset > 8 * 1024 * 1024) {
				buffer = buffer.subarray(scanOffset)
				scanOffset = 0
			}
		}
	}
	return orderedSheetScanAssertions(state, args)
}

function trimByteScanBuffer(buffer: Uint8Array): Uint8Array {
	return buffer.byteLength > 1_048_576 ? buffer.subarray(buffer.byteLength - 1024) : buffer
}

function trimXmlScanBuffer(buffer: string): string {
	return buffer.length > 1_048_576 ? buffer.slice(-1024) : buffer
}

function createOrderedSheetScanState(sheetName: string): OrderedSheetScanState {
	return {
		sheetName,
		cellCount: 0,
		physicalCellCount: 0,
		formulaCount: 0,
		minRow: Number.POSITIVE_INFINITY,
		minCol: Number.POSITIVE_INFINITY,
		maxRow: -1,
		maxCol: -1,
		physicalMinRow: Number.POSITIVE_INFINITY,
		physicalMinCol: Number.POSITIVE_INFINITY,
		physicalMaxRow: -1,
		physicalMaxCol: -1,
		columnLabels: [],
		columnRefPrefixes: [],
		orderedRefs: new OrderedLineHasher(),
		orderedValues: new OrderedLineHasher(),
		orderedFormulas: new OrderedLineHasher(),
	}
}

function orderedSheetScanAssertions(
	state: OrderedSheetScanState,
	args: Args,
): Record<string, string | number | boolean | null> {
	const sheetName = state.sheetName
	const usedRange =
		state.cellCount === 0
			? `${sheetName}!empty`
			: `${sheetName}!${columnLabel(state, state.minCol)}${state.minRow + 1}:${columnLabel(state, state.maxCol)}${state.maxRow + 1}`
	const physicalUsedRange =
		state.physicalCellCount === 0
			? `${sheetName}!empty`
			: `${sheetName}!${columnLabel(state, state.physicalMinCol)}${state.physicalMinRow + 1}:${columnLabel(state, state.physicalMaxCol)}${state.physicalMaxRow + 1}`
	const usedRanges = [usedRange]
	const physicalUsedRanges = [physicalUsedRange]
	return {
		sheetCount: 1,
		sheetNamesHash: hashOrdered([`0:${sheetName}`]),
		cellCount: state.cellCount,
		physicalCellCount: state.physicalCellCount,
		formulaCount: state.formulaCount,
		usedRangeCount: usedRanges.length,
		firstUsedRange: usedRanges[0] ?? null,
		firstPhysicalUsedRange: physicalUsedRanges[0] ?? null,
		usedRangesHash: hashOrdered([...usedRanges].sort()),
		physicalUsedRangesHash: hashOrdered([...physicalUsedRanges].sort()),
		orderedSemanticCellRefsHash: state.orderedRefs.digest(),
		orderedSemanticCellValuesHash: state.orderedValues.digest(),
		orderedFormulaTextHash: state.orderedFormulas.digest(),
		runnerVersion: 'workspace',
		runnerSource: args.source,
		runnerLoadMode: args.mode,
		runnerRichMetadata: args.richMetadata,
		runnerAssertionMode: 'ordered-hashes-stream-direct',
	}
}

interface OrderedSheetScanState {
	readonly sheetName: string
	cellCount: number
	physicalCellCount: number
	formulaCount: number
	minRow: number
	minCol: number
	maxRow: number
	maxCol: number
	physicalMinRow: number
	physicalMinCol: number
	physicalMaxRow: number
	physicalMaxCol: number
	readonly columnLabels: string[]
	readonly columnRefPrefixes: string[]
	readonly orderedRefs: OrderedLineHasher
	readonly orderedValues: OrderedLineHasher
	readonly orderedFormulas: OrderedLineHasher
}

function columnLabel(state: OrderedSheetScanState, col: number): string {
	const cached = state.columnLabels[col]
	if (cached) return cached
	const label = indexToColumn(col)
	state.columnLabels[col] = label
	return label
}

function cellRef(state: OrderedSheetScanState, col: number, rowText: string): string {
	const cached = state.columnRefPrefixes[col]
	if (cached) return cached + rowText
	const prefix = `${state.sheetName}!${columnLabel(state, col)}`
	state.columnRefPrefixes[col] = prefix
	return prefix + rowText
}

function scanOrderedValuesRow(
	xml: string,
	start: number,
	end: number,
	row: number,
	state: OrderedSheetScanState,
	sharedStrings: SharedStringResolver,
): void {
	let cursor = start
	let nextCol = 0
	const fastOut = { col: 0 }
	const rowText = String(row + 1)
	while (cursor < end) {
		const fastNext = scanCanonicalOrderedValuesCell(
			xml,
			cursor,
			end,
			row,
			rowText,
			nextCol,
			state,
			fastOut,
		)
		if (fastNext !== -1) {
			nextCol = fastOut.col + 1
			cursor = fastNext
			continue
		}
		const cellStart = xml.indexOf('<c', cursor)
		if (cellStart === -1 || cellStart >= end) return
		const cellTagEnd = findTagEnd(xml, cellStart)
		if (cellTagEnd === -1 || cellTagEnd >= end) return
		const ref = parseCellRefAttr(xml, cellStart, cellTagEnd)
		const cellRow = ref?.row ?? row
		const cellCol = ref?.col ?? nextCol
		trackPhysicalCell(state, cellRow, cellCol)
		nextCol = cellCol + 1
		const selfClosing = isSelfClosingTag(xml, cellStart, cellTagEnd)
		const cellEnd = selfClosing ? cellTagEnd + 1 : xml.indexOf('</c>', cellTagEnd + 1)
		if (cellEnd === -1 || cellEnd > end) return
		if (!selfClosing) {
			scanOrderedValuesCell(
				xml,
				cellStart,
				cellTagEnd,
				cellTagEnd + 1,
				cellEnd,
				cellRow,
				cellCol,
				state,
				sharedStrings,
			)
		}
		cursor = selfClosing ? cellTagEnd + 1 : cellEnd + 4
	}
}

function scanOrderedValuesRowBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	row: number,
	state: OrderedSheetScanState,
): void {
	let cursor = start
	let nextCol = 0
	const rowText = String(row + 1)
	while (cursor < end) {
		const cellStart = indexOfBytes(bytes, CELL_OPEN_BYTES, cursor)
		if (cellStart === -1 || cellStart >= end) return
		const cellTagEnd = findByteTagEnd(bytes, cellStart)
		if (cellTagEnd === -1 || cellTagEnd >= end) return
		const ref = parseCellRefAttrBytes(bytes, cellStart, cellTagEnd)
		const cellRow = ref?.row ?? row
		const cellCol = ref?.col ?? nextCol
		trackPhysicalCell(state, cellRow, cellCol)
		nextCol = cellCol + 1
		const selfClosing = isSelfClosingByteTag(bytes, cellStart, cellTagEnd)
		const cellEnd = selfClosing
			? cellTagEnd + 1
			: indexOfBytes(bytes, CELL_CLOSE_BYTES, cellTagEnd + 1)
		if (cellEnd === -1 || cellEnd > end) return
		if (!selfClosing) {
			scanOrderedValuesCellBytes(
				bytes,
				cellStart,
				cellTagEnd,
				cellTagEnd + 1,
				cellEnd,
				cellRow,
				cellCol,
				rowText,
				state,
			)
		}
		cursor = selfClosing ? cellTagEnd + 1 : cellEnd + CELL_CLOSE_BYTES.length
	}
}

function scanOrderedValuesCellBytes(
	bytes: Uint8Array,
	cellStart: number,
	cellTagEnd: number,
	bodyStart: number,
	bodyEnd: number,
	row: number,
	col: number,
	rowTextHint: string,
	state: OrderedSheetScanState,
): void {
	const formula = extractFormulaTextBytes(bytes, bodyStart, bodyEnd)
	if (formula !== undefined) {
		state.formulaCount += 1
		const ref = cellRef(state, col, String(row + 1))
		state.orderedFormulas.update(`${ref}=${formula}`)
	}
	const type = cellTypeAttrBytes(bytes, cellStart, cellTagEnd)
	const rawValue = extractTagTextBytes(
		bytes,
		bodyStart,
		bodyEnd,
		VALUE_OPEN_BYTES,
		VALUE_CLOSE_BYTES,
	)
	const value = serializeScannedCellValueBytes(
		type,
		rawValue,
		bytes,
		bodyStart,
		bodyEnd,
		formula !== undefined,
	)
	if (value === undefined) return
	trackSemanticCell(state, row, col)
	const rowText = rowTextHint === String(row + 1) ? rowTextHint : String(row + 1)
	const ref = cellRef(state, col, rowText)
	state.orderedRefs.update(ref)
	state.orderedValues.update(`${ref}\t${value}`)
}

function scanCanonicalOrderedValuesCell(
	xml: string,
	start: number,
	end: number,
	fallbackRow: number,
	fallbackRowText: string,
	fallbackCol: number,
	state: OrderedSheetScanState,
	out: { col: number },
): number {
	if (!xml.startsWith('<c r="', start)) return -1
	const parsed =
		parseExpectedCanonicalCellRef(xml, start + 6, end, fallbackRow, fallbackRowText, fallbackCol) ??
		parseCanonicalCellRef(xml, start + 6, end)
	if (!parsed) return -1
	trackPhysicalCell(state, parsed.row, parsed.col)
	const rowText = parsed.row === fallbackRow ? fallbackRowText : String(parsed.row + 1)
	const ref = cellRef(state, parsed.col, rowText)
	const numericValueStart = parseCanonicalNumericValueStart(xml, parsed.end)
	if (numericValueStart !== -1) {
		const valueEnd = xml.indexOf('</v></c>', numericValueStart)
		if (valueEnd === -1 || valueEnd > end || valueEnd === numericValueStart) return -1
		const raw = xml.slice(numericValueStart, valueEnd)
		const value = Number(raw)
		trackSemanticCell(state, parsed.row, parsed.col)
		state.orderedRefs.update(ref)
		state.orderedValues.update(
			`${ref}\t${Number.isNaN(value) ? `s:${raw}` : `n:${canonicalNumber(value)}`}`,
		)
		out.col = parsed.col
		return valueEnd + 8
	}
	const inlineValueStart = parseCanonicalInlineStringValueStart(xml, parsed.end)
	if (inlineValueStart !== -1) {
		let valueEnd = inlineValueStart
		let hasEntity = false
		while (valueEnd < end) {
			const code = xml.charCodeAt(valueEnd)
			if (code === 60) break
			if (code === 38) hasEntity = true
			valueEnd += 1
		}
		if (!xml.startsWith('</t></is></c>', valueEnd)) return -1
		const raw = xml.slice(inlineValueStart, valueEnd)
		trackSemanticCell(state, parsed.row, parsed.col)
		state.orderedRefs.update(ref)
		state.orderedValues.update(`${ref}\ts:${hasEntity ? decodeXmlText(raw) : raw}`)
		out.col = parsed.col
		return valueEnd + 13
	}
	return -1
}

function parseExpectedCanonicalCellRef(
	xml: string,
	start: number,
	end: number,
	row: number,
	rowText: string,
	col: number,
): { row: number; col: number; end: number } | undefined {
	let rowStart: number
	if (col >= 0 && col < 26) {
		if (start >= end || xml.charCodeAt(start) !== 65 + col) return undefined
		rowStart = start + 1
	} else {
		const colText = indexToColumn(col)
		if (!xml.startsWith(colText, start)) return undefined
		rowStart = start + colText.length
	}
	if (!xml.startsWith(rowText, rowStart)) return undefined
	const quote = rowStart + rowText.length
	if (quote >= end || xml.charCodeAt(quote) !== 34) return undefined
	return { row, col, end: quote }
}

function parseCanonicalNumericValueStart(xml: string, refEnd: number): number {
	if (xml.startsWith('"><v>', refEnd)) return refEnd + 5
	if (!xml.startsWith('" s="', refEnd)) return -1
	const styleEnd = xml.indexOf('"', refEnd + 5)
	if (styleEnd === -1 || !xml.startsWith('><v>', styleEnd + 1)) return -1
	return styleEnd + 5
}

function parseCanonicalInlineStringValueStart(xml: string, refEnd: number): number {
	const prefix = '" t="inlineStr"><is><t>'
	return xml.startsWith(prefix, refEnd) ? refEnd + prefix.length : -1
}

function parseCanonicalCellRef(
	xml: string,
	start: number,
	end: number,
): { row: number; col: number; end: number } | undefined {
	let cursor = start
	let col = 0
	while (cursor < end) {
		const code = xml.charCodeAt(cursor)
		if (code >= 48 && code <= 57) break
		if (code >= 65 && code <= 90) col = col * 26 + (code - 64)
		else if (code >= 97 && code <= 122) col = col * 26 + (code - 96)
		else return undefined
		cursor += 1
	}
	if (cursor === start || col <= 0) return undefined
	let row = 0
	while (cursor < end) {
		const code = xml.charCodeAt(cursor)
		if (code < 48 || code > 57) break
		row = row * 10 + (code - 48)
		cursor += 1
	}
	if (row <= 0 || cursor >= end || xml.charCodeAt(cursor) !== 34) return undefined
	return { row: row - 1, col: col - 1, end: cursor }
}

function scanOrderedValuesCell(
	xml: string,
	cellStart: number,
	cellTagEnd: number,
	bodyStart: number,
	bodyEnd: number,
	row: number,
	col: number,
	state: OrderedSheetScanState,
	sharedStrings: SharedStringResolver,
): void {
	const formula = extractFormulaText(xml, bodyStart, bodyEnd)
	if (formula !== undefined) {
		state.formulaCount += 1
		const ref = cellRef(state, col, String(row + 1))
		state.orderedFormulas.update(`${ref}=${formula}`)
	}
	const type = stringAttr(xml, cellStart, cellTagEnd, 't')
	const rawValue = extractTagText(xml, bodyStart, bodyEnd, 'v')
	const value = serializeScannedCellValue(
		type,
		rawValue,
		xml,
		bodyStart,
		bodyEnd,
		sharedStrings,
		formula !== undefined,
	)
	if (value === undefined) return
	trackSemanticCell(state, row, col)
	const ref = cellRef(state, col, String(row + 1))
	state.orderedRefs.update(ref)
	state.orderedValues.update(`${ref}\t${value}`)
}

function serializeScannedCellValue(
	type: string | undefined,
	rawValue: string | undefined,
	xml: string,
	bodyStart: number,
	bodyEnd: number,
	sharedStrings: SharedStringResolver,
	hasFormula: boolean,
): string | undefined {
	if (type === 's') {
		const index = rawValue === undefined ? -1 : Number.parseInt(rawValue, 10)
		return `s:${Number.isFinite(index) ? (sharedStrings.getString?.(index) ?? '') : ''}`
	}
	if (type === 'inlineStr' || xml.indexOf('<is', bodyStart) !== -1) {
		return `s:${extractInlineStringText(xml, bodyStart, bodyEnd)}`
	}
	if (type === 'b') return `b:${rawValue === '1' ? 'true' : 'false'}`
	if (type === 'e') return `e:${rawValue ?? '#VALUE!'}`
	if (type === 'str') return `s:${rawValue ?? ''}`
	if (rawValue !== undefined && rawValue !== '') {
		const value = Number(rawValue)
		return Number.isNaN(value) ? `s:${rawValue}` : `n:${canonicalNumber(value)}`
	}
	if (hasFormula) return 'empty'
	if (type === 'n') return 'n:0'
	if (type) return 'empty'
	return undefined
}

function extractFormulaText(xml: string, start: number, end: number): string | undefined {
	const cursor = start
	while (cursor < end) {
		const formulaStart = xml.indexOf('<f', cursor)
		if (formulaStart === -1 || formulaStart >= end) return undefined
		const tagEnd = findTagEnd(xml, formulaStart)
		if (tagEnd === -1 || tagEnd >= end) return undefined
		if (isSelfClosingTag(xml, formulaStart, tagEnd)) return ''
		const close = xml.indexOf('</f>', tagEnd + 1)
		if (close === -1 || close > end) return undefined
		const raw = xml.slice(tagEnd + 1, close)
		return raw.includes('&') ? decodeXmlText(raw) : raw
	}
	return undefined
}

function extractInlineStringText(xml: string, start: number, end: number): string {
	let cursor = start
	let value = ''
	while (cursor < end) {
		const textStart = xml.indexOf('<t', cursor)
		if (textStart === -1 || textStart >= end) break
		const tagEnd = findTagEnd(xml, textStart)
		if (tagEnd === -1 || tagEnd >= end) break
		if (isSelfClosingTag(xml, textStart, tagEnd)) {
			cursor = tagEnd + 1
			continue
		}
		const close = xml.indexOf('</t>', tagEnd + 1)
		if (close === -1 || close > end) break
		const raw = xml.slice(tagEnd + 1, close)
		value += raw.includes('&') ? decodeXmlText(raw) : raw
		cursor = close + 4
	}
	return value
}

function extractTagText(
	xml: string,
	start: number,
	end: number,
	tagName: string,
): string | undefined {
	const open = xml.indexOf(`<${tagName}>`, start)
	if (open === -1 || open >= end) return undefined
	const valueStart = open + tagName.length + 2
	const close = xml.indexOf(`</${tagName}>`, valueStart)
	if (close === -1 || close > end) return undefined
	const raw = xml.slice(valueStart, close)
	return raw.includes('&') ? decodeXmlText(raw) : raw
}

function serializeScannedCellValueBytes(
	type: number,
	rawValue: Uint8Array | undefined,
	bytes: Uint8Array,
	bodyStart: number,
	bodyEnd: number,
	hasFormula: boolean,
): string | undefined {
	if (type === TYPE_SHARED_STRING)
		return `s:${rawValue === undefined ? '' : decodeByteText(rawValue)}`
	const inlineStart = indexOfBytes(bytes, INLINE_STRING_OPEN_BYTES, bodyStart)
	if (type === TYPE_INLINE_STRING || (inlineStart !== -1 && inlineStart < bodyEnd)) {
		return `s:${extractInlineStringTextBytes(bytes, bodyStart, bodyEnd)}`
	}
	if (type === TYPE_BOOLEAN)
		return `b:${rawValue && rawValue.byteLength > 0 && rawValue[0] === 49 ? 'true' : 'false'}`
	if (type === TYPE_ERROR)
		return `e:${rawValue === undefined ? '#VALUE!' : decodeByteText(rawValue)}`
	if (type === TYPE_STRING) return `s:${rawValue === undefined ? '' : decodeByteText(rawValue)}`
	if (rawValue !== undefined && rawValue.byteLength > 0) {
		const raw = decodeByteText(rawValue)
		const value = Number(raw)
		return Number.isNaN(value) ? `s:${raw}` : `n:${canonicalNumber(value)}`
	}
	if (hasFormula) return 'empty'
	if (type === TYPE_NUMBER) return 'n:0'
	if (type !== TYPE_NONE) return 'empty'
	return undefined
}

function extractFormulaTextBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
): string | undefined {
	const formulaStart = indexOfBytes(bytes, FORMULA_OPEN_BYTES, start)
	if (formulaStart === -1 || formulaStart >= end) return undefined
	const tagEnd = findByteTagEnd(bytes, formulaStart)
	if (tagEnd === -1 || tagEnd >= end) return undefined
	if (isSelfClosingByteTag(bytes, formulaStart, tagEnd)) return ''
	const close = indexOfBytes(bytes, FORMULA_CLOSE_BYTES, tagEnd + 1)
	if (close === -1 || close > end) return undefined
	return decodeByteText(bytes.subarray(tagEnd + 1, close))
}

function extractInlineStringTextBytes(bytes: Uint8Array, start: number, end: number): string {
	let cursor = start
	let value = ''
	while (cursor < end) {
		const textStart = indexOfBytes(bytes, TEXT_OPEN_BYTES, cursor)
		if (textStart === -1 || textStart >= end) break
		const tagEnd = findByteTagEnd(bytes, textStart)
		if (tagEnd === -1 || tagEnd >= end) break
		if (isSelfClosingByteTag(bytes, textStart, tagEnd)) {
			cursor = tagEnd + 1
			continue
		}
		const close = indexOfBytes(bytes, TEXT_CLOSE_BYTES, tagEnd + 1)
		if (close === -1 || close > end) break
		value += decodeByteText(bytes.subarray(tagEnd + 1, close))
		cursor = close + TEXT_CLOSE_BYTES.length
	}
	return value
}

function extractTagTextBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	openNeedle: Uint8Array,
	closeNeedle: Uint8Array,
): Uint8Array | undefined {
	const open = indexOfBytes(bytes, openNeedle, start)
	if (open === -1 || open >= end) return undefined
	const valueStart = open + openNeedle.length
	const close = indexOfBytes(bytes, closeNeedle, valueStart)
	if (close === -1 || close > end) return undefined
	return bytes.subarray(valueStart, close)
}

function decodeByteText(bytes: Uint8Array): string {
	const text = BYTE_DECODER.decode(bytes)
	return hasByte(bytes, 38) ? decodeXmlText(text) : text
}

function trackSemanticCell(state: OrderedSheetScanState, row: number, col: number): void {
	state.cellCount += 1
	if (row < state.minRow) state.minRow = row
	if (col < state.minCol) state.minCol = col
	if (row > state.maxRow) state.maxRow = row
	if (col > state.maxCol) state.maxCol = col
}

function trackPhysicalCell(state: OrderedSheetScanState, row: number, col: number): void {
	state.physicalCellCount += 1
	if (row < state.physicalMinRow) state.physicalMinRow = row
	if (col < state.physicalMinCol) state.physicalMinCol = col
	if (row > state.physicalMaxRow) state.physicalMaxRow = row
	if (col > state.physicalMaxCol) state.physicalMaxCol = col
}

function parseNonNegativeIntAttr(
	xml: string,
	start: number,
	end: number,
	name: string,
): number | undefined {
	const raw = stringAttr(xml, start, end, name)
	if (!raw) return undefined
	const value = Number.parseInt(raw, 10)
	return Number.isFinite(value) && value >= 0 ? value : undefined
}

function stringAttr(xml: string, start: number, end: number, name: string): string | undefined {
	const needle = `${name}="`
	let cursor = start
	while (cursor < end) {
		const attrStart = xml.indexOf(needle, cursor)
		if (attrStart === -1 || attrStart >= end) return undefined
		const before = attrStart === start ? ' ' : xml[attrStart - 1]
		if (before === ' ' || before === '\n' || before === '\r' || before === '\t') {
			const valueStart = attrStart + needle.length
			const valueEnd = xml.indexOf('"', valueStart)
			if (valueEnd !== -1 && valueEnd <= end) return xml.slice(valueStart, valueEnd)
		}
		cursor = attrStart + needle.length
	}
	return undefined
}

function parseCellRefAttr(
	xml: string,
	start: number,
	end: number,
): { row: number; col: number } | undefined {
	const valueStart = attrValueStart(xml, start, end, 'r')
	if (valueStart === -1) return undefined
	let cursor = valueStart
	let col = 0
	while (cursor < end) {
		const code = xml.charCodeAt(cursor)
		if (code >= 48 && code <= 57) break
		if (code >= 65 && code <= 90) col = col * 26 + (code - 64)
		else if (code >= 97 && code <= 122) col = col * 26 + (code - 96)
		else return undefined
		cursor += 1
	}
	if (cursor === valueStart || col <= 0) return undefined
	let row = 0
	while (cursor < end) {
		const code = xml.charCodeAt(cursor)
		if (code < 48 || code > 57) break
		row = row * 10 + (code - 48)
		cursor += 1
	}
	if (row <= 0 || cursor >= end || xml.charCodeAt(cursor) !== 34) return undefined
	return { row: row - 1, col: col - 1 }
}

function attrValueStart(xml: string, start: number, end: number, name: string): number {
	const needle = `${name}="`
	let cursor = start
	while (cursor < end) {
		const attrStart = xml.indexOf(needle, cursor)
		if (attrStart === -1 || attrStart >= end) return -1
		const before = attrStart === start ? ' ' : xml[attrStart - 1]
		if (before === ' ' || before === '\n' || before === '\r' || before === '\t') {
			return attrStart + needle.length
		}
		cursor = attrStart + needle.length
	}
	return -1
}

function asciiBytes(value: string): Uint8Array {
	const bytes = new Uint8Array(value.length)
	for (let i = 0; i < value.length; i++) bytes[i] = value.charCodeAt(i)
	return bytes
}

function concatBytes(left: Uint8Array, right: Uint8Array): Uint8Array {
	if (left.byteLength === 0) return right
	if (right.byteLength === 0) return left
	const out = new Uint8Array(left.byteLength + right.byteLength)
	out.set(left)
	out.set(right, left.byteLength)
	return out
}

function indexOfBytes(bytes: Uint8Array, needle: Uint8Array, start: number): number {
	if (needle.byteLength === 0) return start
	const first = needle[0]
	const limit = bytes.byteLength - needle.byteLength
	for (let i = Math.max(0, start); i <= limit; i++) {
		if (bytes[i] !== first) continue
		let matches = true
		for (let j = 1; j < needle.byteLength; j++) {
			if (bytes[i + j] === needle[j]) continue
			matches = false
			break
		}
		if (matches) return i
	}
	return -1
}

function findByteTagEnd(bytes: Uint8Array, start: number): number {
	let quote = 0
	for (let i = start; i < bytes.byteLength; i++) {
		const code = bytes[i]
		if (quote !== 0) {
			if (code === quote) quote = 0
			continue
		}
		if (code === 34 || code === 39) {
			quote = code
			continue
		}
		if (code === 62) return i
	}
	return -1
}

function isSelfClosingByteTag(bytes: Uint8Array, start: number, tagEnd: number): boolean {
	let cursor = tagEnd - 1
	while (cursor > start) {
		const code = bytes[cursor]
		if (code === 32 || code === 9 || code === 10 || code === 13) {
			cursor--
			continue
		}
		return code === 47
	}
	return false
}

function parseNonNegativeIntAttrBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
	name: number,
): number | undefined {
	const valueStart = byteAttrValueStart(bytes, start, end, name)
	if (valueStart === -1) return undefined
	let value = 0
	let cursor = valueStart
	while (cursor < end) {
		const code = bytes[cursor]
		if (code === 34) return value
		if (code < 48 || code > 57) return undefined
		value = value * 10 + (code - 48)
		cursor++
	}
	return undefined
}

function parseCellRefAttrBytes(
	bytes: Uint8Array,
	start: number,
	end: number,
): { row: number; col: number } | undefined {
	const valueStart = byteAttrValueStart(bytes, start, end, 114)
	if (valueStart === -1) return undefined
	let cursor = valueStart
	let col = 0
	while (cursor < end) {
		const code = bytes[cursor]
		if (code >= 48 && code <= 57) break
		if (code >= 65 && code <= 90) col = col * 26 + (code - 64)
		else if (code >= 97 && code <= 122) col = col * 26 + (code - 96)
		else return undefined
		cursor += 1
	}
	if (cursor === valueStart || col <= 0) return undefined
	let row = 0
	while (cursor < end) {
		const code = bytes[cursor]
		if (code < 48 || code > 57) break
		row = row * 10 + (code - 48)
		cursor += 1
	}
	if (row <= 0 || cursor >= end || bytes[cursor] !== 34) return undefined
	return { row: row - 1, col: col - 1 }
}

function cellTypeAttrBytes(bytes: Uint8Array, start: number, end: number): number {
	const valueStart = indexOfBytes(bytes, CELL_TYPE_ATTR_BYTES, start)
	if (valueStart === -1 || valueStart >= end) return TYPE_NONE
	const before = valueStart === start ? 32 : bytes[valueStart - 1]
	if (before !== 32 && before !== 9 && before !== 10 && before !== 13) return TYPE_NONE
	const cursor = valueStart + CELL_TYPE_ATTR_BYTES.length
	const first = bytes[cursor]
	const second = bytes[cursor + 1]
	const third = bytes[cursor + 2]
	if (first === 115 && second === 34) return TYPE_SHARED_STRING
	if (first === 98 && second === 34) return TYPE_BOOLEAN
	if (first === 101 && second === 34) return TYPE_ERROR
	if (first === 110 && second === 34) return TYPE_NUMBER
	if (first === 115 && second === 116 && third === 114 && bytes[cursor + 3] === 34) {
		return TYPE_STRING
	}
	if (
		first === 105 &&
		second === 110 &&
		third === 108 &&
		bytes[cursor + 3] === 105 &&
		bytes[cursor + 4] === 110 &&
		bytes[cursor + 5] === 101 &&
		bytes[cursor + 6] === 83 &&
		bytes[cursor + 7] === 116 &&
		bytes[cursor + 8] === 114 &&
		bytes[cursor + 9] === 34
	) {
		return TYPE_INLINE_STRING
	}
	return TYPE_NONE
}

function byteAttrValueStart(bytes: Uint8Array, start: number, end: number, name: number): number {
	let cursor = start
	while (cursor < end) {
		const attrStart = indexOfBytes(
			bytes,
			name === 114 ? CELL_REF_ATTR_BYTES : CELL_TYPE_ATTR_BYTES,
			cursor,
		)
		if (attrStart === -1 || attrStart >= end) return -1
		const before = attrStart === start ? 32 : bytes[attrStart - 1]
		if (before === 32 || before === 9 || before === 10 || before === 13) {
			return attrStart + 3
		}
		cursor = attrStart + 3
	}
	return -1
}

function hasByte(bytes: Uint8Array, value: number): boolean {
	for (let i = 0; i < bytes.byteLength; i++) {
		if (bytes[i] === value) return true
	}
	return false
}

function readFeatureAssertions(
	workbook: ReturnType<Awaited<ReturnType<typeof Ascend.open>>['getWorkbookModel']>,
): {
	readonly readCommentCount: number
	readonly readHyperlinkCount: number
	readonly readDataValidationCount: number
	readonly readConditionalFormatCount: number
	readonly readDefinedNameCount: number
} {
	let readCommentCount = 0
	let readHyperlinkCount = 0
	let readDataValidationCount = 0
	let readConditionalFormatCount = 0
	for (const sheet of workbook.sheets) {
		readCommentCount += sheet.comments.size
		readHyperlinkCount += sheet.hyperlinks.size
		readDataValidationCount += sheet.dataValidations.length
		readConditionalFormatCount += sheet.conditionalFormats.length
	}
	return {
		readCommentCount,
		readHyperlinkCount,
		readDataValidationCount,
		readConditionalFormatCount,
		readDefinedNameCount: workbook.definedNames.size,
	}
}

function memorySample(durationMs: number): {
	readonly durationMs: number
	readonly rssAfterBytes: number
	readonly peakRssBytes: number
	readonly heapUsedBytes: number
	readonly heapTotalBytes: number
	readonly rssAfterGcBytes: number
	readonly heapAfterGcBytes: number
} {
	const memory = process.memoryUsage()
	const rss = typeof memory.rss === 'function' ? memory.rss() : memory.rss
	runGc()
	const afterGc = process.memoryUsage()
	const rssAfterGc = typeof afterGc.rss === 'function' ? afterGc.rss() : afterGc.rss
	return {
		durationMs,
		rssAfterBytes: rss,
		peakRssBytes: Math.max(rss, rssAfterGc),
		heapUsedBytes: memory.heapUsed,
		heapTotalBytes: memory.heapTotal,
		rssAfterGcBytes: rssAfterGc,
		heapAfterGcBytes: afterGc.heapUsed,
	}
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		/* best effort */
	}
}

async function main(): Promise<void> {
	const args = parseArgs()
	const bytes = args.source === 'bytes' ? await Bun.file(args.file).bytes() : undefined
	for (let i = 0; i < args.warmup; i++) {
		if (args.streamOrderedHashes) {
			const assertions = await streamedOrderedReadAssertions(args, bytes)
			if (assertions) continue
		}
		const workbook = await openWorkbookFromSource(args, bytes)
		readAssertions(workbook, args)
	}
	const samples: ReturnType<typeof memorySample>[] = []
	let assertions: Record<string, string | number | boolean | null> | undefined
	for (let i = 0; i < args.repeat; i++) {
		const start = performance.now()
		if (args.streamOrderedHashes) {
			assertions = await streamedOrderedReadAssertions(args, bytes)
			if (!assertions) {
				const workbook = await openWorkbookFromSource(args, bytes)
				assertions = readAssertions(workbook, args)
			}
		} else {
			const workbook = await openWorkbookFromSource(args, bytes)
			assertions = readAssertions(workbook, args)
		}
		const durationMs = performance.now() - start
		samples.push(memorySample(durationMs))
	}
	const payload = { assertions: assertions ?? {}, samples }
	console.log(args.json ? JSON.stringify(payload) : JSON.stringify(payload, null, 2))
}

if (import.meta.main) await main()
