import type { AscendError, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { escapeXml } from '../xml.ts'
import { createZip, encode, StreamingZipBuilder, type ZipCompressionProfile } from './zip.ts'

export type DenseXlsxCellValue = string | number | boolean | null
export type DenseXlsxCellValueType = 'boolean' | 'number' | 'string'
export type DenseXlsxCompressionProfile = ZipCompressionProfile

export interface WriteDenseRowsXlsxOptions {
	readonly rows: number
	readonly cols: number
	readonly sheetName?: string
	readonly valueAt: (row: number, col: number) => DenseXlsxCellValue
	readonly omitCellRefs?: boolean
	/** Skip row indexes for sequential generated rows; empty rows are still emitted to preserve order. */
	readonly omitRowRefs?: boolean
	/** Reuse generated XML for consecutive rows with identical values. */
	readonly cacheRepeatedRows?: boolean
	/** Treat every row as identical to row 0 and call valueAt only for the first row. */
	readonly constantRows?: boolean
	/** Skip XML escaping for generated string values that cannot contain XML-sensitive characters. */
	readonly stringsAreXmlSafe?: boolean
	/** Hint for dense generated values; mismatches fall back to generic cell serialization. */
	readonly valueType?: DenseXlsxCellValueType
	/** Per-column value hints for dense generated rows; mismatches fall back to generic cell serialization. */
	readonly valueTypes?: readonly (DenseXlsxCellValueType | undefined)[]
	/** Hint that every generated cell is present; null mismatches fall back to referenced cells. */
	readonly allCellsPresent?: boolean
	/** ZIP compression policy: fast optimizes wall time; compact optimizes file size. */
	readonly compressionProfile?: DenseXlsxCompressionProfile
}

const DEFAULT_SHEET_NAME = 'Data'
const XML_BYTE_BATCH_TARGET = positiveEnvInt('ASCEND_DENSE_XML_BYTE_BATCH', 256 * 1024)
const XML_ROW_BATCH_TARGET = positiveEnvInt('ASCEND_DENSE_XML_ROW_BATCH', 256)

interface XmlByteSink {
	write(chunk: string): void
	end(): Uint8Array
}

type BunArrayBufferSink = {
	start(options?: { asUint8Array?: boolean; highWaterMark?: number }): void
	write(chunk: string): number | undefined
	end(): ArrayBuffer | Uint8Array
}

type BunArrayBufferSinkConstructor = new () => BunArrayBufferSink

export function writeDenseRowsXlsx(
	options: WriteDenseRowsXlsxOptions,
): Result<Uint8Array, AscendError> {
	try {
		validateDenseRowsOptions(options)
		const parts = buildBaseParts(options.sheetName ?? DEFAULT_SHEET_NAME)
		parts.set('xl/worksheets/sheet1.xml', buildSheetXml(options))
		return ok(createZip(parts, zipOptionsForDenseRows(options)))
	} catch (error) {
		return err(toExportError(error))
	}
}

export async function writeDenseRowsXlsxStreaming(
	options: WriteDenseRowsXlsxOptions,
): Promise<Result<Uint8Array, AscendError>> {
	try {
		validateDenseRowsOptions(options)
		const builder = new StreamingZipBuilder(zipOptionsForDenseRows(options))
		for (const [path, data] of buildBaseParts(options.sheetName ?? DEFAULT_SHEET_NAME)) {
			builder.addEntry(path, data)
		}
		builder.addStreamingEntry('xl/worksheets/sheet1.xml', estimateDenseSheetXmlBytes(options))
		await writeSheetXmlByteChunksAsync(options, (chunk) => builder.writeChunkAsync(chunk))
		await builder.closeEntry()
		return ok(builder.finalize())
	} catch (error) {
		return err(toExportError(error))
	}
}

function validateDenseRowsOptions(options: WriteDenseRowsXlsxOptions): void {
	if (!Number.isInteger(options.rows) || options.rows < 0) {
		throw new Error(`rows must be a non-negative integer, received ${options.rows}`)
	}
	if (!Number.isInteger(options.cols) || options.cols < 0) {
		throw new Error(`cols must be a non-negative integer, received ${options.cols}`)
	}
	if (options.rows > 1_048_576) {
		throw new Error(`rows exceeds Excel worksheet limit: ${options.rows}`)
	}
	if (options.cols > 16_384) {
		throw new Error(`cols exceeds Excel worksheet limit: ${options.cols}`)
	}
	if (typeof options.valueAt !== 'function') {
		throw new Error('valueAt must be a function')
	}
	if (options.valueTypes !== undefined && options.valueTypes.length !== options.cols) {
		throw new Error(
			`valueTypes length must match cols: ${options.valueTypes.length} !== ${options.cols}`,
		)
	}
	if (
		options.compressionProfile !== undefined &&
		options.compressionProfile !== 'fast' &&
		options.compressionProfile !== 'compact'
	) {
		throw new Error(`Unsupported compressionProfile "${options.compressionProfile}"`)
	}
}

function zipOptionsForDenseRows(options: WriteDenseRowsXlsxOptions): {
	readonly compressionProfile?: DenseXlsxCompressionProfile
} {
	return options.compressionProfile === undefined
		? {}
		: { compressionProfile: options.compressionProfile }
}

function toExportError(error: unknown): AscendError {
	const message = error instanceof Error ? error.message : String(error)
	return ascendError('EXPORT_ERROR', `Failed to write dense rows XLSX: ${message}`)
}

function buildSheetXml(options: WriteDenseRowsXlsxOptions): Uint8Array {
	const sink = createArrayBufferXmlSink()
	if (!sink) return buildSheetXmlJoined(options)

	const rowState = createRowXmlState(options)
	sink.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
	sink.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
	sink.write(`<dimension ref="${sheetDimension(options.cols, options.rows)}"/><sheetData>`)
	for (let row = 0; row < options.rows; row++) {
		const rowXml = buildRowXml(options, row, rowState)
		if (rowXml.length > 0) sink.write(rowXml)
	}
	sink.write('</sheetData></worksheet>')
	return sink.end()
}

function buildSheetXmlJoined(options: WriteDenseRowsXlsxOptions): Uint8Array {
	const rowState = createRowXmlState(options)
	const out: string[] = [
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
		'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
		`<dimension ref="${sheetDimension(options.cols, options.rows)}"/>`,
		'<sheetData>',
	]
	for (let row = 0; row < options.rows; row++) {
		const rowXml = buildRowXml(options, row, rowState)
		if (rowXml.length > 0) out.push(rowXml)
	}
	out.push('</sheetData></worksheet>')
	return encode(out.join(''))
}

function createXmlByteSink(highWaterMark?: number): XmlByteSink {
	return createArrayBufferXmlSink(highWaterMark) ?? createJoinedXmlSink()
}

function createJoinedXmlSink(): XmlByteSink {
	const chunks: string[] = []
	return {
		write: (chunk) => {
			chunks.push(chunk)
		},
		end: () => encode(chunks.join('')),
	}
}

function createArrayBufferXmlSink(highWaterMark?: number): XmlByteSink | undefined {
	const sink = createBunArrayBufferSink()
	if (!sink) return undefined
	sink.start(
		highWaterMark === undefined ? { asUint8Array: true } : { asUint8Array: true, highWaterMark },
	)
	return {
		write: (chunk) => {
			sink.write(chunk)
		},
		end: () => {
			const bytes = sink.end()
			return bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
		},
	}
}

function createBunArrayBufferSink(): BunArrayBufferSink | undefined {
	const ctor = (
		globalThis as { readonly Bun?: { readonly ArrayBufferSink?: BunArrayBufferSinkConstructor } }
	).Bun?.ArrayBufferSink
	if (typeof ctor !== 'function') return undefined
	return new ctor()
}

function positiveEnvInt(name: string, fallback: number): number {
	const raw = (
		globalThis as { readonly process?: { readonly env?: Record<string, string | undefined> } }
	).process?.env?.[name]
	if (!raw) return fallback
	const value = Number.parseInt(raw, 10)
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function estimateDenseSheetXmlBytes(options: WriteDenseRowsXlsxOptions): number {
	const cellCount = options.rows * options.cols
	const rowOverhead = options.rows * 18
	const cellOverhead =
		options.valueType === 'string'
			? 36
			: options.valueType === 'number'
				? 20
				: options.valueType === 'boolean'
					? 18
					: 28
	return 512 + rowOverhead + cellCount * cellOverhead
}

async function writeSheetXmlByteChunksAsync(
	options: WriteDenseRowsXlsxOptions,
	onChunk: (chunk: Uint8Array) => Promise<void>,
): Promise<void> {
	const rowState = createRowXmlState(options)
	await onChunk(encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'))
	await onChunk(
		encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'),
	)
	await onChunk(
		encode(`<dimension ref="${sheetDimension(options.cols, options.rows)}"/><sheetData>`),
	)
	let batch = createXmlByteSink(XML_BYTE_BATCH_TARGET)
	let batchChars = 0
	let batchRows = 0
	for (let row = 0; row < options.rows; row++) {
		const rowXml = buildRowXml(options, row, rowState)
		if (rowXml.length === 0) continue
		batch.write(rowXml)
		batchChars += rowXml.length
		batchRows++
		if (batchRows >= XML_ROW_BATCH_TARGET || batchChars >= XML_BYTE_BATCH_TARGET) {
			await onChunk(batch.end())
			batch = createXmlByteSink(XML_BYTE_BATCH_TARGET)
			batchChars = 0
			batchRows = 0
		}
	}
	if (batchRows > 0) await onChunk(batch.end())
	await onChunk(encode('</sheetData></worksheet>'))
}

interface RowXmlState {
	constantRows: boolean
	constantBody: string | null | undefined
	constantFallback: boolean
	lastValues: DenseXlsxCellValue[] | null
	lastBody: string
}

function createRowXmlState(options: WriteDenseRowsXlsxOptions): RowXmlState | undefined {
	return (options.cacheRepeatedRows === true || options.constantRows === true) &&
		options.omitCellRefs === true
		? {
				constantRows: options.constantRows === true,
				constantBody: undefined,
				constantFallback: false,
				lastValues: null,
				lastBody: '',
			}
		: undefined
}

function buildRowXml(options: WriteDenseRowsXlsxOptions, row: number, state?: RowXmlState): string {
	if (options.omitCellRefs === true) return buildDenseRowXmlWithoutRefs(options, row, state)

	const cells: string[] = []
	const values: DenseXlsxCellValue[] = []
	let hasValue = false
	for (let col = 0; col < options.cols; col++) {
		const value = options.valueAt(row, col)
		values.push(value)
		if (value !== null) hasValue = true
	}
	for (let col = 0; col < values.length; col++) {
		const value = values[col] ?? null
		const xml = cellXml(
			value,
			`${columnName(col)}${row + 1}`,
			options.stringsAreXmlSafe,
			options.valueType,
		)
		if (xml.length > 0) cells.push(xml)
	}
	if (!hasValue) return emptyRowXml(options)
	return `${rowOpenXml(options, row)}${cells.join('')}</row>`
}

function buildDenseRowXmlWithoutRefs(
	options: WriteDenseRowsXlsxOptions,
	row: number,
	state?: RowXmlState,
): string {
	if (state?.constantRows === true && row > 0 && !state.constantFallback) {
		if (state.constantBody === null) return ''
		if (state.constantBody !== undefined) {
			return `${rowOpenXml(options, row)}${state.constantBody}</row>`
		}
	}
	if (state === undefined && options.allCellsPresent === true) {
		return buildPresentDenseRowXmlWithoutRefs(options, row)
	}

	let body = ''
	const values: DenseXlsxCellValue[] | undefined = state ? [] : undefined
	let hasValue = false
	let hasNull = false
	for (let col = 0; col < options.cols; col++) {
		const value = options.valueAt(row, col)
		values?.push(value)
		if (value === null) {
			hasNull = true
		} else {
			hasValue = true
		}
		body += cellXml(value, undefined, options.stringsAreXmlSafe, options.valueType)
	}
	if (!hasValue) {
		if (state?.constantRows === true) state.constantBody = null
		return emptyRowXml(options)
	}
	if (hasNull) {
		if (state?.constantRows === true) state.constantFallback = true
		return buildRowXmlWithRefs(options, row)
	}
	if (state && values && sameValues(values, state.lastValues)) {
		return `${rowOpenXml(options, row)}${state.lastBody}</row>`
	}
	if (state && values) {
		state.lastValues = values
		state.lastBody = body
		if (state.constantRows) state.constantBody = body
	}
	return `${rowOpenXml(options, row)}${body}</row>`
}

function buildPresentDenseRowXmlWithoutRefs(
	options: WriteDenseRowsXlsxOptions,
	row: number,
): string {
	if (options.valueTypes !== undefined) {
		return buildPresentColumnTypedRowXmlWithoutRefs(options, row, options.valueTypes)
	}
	if (options.valueType === 'string' && options.stringsAreXmlSafe === true) {
		return buildPresentSafeStringRowXmlWithoutRefs(options, row)
	}
	if (options.valueType === 'number') return buildPresentNumberRowXmlWithoutRefs(options, row)
	if (options.valueType === 'boolean') return buildPresentBooleanRowXmlWithoutRefs(options, row)
	return buildPresentGenericRowXmlWithoutRefs(options, row)
}

function buildPresentGenericRowXmlWithoutRefs(
	options: WriteDenseRowsXlsxOptions,
	row: number,
): string {
	let body = ''
	for (let col = 0; col < options.cols; col++) {
		const value = options.valueAt(row, col)
		if (value === null) return buildRowXmlWithRefs(options, row)
		body += cellXml(value, undefined, options.stringsAreXmlSafe, options.valueType)
	}
	return body.length === 0 ? '' : `${rowOpenXml(options, row)}${body}</row>`
}

function buildPresentColumnTypedRowXmlWithoutRefs(
	options: WriteDenseRowsXlsxOptions,
	row: number,
	valueTypes: readonly (DenseXlsxCellValueType | undefined)[],
): string {
	let body = ''
	for (let col = 0; col < options.cols; col++) {
		const value = options.valueAt(row, col)
		if (value === null) return buildRowXmlWithRefs(options, row)
		const valueType = valueTypes[col]
		if (valueType === 'number') {
			if (typeof value !== 'number') return buildPresentGenericRowXmlWithoutRefs(options, row)
			// biome-ignore lint/style/useTemplate: Bun/JSC is measurably faster with concatenation here.
			body += '<c><v>' + value + '</v></c>'
			continue
		}
		if (valueType === 'string') {
			if (typeof value !== 'string') return buildPresentGenericRowXmlWithoutRefs(options, row)
			body +=
				'<c t="str"><v>' +
				(options.stringsAreXmlSafe === true ? value : escapeXml(value)) +
				'</v></c>'
			continue
		}
		if (valueType === 'boolean') {
			if (typeof value !== 'boolean') return buildPresentGenericRowXmlWithoutRefs(options, row)
			body += value ? '<c t="b"><v>1</v></c>' : '<c t="b"><v>0</v></c>'
			continue
		}
		body += cellXml(value, undefined, options.stringsAreXmlSafe, options.valueType)
	}
	return body.length === 0 ? '' : `${rowOpenXml(options, row)}${body}</row>`
}

function buildPresentSafeStringRowXmlWithoutRefs(
	options: WriteDenseRowsXlsxOptions,
	row: number,
): string {
	let body = ''
	for (let col = 0; col < options.cols; col++) {
		const value = options.valueAt(row, col)
		if (value === null) return buildRowXmlWithRefs(options, row)
		if (typeof value !== 'string') return buildPresentGenericRowXmlWithoutRefs(options, row)
		// biome-ignore lint/style/useTemplate: Bun/JSC is measurably faster with concatenation here.
		body += '<c t="str"><v>' + value + '</v></c>'
	}
	return body.length === 0 ? '' : `${rowOpenXml(options, row)}${body}</row>`
}

function buildPresentNumberRowXmlWithoutRefs(
	options: WriteDenseRowsXlsxOptions,
	row: number,
): string {
	let body = ''
	for (let col = 0; col < options.cols; col++) {
		const value = options.valueAt(row, col)
		if (value === null) return buildRowXmlWithRefs(options, row)
		if (typeof value !== 'number') return buildPresentGenericRowXmlWithoutRefs(options, row)
		// biome-ignore lint/style/useTemplate: Bun/JSC is measurably faster with concatenation here.
		body += '<c><v>' + value + '</v></c>'
	}
	return body.length === 0 ? '' : `${rowOpenXml(options, row)}${body}</row>`
}

function buildPresentBooleanRowXmlWithoutRefs(
	options: WriteDenseRowsXlsxOptions,
	row: number,
): string {
	let body = ''
	for (let col = 0; col < options.cols; col++) {
		const value = options.valueAt(row, col)
		if (value === null) return buildRowXmlWithRefs(options, row)
		if (typeof value !== 'boolean') return buildPresentGenericRowXmlWithoutRefs(options, row)
		body += value ? '<c t="b"><v>1</v></c>' : '<c t="b"><v>0</v></c>'
	}
	return body.length === 0 ? '' : `${rowOpenXml(options, row)}${body}</row>`
}

function buildRowXmlWithRefs(options: WriteDenseRowsXlsxOptions, row: number): string {
	const cells: string[] = []
	let hasValue = false
	for (let col = 0; col < options.cols; col++) {
		const value = options.valueAt(row, col)
		if (value === null) continue
		hasValue = true
		cells.push(
			cellXml(value, `${columnName(col)}${row + 1}`, options.stringsAreXmlSafe, options.valueType),
		)
	}
	if (!hasValue) return emptyRowXml(options)
	return `${rowOpenXml(options, row)}${cells.join('')}</row>`
}

function rowOpenXml(options: WriteDenseRowsXlsxOptions, row: number): string {
	return options.omitRowRefs === true ? '<row>' : `<row r="${row + 1}">`
}

function emptyRowXml(options: WriteDenseRowsXlsxOptions): string {
	return options.omitRowRefs === true ? '<row></row>' : ''
}

function sameValues(
	values: readonly DenseXlsxCellValue[],
	previous: readonly DenseXlsxCellValue[] | null,
): boolean {
	if (!previous || values.length !== previous.length) return false
	for (let i = 0; i < values.length; i++) {
		if (values[i] !== previous[i]) return false
	}
	return true
}

function cellXml(
	value: DenseXlsxCellValue,
	ref: string | undefined,
	stringsAreXmlSafe = false,
	valueType?: DenseXlsxCellValueType,
): string {
	if (value === null) return ref === undefined ? '<c/>' : ''
	const r = ref === undefined ? '' : ` r="${ref}"`
	if (valueType === 'number' && typeof value === 'number') return `<c${r}><v>${value}</v></c>`
	if (valueType === 'boolean' && typeof value === 'boolean') {
		return `<c${r} t="b"><v>${value ? 1 : 0}</v></c>`
	}
	if (valueType === 'string' && typeof value === 'string') {
		return `<c${r} t="str"><v>${stringsAreXmlSafe ? value : escapeXml(value)}</v></c>`
	}
	if (typeof value === 'number') return `<c${r}><v>${value}</v></c>`
	if (typeof value === 'boolean') return `<c${r} t="b"><v>${value ? 1 : 0}</v></c>`
	return `<c${r} t="str"><v>${stringsAreXmlSafe ? value : escapeXml(value)}</v></c>`
}

function sheetDimension(cols: number, rows: number): string {
	return `A1:${columnName(Math.max(0, cols - 1))}${Math.max(1, rows)}`
}

function columnName(index: number): string {
	let n = index + 1
	let column = ''
	while (n > 0) {
		const rem = (n - 1) % 26
		column = String.fromCharCode(65 + rem) + column
		n = Math.floor((n - 1) / 26)
	}
	return column
}

function buildBaseParts(sheetName: string): Map<string, Uint8Array> {
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
<sheets><sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/></sheets>
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
