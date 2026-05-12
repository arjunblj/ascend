import type { StyleId } from '@ascend/core'
import type { AscendError } from '@ascend/schema'
import { ascendError, err, ok, type Result } from '@ascend/schema'
import {
	getRelsPath,
	parseRelationships,
	REL_OFFICE_DOC,
	REL_SHARED_STRINGS,
	REL_STYLES,
	resolvePath,
} from './relationships.ts'
import { emptySharedStrings, parseSharedStrings } from './shared-strings.ts'
import {
	type StreamedSheetRow,
	streamSheetRowsBytes,
	streamSheetRowsTextChunks,
	streamSheetRowsXml,
	ValueInternPool,
} from './sheet.ts'
import { parseStylesLite } from './styles.ts'
import { parseWorkbookXml } from './workbook.ts'
import { extractZip, type ZipArchive } from './zip.ts'

export type XlsxByteSource = Uint8Array | AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>

export interface StreamXlsxRowsOptions {
	readonly sheet?: string | number
	readonly mode?: 'values' | 'formula'
	readonly chunkedSheetXml?: boolean
}

export async function readXlsxRowsStream(
	source: XlsxByteSource,
	options: StreamXlsxRowsOptions = {},
): Promise<Result<AsyncGenerator<StreamedSheetRow>, AscendError>> {
	const bytes = source instanceof Uint8Array ? source : await collectBytes(source)
	let archive: ZipArchive
	try {
		archive = extractZip(bytes)
	} catch (error) {
		return err(
			ascendError(
				'CORRUPT_FILE',
				`Invalid ZIP: ${error instanceof Error ? error.message : 'unknown'}`,
			),
		)
	}
	const bootstrap = await archive.readTextsAsync(['[Content_Types].xml', '_rels/.rels'])
	const rootRelsXml = bootstrap.get('_rels/.rels')
	if (!rootRelsXml) {
		return err(ascendError('CORRUPT_FILE', 'Missing _rels/.rels'))
	}
	const rootRels = parseRelationships(rootRelsXml)
	const docRel = rootRels.find((rel) => rel.type === REL_OFFICE_DOC)
	if (!docRel) {
		return err(ascendError('CORRUPT_FILE', 'No officeDocument relationship found'))
	}
	const workbookPath = docRel.target.replace(/^\//, '')
	const workbookParts = await archive.readTextsAsync([workbookPath, getRelsPath(workbookPath)])
	const workbookXml = workbookParts.get(workbookPath)
	if (!workbookXml) {
		return err(ascendError('CORRUPT_FILE', `Missing workbook: ${workbookPath}`))
	}
	const workbookInfo = parseWorkbookXml(workbookXml)
	const workbookRels = parseRelationships(workbookParts.get(getRelsPath(workbookPath)) ?? '')
	const relMap = new Map(workbookRels.map((rel) => [rel.id, rel]))
	const sheetEntry = resolveSheet(workbookInfo.sheets, options.sheet)
	if (!sheetEntry) {
		return err(ascendError('INVALID_ARGUMENT', 'Requested sheet was not found'))
	}
	const sheetRel = relMap.get(sheetEntry.rId)
	if (!sheetRel) {
		return err(ascendError('CORRUPT_FILE', `Missing relationship for sheet ${sheetEntry.name}`))
	}
	const sheetPath = resolvePath(workbookPath, sheetRel.target)
	const sharedStringsRel = workbookRels.find((rel) => rel.type === REL_SHARED_STRINGS)
	const stylesRel = workbookRels.find((rel) => rel.type === REL_STYLES)
	const sharedStringsPath = sharedStringsRel
		? resolvePath(workbookPath, sharedStringsRel.target)
		: null
	const stylesPath = stylesRel ? resolvePath(workbookPath, stylesRel.target) : null
	if (!archive.has(sheetPath))
		return err(ascendError('CORRUPT_FILE', `Missing worksheet: ${sheetPath}`))
	const mode = options.mode ?? 'values'
	const chunkedSheetXml = options.chunkedSheetXml === true
	const sheetZipEntry = archive.get(sheetPath)
	const shouldChunkSheetXml =
		chunkedSheetXml &&
		(sheetZipEntry?.uncompressedSize ?? Number.POSITIVE_INFINITY) > 64 * 1024 * 1024
	const shouldParseSheetBytes = mode === 'values' && !shouldChunkSheetXml
	const parsePaths =
		mode === 'values' && (shouldChunkSheetXml || shouldParseSheetBytes)
			? [sharedStringsPath, stylesPath].filter((path): path is string => path !== null)
			: [sheetPath, sharedStringsPath, stylesPath].filter((path): path is string => path !== null)
	const parseParts = await archive.readTextsAsync(parsePaths)
	const valuePool = mode === 'values' ? undefined : new ValueInternPool()
	const sharedStringsXml = sharedStringsPath ? parseParts.get(sharedStringsPath) : undefined
	const sharedStrings = sharedStringsXml
		? parseSharedStrings(sharedStringsXml, {
				normalize: (value) => valuePool?.internValue(value) ?? value,
				lazy: false,
			})
		: emptySharedStrings()
	const stylesXml = stylesPath ? parseParts.get(stylesPath) : undefined
	const stylesLite = stylesXml
		? parseStylesLite(stylesXml)
		: { isDateFormat: [false], metadata: undefined }
	const styleCount = Math.max(1, stylesLite.isDateFormat.length)
	const styleIds = Array.from({ length: styleCount }, (_, index) => index as StyleId)
	const ctx = {
		sharedStrings,
		styleIds,
		isDateFormat: stylesLite.isDateFormat,
		hasDateStyles: stylesLite.isDateFormat.some(Boolean),
		...(valuePool ? { valuePool } : {}),
		valuesOnly: mode === 'values',
		formulaOnly: mode === 'formula',
	}
	const sheetBytes = shouldParseSheetBytes ? archive.readBytes(sheetPath) : undefined
	const iterator =
		mode === 'values' && shouldChunkSheetXml
			? streamSheetRowsTextChunks(sheetEntry.name, archive.readTextChunks(sheetPath), ctx)
			: sheetBytes
				? streamSheetRowsBytes(sheetEntry.name, sheetBytes, ctx)
				: streamSheetRowsXml(
						sheetEntry.name,
						parseParts.get(sheetPath) ?? archive.readText(sheetPath) ?? '',
						ctx,
					)
	return ok(
		(async function* () {
			for (const row of iterator) yield row
		})(),
	)
}

async function collectBytes(
	source: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>,
): Promise<Uint8Array> {
	const chunks: Uint8Array[] = []
	let total = 0
	for await (const chunk of toAsyncIterable(source)) {
		chunks.push(chunk)
		total += chunk.byteLength
	}
	const bytes = new Uint8Array(total)
	let offset = 0
	for (let i = 0; i < chunks.length; i++) {
		const chunk = chunks[i] as Uint8Array
		bytes.set(chunk, offset)
		offset += chunk.byteLength
	}
	return bytes
}

async function* toAsyncIterable(
	source: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>,
): AsyncGenerator<Uint8Array> {
	if (isAsyncIterable(source)) {
		yield* source
		return
	}
	const reader = (source as ReadableStream<Uint8Array>).getReader()
	try {
		while (true) {
			const next = await reader.read()
			if (next.done) return
			yield next.value
		}
	} finally {
		reader.releaseLock()
	}
}

function isAsyncIterable(
	source: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>,
): source is AsyncIterable<Uint8Array> {
	return typeof (source as AsyncIterable<Uint8Array>)[Symbol.asyncIterator] === 'function'
}

function resolveSheet(
	sheets: readonly { readonly name: string; readonly rId: string }[],
	selector: string | number | undefined,
): { readonly name: string; readonly rId: string } | undefined {
	if (selector === undefined) return sheets[0]
	if (typeof selector === 'number') return sheets[selector]
	const target = selector.toLowerCase()
	return sheets.find((sheet) => sheet.name.toLowerCase() === target)
}
