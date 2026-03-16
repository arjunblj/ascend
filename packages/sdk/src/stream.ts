import { readFile } from 'node:fs/promises'
import {
	readXlsxRowsStream,
	type StreamedSheetRow,
	type StreamXlsxRowsOptions,
	type XlsxByteSource,
} from '@ascend/io-xlsx'
import { AscendException } from '@ascend/schema'

export type WorkbookRowStreamSource = string | XlsxByteSource

export async function streamWorkbookRows(
	source: WorkbookRowStreamSource,
	options: StreamXlsxRowsOptions = {},
): Promise<AsyncGenerator<StreamedSheetRow>> {
	const normalized = typeof source === 'string' ? await loadBytes(source) : source
	const result = await readXlsxRowsStream(normalized, options)
	if (!result.ok) throw new AscendException(result.error)
	return result.value
}

async function loadBytes(path: string): Promise<Uint8Array> {
	if (typeof Bun !== 'undefined') {
		return Bun.file(path).bytes()
	}
	return new Uint8Array(await readFile(path))
}
