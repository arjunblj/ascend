import { readFile } from 'node:fs/promises'
import type { Workbook } from '@ascend/core'
import { readCsv } from '@ascend/io-csv'
import { type PreservationCapsule, type ReadXlsxLoadInfo, readXlsx } from '@ascend/io-xlsx'
import {
	AscendException,
	type CompatibilityReport,
	type CsvDialect,
	emptyReport,
} from '@ascend/schema'
import type { WorkbookLoadInfo } from './types.ts'

const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04]
const COMPOUND_FILE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

export interface LoadedWorkbookSource {
	readonly workbook: Workbook
	readonly capsules: readonly PreservationCapsule[]
	readonly report: CompatibilityReport
	readonly loadInfo: WorkbookLoadInfo
	readonly originalBytes: Uint8Array | null
}

export async function openWorkbookSource(
	pathOrBytes: string | Uint8Array,
	options?: {
		mode?: 'full' | 'metadata-only' | 'values' | 'formula'
		sheets?: readonly string[]
		maxRows?: number
		richMetadata?: boolean
		password?: string
		pivotCacheRecordMaterializeLimit?: number | 'all'
	},
): Promise<LoadedWorkbookSource> {
	let bytes: Uint8Array
	let ext = ''

	if (typeof pathOrBytes === 'string') {
		ext = pathOrBytes.split('.').pop()?.toLowerCase() ?? ''
		bytes =
			typeof Bun !== 'undefined'
				? await Bun.file(pathOrBytes).bytes()
				: new Uint8Array(await readFile(pathOrBytes))
	} else {
		bytes = pathOrBytes
	}

	if (ext === 'csv' || ext === 'tsv') {
		const text = new TextDecoder().decode(bytes)
		const dialect: Partial<CsvDialect> | undefined = ext === 'tsv' ? { delimiter: '\t' } : undefined
		const result = readCsv(text, dialect)
		if (!result.ok) throw new AscendException(result.error)
		return {
			workbook: result.value,
			capsules: [],
			report: emptyReport('csv'),
			loadInfo: buildWorkbookLoadInfo({
				mode: 'full',
				isPartial: false,
				cellsHydrated: true,
				richSheetMetadataHydrated: true,
				hasAllSheets: true,
				sourceSheetNames: result.value.sheets.map((sheet) => sheet.name),
				loadedSheetNames: result.value.sheets.map((sheet) => sheet.name),
			}),
			originalBytes: null,
		}
	}

	if (ext === 'xlsx' || ext === 'xlsm' || isZip(bytes) || isCompoundFile(bytes)) {
		const result = readXlsx(bytes, options)
		if (!result.ok) throw new AscendException(result.error)
		const loadInfo = buildWorkbookLoadInfo(result.value.loadInfo)
		if (loadInfo.isPartial) {
			result.value.workbook.sourceArchiveBytes = null
		}
		return {
			workbook: result.value.workbook,
			capsules: result.value.capsules,
			report: result.value.report,
			loadInfo,
			originalBytes: loadInfo.isPartial
				? null
				: (result.value.workbook.sourceArchiveBytes ?? bytes),
		}
	}

	const text = new TextDecoder().decode(bytes)
	const result = readCsv(text)
	if (!result.ok) throw new AscendException(result.error)
	return {
		workbook: result.value,
		capsules: [],
		report: emptyReport('csv'),
		loadInfo: buildWorkbookLoadInfo({
			mode: 'full',
			isPartial: false,
			cellsHydrated: true,
			richSheetMetadataHydrated: true,
			hasAllSheets: true,
			sourceSheetNames: result.value.sheets.map((sheet) => sheet.name),
			loadedSheetNames: result.value.sheets.map((sheet) => sheet.name),
		}),
		originalBytes: null,
	}
}

function isZip(bytes: Uint8Array): boolean {
	if (bytes.length < 4) return false
	return (
		bytes[0] === ZIP_MAGIC[0] &&
		bytes[1] === ZIP_MAGIC[1] &&
		bytes[2] === ZIP_MAGIC[2] &&
		bytes[3] === ZIP_MAGIC[3]
	)
}

function isCompoundFile(bytes: Uint8Array): boolean {
	if (bytes.length < COMPOUND_FILE_MAGIC.length) return false
	return COMPOUND_FILE_MAGIC.every((byte, index) => bytes[index] === byte)
}

export function buildWorkbookLoadInfo(info: ReadXlsxLoadInfo): WorkbookLoadInfo {
	return {
		mode: info.mode,
		isPartial: info.isPartial,
		cellsHydrated: info.cellsHydrated,
		richSheetMetadataHydrated: info.richSheetMetadataHydrated,
		hasAllSheets: info.hasAllSheets,
		sourceSheets: info.sourceSheetNames,
		loadedSheets: info.loadedSheetNames,
	}
}
