export type { PreservationCapsule } from './preserve.ts'
export type { ReadXlsxLoadInfo, ReadXlsxOptions, ReadXlsxResult } from './reader/index.ts'
export { readXlsx } from './reader/index.ts'
export type { StreamedSheetRow } from './reader/sheet.ts'
export type { StreamXlsxRowsOptions, XlsxByteSource } from './reader/stream.ts'
export { readXlsxRowsStream } from './reader/stream.ts'
export type { ZipArchive } from './reader/zip.ts'
export { extractZip } from './reader/zip.ts'
export type {
	DenseXlsxCellValue,
	DenseXlsxCompressionProfile,
	WriteDenseRowsXlsxOptions,
} from './writer/dense-rows.ts'
export { writeDenseRowsXlsx, writeDenseRowsXlsxStreaming } from './writer/dense-rows.ts'
export type { WriteXlsxOptions } from './writer/index.ts'
export {
	planWriteXlsx,
	summarizePlannedWrite,
	writeXlsx,
	writeXlsxStreaming,
} from './writer/index.ts'
export type { WritePlanResult, WritePlanSummary } from './writer/plan.ts'
