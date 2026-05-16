export type {
	InspectXlsxPackageGraphOptions,
	XlsxPackageContentTypeDefault,
	XlsxPackageContentTypeOverride,
	XlsxPackageContentTypeSource,
	XlsxPackageGraph,
	XlsxPackageGraphPart,
	XlsxPackageGraphRelationship,
	XlsxPackageLossPolicy,
	XlsxPackageOwnerScope,
} from './package-graph.ts'
export { classifyPackageFeatureFamily, inspectXlsxPackageGraph } from './package-graph.ts'
export type {
	XlsxPackageGraphFidelityIssue,
	XlsxPackageGraphFidelityIssueCode,
	XlsxPackageGraphReadIntegrityOptions,
	XlsxPackageGraphSafeEditIntegrityOptions,
} from './package-graph-fidelity.ts'
export {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
} from './package-graph-fidelity.ts'
export type { PreservationCapsule } from './preserve.ts'
export type { ReadXlsxLoadInfo, ReadXlsxOptions, ReadXlsxResult } from './reader/index.ts'
export { readXlsx, readXlsxArchive } from './reader/index.ts'
export type { StreamedSheetRow } from './reader/sheet.ts'
export type { StreamXlsxRowsOptions, XlsxByteSource } from './reader/stream.ts'
export { readXlsxRowsStream } from './reader/stream.ts'
export type { ZipArchive } from './reader/zip.ts'
export { extractZip, extractZipFromFile } from './reader/zip.ts'
export type {
	DenseXlsxCellValue,
	DenseXlsxCompressionProfile,
	WriteDenseRowsXlsxOptions,
} from './writer/dense-rows.ts'
export { writeDenseRowsXlsx, writeDenseRowsXlsxStreaming } from './writer/dense-rows.ts'
export type { DirtyCellPatch, WriteXlsxOptions } from './writer/index.ts'
export {
	planWriteXlsx,
	summarizePlannedWrite,
	writeXlsx,
	writeXlsxStreaming,
} from './writer/index.ts'
export type { WritePlanResult, WritePlanSummary } from './writer/plan.ts'
export type { ZipCompressionProfile } from './writer/zip.ts'
