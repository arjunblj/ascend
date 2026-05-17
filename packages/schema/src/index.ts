export type {
	CompatibilityReport,
	CompatibilityStatus,
	CompatibilityTier,
	FeatureReport,
} from './compatibility.ts'
export { emptyReport } from './compatibility.ts'
export type {
	CalcSettings,
	CalcSettingsAttribute,
	CsvDialect,
	ExportOptions,
	ImportOptions,
} from './config.ts'
export { DEFAULT_CALC_SETTINGS, DEFAULT_CSV_DIALECT } from './config.ts'
export type {
	AscendError,
	ErrorCode,
	Result,
} from './errors.ts'
export { AscendException, ascendError, assertUnreachable, err, ok } from './errors.ts'
export type { ExcelNameValidationIssue } from './excel-names.ts'
export {
	validateExcelDefinedName,
	validateExcelTableName,
	validateExcelWorksheetName,
} from './excel-names.ts'
export { levenshtein } from './levenshtein.ts'
export type { MachineEnvelope, MachineFailure, MachineSuccess } from './machine.ts'
export { MACHINE_FORMAT_VERSION, machineFailure, machineSuccess } from './machine.ts'
export type {
	CellUpdate,
	ConditionalFormatRule,
	DataValidationRule,
	Operation,
	PasteMode,
	SortSpec,
	StyleInput,
} from './operations.ts'
export type {
	ArrayValue,
	CellValue,
	ExcelError,
	InputValue,
	RichTextColor,
	RichTextRun,
	ScalarCellValue,
} from './values.ts'
export {
	arrayValue,
	booleanValue,
	coerceCellValueToString,
	dateValue,
	EMPTY,
	errorValue,
	isArrayValue,
	isEmpty,
	isError,
	numberValue,
	richTextValue,
	stringValue,
	topLeftScalar,
	valuesEqual,
} from './values.ts'
