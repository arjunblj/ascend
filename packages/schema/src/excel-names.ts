export interface ExcelNameValidationIssue {
	readonly message: string
	readonly suggestedFix: string
}

const WORKSHEET_NAME_SUGGESTED_FIX =
	'Use a worksheet name that is 1 to 31 characters and does not contain colon, backslash, slash, question mark, asterisk, left bracket, right bracket, or leading/trailing apostrophes.'

const TABLE_NAME_SUGGESTED_FIX =
	'Use a table name that starts with a letter, underscore, or backslash; uses only letters, numbers, periods, and underscores after that; is not C, R, A1-style, or R1C1-style; and is 255 characters or fewer.'

const DEFINED_NAME_SUGGESTED_FIX =
	'Use a defined name that starts with a letter, underscore, or backslash; uses only letters, numbers, periods, and underscores after that; is not C, R, A1-style, or R1C1-style; and is 255 characters or fewer.'

const INVALID_WORKSHEET_NAME_CHARS = /[:\\/?*[\]]/

export function validateExcelWorksheetName(name: string): ExcelNameValidationIssue | null {
	if (name.length === 0) {
		return issue('Sheet name cannot be empty', WORKSHEET_NAME_SUGGESTED_FIX)
	}
	if (name.length > 31) {
		return issue(`Sheet name "${name}" exceeds 31 characters`, WORKSHEET_NAME_SUGGESTED_FIX)
	}
	if (INVALID_WORKSHEET_NAME_CHARS.test(name)) {
		return issue(`Sheet name "${name}" contains invalid characters`, WORKSHEET_NAME_SUGGESTED_FIX)
	}
	if (name.startsWith("'") || name.endsWith("'")) {
		return issue(
			`Sheet name "${name}" cannot start or end with an apostrophe`,
			WORKSHEET_NAME_SUGGESTED_FIX,
		)
	}
	return null
}

export function validateExcelTableName(name: string): ExcelNameValidationIssue | null {
	if (name.length === 0) {
		return issue('Table name cannot be empty', TABLE_NAME_SUGGESTED_FIX)
	}
	if (name.length > 255) {
		return issue(`Table name "${name}" exceeds 255 characters`, TABLE_NAME_SUGGESTED_FIX)
	}
	if (/^[cr]$/i.test(name)) {
		return issue(`Table name "${name}" is reserved`, TABLE_NAME_SUGGESTED_FIX)
	}
	if (isA1StyleReference(name) || /^R\d+C\d+$/i.test(name)) {
		return issue(`Table name "${name}" cannot be a cell reference`, TABLE_NAME_SUGGESTED_FIX)
	}
	if (!/^[\p{L}_\\][\p{L}\p{N}._]*$/u.test(name)) {
		return issue(`Table name "${name}" contains invalid characters`, TABLE_NAME_SUGGESTED_FIX)
	}
	return null
}

export function validateExcelDefinedName(name: string): ExcelNameValidationIssue | null {
	if (name.length === 0) {
		return issue('Defined name cannot be empty', DEFINED_NAME_SUGGESTED_FIX)
	}
	if (name.length > 255) {
		return issue(`Defined name "${name}" exceeds 255 characters`, DEFINED_NAME_SUGGESTED_FIX)
	}
	if (/^[cr]$/i.test(name)) {
		return issue(`Defined name "${name}" is reserved`, DEFINED_NAME_SUGGESTED_FIX)
	}
	if (isA1StyleReference(name) || /^R\d+C\d+$/i.test(name)) {
		return issue(`Defined name "${name}" cannot be a cell reference`, DEFINED_NAME_SUGGESTED_FIX)
	}
	if (!/^[\p{L}_\\][\p{L}\p{N}._]*$/u.test(name)) {
		return issue(`Defined name "${name}" contains invalid characters`, DEFINED_NAME_SUGGESTED_FIX)
	}
	return null
}

function issue(message: string, suggestedFix: string): ExcelNameValidationIssue {
	return { message, suggestedFix }
}

function isA1StyleReference(name: string): boolean {
	const match = /^([A-Za-z]+)(\d+)$/.exec(name)
	if (!match?.[1] || !match[2]) return false
	const row = Number.parseInt(match[2], 10)
	const col = columnToIndex(match[1])
	return col >= 0 && col <= 16_383 && row >= 1 && row <= 1_048_576
}

function columnToIndex(label: string): number {
	let index = 0
	for (const char of label.toUpperCase()) {
		const code = char.charCodeAt(0)
		if (code < 65 || code > 90) return -1
		index = index * 26 + (code - 64)
	}
	return index - 1
}
