export type ErrorCode =
	| 'FORMULA_PARSE_ERROR'
	| 'FORMULA_EVAL_ERROR'
	| 'INVALID_REF'
	| 'INVALID_RANGE'
	| 'SHEET_NOT_FOUND'
	| 'TABLE_NOT_FOUND'
	| 'NAME_NOT_FOUND'
	| 'NAME_CONFLICT'
	| 'IMPORT_ERROR'
	| 'EXPORT_ERROR'
	| 'UNSUPPORTED_FORMAT'
	| 'CORRUPT_FILE'
	| 'CAPSULE_INVALID'
	| 'CIRCULAR_REF'
	| 'VALIDATION_ERROR'
	| 'PROTECTION_ERROR'
	| 'MERGE_CONFLICT'
	| 'STYLE_ERROR'

export interface AscendError {
	readonly code: ErrorCode
	readonly message: string
	readonly retryable: boolean
	readonly refs?: readonly string[]
	readonly details?: Record<string, unknown>
	readonly suggestedFix?: string
}

export type Result<T, E = AscendError> =
	| { readonly ok: true; readonly value: T }
	| { readonly ok: false; readonly error: E }

export function ok<T>(value: T): Result<T, never> {
	return { ok: true, value }
}

export function err<E>(error: E): Result<never, E> {
	return { ok: false, error }
}

export function ascendError(
	code: ErrorCode,
	message: string,
	opts?: Partial<Pick<AscendError, 'retryable' | 'refs' | 'details' | 'suggestedFix'>>,
): AscendError {
	return {
		code,
		message,
		retryable: opts?.retryable ?? false,
		...(opts?.refs !== undefined ? { refs: opts.refs } : {}),
		...(opts?.details !== undefined ? { details: opts.details } : {}),
		...(opts?.suggestedFix !== undefined ? { suggestedFix: opts.suggestedFix } : {}),
	}
}
