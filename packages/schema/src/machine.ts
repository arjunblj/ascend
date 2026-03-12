import type { AscendError } from './errors.ts'

export const MACHINE_FORMAT_VERSION = 1 as const

export interface MachineSuccess<T> {
	readonly formatVersion: typeof MACHINE_FORMAT_VERSION
	readonly ok: true
	readonly data: T
}

export interface MachineFailure {
	readonly formatVersion: typeof MACHINE_FORMAT_VERSION
	readonly ok: false
	readonly error: {
		readonly message: string
		readonly code?: string
		readonly retryable?: boolean
		readonly refs?: readonly string[]
		readonly details?: Record<string, unknown>
		readonly suggestedFix?: string
	}
}

export type MachineEnvelope<T> = MachineSuccess<T> | MachineFailure

export function machineSuccess<T>(data: T): MachineSuccess<T> {
	return {
		formatVersion: MACHINE_FORMAT_VERSION,
		ok: true,
		data,
	}
}

export function machineFailure(error: string | AscendError): MachineFailure {
	if (typeof error === 'string') {
		return {
			formatVersion: MACHINE_FORMAT_VERSION,
			ok: false,
			error: { message: error },
		}
	}
	return {
		formatVersion: MACHINE_FORMAT_VERSION,
		ok: false,
		error: {
			message: error.message,
			code: error.code,
			retryable: error.retryable,
			...(error.refs ? { refs: error.refs } : {}),
			...(error.details ? { details: error.details } : {}),
			...(error.suggestedFix ? { suggestedFix: error.suggestedFix } : {}),
		},
	}
}
