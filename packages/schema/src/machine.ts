/**
 * IO-3 Binary snapshot format assessment:
 *
 * The MachineEnvelope currently uses JSON serialization. For large workbooks,
 * a binary format could reduce serialization overhead and payload size.
 *
 * Candidate approaches:
 *
 * 1. MessagePack / CBOR: Drop-in JSON-compatible binary encodings. ~30-50%
 *    smaller payloads, faster parse for numeric-heavy data. No schema needed.
 *    Lowest integration cost -- swap JSON.stringify/parse for msgpack encode/decode.
 *
 * 2. FlatBuffers / Cap'n Proto: Zero-copy deserialization. The envelope header
 *    (formatVersion, ok, error) and CellValue discriminated unions map well to
 *    flat schemas. Requires .fbs schema file and codegen step.
 *
 * 3. Custom binary layout: Header (magic bytes, version u16, flags u16) + sections:
 *    - String table: length-prefixed UTF-8 strings referenced by varint index
 *    - Style table: packed style records (already deduplicated by StyleRegistry)
 *    - Cell data per sheet: column-oriented encoding with varint row indices and
 *      typed value encoding (1-byte tag + payload: float64 for numbers, varint
 *      string-table index for strings, 0-byte for empty, 1-byte for booleans,
 *      varint error code for errors, float64 serial for dates)
 *    - Formula strings: varint index into string table
 *    Most compact but highest maintenance cost.
 *
 * Recommended path: Start with MessagePack for the MachineEnvelope wrapper
 * (low effort, good gains for API responses). If workbook snapshots become a
 * bottleneck, add a dedicated binary cell-data format (approach 3) behind a
 * content-type negotiation flag (application/x-ascend-binary vs application/json).
 *
 * Key considerations:
 * - Binary envelope needs length-prefix or chunked framing since JSON's
 *   self-delimiting property is lost.
 * - Version field in the header allows format evolution; clients negotiate
 *   format via Accept header.
 * - CellValue union (empty|number|string|boolean|error|date|array|richText)
 *   maps to a 1-byte type tag + variable payload.
 * - Formula dedup via string table could save 20-40% on formula-heavy sheets.
 */
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
