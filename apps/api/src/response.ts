import type { AscendError } from '@ascend/schema'
import { machineFailure, machineSuccess } from '@ascend/schema'

const JSON_HEADERS = { 'Content-Type': 'application/json' }

export function jsonSuccess(body: unknown, status = 200): Response {
	return new Response(JSON.stringify(machineSuccess(body)), {
		status,
		headers: JSON_HEADERS,
	})
}

export function jsonFailure(message: string, status: number): Response {
	return new Response(JSON.stringify(machineFailure(message)), {
		status,
		headers: JSON_HEADERS,
	})
}

export function jsonFailureError(error: string | AscendError, status: number): Response {
	return new Response(JSON.stringify(machineFailure(error)), {
		status,
		headers: JSON_HEADERS,
	})
}

export function binaryResponse(body: Uint8Array, contentType: string): Response {
	return new Response(body, {
		headers: {
			'Content-Type': contentType,
		},
	})
}
