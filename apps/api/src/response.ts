import { type AscendError, ascendError, machineFailure, machineSuccess } from '@ascend/schema'

const JSON_HEADERS = { 'Content-Type': 'application/json' }

export function jsonSuccess(body: unknown, status = 200): Response {
	return new Response(JSON.stringify(machineSuccess(body)), {
		status,
		headers: JSON_HEADERS,
	})
}

export function jsonFailure(message: string, status: number): Response {
	return new Response(JSON.stringify(machineFailure(structuredStringFailure(message, status))), {
		status,
		headers: JSON_HEADERS,
	})
}

export function jsonFailureError(error: string | AscendError, status: number): Response {
	const structuredError = typeof error === 'string' ? structuredStringFailure(error, status) : error
	return new Response(JSON.stringify(machineFailure(structuredError)), {
		status,
		headers: JSON_HEADERS,
	})
}

export function binaryResponse(body: Uint8Array, contentType: string): Response {
	const bytes = new Uint8Array(body.byteLength)
	bytes.set(body)
	return new Response(bytes.buffer as ArrayBuffer, {
		headers: {
			'Content-Type': contentType,
		},
	})
}

function structuredStringFailure(message: string, status: number): AscendError {
	if (status >= 500) {
		return ascendError('INTERNAL_ERROR', message, {
			retryable: false,
			retryStrategy: 'none',
			suggestedFix: 'Inspect server logs, fix the server-side failure, and retry the request.',
		})
	}
	return ascendError('INVALID_ARGUMENT', message, {
		retryable: true,
		retryStrategy: 'modified',
		suggestedFix:
			status === 404
				? 'Use a supported Ascend API endpoint and retry the request.'
				: 'Adjust the request body or endpoint arguments and retry.',
	})
}
