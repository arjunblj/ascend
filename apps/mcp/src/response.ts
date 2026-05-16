import { type AscendError, ascendError, machineFailure, machineSuccess } from '@ascend/schema'

export function okResponse<T>(data: T, summary: string) {
	return {
		content: [{ type: 'text' as const, text: summary }],
		structuredContent: machineSuccess(data) as unknown as Record<string, unknown>,
	}
}

export function errorResponse(error: string | AscendError) {
	const structuredError =
		typeof error === 'string'
			? ascendError('INVALID_ARGUMENT', error, {
					retryable: true,
					retryStrategy: 'modified',
					suggestedFix: 'Adjust the tool arguments and retry.',
				})
			: error
	return {
		content: [{ type: 'text' as const, text: structuredError.message }],
		structuredContent: machineFailure(structuredError) as unknown as Record<string, unknown>,
		isError: true,
	}
}
