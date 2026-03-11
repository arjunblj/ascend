import type { AscendError } from '@ascend/schema'
import { machineFailure, machineSuccess } from '@ascend/schema'

export function okResponse<T>(data: T, summary: string) {
	return {
		content: [{ type: 'text' as const, text: summary }],
		structuredContent: machineSuccess(data) as unknown as Record<string, unknown>,
	}
}

export function errorResponse(error: string | AscendError) {
	const message = typeof error === 'string' ? error : error.message
	return {
		content: [{ type: 'text' as const, text: message }],
		structuredContent: machineFailure(error) as unknown as Record<string, unknown>,
		isError: true,
	}
}
