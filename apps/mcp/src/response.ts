import { machineFailure, machineSuccess } from '@ascend/schema'

export function okResponse<T>(data: T, summary: string) {
	return {
		content: [{ type: 'text' as const, text: summary }],
		structuredContent: machineSuccess(data) as unknown as Record<string, unknown>,
	}
}

export function errorResponse(message: string) {
	return {
		content: [{ type: 'text' as const, text: message }],
		structuredContent: machineFailure(message) as unknown as Record<string, unknown>,
		isError: true,
	}
}
