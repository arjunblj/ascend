import { type AscendError, ascendError, machineFailure, machineSuccess } from '@ascend/schema'

export function jsonOut(data: unknown): string {
	return JSON.stringify(machineSuccess(data), null, 2)
}

export function jsonErr(error: string | AscendError): string {
	return JSON.stringify(machineFailure(error), null, 2)
}

export function cliError(error: string | AscendError, flags: Map<string, string>): void {
	const structuredError =
		typeof error === 'string'
			? ascendError('INVALID_ARGUMENT', error, {
					retryable: true,
					retryStrategy: 'modified',
					suggestedFix: 'Adjust the command arguments or flags and retry.',
				})
			: error
	if (flags.has('json')) {
		console.log(jsonErr(structuredError))
	} else {
		console.error(structuredError.message)
	}
}
