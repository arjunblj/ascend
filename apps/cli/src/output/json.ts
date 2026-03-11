import type { AscendError } from '@ascend/schema'
import { machineFailure, machineSuccess } from '@ascend/schema'

export function jsonOut(data: unknown): string {
	return JSON.stringify(machineSuccess(data), null, 2)
}

export function jsonErr(error: string | AscendError): string {
	return JSON.stringify(machineFailure(error), null, 2)
}

export function cliError(error: string | AscendError, flags: Map<string, string>): void {
	if (flags.has('json')) {
		console.log(jsonErr(error))
	} else {
		console.error(typeof error === 'string' ? error : error.message)
	}
}
