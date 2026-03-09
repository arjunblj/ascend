import { machineSuccess } from '@ascend/schema'

export function jsonOut(data: unknown): string {
	return JSON.stringify(machineSuccess(data), null, 2)
}
