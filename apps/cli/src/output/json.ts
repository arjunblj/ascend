export function jsonOut(data: unknown): string {
	return JSON.stringify(data, null, 2)
}
