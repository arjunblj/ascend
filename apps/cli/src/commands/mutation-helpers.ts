import type { Operation } from '@ascend/schema'
import type { AscendWorkbook } from '@ascend/sdk'

export interface WriteSelector {
	readonly ref: string
	readonly sheet?: string
}

export function parseWriteSelector(
	selector: string,
	requestedSheet: string | undefined,
): WriteSelector {
	const bang = selector.lastIndexOf('!')
	if (bang === -1) {
		return requestedSheet ? { sheet: requestedSheet, ref: selector } : { ref: selector }
	}
	const sheet = selector.slice(0, bang).replace(/^'|'$/g, '')
	const ref = selector.slice(bang + 1)
	return { sheet, ref }
}

export function resolveSheetName(
	wb: Pick<AscendWorkbook, 'sheets'>,
	explicitSheet: string | undefined,
): string | undefined {
	if (explicitSheet) return explicitSheet
	return wb.sheets.length === 1 ? wb.sheets[0] : undefined
}

export function buildSetCellOps(
	sheetName: string,
	ref: string,
	values: unknown,
): readonly Operation[] {
	const updates = Array.isArray(values)
		? values.map((value: unknown, index: number) => ({
				ref: offsetRef(ref, index),
				value: value as string | number | boolean | null,
			}))
		: [{ ref, value: values as string | number | boolean | null }]
	return [{ op: 'setCells', sheet: sheetName, updates }]
}

function offsetRef(baseRef: string, offset: number): string {
	const match = baseRef.match(/^([A-Za-z]+)(\d+)$/)
	if (!match) return baseRef
	const col = match[1]
	const row = Number.parseInt(match[2] ?? '1', 10) + offset
	return `${col}${row}`
}
