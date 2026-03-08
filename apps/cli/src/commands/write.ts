import { readFile } from 'node:fs/promises'
import type { Operation } from '@ascend/schema'
import { AscendWorkbook } from '@ascend/sdk'
import { jsonOut } from '../output/json.ts'

export async function writeCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		console.error('Usage: ascend write <file> <range> <json-values> [--sheet <name>]')
		console.error('       ascend write <file> --ops <file.json>')
		return 1
	}

	const wb = await AscendWorkbook.open(file)

	const opsFile = flags.get('ops')
	if (opsFile) {
		const raw = await readFile(opsFile, 'utf-8')
		const ops: Operation[] = JSON.parse(raw)
		const result = wb.apply(ops)
		if (result.errors.length > 0) {
			for (const e of result.errors) console.error(e.message)
			return 1
		}
		await wb.save(file)
		if (flags.has('json')) {
			console.log(jsonOut(result))
		} else {
			console.log(`Applied ${ops.length} operation(s) to ${file}`)
		}
		return 0
	}

	const range = args[1]
	const valuesStr = args[2]
	if (!range || !valuesStr) {
		console.error('Usage: ascend write <file> <range> <json-values> [--sheet <name>]')
		return 1
	}

	const values: unknown = JSON.parse(valuesStr)
	const sheetName = flags.get('sheet') ?? wb.sheets[0]
	if (!sheetName) {
		console.error('No sheets in workbook')
		return 1
	}

	const updates = Array.isArray(values)
		? values.map((v: unknown, i: number) => ({
				ref: offsetRef(range, i),
				value: v as string | number | boolean | null,
			}))
		: [{ ref: range, value: values as string | number | boolean | null }]

	const ops: Operation[] = [{ op: 'setCells', sheet: sheetName, updates }]
	const result = wb.apply(ops)
	if (result.errors.length > 0) {
		for (const e of result.errors) console.error(e.message)
		return 1
	}

	await wb.save(file)
	if (flags.has('json')) {
		console.log(jsonOut(result))
	} else {
		console.log(`Wrote ${updates.length} cell(s) to ${file}`)
	}
	return 0
}

function offsetRef(baseRef: string, offset: number): string {
	const match = baseRef.match(/^([A-Za-z]+)(\d+)$/)
	if (!match) return baseRef
	const col = match[1]
	const row = Number.parseInt(match[2] ?? '1', 10) + offset
	return `${col}${row}`
}
