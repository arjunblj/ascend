#!/usr/bin/env bun
/**
 * Open an existing XLSX, tweak cells, recalc, save.
 * Usage: bun run examples/read-modify-save.ts <input.xlsx> [output.xlsx]
 */
import { Ascend } from '@ascend/sdk'

const input = process.argv[2]
if (!input) {
	console.error('Usage: bun run examples/read-modify-save.ts <input.xlsx> [output.xlsx]')
	process.exit(1)
}
const output = process.argv[3] ?? input.replace(/\.xlsx$/i, '-out.xlsx')

const wb = await Ascend.open(input)
const first = wb.inspect().sheets[0]?.name ?? 'Sheet1'
wb.apply([{ op: 'setCells', sheet: first, updates: [{ ref: 'A1', value: 'Modified' }] }])
wb.recalc()
await wb.save(output)
console.log(`Read ${input}, wrote ${output}`)
