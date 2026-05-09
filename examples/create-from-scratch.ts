#!/usr/bin/env bun
/**
 * Build a small workbook in memory and save as XLSX.
 * Usage: bun run examples/create-from-scratch.ts [out.xlsx]
 */
import { Ascend } from '@ascend/sdk'

const out = process.argv[2] ?? 'examples/out-create.xlsx'
const wb = Ascend.create()
wb.apply([
	{
		op: 'setCells',
		sheet: 'Sheet1',
		updates: [
			{ ref: 'A1', value: 'Hello' },
			{ ref: 'B1', value: 42 },
		],
	},
	{ op: 'setFormula', sheet: 'Sheet1', ref: 'C1', formula: '=B1*2' },
])
wb.recalc()
await wb.save(out)
console.log(`Wrote ${out}`)
