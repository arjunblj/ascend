#!/usr/bin/env bun
/**
 * BatchBuilder + apply pattern.
 * Usage: bun run examples/batch-ops.ts [out.xlsx]
 */
import { Ascend, BatchBuilder } from '@ascend/sdk'

const out = process.argv[2] ?? 'examples/out-batch.xlsx'
const wb = Ascend.create()
new BatchBuilder(wb)
	.set('Sheet1!A1', 'Qty')
	.set('Sheet1!B1', 'Price')
	.set('Sheet1!A2', 3)
	.set('Sheet1!B2', 12.5)
	.formula('Sheet1!C2', '=A2*B2')
	.commitAndRecalc()
await wb.save(out)
console.log(`Wrote ${out}`)
