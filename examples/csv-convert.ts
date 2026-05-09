#!/usr/bin/env bun
/**
 * CSV round-trip via the SDK.
 * Usage: bun run examples/csv-convert.ts [out.xlsx]
 */
import { Ascend } from '@ascend/sdk'

const csv = `name,score
alice,10
bob,20
`
const wb = Ascend.fromCsv(csv)
wb.recalc()
const out = process.argv[2] ?? 'examples/out-csv.xlsx'
await wb.save(out)
console.log(`Wrote ${out} from inline CSV`)
