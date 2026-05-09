#!/usr/bin/env bun
/**
 * Evaluate a formula string against an empty workbook (utility demo).
 * Usage: bun run examples/formula-eval.ts '=1+2'
 */
import { Ascend } from '@ascend/sdk'

const arg = process.argv[2] ?? '=SUM(1,2,3)'
const wb = Ascend.create()
const v = wb.eval(arg)
console.log(JSON.stringify(v))
