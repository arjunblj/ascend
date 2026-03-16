import { Workbook } from '@ascend/core'
import { defaultCalcContext, recalculate } from '@ascend/engine'
import { numberValue } from '@ascend/schema'

const FORMULA_COUNT = 5000
const REPEAT = 5

function main(): void {
	const samples: number[] = []
	let checksum = 0
	for (let i = 0; i < REPEAT; i++) {
		const wb = buildWorkbook()
		const start = performance.now()
		const result = recalculate(wb, defaultCalcContext())
		if (result.errors.length > 0)
			throw new Error(result.errors[0]?.error.message ?? 'recalc failed')
		samples.push(performance.now() - start)
		checksum ^= readChecksum(wb)
	}
	samples.sort((a, b) => a - b)
	const medianMs = samples[Math.floor(samples.length / 2)] ?? 0
	console.log('Indirect Literal Benchmark')
	console.log('='.repeat(64))
	console.log(`formulas=${FORMULA_COUNT} medianMs=${medianMs.toFixed(2)} checksum=${checksum}`)
	console.log(`throughputPerSec=${((FORMULA_COUNT / medianMs) * 1000).toFixed(2)}`)
}

function buildWorkbook(): Workbook {
	const wb = new Workbook()
	const input = wb.addSheet('Input')
	const calc = wb.addSheet('Calc')
	for (let row = 0; row < FORMULA_COUNT; row++) {
		input.cells.set(row, 0, { value: numberValue(row + 1), formula: null, styleId: 0 })
		calc.cells.set(row, 0, {
			value: numberValue(row + 1),
			formula: `INDIRECT("Input!A${row + 1}")`,
			styleId: 0,
		})
	}
	return wb
}

function readChecksum(wb: Workbook): number {
	const sheet = wb.sheets[1]
	if (!sheet) return 0
	let sum = 0
	for (let row = 0; row < FORMULA_COUNT; row++) {
		const value = sheet.cells.get(row, 0)?.value
		if (value?.kind === 'number') sum += value.value
	}
	return sum
}

main()
