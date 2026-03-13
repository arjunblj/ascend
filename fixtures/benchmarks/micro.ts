import {
	createWorkbook,
	SparseGrid,
	type StyleId,
	StyleRegistry,
} from '../../packages/core/src/index.ts'
import type { CellStyle } from '../../packages/core/src/style.ts'
import type { RangeDependency } from '../../packages/engine/src/dep-graph.ts'
import {
	cellKey,
	compileFormula,
	DependencyGraph,
	defaultCalcContext,
	type EvalContext,
	evaluateCompiled,
} from '../../packages/engine/src/index.ts'
import { clearGlobalParseCache, parseFormula } from '../../packages/formulas/src/index.ts'
import { numberValue, stringValue } from '../../packages/schema/src/index.ts'

const SID = 0 as StyleId

interface MicroBenchmark {
	readonly name: string
	readonly targetOpsPerSec: number
	run(): number
}

function createDeterministicRandom(seed: number): () => number {
	let state = seed >>> 0
	return () => {
		state = (Math.imul(state, 1664525) + 1013904223) >>> 0
		return state / 0x1_0000_0000
	}
}

const benchmarks: readonly MicroBenchmark[] = [
	{
		name: 'SparseGrid.readValue (100k reads)',
		targetOpsPerSec: 500_000,
		run() {
			const grid = new SparseGrid()
			const random = createDeterministicRandom(0xbe_ac01)
			for (let r = 0; r < 1000; r++) {
				for (let c = 0; c < 10; c++) {
					grid.setResolved(r, c, numberValue(r * 10 + c), null, SID)
				}
			}
			const count = 100_000
			let checksum = 0
			for (let i = 0; i < count; i++) {
				const r = (random() * 1000) | 0
				const c = (random() * 10) | 0
				const v = grid.readValue(r, c)
				if (v.kind === 'number') checksum += v.value
			}
			void checksum
			return count
		},
	},
	{
		name: 'SparseGrid.setResolved (100k writes)',
		targetOpsPerSec: 300_000,
		run() {
			const grid = new SparseGrid()
			const count = 100_000
			for (let i = 0; i < count; i++) {
				const r = (i / 10) | 0
				const c = i % 10
				grid.setResolved(r, c, numberValue(i), null, SID)
			}
			return count
		},
	},
	{
		name: 'Formula parse (1000 formulas)',
		targetOpsPerSec: 50_000,
		run() {
			const formulas: string[] = []
			for (let i = 0; i < 1000; i++) {
				switch (i % 5) {
					case 0:
						formulas.push(`SUM(A${i + 1}:A${i + 100})`)
						break
					case 1:
						formulas.push(`IF(B${i + 1}>0,C${i + 1}*2,D${i + 1}+1)`)
						break
					case 2:
						formulas.push(`VLOOKUP(E${i + 1},A$1:D$1000,3,FALSE)`)
						break
					case 3:
						formulas.push(`INDEX(B$1:B$5000,MATCH(F${i + 1},A$1:A$5000,0))`)
						break
					default:
						formulas.push(`SUMIFS(C$1:C$1000,A$1:A$1000,"cat${i % 5}")`)
						break
				}
			}
			clearGlobalParseCache()
			for (const f of formulas) {
				parseFormula(f)
			}
			return formulas.length
		},
	},
	{
		name: 'Compiled eval (10k formulas)',
		targetOpsPerSec: 100_000,
		run() {
			const workbook = createWorkbook()
			workbook.addSheet('Sheet1')
			const sheet = workbook.sheets[0]
			if (!sheet) throw new Error('no sheet')
			for (let r = 0; r < 100; r++) {
				sheet.cells.set(r, 0, { value: numberValue(r + 1), formula: null, styleId: SID })
			}
			const formula = 'A1+A2'
			const parsed = parseFormula(formula)
			if (!parsed.ok) throw new Error('Failed to parse formula for compiled eval benchmark')
			const compiled = compileFormula(parsed.value)
			if (!compiled) throw new Error('Failed to compile formula for benchmark')
			const ctx: EvalContext = {
				workbook,
				calcContext: defaultCalcContext(),
				sheetIndex: 0,
				row: 0,
				col: 1,
			}
			const count = 10_000
			for (let i = 0; i < count; i++) {
				evaluateCompiled(compiled, ctx)
			}
			return count
		},
	},
	{
		name: 'CellValue creation (100k numberValue + stringValue)',
		targetOpsPerSec: 2_000_000,
		run() {
			const count = 100_000
			let checksum = 0
			for (let i = 0; i < count; i++) {
				const n = numberValue(i)
				if (n.kind === 'number') checksum += n.value
				const s = stringValue(`val-${i}`)
				if (s.kind === 'string') checksum += s.value.length
			}
			void checksum
			return count * 2
		},
	},
	{
		name: 'DependencyGraph query (1000 range deps)',
		targetOpsPerSec: 50_000,
		run() {
			const graph = new DependencyGraph()
			const rangeDeps: RangeDependency[] = []
			for (let i = 0; i < 1000; i++) {
				const fk = cellKey(0, 5000 + i, 0)
				const startRow = (i * 7) % 4000
				const endRow = startRow + 100
				rangeDeps.length = 0
				rangeDeps.push({
					sheetIndex: 0,
					startRow,
					startCol: 0,
					endRow,
					endCol: 5,
				})
				graph.addFormula(fk, [], false, rangeDeps)
			}
			const random = createDeterministicRandom(0xde_0001)
			const count = 1000
			let totalDeps = 0
			for (let i = 0; i < count; i++) {
				const row = (random() * 5000) | 0
				const col = (random() * 6) | 0
				const deps = graph.getDependents(cellKey(0, row, col))
				totalDeps += deps.length
			}
			void totalDeps
			return count
		},
	},
	{
		name: 'StyleRegistry register (1000 styles)',
		targetOpsPerSec: 20_000,
		run() {
			const registry = new StyleRegistry()
			const count = 1000
			const fonts = ['Arial', 'Calibri', 'Helvetica', 'Times New Roman', 'Courier']
			const colors = ['FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF']
			for (let i = 0; i < count; i++) {
				const style: CellStyle = {
					font: {
						name: fonts[i % fonts.length],
						size: 8 + (i % 20),
						bold: i % 3 === 0,
						italic: i % 5 === 0,
						color: { kind: 'rgb', rgb: colors[i % colors.length] ?? '#000000' },
					},
					fill: {
						pattern: i % 2 === 0 ? 'solid' : 'none',
						fgColor: { kind: 'rgb', rgb: colors[(i + 1) % colors.length] ?? '#000000' },
					},
					numberFormat: i % 4 === 0 ? '#,##0.00' : undefined,
				}
				registry.register(style)
			}
			return count
		},
	},
]

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {}
}

function formatRate(rate: number): string {
	if (rate >= 1_000_000) return `${(rate / 1_000_000).toFixed(2)}M/s`
	if (rate >= 1_000) return `${(rate / 1_000).toFixed(1)}K/s`
	return `${rate.toFixed(1)}/s`
}

interface MicroResult {
	readonly name: string
	readonly ops: number
	readonly durationMs: number
	readonly opsPerSec: number
	readonly targetOpsPerSec: number
	readonly pass: boolean
}

function runMicro(bench: MicroBenchmark, warmupRuns: number, measureRuns: number): MicroResult {
	for (let i = 0; i < warmupRuns; i++) {
		bench.run()
	}
	runGc()
	const durations: number[] = []
	let totalOps = 0
	for (let i = 0; i < measureRuns; i++) {
		runGc()
		const start = performance.now()
		const ops = bench.run()
		const elapsed = performance.now() - start
		durations.push(elapsed)
		totalOps = ops
	}
	durations.sort((a, b) => a - b)
	const medianMs = durations[Math.floor(durations.length / 2)] ?? durations[0] ?? 0
	const opsPerSec = medianMs > 0 ? (totalOps / medianMs) * 1000 : Number.POSITIVE_INFINITY
	return {
		name: bench.name,
		ops: totalOps,
		durationMs: medianMs,
		opsPerSec,
		targetOpsPerSec: bench.targetOpsPerSec,
		pass: opsPerSec >= bench.targetOpsPerSec,
	}
}

async function main(): Promise<void> {
	const json = process.argv.includes('--json')
	const warmup = 2
	const measure = 5
	const results: MicroResult[] = []
	for (const bench of benchmarks) {
		if (!json) process.stdout.write(`  ${bench.name} ... `)
		const result = runMicro(bench, warmup, measure)
		results.push(result)
		if (!json) {
			const status = result.pass ? 'PASS' : 'FAIL'
			console.log(
				`${status}  ${formatRate(result.opsPerSec)} (target: ${formatRate(result.targetOpsPerSec)})  ${result.durationMs.toFixed(2)}ms`,
			)
		}
	}
	if (json) {
		console.log(JSON.stringify(results, null, 2))
		return
	}
	const passed = results.filter((r) => r.pass).length
	console.log(`\n${passed}/${results.length} microbenchmarks met targets`)
}

await main()
