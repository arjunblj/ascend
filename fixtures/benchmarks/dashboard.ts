import { readFileSync } from 'node:fs'
import { type BenchmarkSuiteResult, formatRate } from './results.ts'

function loadSuite(path: string): BenchmarkSuiteResult {
	const raw = readFileSync(path, 'utf-8')
	return JSON.parse(raw) as BenchmarkSuiteResult
}

function pctStr(baseline: number, candidate: number): string {
	if (baseline === 0) return 'n/a'
	const pct = ((candidate - baseline) / baseline) * 100
	const sign = pct > 0 ? '+' : ''
	return `${sign}${pct.toFixed(1)}%`
}

function padRight(s: string, w: number): string {
	return s + ' '.repeat(Math.max(0, w - s.length))
}

function padLeft(s: string, w: number): string {
	return ' '.repeat(Math.max(0, w - s.length)) + s
}

function main(): void {
	const args = process.argv.slice(2).filter((a) => !a.startsWith('-'))
	if (args.length === 0) {
		console.error(
			'Usage: bun run fixtures/benchmarks/dashboard.ts results1.json [results2.json ...]',
		)
		process.exit(1)
	}

	const suites = args.map((path) => ({ path, suite: loadSuite(path) }))
	const baseline = suites[0]
	if (!baseline) {
		console.error('No results files provided')
		process.exit(1)
	}

	const allScenarios = new Set<string>()
	for (const { suite } of suites) {
		for (const c of suite.cases) allScenarios.add(c.name)
	}
	const scenarioNames = [...allScenarios].sort()

	const headers = ['scenario', 'median (ms)', 'throughput']
	if (suites.length > 1) {
		for (let i = 1; i < suites.length; i++) {
			const label = suites[i]?.path.replace(/.*\//, '').replace(/\.json$/, '') ?? `run${i + 1}`
			headers.push(`median-${label}`, `delta%`)
		}
	}

	const rows: string[][] = []
	for (const name of scenarioNames) {
		const baseCase = baseline.suite.cases.find((c) => c.name === name)
		const baseMedian = baseCase?.metrics.medianMs ?? 0
		const baseThroughput = baseCase?.metrics.throughputPerSec

		const row: string[] = [
			name,
			baseMedian.toFixed(2),
			baseThroughput !== undefined ? formatRate(baseThroughput) : 'n/a',
		]

		for (let i = 1; i < suites.length; i++) {
			const s = suites[i]
			if (!s) {
				row.push('n/a', 'n/a')
				continue
			}
			const candCase = s.suite.cases.find((c) => c.name === name)
			if (!candCase) {
				row.push('n/a', 'n/a')
				continue
			}
			row.push(candCase.metrics.medianMs.toFixed(2), pctStr(baseMedian, candCase.metrics.medianMs))
		}
		rows.push(row)
	}

	const colWidths = headers.map((h, i) => Math.max(h.length, ...rows.map((r) => r[i]?.length ?? 0)))
	const rightAlignCols = new Set(headers.map((_, i) => i).filter((i) => i > 0))

	const pad = (s: string, i: number) =>
		rightAlignCols.has(i) ? padLeft(s, colWidths[i] ?? 0) : padRight(s, colWidths[i] ?? 0)

	const lines: string[] = []
	lines.push('')
	lines.push(`# Benchmark Dashboard`)
	lines.push('')
	if (baseline.suite.git.sha) {
		lines.push(
			`Baseline: \`${baseline.suite.git.sha.slice(0, 8)}\` (${baseline.suite.git.branch ?? 'unknown'})`,
		)
	}
	lines.push(`Generated: ${new Date().toISOString()}`)
	lines.push('')
	lines.push(`| ${headers.map((h, i) => pad(h, i)).join(' | ')} |`)
	lines.push(`| ${colWidths.map((w) => '-'.repeat(w)).join(' | ')} |`)
	for (const row of rows) {
		lines.push(`| ${row.map((cell, i) => pad(cell, i)).join(' | ')} |`)
	}
	lines.push('')

	if (suites.length > 1) {
		const lastSuite = suites[suites.length - 1]
		if (lastSuite) {
			let improved = 0
			let regressed = 0
			let unchanged = 0
			for (const name of scenarioNames) {
				const baseCase = baseline.suite.cases.find((c) => c.name === name)
				const candCase = lastSuite.suite.cases.find((c) => c.name === name)
				if (!baseCase || !candCase) continue
				const baseMs = baseCase.metrics.medianMs
				const candMs = candCase.metrics.medianMs
				if (baseMs === 0) {
					unchanged++
					continue
				}
				const delta = (candMs - baseMs) / baseMs
				if (delta < -0.1) improved++
				else if (delta > 0.1) regressed++
				else unchanged++
			}
			lines.push(
				`**Summary vs baseline:** ${improved} improved, ${regressed} regressed, ${unchanged} unchanged`,
			)
			lines.push('')
		}
	}

	console.log(lines.join('\n'))
}

main()
