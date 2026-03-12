import { readFileSync } from 'node:fs'
import { type BenchmarkSuiteResult, formatBytes, formatRate } from './results.ts'

function readSuite(path: string): BenchmarkSuiteResult {
	return JSON.parse(readFileSync(path, 'utf-8')) as BenchmarkSuiteResult
}

function deltaPct(baseline: number, current: number): string {
	if (baseline === 0) return 'n/a'
	const pct = ((current - baseline) / baseline) * 100
	const sign = pct > 0 ? '+' : ''
	return `${sign}${pct.toFixed(1)}%`
}

function pad(value: string, width: number): string {
	return value + ' '.repeat(Math.max(0, width - value.length))
}

function mdTable(headers: string[], rows: string[][]): string {
	const widths = headers.map((h, i) => Math.max(h.length, ...rows.map((r) => r[i]?.length ?? 0)))
	const headerLine = `| ${headers.map((h, i) => pad(h, widths[i] ?? 0)).join(' | ')} |`
	const sepLine = `| ${widths.map((w) => '-'.repeat(w)).join(' | ')} |`
	const bodyLines = rows.map(
		(row) => `| ${row.map((c, i) => pad(c, widths[i] ?? 0)).join(' | ')} |`,
	)
	return [headerLine, sepLine, ...bodyLines].join('\n')
}

async function main(): Promise<void> {
	const files = process.argv.slice(2)
	if (files.length === 0) {
		console.error(
			'Usage: bun run fixtures/benchmarks/dashboard.ts results1.json [results2.json ...]',
		)
		process.exit(1)
	}

	const suites = files.map((f) => readSuite(f))
	const baseline = suites[0]!

	const allNames = new Set<string>()
	for (const suite of suites) {
		for (const c of suite.cases) allNames.add(c.name)
	}
	const sortedNames = [...allNames].sort()

	const headers = ['Scenario', 'Median (ms)', 'Throughput']
	for (let i = 1; i < suites.length; i++) {
		const label = files.length === 2 ? '' : ` (${i + 1})`
		headers.push(`Δ% median${label}`, `Δ% throughput${label}`)
	}

	const rows: string[][] = []
	for (const name of sortedNames) {
		const baseCase = baseline.cases.find((c) => c.name === name)
		const row: string[] = [name]
		if (!baseCase) {
			row.push('—', '—')
			for (let i = 1; i < suites.length; i++) row.push('new', 'new')
		} else {
			row.push(baseCase.metrics.medianMs.toFixed(2))
			row.push(
				baseCase.metrics.throughputPerSec !== undefined
					? formatRate(baseCase.metrics.throughputPerSec)
					: '—',
			)
			for (let i = 1; i < suites.length; i++) {
				const caseI = suites[i]!.cases.find((c) => c.name === name)
				if (!caseI) {
					row.push('missing', 'missing')
				} else {
					row.push(deltaPct(baseCase.metrics.medianMs, caseI.metrics.medianMs))
					if (
						baseCase.metrics.throughputPerSec !== undefined &&
						caseI.metrics.throughputPerSec !== undefined
					) {
						row.push(deltaPct(baseCase.metrics.throughputPerSec, caseI.metrics.throughputPerSec))
					} else {
						row.push('—')
					}
				}
			}
		}
		rows.push(row)
	}

	const labels = files.map((f, i) => `${i + 1}. \`${f}\`${i === 0 ? ' (baseline)' : ''}`)
	console.log(`## Benchmark Comparison\n`)
	console.log(`${labels.join('\n')}\n`)

	if (baseline.git.sha) {
		const parts: string[] = []
		if (baseline.git.branch) parts.push(`branch: ${baseline.git.branch}`)
		parts.push(`sha: ${baseline.git.sha.slice(0, 8)}`)
		console.log(`> ${parts.join(', ')}\n`)
	}

	console.log(mdTable(headers, rows))

	if (suites.length >= 2) {
		const candidate = suites[1]!
		let regressions = 0
		let improvements = 0
		for (const name of sortedNames) {
			const b = baseline.cases.find((c) => c.name === name)
			const n = candidate.cases.find((c) => c.name === name)
			if (!b || !n) continue
			const pct =
				b.metrics.medianMs === 0
					? 0
					: (n.metrics.medianMs - b.metrics.medianMs) / b.metrics.medianMs
			if (pct > 0.1) regressions++
			else if (pct < -0.1) improvements++
		}
		console.log(`\n**Summary:** ${improvements} improved, ${regressions} regressed`)
	}
}

await main()
