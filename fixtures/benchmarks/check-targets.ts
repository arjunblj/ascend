import { readFileSync } from 'node:fs'
import type { BenchmarkSuiteResult } from './results.ts'
import { checkThroughputTargets, formatTargetResults } from './targets.ts'

function main(): void {
	const inputPath = process.argv[2]
	if (!inputPath) {
		console.error('Usage: bun fixtures/benchmarks/check-targets.ts <benchmark-json-path>')
		process.exit(2)
	}

	const raw = readFileSync(inputPath, 'utf8')
	const suite = JSON.parse(raw) as BenchmarkSuiteResult
	const results = checkThroughputTargets(suite)
	console.log(formatTargetResults(results))

	const failed = results.filter((entry) => !entry.passed)
	if (failed.length === 0) return

	console.error('\nThroughput target check failed.')
	for (const item of failed) {
		const actual =
			item.actualCellsPerSec === null ? 'no data' : `${item.actualCellsPerSec.toFixed(1)}/s`
		console.error(
			`- ${item.target.metric}: actual ${actual}, required ${item.target.minCellsPerSec.toFixed(1)}/s`,
		)
	}
	process.exit(1)
}

main()
