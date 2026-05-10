import { readFileSync } from 'node:fs'
import type { BenchmarkSuiteResult } from './results.ts'
import { checkThroughputTargets, formatTargetResults } from './targets.ts'

function main(): void {
	const inputPath = process.argv.find(
		(arg) => !arg.startsWith('--') && arg !== process.argv[0] && arg !== process.argv[1],
	)
	if (!inputPath) {
		console.error(
			'Usage: bun fixtures/benchmarks/check-targets.ts <benchmark-json-path> [--min-ratio <0..1>]',
		)
		process.exit(2)
	}
	const minRatio = parseMinRatio(process.argv)

	const raw = readFileSync(inputPath, 'utf8')
	const suite = JSON.parse(raw) as BenchmarkSuiteResult
	const results = checkThroughputTargets(suite, { minRatio })
	if (minRatio !== 1) {
		console.log(`Applying throughput target ratio: ${(minRatio * 100).toFixed(0)}%`)
	}
	console.log(formatTargetResults(results))

	const failed = results.filter((entry) => !entry.passed)
	if (failed.length === 0) return

	console.error('\nThroughput target check failed.')
	for (const item of failed) {
		const actual =
			item.actualCellsPerSec === null ? 'no data' : `${item.actualCellsPerSec.toFixed(1)}/s`
		console.error(
			`- ${item.target.metric}: actual ${actual}, required ${item.requiredCellsPerSec.toFixed(1)}/s`,
		)
	}
	process.exit(1)
}

function parseMinRatio(args: readonly string[]): number {
	const index = args.indexOf('--min-ratio')
	if (index === -1) return 1
	const raw = args[index + 1]
	const ratio = raw === undefined ? Number.NaN : Number(raw)
	if (!Number.isFinite(ratio) || ratio <= 0 || ratio > 1) {
		console.error('--min-ratio must be a number greater than 0 and no more than 1')
		process.exit(2)
	}
	return ratio
}

main()
