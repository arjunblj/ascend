import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'

interface FileCoverage {
	file: string
	covered: number
	total: number
	uncoveredLines: number[]
}

const DEFAULT_TARGETS = [
	'packages/io-xlsx/src/reader/',
	'packages/io-xlsx/src/writer/',
	'packages/engine/src/evaluator.ts',
	'packages/engine/src/codegen.ts',
]

function main(): void {
	const lcovPathArg = process.argv[2] ?? '.coverage/lcov.info'
	const lcovPath = resolve(process.cwd(), lcovPathArg)
	if (!existsSync(lcovPath)) {
		console.error(`Coverage file not found: ${lcovPath}`)
		process.exit(2)
	}

	const targetArgs = process.argv.slice(3)
	const targets = targetArgs.length > 0 ? targetArgs : DEFAULT_TARGETS
	const allFiles = parseLcov(readFileSync(lcovPath, 'utf8'))
	const selected = allFiles.filter((entry) => targets.some((target) => entry.file.includes(target)))

	if (selected.length === 0) {
		console.log('No matching files in lcov report.')
		process.exit(1)
	}

	selected.sort((a, b) => coveragePct(a) - coveragePct(b))
	console.log('Coverage Audit')
	console.log('='.repeat(96))
	console.log(
		[
			'File'.padEnd(58),
			'Coverage'.padStart(10),
			'Covered'.padStart(10),
			'Total'.padStart(8),
			'Uncovered'.padStart(10),
		].join(' '),
	)
	console.log('-'.repeat(96))
	for (const entry of selected) {
		const pct = `${coveragePct(entry).toFixed(1)}%`
		console.log(
			[
				truncate(entry.file, 58).padEnd(58),
				pct.padStart(10),
				String(entry.covered).padStart(10),
				String(entry.total).padStart(8),
				String(entry.uncoveredLines.length).padStart(10),
			].join(' '),
		)
	}
	console.log('-'.repeat(96))

	const lowest = selected[0]
	if (lowest) {
		const sample = lowest.uncoveredLines.slice(0, 30).join(', ')
		console.log(`Lowest coverage file: ${lowest.file}`)
		console.log(`Sample uncovered lines: ${sample || 'none'}`)
	}
}

function parseLcov(content: string): FileCoverage[] {
	const out: FileCoverage[] = []
	let currentFile: string | null = null
	let lines = new Map<number, number>()

	const flush = () => {
		if (!currentFile) return
		const total = lines.size
		let covered = 0
		const uncoveredLines: number[] = []
		for (const [line, hits] of lines) {
			if (hits > 0) covered++
			else uncoveredLines.push(line)
		}
		uncoveredLines.sort((a, b) => a - b)
		out.push({ file: normalizePath(currentFile), covered, total, uncoveredLines })
	}

	for (const raw of content.split('\n')) {
		const line = raw.trim()
		if (line.startsWith('SF:')) {
			flush()
			currentFile = line.slice(3)
			lines = new Map()
			continue
		}
		if (line.startsWith('DA:')) {
			const payload = line.slice(3)
			const [lineNoRaw, hitsRaw] = payload.split(',')
			const lineNo = Number.parseInt(lineNoRaw ?? '', 10)
			const hits = Number.parseInt(hitsRaw ?? '', 10)
			if (!Number.isNaN(lineNo) && !Number.isNaN(hits)) lines.set(lineNo, hits)
			continue
		}
		if (line === 'end_of_record') {
			flush()
			currentFile = null
			lines = new Map()
		}
	}
	flush()
	return out
}

function coveragePct(entry: FileCoverage): number {
	if (entry.total === 0) return 100
	return (entry.covered / entry.total) * 100
}

function normalizePath(path: string): string {
	const cwd = process.cwd().replaceAll('\\', '/')
	return path.replaceAll('\\', '/').replace(`${cwd}/`, '')
}

function truncate(value: string, max: number): string {
	if (value.length <= max) return value
	return `...${value.slice(value.length - max + 3)}`
}

main()
