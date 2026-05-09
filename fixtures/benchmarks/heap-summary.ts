#!/usr/bin/env bun

interface HeapSnapshot {
	readonly version: number
	readonly type: string
	readonly nodes: readonly number[]
	readonly nodeClassNames: readonly string[]
}

interface HeapClassSummary {
	readonly className: string
	readonly count: number
	readonly selfSizeBytes: number
	readonly averageSelfSizeBytes: number
}

interface HeapSummary {
	readonly file: string
	readonly version: number
	readonly type: string
	readonly nodeCount: number
	readonly totalSelfSizeBytes: number
	readonly topClasses: readonly HeapClassSummary[]
}

const NODE_STRIDE = 4
const NODE_SELF_SIZE_OFFSET = 1
const NODE_CLASS_OFFSET = 2

function readOption(args: readonly string[], name: string): string | undefined {
	const index = args.indexOf(name)
	return index >= 0 ? args[index + 1] : undefined
}

function hasFlag(args: readonly string[], name: string): boolean {
	return args.includes(name)
}

function positiveInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function formatBytes(bytes: number): string {
	const units = ['B', 'KiB', 'MiB', 'GiB']
	let value = bytes
	let unit = 0
	while (value >= 1024 && unit < units.length - 1) {
		value /= 1024
		unit++
	}
	return `${value.toFixed(unit === 0 ? 0 : 2)} ${units[unit]}`
}

async function summarizeHeap(file: string, top: number): Promise<HeapSummary> {
	const snapshot = (await Bun.file(file).json()) as HeapSnapshot
	if (!Array.isArray(snapshot.nodes) || !Array.isArray(snapshot.nodeClassNames)) {
		throw new Error(`${file} is not a Bun Inspector heap snapshot`)
	}
	if (snapshot.nodes.length % NODE_STRIDE !== 0) {
		throw new Error(`${file} has an unsupported node stride`)
	}

	let totalSelfSizeBytes = 0
	const byClass = new Map<string, { count: number; selfSizeBytes: number }>()
	for (let offset = 0; offset < snapshot.nodes.length; offset += NODE_STRIDE) {
		const className =
			snapshot.nodeClassNames[snapshot.nodes[offset + NODE_CLASS_OFFSET] ?? -1] ?? '<unknown>'
		const selfSizeBytes = snapshot.nodes[offset + NODE_SELF_SIZE_OFFSET] ?? 0
		totalSelfSizeBytes += selfSizeBytes
		const current = byClass.get(className) ?? { count: 0, selfSizeBytes: 0 }
		current.count++
		current.selfSizeBytes += selfSizeBytes
		byClass.set(className, current)
	}

	const topClasses = [...byClass.entries()]
		.map(([className, value]) => ({
			className,
			count: value.count,
			selfSizeBytes: value.selfSizeBytes,
			averageSelfSizeBytes: value.selfSizeBytes / value.count,
		}))
		.sort((a, b) => b.selfSizeBytes - a.selfSizeBytes || b.count - a.count)
		.slice(0, top)

	return {
		file,
		version: snapshot.version,
		type: snapshot.type,
		nodeCount: snapshot.nodes.length / NODE_STRIDE,
		totalSelfSizeBytes,
		topClasses,
	}
}

function printSummary(summary: HeapSummary): void {
	console.log(`${summary.file}`)
	console.log(
		`  nodes=${summary.nodeCount.toLocaleString()} self=${formatBytes(summary.totalSelfSizeBytes)} type=${summary.type} v${summary.version}`,
	)
	console.log('  top classes by self size:')
	for (const entry of summary.topClasses) {
		console.log(
			`    ${entry.className.padEnd(32)} ${formatBytes(entry.selfSizeBytes).padStart(12)} count=${entry.count.toLocaleString()} avg=${formatBytes(entry.averageSelfSizeBytes)}`,
		)
	}
}

const args = process.argv.slice(2)
const top = positiveInt(readOption(args, '--top'), 20)
const json = hasFlag(args, '--json')
const files = args.filter((arg, index) => {
	if (arg === '--json') return false
	if (arg === '--top') return false
	if (args[index - 1] === '--top') return false
	return true
})
if (files.length === 0) {
	throw new Error(
		'Usage: bun run fixtures/benchmarks/heap-summary.ts <snapshot.json>... [--top N] [--json]',
	)
}

const summaries = await Promise.all(files.map((file) => summarizeHeap(file, top)))
if (json) {
	console.log(JSON.stringify({ tool: 'heap-summary', summaries }, null, 2))
} else {
	for (const summary of summaries) printSummary(summary)
}
