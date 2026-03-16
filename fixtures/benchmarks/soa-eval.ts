interface ResultRow {
	readonly label: string
	readonly medianMs: number
	readonly checksum: number
}

const SLOT_COUNT = 300_000
const MUTATION_COUNT = 500_000
const REPEAT = 5

function main(): void {
	const objectResult = benchmark('object-layout', runObjectLayout)
	const soaResult = benchmark('soa-layout', runSoaLayout)

	console.log('Formula Storage Layout A/B')
	console.log('='.repeat(72))
	for (const row of [objectResult, soaResult]) {
		console.log(
			`${row.label.padEnd(16)} median=${row.medianMs.toFixed(2)}ms checksum=${row.checksum}`,
		)
	}
	const delta = pctDelta(soaResult.medianMs, objectResult.medianMs)
	console.log('-'.repeat(72))
	console.log(`SoA vs object runtime delta: ${delta}`)
}

function benchmark(label: string, fn: () => number): ResultRow {
	const samples: number[] = []
	let checksum = 0
	for (let i = 0; i < REPEAT; i++) {
		const start = performance.now()
		checksum ^= fn()
		samples.push(performance.now() - start)
	}
	samples.sort((a, b) => a - b)
	return {
		label,
		medianMs: samples[Math.floor(samples.length / 2)] ?? 0,
		checksum,
	}
}

function runObjectLayout(): number {
	type Slot = {
		formula: string | null
		sharedIndex: number
		flags: number
	}
	const slots: Slot[] = Array.from({ length: SLOT_COUNT }, (_, i) => ({
		formula: i % 7 === 0 ? `A${i + 1}*2` : null,
		sharedIndex: i % 37,
		flags: i % 2,
	}))
	const rand = lcg(0x1a2b3c4d)
	for (let i = 0; i < MUTATION_COUNT; i++) {
		const idx = Math.floor(rand() * SLOT_COUNT)
		const slot = slots[idx]
		if (!slot) continue
		if ((i & 1) === 0) {
			slot.formula = null
			slot.flags = 0
		} else {
			slot.formula = `SUM(A${idx + 1}:B${idx + 1})`
			slot.flags = 1
			slot.sharedIndex = idx % 101
		}
	}
	let checksum = 0
	for (let i = 0; i < slots.length; i++) {
		const slot = slots[i]
		if (!slot) continue
		if (slot.flags === 1) checksum += slot.sharedIndex + (slot.formula?.length ?? 0)
	}
	return checksum
}

function runSoaLayout(): number {
	const formulas = new Array<string | null>(SLOT_COUNT)
	const sharedIndex = new Int32Array(SLOT_COUNT)
	const flags = new Uint8Array(SLOT_COUNT)
	for (let i = 0; i < SLOT_COUNT; i++) {
		formulas[i] = i % 7 === 0 ? `A${i + 1}*2` : null
		sharedIndex[i] = i % 37
		flags[i] = i % 2
	}
	const rand = lcg(0x1a2b3c4d)
	for (let i = 0; i < MUTATION_COUNT; i++) {
		const idx = Math.floor(rand() * SLOT_COUNT)
		if ((i & 1) === 0) {
			formulas[idx] = null
			flags[idx] = 0
		} else {
			formulas[idx] = `SUM(A${idx + 1}:B${idx + 1})`
			flags[idx] = 1
			sharedIndex[idx] = idx % 101
		}
	}
	let checksum = 0
	for (let i = 0; i < SLOT_COUNT; i++) {
		if (flags[i] === 1) checksum += sharedIndex[i] + (formulas[i]?.length ?? 0)
	}
	return checksum
}

function lcg(seed: number): () => number {
	let state = seed >>> 0
	return () => {
		state = (Math.imul(1664525, state) + 1013904223) >>> 0
		return state / 0x1_0000_0000
	}
}

function pctDelta(candidate: number, baseline: number): string {
	if (baseline === 0) return 'n/a'
	const pct = ((candidate - baseline) / baseline) * 100
	return `${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%`
}

main()
