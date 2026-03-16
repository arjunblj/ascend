import wabtInit from 'wabt'

interface BenchRow {
	readonly label: string
	readonly medianMs: number
	readonly checksum: bigint
}

const LENGTH = 1_000_000
const REPEAT = 7

async function main(): Promise<void> {
	const values = buildValues(LENGTH)
	const jsRow = bench('js-sum', () => jsSum(values))
	const wasm = await createWasmSum()
	const wasmRow = bench('wasm-sum', () => wasm(values))

	console.log('WASM Aggregate A/B')
	console.log('='.repeat(72))
	console.log(
		`${jsRow.label.padEnd(12)} median=${jsRow.medianMs.toFixed(2)}ms checksum=${jsRow.checksum}`,
	)
	console.log(
		`${wasmRow.label.padEnd(12)} median=${wasmRow.medianMs.toFixed(2)}ms checksum=${wasmRow.checksum}`,
	)
	console.log('-'.repeat(72))
	console.log(`WASM vs JS runtime delta: ${pctDelta(wasmRow.medianMs, jsRow.medianMs)}`)
}

function bench(label: string, fn: () => bigint): BenchRow {
	const samples: number[] = []
	let checksum = 0n
	for (let i = 0; i < REPEAT; i++) {
		const start = performance.now()
		checksum ^= fn()
		samples.push(performance.now() - start)
	}
	samples.sort((a, b) => a - b)
	return { label, medianMs: samples[Math.floor(samples.length / 2)] ?? 0, checksum }
}

function buildValues(length: number): Int32Array {
	const values = new Int32Array(length)
	for (let i = 0; i < values.length; i++) {
		values[i] = ((i * 17) % 2000) - 1000
	}
	return values
}

function jsSum(values: Int32Array): bigint {
	let total = 0n
	for (let i = 0; i < values.length; i++) total += BigInt(values[i] as number)
	return total
}

async function createWasmSum(): Promise<(values: Int32Array) => bigint> {
	const wabt = await wabtInit()
	const module = wabt.parseWat(
		'sum_i32.wat',
		`(module
  (memory (export "memory") 64)
  (func (export "sum_i32") (param $ptr i32) (param $len i32) (result i64)
    (local $i i32)
    (local $sum i64)
    (block $exit
      (loop $loop
        (br_if $exit (i32.ge_u (local.get $i) (local.get $len)))
        (local.set $sum
          (i64.add
            (local.get $sum)
            (i64.extend_i32_s
              (i32.load
                (i32.add
                  (local.get $ptr)
                  (i32.shl (local.get $i) (i32.const 2))
                )
              )
            )
          )
        )
        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $loop)
      )
    )
    (local.get $sum)
  )
)`,
	)
	const { buffer } = module.toBinary({})
	const instance = await WebAssembly.instantiate(buffer)
	const exports = instance.instance.exports as {
		readonly memory: WebAssembly.Memory
		readonly sum_i32: (ptr: number, len: number) => bigint
	}
	return (values: Int32Array): bigint => {
		const mem = new Int32Array(exports.memory.buffer, 0, values.length)
		mem.set(values)
		return exports.sum_i32(0, values.length)
	}
}

function pctDelta(candidate: number, baseline: number): string {
	if (baseline === 0) return 'n/a'
	const pct = ((candidate - baseline) / baseline) * 100
	return `${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%`
}

await main()
