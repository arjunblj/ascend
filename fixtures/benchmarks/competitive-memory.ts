/**
 * Rough RSS / heap comparison: Ascend vs SheetJS vs ExcelJS after loading a workbook.
 * Run: bun run fixtures/benchmarks/competitive-memory.ts
 */
import { writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createWorkbook, type StyleId } from '../../packages/core/src/index.ts'
import { writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { Ascend } from '../../packages/sdk/src/index.ts'

const ROWS = 5000
const COLS = 20
const SID = 0 as StyleId

function fmtMB(bytes: number): string {
	return `${(bytes / 1_048_576).toFixed(1)} MB`
}

function gc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		/* best effort */
	}
}

async function buildXlsx(): Promise<string> {
	const wb = createWorkbook()
	const sh = wb.addSheet('Data')
	for (let r = 0; r < ROWS; r++) {
		for (let c = 0; c < COLS; c++) {
			sh.cells.set(r, c, {
				value: { kind: 'number', value: r * COLS + c },
				formula: null,
				styleId: SID,
			})
		}
	}
	const res = writeXlsx(wb)
	if (!res.ok) throw new Error(res.error.message)
	const p = join(tmpdir(), `ascend-mem-${String(Date.now())}.xlsx`)
	await writeFile(p, res.value)
	return p
}

async function main(): Promise<void> {
	const path = await buildXlsx()
	const cells = ROWS * COLS
	console.log(`Workbook: ${ROWS}×${COLS} = ${String(cells)} cells\n`)

	gc()
	const baseRss = process.memoryUsage.rss()

	gc()
	await Ascend.open(path)
	gc()
	const ascendRss = process.memoryUsage.rss() - baseRss
	console.log(`Ascend      RSS delta: ${fmtMB(ascendRss)}`)

	let xlsxMod: typeof import('xlsx') | null = null
	let ExcelJS: typeof import('exceljs') | null = null
	try {
		xlsxMod = await import('xlsx')
	} catch {
		/* skip */
	}
	try {
		ExcelJS = await import('exceljs')
	} catch {
		/* skip */
	}

	if (xlsxMod) {
		gc()
		const rss0 = process.memoryUsage.rss()
		const fs = await import('node:fs')
		xlsxMod.read(fs.readFileSync(path), { type: 'buffer' })
		gc()
		console.log(`SheetJS     RSS delta: ${fmtMB(process.memoryUsage.rss() - rss0)}`)
	}

	if (ExcelJS) {
		gc()
		const rss0 = process.memoryUsage.rss()
		const wb = new ExcelJS.Workbook()
		await wb.xlsx.readFile(path)
		gc()
		console.log(`ExcelJS     RSS delta: ${fmtMB(process.memoryUsage.rss() - rss0)}`)
	}

	console.log('\nNote: RSS delta is approximate; GC timing may shift results.')
}

await main()
