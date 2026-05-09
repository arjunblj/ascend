import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { applyOperation } from '../../../engine/src/operations.ts'
import { readXlsx } from '../reader/index.ts'
import { writeXlsx } from './index.ts'

const openpyxlAvailable = Bun.spawnSync(['python3', '-c', 'import openpyxl']).exitCode === 0
const ossTest = openpyxlAvailable ? test : test.skip

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('chart OSS compatibility', () => {
	ossTest('edits OpenPyXL chart series sources and remains readable by OpenPyXL', () => {
		const dir = mkdtempSync(join(tmpdir(), 'ascend-chart-oss-'))
		try {
			const input = join(dir, 'openpyxl-chart.xlsx')
			const output = join(dir, 'ascend-chart-edited.xlsx')
			runPython(CREATE_OPENPYXL_CHART, input)

			const read = readXlsx(readFileSync(input))
			expectOk(read)
			const chart = read.value.workbook.chartParts[0]
			expect(chart).toMatchObject({
				partPath: 'xl/charts/chart1.xml',
				chartType: 'barChart',
			})
			if (!chart) throw new Error('Expected OpenPyXL chart to be parsed')
			expect(chart?.series[0]?.valueRef).toBe("'Data'!$B$2:$B$5")

			const applied = applyOperation(read.value.workbook, {
				op: 'setChartSeriesSource',
				partPath: chart.partPath,
				seriesIndex: 0,
				valueRef: 'Data!$C$2:$C$5',
			})
			expectOk(applied)
			const written = writeXlsx(read.value.workbook, read.value.capsules)
			expectOk(written)
			writeFileSync(output, written.value)

			const external = runPython(ASSERT_OPENPYXL_CHART_SOURCE, output)
			expect(external.stdout).toContain('Data!$C$2:$C$5')
		} finally {
			rmSync(dir, { recursive: true, force: true })
		}
	})
})

function runPython(script: string, path: string): { stdout: string } {
	const result = Bun.spawnSync(['python3', '-c', script, path])
	if (result.exitCode !== 0) {
		throw new Error(new TextDecoder().decode(result.stderr))
	}
	return { stdout: new TextDecoder().decode(result.stdout) }
}

const CREATE_OPENPYXL_CHART = `
import sys
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

path = sys.argv[1]
wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.append(["Month", "Revenue", "Margin"])
for row in [["Jan", 10, 3], ["Feb", 20, 7], ["Mar", 30, 9], ["Apr", 40, 11]]:
    ws.append(row)
chart = BarChart()
chart.title = "Revenue"
chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=5))
ws.add_chart(chart, "E2")
wb.save(path)
`

const ASSERT_OPENPYXL_CHART_SOURCE = `
import sys
from openpyxl import load_workbook

path = sys.argv[1]
wb = load_workbook(path)
ws = wb["Data"]
assert len(ws._charts) == 1
source = ws._charts[0].series[0].val.numRef.f
assert source == "Data!$C$2:$C$5", source
print(source)
`
