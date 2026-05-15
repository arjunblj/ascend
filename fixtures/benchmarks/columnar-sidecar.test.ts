import { describe, expect, test } from 'bun:test'
import { createWorkbook } from '../../packages/core/src/index.ts'
import {
	buildNumericColumnSidecar,
	columnarSidecarClaimReport,
	columnarSidecarClaimReportMarkdown,
	runColumnarSidecarBenchmark,
	sumSidecarColumn,
} from './columnar-sidecar.ts'

describe('columnar sidecar benchmark harness', () => {
	test('checks sidecar and grid checksums without relying on timing thresholds', () => {
		const result = runColumnarSidecarBenchmark({ rows: 128, cols: 4, repeats: 3 })

		expect(result.range).toBe('A1:D128')
		expect(result.cells).toBe(512)
		expect(result.populatedCount).toBe(512)
		expect(result.numericCount).toBe(512)
		expect(result.estimatedSidecarPayloadBytes).toBe(512 * 9)
		expect(result.checksum).toBeGreaterThan(0)
		expect(result.gridRepeatedScanMs).toBeGreaterThanOrEqual(0)
		expect(result.sidecarRepeatedScanMs).toBeGreaterThanOrEqual(0)
	})

	test('records generation and validity metadata for mixed numeric ranges', () => {
		const workbook = createWorkbook()
		const sheet = workbook.addSheet('Data')
		sheet.cells.setPlainNumberSpan(0, 0, [1, 2])
		sheet.cells.setPlainNumberSpan(1, 0, [3, 4])
		sheet.cells.setPlainString(1, 1, 'skip')

		const sidecar = buildNumericColumnSidecar(sheet, 'A1:B2', 7)

		expect(sidecar.generation).toBe(7)
		expect(sidecar.populatedCount).toBe(4)
		expect(sidecar.numericCount).toBe(3)
		expect(sidecar.checksum).toBe(6)
		expect(sumSidecarColumn(sidecar, 0)).toBe(4)
		expect(sumSidecarColumn(sidecar, 1)).toBe(2)
	})

	test('renders claim-safe sidecar proof boundaries', () => {
		const result = runColumnarSidecarBenchmark({ rows: 512, cols: 4, repeats: 8 })
		const report = columnarSidecarClaimReport(result)
		const markdown = columnarSidecarClaimReportMarkdown(report)

		expect(report.allowedClaim).toContain('disposable numeric columnar sidecar')
		expect(report.boundary).toContain('not a production cache')
		expect(report.killCriterion).toContain('real workbook tables')
		expect(markdown).toContain('Do not promote yet')
		expect(markdown).toContain('Estimated sidecar payload bytes')
	})
})
