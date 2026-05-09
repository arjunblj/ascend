import { describe, expect, test } from 'bun:test'
import { HydrationService } from './hydration-service.ts'
import type { WorkbookSessionController } from './session-controller.ts'

describe('HydrationService', () => {
	test('caches viewport reads with bounded LRU eviction', () => {
		const reads: string[] = []
		const service = new HydrationService(fakeSession(reads), 2)
		const viewport = {
			topRow: 0,
			leftCol: 0,
			visibleRows: 2,
			visibleCols: 2,
			columnWidths: [10, 10],
			overscanRows: 0,
			overscanCols: 0,
		}

		service.readViewport('Sheet1', viewport)
		service.readViewport('Sheet1', viewport)
		expect(reads).toEqual(['Sheet1:A1:B2'])
		expect(service.cacheHitRate()).toBe(0.5)

		service.readViewport('Sheet1', { ...viewport, topRow: 2 })
		service.readViewport('Sheet1', { ...viewport, topRow: 4 })
		expect(service.cacheSize()).toBe(2)

		service.readViewport('Sheet1', viewport)
		expect(reads).toContain('Sheet1:A1:B2')
		expect(reads.filter((read) => read === 'Sheet1:A1:B2')).toHaveLength(2)
	})
})

function fakeSession(reads: string[]): WorkbookSessionController {
	return {
		requireWorkbook() {
			return {
				readRangeCompact(sheet: string, range: string) {
					reads.push(`${sheet}:${range}`)
					return undefined
				},
			}
		},
	} as unknown as WorkbookSessionController
}
