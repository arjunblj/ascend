import type { CompactRangeInfo } from '@ascend/sdk'
import { indexToColumn } from '@ascend/sdk'
import { viewportRange } from '../model/viewport.ts'
import type { ViewportState } from '../runtime/types.ts'
import type { WorkbookSessionController } from './session-controller.ts'

export class HydrationService {
	private readonly rawTileCache = new Map<string, CompactRangeInfo | undefined>()
	private hits = 0
	private misses = 0

	constructor(
		private readonly session: WorkbookSessionController,
		private readonly maxEntries = 256,
	) {}

	readViewport(sheetName: string, viewport: ViewportState): CompactRangeInfo | undefined {
		const range = viewportRange(viewport, indexToColumn)
		const key = `${sheetName}:${range}`
		if (this.rawTileCache.has(key)) {
			const cached = this.rawTileCache.get(key)
			this.rawTileCache.delete(key)
			this.rawTileCache.set(key, cached)
			this.hits += 1
			return cached
		}
		this.misses += 1
		const data = this.session.requireWorkbook().readRangeCompact(sheetName, range, {
			includeRefs: true,
		})
		this.rawTileCache.set(key, data)
		this.evictOldest()
		return data
	}

	cacheHitRate(): number {
		const total = this.hits + this.misses
		return total === 0 ? 0 : this.hits / total
	}

	cacheSize(): number {
		return this.rawTileCache.size
	}

	invalidate(): void {
		this.rawTileCache.clear()
		this.hits = 0
		this.misses = 0
	}

	private evictOldest(): void {
		while (this.rawTileCache.size > this.maxEntries) {
			const oldest = this.rawTileCache.keys().next().value
			if (oldest === undefined) return
			this.rawTileCache.delete(oldest)
		}
	}
}
