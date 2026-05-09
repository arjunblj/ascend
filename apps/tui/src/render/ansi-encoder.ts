import type { RenderFrame, RenderPatch } from '../runtime/types.ts'
import { diffFrames } from './diff.ts'
import { ANSI, moveTo } from './styles.ts'

export interface EncodedAnsiFrame {
	readonly patch: RenderPatch
	readonly encoded: string
	readonly diffMs: number
	readonly encodeMs: number
}

export function encodeAnsiFrame(previous: RenderFrame | null, next: RenderFrame): EncodedAnsiFrame {
	const diffStart = performance.now()
	const patch = diffFrames(previous, next)
	const diffMs = performance.now() - diffStart
	const encodeStart = performance.now()
	const chunks: string[] = []
	if (patch.fullRedraw) chunks.push(ANSI.home)
	for (const line of patch.lines) {
		chunks.push(moveTo(line.row, 1), line.text)
	}
	if (next.cursor?.visible) {
		chunks.push(
			moveTo(clamp(next.cursor.row, 1, next.size.rows), clamp(next.cursor.col, 1, next.size.cols)),
			ANSI.showCursor,
		)
	} else {
		chunks.push(ANSI.hideCursor)
	}
	return {
		patch,
		encoded: chunks.join(''),
		diffMs,
		encodeMs: performance.now() - encodeStart,
	}
}

function clamp(value: number, min: number, max: number): number {
	return Math.min(max, Math.max(min, value))
}
