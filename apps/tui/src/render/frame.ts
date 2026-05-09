import type { FrameStats, RenderFrame, TerminalSize } from '../runtime/types.ts'
import { fitAnsi } from './ansi-text.ts'

export function createFrame(
	size: TerminalSize,
	lines: readonly string[],
	cursor?: RenderFrame['cursor'],
): RenderFrame {
	const normalized = normalizeLines(lines, size)
	const stats: FrameStats = {
		fullFrameCells: size.rows * size.cols,
		dirtyCells: size.rows * size.cols,
		dirtyRows: normalized.length,
		bytes: normalized.join('\n').length,
	}
	return {
		size,
		lines: normalized,
		...(cursor ? { cursor } : {}),
		stats,
	}
}

function normalizeLines(lines: readonly string[], size: TerminalSize): readonly string[] {
	const out: string[] = []
	for (let row = 0; row < size.rows; row++) {
		const line = lines[row] ?? ''
		out.push(fitAnsi(line, size.cols))
	}
	return out
}
