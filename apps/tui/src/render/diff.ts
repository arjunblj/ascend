import type { RenderFrame, RenderPatch } from '../runtime/types.ts'

export function diffFrames(previous: RenderFrame | null, next: RenderFrame): RenderPatch {
	if (!previous || previous.size.rows !== next.size.rows || previous.size.cols !== next.size.cols) {
		return {
			fullRedraw: true,
			lines: next.lines.map((text, index) => ({ row: index + 1, text })),
			bytes: next.lines.join('\n').length,
		}
	}
	const lines: Array<{ row: number; text: string }> = []
	let bytes = 0
	for (let i = 0; i < next.lines.length; i++) {
		const text = next.lines[i] ?? ''
		if (previous.lines[i] !== text) {
			lines.push({ row: i + 1, text })
			bytes += text.length
		}
	}
	return {
		fullRedraw: false,
		lines,
		bytes,
	}
}
