import { WorkbookTuiEngine } from './engine.ts'
import type { InputEvent, TerminalSize, TraceResult } from './types.ts'

export async function runHeadlessTrace(input: {
	readonly path?: string
	readonly trace: readonly InputEvent[]
	readonly size?: TerminalSize
}): Promise<TraceResult> {
	const size = input.size ?? { rows: 32, cols: 120 }
	const engine = await WorkbookTuiEngine.create(input.path ? { path: input.path, size } : { size })
	return engine.runHeadless(input.trace, { size, includeFrames: true })
}
