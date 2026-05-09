import type {
	RenderFrame,
	RenderStats,
	TerminalCapabilities,
	TerminalRenderer,
} from '../runtime/types.ts'
import { encodeAnsiFrame } from './ansi-encoder.ts'
import { ANSI } from './styles.ts'

export class AnsiRenderer implements TerminalRenderer {
	private previous: RenderFrame | null = null
	private initialized = false
	private lastDraw = 0

	async init(capabilities: TerminalCapabilities): Promise<void> {
		if (this.initialized || !capabilities.isTty) return
		process.stdout.write(
			`${ANSI.altScreen}${ANSI.hideCursor}${ANSI.clear}${ANSI.home}${ANSI.enableMouse}${ANSI.enableBracketedPaste}`,
		)
		this.initialized = true
	}

	async draw(frame: RenderFrame): Promise<RenderStats> {
		const encoded = encodeAnsiFrame(this.previous, frame)
		const output = encoded.encoded
		const writeStart = performance.now()
		if (output.length > 0) process.stdout.write(output)
		const writeMs = performance.now() - writeStart
		this.previous = frame
		const now = performance.now()
		const fps = this.lastDraw > 0 ? 1000 / Math.max(1, now - this.lastDraw) : 0
		this.lastDraw = now
		return {
			frameBuildMs: 0,
			frameDiffMs: encoded.diffMs,
			encodeMs: encoded.encodeMs,
			writeMs,
			changedCells: encoded.patch.lines.length * frame.size.cols,
			bytesOut: output.length,
			droppedFrames: 0,
			fps,
		}
	}

	async shutdown(): Promise<void> {
		if (!this.initialized) return
		process.stdout.write(
			`${ANSI.disableBracketedPaste}${ANSI.disableMouse}${ANSI.showCursor}${ANSI.reset}${ANSI.clear}${ANSI.home}${ANSI.mainScreen}`,
		)
		this.initialized = false
	}
}
