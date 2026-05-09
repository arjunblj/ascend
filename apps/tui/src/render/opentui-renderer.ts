import { type CliRenderer, createCliRenderer, Text, type TextRenderable } from '@opentui/core'
import stripAnsi from 'strip-ansi'
import type {
	RenderFrame,
	RenderStats,
	TerminalCapabilities,
	TerminalRenderer,
} from '../runtime/types.ts'

export class OpenTuiRenderer implements TerminalRenderer {
	private renderer: CliRenderer | null = null
	private text: TextRenderable | null = null
	private lastDraw = 0

	async init(capabilities: TerminalCapabilities): Promise<void> {
		if (this.renderer) return
		const renderer = await createCliRenderer({
			exitOnCtrlC: false,
			testing: !capabilities.isTty,
			consoleMode: 'disabled',
			clearOnShutdown: true,
			gatherStats: true,
			targetFps: 60,
			maxFps: 120,
			useMouse: capabilities.mouse,
			useKittyKeyboard:
				capabilities.keyboardProtocol === 'kitty'
					? { disambiguate: true, alternateKeys: true, reportText: true }
					: null,
		})
		renderer.root.add(
			Text({
				content: '',
				width: '100%',
				height: '100%',
			}),
		)
		this.renderer = renderer
		this.text = (renderer.root.getChildren()[0] as TextRenderable | undefined) ?? null
	}

	async draw(frame: RenderFrame): Promise<RenderStats> {
		if (!this.renderer || !this.text) {
			await this.init({
				isTty: false,
				color: 'truecolor',
				unicode: true,
				mouse: false,
				bracketedPaste: false,
				hyperlinks: false,
				graphics: 'off',
				keyboardProtocol: 'legacy',
				profile: 'legacy',
			})
		}
		if (!this.renderer || !this.text) throw new Error('OpenTUI renderer failed to initialize.')
		const started = performance.now()
		const content = frame.lines.map((line) => stripAnsi(line)).join('\n')
		this.text.content = content
		this.renderer.requestRender()
		this.renderer.intermediateRender()
		const now = performance.now()
		const stats = this.renderer.getStats()
		const drawMs = now - started
		const fps = this.lastDraw > 0 ? 1000 / Math.max(1, now - this.lastDraw) : stats.fps
		this.lastDraw = now
		return {
			frameBuildMs: drawMs,
			frameDiffMs: 0,
			encodeMs: 0,
			writeMs: drawMs,
			changedCells: frame.size.rows * frame.size.cols,
			bytesOut: Buffer.byteLength(content),
			droppedFrames: 0,
			fps,
		}
	}

	async shutdown(): Promise<void> {
		this.text = null
		this.renderer?.destroy()
		this.renderer = null
	}
}
