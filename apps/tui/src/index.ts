import { TerminalInputParser } from './input/keymap.ts'
import { AnsiRenderer } from './render/ansi-renderer.ts'
import { detectTerminalCapabilities } from './render/terminal-capabilities.ts'
import { WorkbookTuiEngine } from './runtime/engine.ts'
import type { InputEvent, TelemetrySample, TerminalSize } from './runtime/types.ts'

const DEFAULT_OPEN_PREVIEW_ROWS = 500

export { WorkbookTuiEngine } from './runtime/engine.ts'
export { runHeadlessTrace } from './runtime/headless-trace.ts'
export type {
	CommandDescriptor,
	DispatchResult,
	InputEvent,
	RenderFrame,
	TerminalRenderer,
	TerminalSize,
	TraceResult,
	TuiEngine,
	TuiStateSnapshot,
} from './runtime/types.ts'

export const usage = `ascend tui [file] [flags]

Interactive terminal spreadsheet.

Flags:
  --sheet <name>          Start on a sheet
  --preview-rows <n>      Open the first n rows in read-only values mode
  --renderer <name>       ansi (default) or opentui
  --calibrate             Open terminal capability calibration on start
  --telemetry-json        Print telemetry JSON after the session exits
`

export const openUsage = `ascend open [file] [flags]

Friendly preview-first entrypoint for the Ascend terminal spreadsheet.

Flags:
  --sheet <name>          Start on a sheet
  --preview-rows <n>      Open the first n rows in read-only values mode (default: 500 for files)
  --renderer <name>       ansi (default) or opentui
  --calibrate             Open terminal capability calibration on start
  --telemetry-json        Print telemetry JSON after the session exits
`

export async function tuiCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const sheet = flags.get('sheet')
	const previewRows = parsePreviewRows(flags.get('preview-rows'))
	const telemetry = await runTui({
		...(args[0] ? { path: args[0] } : {}),
		...(sheet ? { sheet } : {}),
		...(previewRows !== undefined ? { previewRows } : {}),
		renderer: flags.get('renderer') === 'opentui' ? 'opentui' : 'ansi',
		calibrate: flags.has('calibrate'),
	})
	if (flags.has('telemetry-json')) console.log(JSON.stringify(telemetry, null, 2))
	return 0
}

export async function openCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const previewFlags = new Map(flags)
	if (args[0] && !previewFlags.has('preview-rows')) {
		previewFlags.set('preview-rows', String(DEFAULT_OPEN_PREVIEW_ROWS))
	}
	return tuiCommand(args, previewFlags)
}

export async function runTui(input: {
	readonly path?: string
	readonly sheet?: string
	readonly trace?: readonly InputEvent[]
	readonly size?: TerminalSize
	readonly renderer?: 'ansi' | 'opentui'
	readonly calibrate?: boolean
	readonly recentStorePath?: string
	readonly previewRows?: number
}): Promise<readonly TelemetrySample[]> {
	const size = input.size ?? currentTerminalSize()
	const engine = await WorkbookTuiEngine.create({
		...(input.path ? { path: input.path } : {}),
		...(input.sheet ? { sheet: input.sheet } : {}),
		...(input.recentStorePath ? { recentStorePath: input.recentStorePath } : {}),
		...(input.previewRows !== undefined
			? { loadOptions: { mode: 'values', maxRows: input.previewRows } }
			: {}),
		persistState: process.stdin.isTTY && process.stdout.isTTY,
		size,
	})
	if (input.calibrate) await engine.dispatch({ kind: 'command', command: 'calibrate' })
	if (input.trace) {
		const result = await engine.runHeadless(input.trace, { size, includeFrames: true })
		return result.telemetry
	}

	const frame = engine.render(size)
	if (!process.stdin.isTTY || !process.stdout.isTTY) {
		console.log(frame.lines.join('\n'))
		return engine.state().telemetry
	}

	const renderer =
		input.renderer === 'opentui'
			? new (await import('./render/opentui-renderer.ts')).OpenTuiRenderer()
			: new AnsiRenderer()

	return new Promise((resolve) => {
		let closed = false
		const inputParser = new TerminalInputParser()
		let drawQueue = Promise.resolve()
		let inputQueue = Promise.resolve()
		let drawScheduled = false
		let escapeFlushTimer: ReturnType<typeof setTimeout> | undefined
		const cleanup = async (): Promise<void> => {
			if (closed) return
			closed = true
			if (escapeFlushTimer) clearTimeout(escapeFlushTimer)
			process.stdin.off('data', onData)
			process.stdout.off('resize', onResize)
			process.off('SIGINT', onSigint)
			process.off('SIGTERM', onSignal)
			process.off('SIGHUP', onSignal)
			process.off('uncaughtException', onFatal)
			process.off('unhandledRejection', onFatal)
			if (typeof process.stdin.setRawMode === 'function') process.stdin.setRawMode(false)
			process.stdin.pause()
			await renderer.shutdown()
			resolve(engine.state().telemetry)
		}
		const cleanupAfterError = async (error: unknown) => {
			console.error(error instanceof Error ? error.message : String(error))
			await cleanup()
		}
		const draw = async () => {
			engine.recordRendererStats(await renderer.draw(engine.render(currentTerminalSize())))
		}
		const enqueueDraw = () => {
			if (drawScheduled) return drawQueue
			drawScheduled = true
			drawQueue = drawQueue
				.then(async () => {
					while (drawScheduled && !closed) {
						drawScheduled = false
						await draw()
					}
				})
				.catch(cleanupAfterError)
			return drawQueue
		}
		const dispatchEvents = async (events: readonly InputEvent[]) => {
			let shouldRender = false
			let shouldExit = false
			for (const event of events) {
				const result = await engine.dispatch(event)
				shouldRender ||= result.shouldRender
				shouldExit ||= result.shouldExit ?? false
				if (shouldExit) break
			}
			if (shouldRender) await enqueueDraw()
			if (shouldExit) await cleanup()
		}
		const enqueueInput = (task: () => Promise<void>) => {
			inputQueue = inputQueue.then(task).catch(cleanupAfterError)
			return inputQueue
		}
		const scheduleAmbiguousEscapeFlush = () => {
			if (escapeFlushTimer) clearTimeout(escapeFlushTimer)
			if (!inputParser.hasPending()) return
			escapeFlushTimer = setTimeout(() => {
				void enqueueInput(async () => {
					const events = inputParser.flush()
					if (events.length > 0) await dispatchEvents(events)
				})
			}, 25)
		}
		const onData = (buffer: Buffer) => {
			void enqueueInput(async () => {
				if (escapeFlushTimer) clearTimeout(escapeFlushTimer)
				let shouldRender = false
				let shouldExit = false
				for (const event of inputParser.push(buffer)) {
					const result = await engine.dispatch(event)
					shouldRender ||= result.shouldRender
					shouldExit ||= result.shouldExit ?? false
					if (shouldExit) break
				}
				if (shouldRender) await enqueueDraw()
				if (shouldExit) await cleanup()
				scheduleAmbiguousEscapeFlush()
			})
		}
		const onResize = () => {
			void (async () => {
				await engine.dispatch({ kind: 'resize', size: currentTerminalSize() })
				await enqueueDraw()
			})().catch(cleanupAfterError)
		}
		const onSigint = () => {
			void cleanup()
		}
		const onSignal = () => {
			void cleanup()
		}
		const onFatal = (error: unknown) => {
			void cleanup().then(() => {
				setImmediate(() => {
					throw error instanceof Error ? error : new Error(String(error))
				})
			})
		}

		process.on('SIGINT', onSigint)
		process.on('SIGTERM', onSignal)
		process.on('SIGHUP', onSignal)
		process.on('uncaughtException', onFatal)
		process.on('unhandledRejection', onFatal)

		void (async () => {
			await renderer.init(detectTerminalCapabilities())
			engine.recordRendererStats(await renderer.draw(frame))
			if (typeof process.stdin.setRawMode === 'function') process.stdin.setRawMode(true)
			process.stdin.resume()
			process.stdin.on('data', onData)
			process.stdout.on('resize', onResize)
		})().catch(cleanupAfterError)
	})
}

function currentTerminalSize(): TerminalSize {
	return {
		rows: process.stdout.rows || 32,
		cols: process.stdout.columns || 120,
	}
}

function parsePreviewRows(value: string | undefined): number | undefined {
	if (value === undefined) return undefined
	const parsed = Number.parseInt(value, 10)
	if (!Number.isFinite(parsed) || parsed < 1) {
		throw new Error('--preview-rows must be a positive integer')
	}
	return parsed
}
