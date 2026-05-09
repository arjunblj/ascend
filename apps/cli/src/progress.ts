import {
	type AgentWorkflowProgressHandler,
	type AgentWorkflowTrace,
	AscendWorkbook,
	WorkbookDocument,
} from '@ascend/sdk'

const SPINNER_FRAMES = ['-', '\\', '|', '/'] as const
const DEFAULT_DELAY_MS = 350
const DEFAULT_TICK_MS = 125

export async function withProgress<T>(
	label: string,
	task: () => T | Promise<T>,
): Promise<{ value: T; durationMs: number }> {
	if (!process.stderr.isTTY) {
		const startedAt = performance.now()
		const value = await Promise.resolve(task())
		return { value, durationMs: performance.now() - startedAt }
	}

	let timer: ReturnType<typeof setTimeout> | undefined
	let interval: ReturnType<typeof setInterval> | undefined
	let shown = false
	let frame = 0
	const startedAt = performance.now()

	const render = (): void => {
		const elapsedMs = performance.now() - startedAt
		const spinner = SPINNER_FRAMES[frame % SPINNER_FRAMES.length] ?? '-'
		frame += 1
		process.stderr.write(`\r${spinner} ${label} ${formatElapsed(elapsedMs)}`)
	}

	timer = setTimeout(() => {
		shown = true
		render()
		interval = setInterval(render, DEFAULT_TICK_MS)
	}, DEFAULT_DELAY_MS)

	try {
		const value = await Promise.resolve(task())
		const durationMs = performance.now() - startedAt
		return { value, durationMs }
	} finally {
		if (timer) clearTimeout(timer)
		if (interval) clearInterval(interval)
		if (shown) clearStatusLine()
	}
}

export async function openWorkbookWithProgress(
	file: string | Uint8Array,
	options?: {
		mode?: 'full' | 'metadata-only' | 'values' | 'formula'
		sheets?: readonly string[]
		richMetadata?: boolean
	},
): Promise<{ workbook: AscendWorkbook; durationMs: number }> {
	const label = typeof file === 'string' ? `Opening ${file}` : 'Opening workbook'
	const { value, durationMs } = await withProgress(label, () => AscendWorkbook.open(file, options))
	return { workbook: value, durationMs }
}

export async function openWorkbookDocumentWithProgress(
	file: string,
	options?: {
		mode?: 'full' | 'metadata-only' | 'values' | 'formula'
		sheets?: readonly string[]
		richMetadata?: boolean
	},
): Promise<{ document: WorkbookDocument; durationMs: number }> {
	const label = `Opening ${file}`
	const { value, durationMs } = await withProgress(label, () =>
		WorkbookDocument.open(file, options),
	)
	return { document: value, durationMs }
}

export type CliProgressWriter = (event: object) => void

export function createAgentProgressReporter(
	flags: Map<string, string>,
): AgentWorkflowProgressHandler | undefined {
	const write = createJsonlProgressWriter(flags)
	if (!write) return undefined
	return (event) => write(event)
}

export function createJsonlProgressWriter(
	flags: Map<string, string>,
): CliProgressWriter | undefined {
	const mode = flags.get('progress')
	if (!mode) return undefined
	if (mode !== 'jsonl') {
		throw new Error(`Unsupported --progress mode: ${mode}. Use --progress jsonl.`)
	}
	return (event) => {
		process.stderr.write(`${JSON.stringify({ type: 'progress', ...event })}\n`)
	}
}

export function emitCliProgress(
	write: CliProgressWriter | undefined,
	event: {
		readonly sequence: number
		readonly kind: AgentWorkflowTrace['kind'] | 'check'
		readonly phase: string
		readonly status: 'started' | 'ok' | 'warning' | 'blocked' | 'failed' | 'skipped'
		readonly summary: string
		readonly count?: number
		readonly details?: unknown
	},
): void {
	write?.({ formatVersion: 1, ...event })
}

function clearStatusLine(): void {
	process.stderr.write('\r')
	process.stderr.write(' '.repeat(80))
	process.stderr.write('\r')
}

function formatElapsed(elapsedMs: number): string {
	if (elapsedMs < 1000) return `${Math.round(elapsedMs)}ms`
	return `${(elapsedMs / 1000).toFixed(1)}s`
}
