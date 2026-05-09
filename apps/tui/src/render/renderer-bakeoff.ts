import type { RenderFrame } from '../runtime/types.ts'
import { encodeAnsiFrame } from './ansi-encoder.ts'

export type RendererCandidateId = 'ansi' | 'opentui'

export interface RendererBakeoffResult {
	readonly candidate: RendererCandidateId
	readonly status: 'passed' | 'unavailable' | 'failed'
	readonly frames: number
	readonly p95Ms: number | null
	readonly medianMs: number | null
	readonly bytesOut: number
	readonly reason?: string
}

export async function runRendererBakeoff(
	frames: readonly RenderFrame[],
	candidates: readonly RendererCandidateId[] = ['ansi', 'opentui'],
): Promise<readonly RendererBakeoffResult[]> {
	const results: RendererBakeoffResult[] = []
	for (const candidate of candidates) {
		switch (candidate) {
			case 'ansi':
				results.push(runAnsiBakeoff(frames))
				break
			case 'opentui':
				results.push(await runOpenTuiBakeoff(frames))
				break
		}
	}
	return results
}

function runAnsiBakeoff(frames: readonly RenderFrame[]): RendererBakeoffResult {
	let previous: RenderFrame | null = null
	const samples: number[] = []
	let bytesOut = 0
	for (const frame of frames) {
		const start = performance.now()
		const encoded = encodeAnsiFrame(previous, frame)
		samples.push(performance.now() - start)
		bytesOut += encoded.encoded.length
		previous = frame
	}
	return {
		candidate: 'ansi',
		status: 'passed',
		frames: frames.length,
		p95Ms: percentile(samples, 0.95),
		medianMs: percentile(samples, 0.5),
		bytesOut,
	}
}

async function runOpenTuiBakeoff(frames: readonly RenderFrame[]): Promise<RendererBakeoffResult> {
	const { OpenTuiRenderer } = await import('./opentui-renderer.ts')
	const renderer = new OpenTuiRenderer()
	const samples: number[] = []
	let bytesOut = 0
	try {
		await renderer.init({
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
		for (const frame of frames) {
			const start = performance.now()
			const stats = await renderer.draw(frame)
			samples.push(performance.now() - start)
			bytesOut += stats.bytesOut
		}
		return {
			candidate: 'opentui',
			status: 'passed',
			frames: frames.length,
			p95Ms: percentile(samples, 0.95),
			medianMs: percentile(samples, 0.5),
			bytesOut,
		}
	} catch (error) {
		return {
			candidate: 'opentui',
			status: 'failed',
			frames: frames.length,
			p95Ms: null,
			medianMs: null,
			bytesOut,
			reason: error instanceof Error ? error.message : String(error),
		}
	} finally {
		await renderer.shutdown()
	}
}

function percentile(samples: readonly number[], p: number): number | null {
	if (samples.length === 0) return null
	const sorted = [...samples].sort((a, b) => a - b)
	const index = Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * p) - 1))
	return sorted[index] ?? null
}
