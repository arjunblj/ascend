import type { TelemetrySample } from './types.ts'

export class TelemetryBuffer {
	private readonly samples: TelemetrySample[] = []

	constructor(private readonly capacity = 512) {}

	record(sample: Omit<TelemetrySample, 'timestamp'>): void {
		this.samples.push({ timestamp: Date.now(), ...sample })
		if (this.samples.length > this.capacity) this.samples.shift()
	}

	all(): readonly TelemetrySample[] {
		return [...this.samples]
	}

	latest(): TelemetrySample | undefined {
		return this.samples.at(-1)
	}
}
