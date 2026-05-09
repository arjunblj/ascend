export interface CalcGeneration {
	readonly generation: number
	readonly staleRefs: readonly string[]
}

export class CalcWorker {
	private generation = 0

	next(staleRefs: readonly string[] = []): CalcGeneration {
		return {
			generation: ++this.generation,
			staleRefs,
		}
	}
}
