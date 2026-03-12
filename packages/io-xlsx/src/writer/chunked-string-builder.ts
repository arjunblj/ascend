export class ChunkedStringBuilder {
	private chunks: string[] = []

	push(s: string): void {
		this.chunks.push(s)
	}

	toString(): string {
		return this.chunks.join('')
	}

	get length(): number {
		return this.chunks.length
	}
}
