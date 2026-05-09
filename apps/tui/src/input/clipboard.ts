export interface GridClipboard {
	readonly text: string
	readonly mode: 'copy' | 'cut'
}

export interface ClipboardReadResult {
	readonly text: string
	readonly source: 'system' | 'internal' | 'empty'
}

export interface ClipboardWriteResult {
	readonly system: boolean
	readonly osc52: string
}

export class ClipboardController {
	private gridClipboard: GridClipboard | null = null

	async writeText(entry: GridClipboard): Promise<ClipboardWriteResult> {
		this.gridClipboard = entry
		let system = false
		try {
			const clipboard = await loadClipboard()
			await clipboard.write(entry.text)
			system = true
		} catch {
			system = false
		}
		return { system, osc52: osc52Clipboard(entry.text) }
	}

	async readText(): Promise<ClipboardReadResult> {
		try {
			const clipboard = await loadClipboard()
			const text = await clipboard.read()
			if (text.length > 0) return { text, source: 'system' }
		} catch {
			// Fall back to the in-process clipboard when the host denies clipboard access.
		}
		if (this.gridClipboard) return { text: this.gridClipboard.text, source: 'internal' }
		return { text: '', source: 'empty' }
	}

	setGridClipboard(entry: GridClipboard): void {
		this.gridClipboard = entry
	}

	getGridClipboard(): GridClipboard | null {
		return this.gridClipboard
	}
}

export function osc52Clipboard(text: string): string {
	return `\u001B]52;c;${Buffer.from(text).toString('base64')}\u0007`
}

async function loadClipboard(): Promise<typeof import('clipboardy').default> {
	return (await import('clipboardy')).default
}
