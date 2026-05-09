import type { Operation } from '@ascend/schema'
import type { JournalEntry, SelectionState } from '../runtime/types.ts'

export class EditJournal {
	private generation = 0
	private readonly entries: JournalEntry[] = []
	private readonly undone: JournalEntry[] = []

	append(input: {
		readonly commandId: string
		readonly selectionBefore: SelectionState
		readonly selectionAfter: SelectionState
		readonly ops: readonly Operation[]
		readonly inverseOps?: readonly Operation[]
		readonly affectedCells?: readonly string[]
		readonly recalcDirtyRefs?: readonly string[]
	}): JournalEntry {
		const entry: JournalEntry = {
			id: `j${Date.now().toString(36)}-${this.generation + 1}`,
			generation: ++this.generation,
			commandId: input.commandId,
			selectionBefore: input.selectionBefore,
			selectionAfter: input.selectionAfter,
			ops: input.ops,
			inverseOps: input.inverseOps ?? [],
			preimageRanges: [],
			affectedCells: input.affectedCells ?? [],
			recalcDirtyRefs: input.recalcDirtyRefs ?? [],
			timestamp: Date.now(),
		}
		this.entries.push(entry)
		this.undone.length = 0
		return entry
	}

	undo(): JournalEntry | undefined {
		const entry = this.entries.pop()
		if (entry) this.undone.push(entry)
		return entry
	}

	redo(): JournalEntry | undefined {
		const entry = this.undone.pop()
		if (entry) this.entries.push(entry)
		return entry
	}

	clear(): void {
		this.entries.length = 0
		this.undone.length = 0
	}

	all(): readonly JournalEntry[] {
		return [...this.entries]
	}
}
