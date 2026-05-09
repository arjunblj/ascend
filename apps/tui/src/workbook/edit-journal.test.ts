import { describe, expect, test } from 'bun:test'
import { createSelection } from '../model/selection.ts'
import { EditJournal } from './edit-journal.ts'

describe('EditJournal', () => {
	test('undo/redo returns entries and append clears redo stack', () => {
		const journal = new EditJournal()
		const selection = createSelection()
		const first = journal.append({
			commandId: 'one',
			selectionBefore: selection,
			selectionAfter: selection,
			ops: [],
		})
		expect(journal.undo()?.id).toBe(first.id)
		expect(journal.redo()?.id).toBe(first.id)
		expect(journal.undo()?.id).toBe(first.id)
		journal.append({
			commandId: 'two',
			selectionBefore: selection,
			selectionAfter: selection,
			ops: [],
		})
		expect(journal.redo()).toBeUndefined()
		expect(journal.all().map((entry) => entry.commandId)).toEqual(['two'])
	})
})
