import type { InputValue, Operation } from '@ascend/schema'

export function setCell(sheet: string, ref: string, value: InputValue): Operation {
	return { op: 'setCells', sheet, updates: [{ ref, value }] }
}

export function setFormula(sheet: string, ref: string, formula: string): Operation {
	return { op: 'setFormula', sheet, ref, formula }
}

export function addSheet(name: string): Operation {
	return { op: 'addSheet', name }
}

export function deleteSheet(sheet: string): Operation {
	return { op: 'deleteSheet', sheet }
}

export function renameSheet(sheet: string, newName: string): Operation {
	return { op: 'renameSheet', sheet, newName }
}

export function insertRows(sheet: string, at: number, count: number): Operation {
	return { op: 'insertRows', sheet, at, count }
}

export function deleteRows(sheet: string, at: number, count: number): Operation {
	return { op: 'deleteRows', sheet, at, count }
}

export function insertCols(sheet: string, at: number, count: number): Operation {
	return { op: 'insertCols', sheet, at, count }
}

export function deleteCols(sheet: string, at: number, count: number): Operation {
	return { op: 'deleteCols', sheet, at, count }
}
