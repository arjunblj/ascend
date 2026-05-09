import type { FileHubEntry, FileHubState, OpenWorkbook, WorkbookWorkspace } from './types.ts'

export function createFileHub(entries: readonly FileHubEntry[] = []): FileHubState {
	return {
		visible: true,
		section: 'recent',
		query: '',
		selectedIndex: 0,
		entries,
	}
}

export function createWorkspace(document?: OpenWorkbook): WorkbookWorkspace {
	return {
		documents: document ? [document] : [],
		activeWorkbookId: document?.id ?? null,
		fileHub: document
			? { ...createFileHub(), visible: false }
			: createFileHub(defaultFileHubEntries()),
		focusedRegion: document ? 'grid' : 'fileHub',
	}
}

export function setActiveDocument(
	workspace: WorkbookWorkspace,
	document: OpenWorkbook,
): WorkbookWorkspace {
	const documents = upsertDocument(workspace.documents, document)
	return {
		...workspace,
		documents,
		activeWorkbookId: document.id,
		fileHub: { ...workspace.fileHub, visible: false },
		focusedRegion: 'grid',
	}
}

export function updateActiveDocument(
	workspace: WorkbookWorkspace,
	update: Partial<OpenWorkbook>,
): WorkbookWorkspace {
	const activeId = workspace.activeWorkbookId
	if (!activeId) return workspace
	return updateDocument(workspace, activeId, update)
}

export function updateDocument(
	workspace: WorkbookWorkspace,
	documentId: string,
	update: Partial<OpenWorkbook>,
): WorkbookWorkspace {
	return {
		...workspace,
		documents: workspace.documents.map((doc) =>
			doc.id === documentId ? { ...doc, ...update } : doc,
		),
	}
}

export function activeDocument(workspace: WorkbookWorkspace): OpenWorkbook | undefined {
	return workspace.documents.find((doc) => doc.id === workspace.activeWorkbookId)
}

export function showFileHub(workspace: WorkbookWorkspace, visible = true): WorkbookWorkspace {
	return {
		...workspace,
		fileHub: { ...workspace.fileHub, visible },
		focusedRegion: visible ? 'fileHub' : 'grid',
	}
}

function upsertDocument(
	documents: readonly OpenWorkbook[],
	document: OpenWorkbook,
): readonly OpenWorkbook[] {
	const existing = documents.findIndex((doc) => doc.id === document.id)
	if (existing < 0) return [...documents, document]
	return documents.map((doc, index) => (index === existing ? document : doc))
}

function defaultFileHubEntries(): readonly FileHubEntry[] {
	return [
		{ label: 'Open workbook', detail: 'Ctrl+O or :open <path>' },
		{ label: 'New workbook', detail: ':new' },
		{ label: 'Recover edits', detail: ':recover' },
		{ label: 'Performance traces', detail: ':perf' },
	]
}
