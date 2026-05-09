export type SaveAction = 'save' | 'save-as' | 'save-copy' | 'export'

export function describeSaveAction(action: SaveAction): string {
	switch (action) {
		case 'save':
			return 'Save updates the current workbook.'
		case 'save-as':
			return 'Save As writes a new file and continues editing it.'
		case 'save-copy':
			return 'Save a Copy writes another file and keeps editing the original.'
		case 'export':
			return 'Export writes another format and may lose workbook features.'
	}
}
