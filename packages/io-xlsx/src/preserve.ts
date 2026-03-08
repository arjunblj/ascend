export interface PreservationCapsule {
	partPath: string
	contentType: string
	relationships: Array<{ id: string; type: string; target: string }>
	content: Uint8Array
	anchor: { kind: 'workbook' } | { kind: 'sheet'; sheetName: string }
	relType?: string
}
