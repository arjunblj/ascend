export interface PreservationCapsule {
	partPath: string
	contentType: string
	contentTypeSource?: 'override' | 'default' | 'fallback'
	relationships: Array<{ id: string; type: string; target: string; targetMode?: string }>
	content?: Uint8Array
	anchor: { kind: 'workbook' } | { kind: 'sheet'; sheetId: string; sheetName?: string }
	relType?: string
}
