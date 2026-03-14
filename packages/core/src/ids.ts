declare const __brand: unique symbol
type Brand<T, B extends string> = T & { readonly [__brand]: B }

export type WorkbookId = Brand<string, 'WorkbookId'>
export type SheetId = Brand<string, 'SheetId'>
export type TableId = Brand<string, 'TableId'>
export type StyleId = Brand<number, 'StyleId'>

export const DEFAULT_STYLE_ID = 0 as StyleId

export function createWorkbookId(): WorkbookId {
	return crypto.randomUUID() as WorkbookId
}

export function createSheetId(): SheetId {
	return crypto.randomUUID() as SheetId
}

export function createTableId(): TableId {
	return crypto.randomUUID() as TableId
}
