export interface CalcSettings {
	readonly calcMode: 'auto' | 'manual'
	readonly fullCalcOnLoad: boolean
	readonly dateSystem: '1900' | '1904'
	readonly iterativeCalc: {
		readonly enabled: boolean
		readonly maxIterations: number
		readonly maxChange: number
	}
}

export const DEFAULT_CALC_SETTINGS: CalcSettings = {
	calcMode: 'auto',
	fullCalcOnLoad: false,
	dateSystem: '1900',
	iterativeCalc: {
		enabled: false,
		maxIterations: 100,
		maxChange: 0.001,
	},
}

export interface ImportOptions {
	readonly password?: string
	readonly locale?: string
}

export interface ExportOptions {
	readonly recalcBeforeExport?: boolean
	readonly preserveCapsules?: boolean
}

export interface CsvDialect {
	readonly delimiter: string
	readonly quote: string
	readonly escape: string
	readonly lineEnding: '\n' | '\r\n'
	readonly encoding: string
	readonly hasHeader: boolean
	readonly bom: boolean
}

export const DEFAULT_CSV_DIALECT: CsvDialect = {
	delimiter: ',',
	quote: '"',
	escape: '"',
	lineEnding: '\n',
	encoding: 'utf-8',
	hasHeader: true,
	bom: false,
}
