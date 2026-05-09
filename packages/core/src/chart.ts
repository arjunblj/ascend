export interface ChartSeriesInfo {
	readonly nameRef?: string
	readonly nameText?: string
	readonly categoryRef?: string
	readonly valueRef?: string
}

export interface ChartPartInfo {
	readonly partPath: string
	readonly sheetName?: string
	readonly chartType?: string
	readonly title?: string
	readonly series: readonly ChartSeriesInfo[]
}
