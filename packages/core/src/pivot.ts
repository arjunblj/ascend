export interface PivotTableInfo {
	readonly partPath: string
	readonly sheetName: string
	readonly name?: string
	readonly cacheId?: number
	readonly locationRef?: string
	readonly location?: PivotTableLocationInfo
	readonly options?: PivotTableOptionsInfo
	readonly style?: PivotTableStyleInfo
	readonly fields: readonly PivotFieldInfo[]
	readonly rowFields: readonly PivotFieldReference[]
	readonly columnFields: readonly PivotFieldReference[]
	readonly pageFields: readonly PivotFieldReference[]
	readonly dataFields: readonly PivotDataFieldInfo[]
	readonly rowItems?: readonly PivotAxisItemInfo[]
	readonly columnItems?: readonly PivotAxisItemInfo[]
	readonly formats?: readonly PivotFormatInfo[]
	readonly chartFormats?: readonly PivotChartFormatInfo[]
}

export interface PivotTableLocationInfo {
	readonly ref?: string
	readonly firstHeaderRow?: number
	readonly firstDataRow?: number
	readonly firstDataCol?: number
	readonly rowPageCount?: number
	readonly colPageCount?: number
}

export interface PivotTableOptionsInfo {
	readonly applyAlignmentFormats?: boolean
	readonly applyBorderFormats?: boolean
	readonly applyFontFormats?: boolean
	readonly applyNumberFormats?: boolean
	readonly applyPatternFormats?: boolean
	readonly applyWidthHeightFormats?: boolean
	readonly colGrandTotals?: boolean
	readonly rowGrandTotals?: boolean
	readonly compact?: boolean
	readonly compactData?: boolean
	readonly dataOnRows?: boolean
	readonly enableDrill?: boolean
	readonly enableEdit?: boolean
	readonly gridDropZones?: boolean
	readonly hideValuesRow?: boolean
	readonly itemPrintTitles?: boolean
	readonly multipleFieldFilters?: boolean
	readonly outline?: boolean
	readonly outlineData?: boolean
	readonly showItems?: boolean
	readonly showMemberPropertyTips?: boolean
	readonly showMultipleLabel?: boolean
	readonly useAutoFormatting?: boolean
	readonly indent?: number
	readonly createdVersion?: number
	readonly updatedVersion?: number
	readonly minRefreshableVersion?: number
	readonly dataPosition?: number
	readonly chartFormat?: number
	readonly dataCaption?: string
	readonly rowHeaderCaption?: string
	readonly colHeaderCaption?: string
	readonly fillDownLabelsDefault?: boolean
	readonly enabledSubtotalsDefault?: boolean
	readonly subtotalsOnTopDefault?: boolean
}

export interface PivotTableStyleInfo {
	readonly name?: string
	readonly showRowHeaders?: boolean
	readonly showColHeaders?: boolean
	readonly showRowStripes?: boolean
	readonly showColStripes?: boolean
	readonly showLastColumn?: boolean
}

export interface PivotCacheInfo {
	readonly partPath: string
	readonly cacheId?: number
	readonly relId?: string
	readonly recordCount?: number
	readonly refreshedVersion?: number
	readonly minRefreshableVersion?: number
	readonly createdVersion?: number
	readonly refreshedBy?: string
	readonly refreshedDate?: number
	readonly refreshOnLoad?: boolean
	readonly enableRefresh?: boolean
	readonly invalid?: boolean
	readonly saveData?: boolean
	readonly optimizeMemory?: boolean
	readonly upgradeOnRefresh?: boolean
	readonly extensionCacheId?: number
	readonly sourceType?: string
	readonly sourceSheet?: string
	readonly sourceRef?: string
	readonly sourceName?: string
	readonly recordsPartPath?: string
	readonly records?: PivotCacheRecordsInfo
	readonly fields: readonly PivotCacheFieldInfo[]
}

export interface PivotCacheRecordsInfo {
	readonly partPath: string
	readonly declaredCount?: number
	readonly parsedCount: number
	readonly preview: readonly PivotCacheRecordInfo[]
	readonly materializedRecords?: readonly PivotCacheRecordInfo[]
	readonly materializedCount?: number
	readonly materializedComplete?: boolean
	readonly valueKindCounts: readonly PivotCacheRecordValueKindCount[]
}

export interface PivotCacheRecordInfo {
	readonly index: number
	readonly values: readonly PivotCacheRecordValueInfo[]
}

export interface PivotCacheRecordValueInfo {
	readonly index: number
	readonly kind: PivotCacheRecordValueKind
	readonly value?: string
	readonly sharedItemIndex?: number
}

export interface PivotCacheRecordValueKindCount {
	readonly kind: PivotCacheRecordValueKind
	readonly count: number
}

export type PivotCacheRecordValueKind =
	| 'string'
	| 'number'
	| 'date'
	| 'boolean'
	| 'error'
	| 'missing'
	| 'sharedItem'
	| 'unknown'

export interface PivotCacheFieldInfo {
	readonly index: number
	readonly name?: string
	readonly databaseField?: boolean
	readonly numFmtId?: number
	readonly formula?: string
	readonly sharedItemsInfo?: PivotCacheSharedItemsInfo
	readonly sharedItems?: readonly PivotCacheSharedItemInfo[]
	readonly fieldGroup?: PivotCacheFieldGroupInfo
}

export interface PivotCacheSharedItemInfo {
	readonly index: number
	readonly kind: 'string' | 'number' | 'date' | 'boolean' | 'error' | 'missing'
	readonly value?: string
}

export interface PivotCacheSharedItemsInfo {
	readonly count?: number
	readonly containsBlank?: boolean
	readonly containsDate?: boolean
	readonly containsNonDate?: boolean
	readonly containsNumber?: boolean
	readonly containsInteger?: boolean
	readonly containsString?: boolean
	readonly containsMixedTypes?: boolean
	readonly containsSemiMixedTypes?: boolean
	readonly minValue?: number
	readonly maxValue?: number
	readonly minDate?: string
	readonly maxDate?: string
}

export interface PivotCacheFieldGroupInfo {
	readonly base?: number
	readonly parent?: number
	readonly range?: PivotCacheFieldGroupRangeInfo
	readonly discreteItems?: readonly PivotCacheFieldGroupDiscreteItemInfo[]
	readonly groupItems?: readonly PivotCacheSharedItemInfo[]
}

export interface PivotCacheFieldGroupRangeInfo {
	readonly groupBy?: string
	readonly startDate?: string
	readonly endDate?: string
	readonly startNumber?: number
	readonly endNumber?: number
	readonly groupInterval?: number
	readonly autoStart?: boolean
	readonly autoEnd?: boolean
}

export interface PivotCacheFieldGroupDiscreteItemInfo {
	readonly index: number
	readonly value?: number
}

export interface PivotFieldInfo {
	readonly index: number
	readonly axis?: string
	readonly name?: string
	readonly numFmtId?: number
	readonly hidden?: boolean
	readonly dataField?: boolean
	readonly defaultSubtotal?: boolean
	readonly showAll?: boolean
	readonly multipleItemSelectionAllowed?: boolean
	readonly compact?: boolean
	readonly outline?: boolean
	readonly subtotalTop?: boolean
	readonly dragToRow?: boolean
	readonly dragToCol?: boolean
	readonly dragToPage?: boolean
	readonly includeNewItemsInFilter?: boolean
	readonly fillDownLabels?: boolean
	readonly itemPageCount?: number
	readonly sortType?: string
	readonly items?: readonly PivotFieldItemInfo[]
}

export interface PivotFieldItemInfo {
	readonly index: number
	readonly cacheIndex?: number
	readonly itemType?: string
	readonly caption?: string
	readonly hidden?: boolean
	readonly manualFilter?: boolean
	readonly showDetails?: boolean
	readonly calculated?: boolean
	readonly missing?: boolean
	readonly childItems?: boolean
	readonly expanded?: boolean
	readonly drillAcrossAttributes?: boolean
}

export interface PivotFieldReference {
	readonly index: number
	readonly name?: string
	readonly item?: number
	readonly hierarchy?: number
	readonly caption?: string
}

export interface PivotDataFieldInfo {
	readonly fieldIndex: number
	readonly name?: string
	readonly subtotal?: string
	readonly numFmtId?: number
	readonly showDataAs?: string
	readonly baseField?: number
	readonly baseItem?: number
}

export interface PivotAxisItemInfo {
	readonly index: number
	readonly itemType?: string
	readonly repeatedItemCount?: number
	readonly dataFieldIndex?: number
	readonly fieldItems: readonly PivotAxisFieldItemInfo[]
}

export interface PivotAxisFieldItemInfo {
	readonly index: number
	readonly item?: number
}

export interface PivotFormatInfo {
	readonly index: number
	readonly dxfId?: number
	readonly action?: string
	readonly area?: PivotAreaInfo
}

export interface PivotChartFormatInfo {
	readonly index: number
	readonly chart?: number
	readonly formatId?: number
	readonly series?: boolean
	readonly area?: PivotAreaInfo
}

export interface PivotAreaInfo {
	readonly type?: string
	readonly axis?: string
	readonly field?: number
	readonly fieldPosition?: number
	readonly dataOnly?: boolean
	readonly labelOnly?: boolean
	readonly grandRow?: boolean
	readonly grandCol?: boolean
	readonly cacheIndex?: boolean
	readonly outline?: boolean
	readonly collapsedLevelsAreSubtotals?: boolean
	readonly references?: readonly PivotAreaReferenceInfo[]
}

export interface PivotAreaReferenceInfo {
	readonly index: number
	readonly field?: number
	readonly itemCount?: number
	readonly selected?: boolean
	readonly items: readonly PivotAxisFieldItemInfo[]
}

export function clonePivotCacheInfo(entry: PivotCacheInfo): PivotCacheInfo {
	return {
		...entry,
		...(entry.records
			? {
					records: {
						...entry.records,
						preview: entry.records.preview.map((record) => ({
							...record,
							values: record.values.map((value) => ({ ...value })),
						})),
						...(entry.records.materializedRecords
							? {
									materializedRecords: entry.records.materializedRecords.map((record) => ({
										...record,
										values: record.values.map((value) => ({ ...value })),
									})),
								}
							: {}),
						valueKindCounts: entry.records.valueKindCounts.map((count) => ({ ...count })),
					},
				}
			: {}),
		fields: entry.fields.map((field) => ({
			...field,
			...(field.sharedItemsInfo ? { sharedItemsInfo: { ...field.sharedItemsInfo } } : {}),
			...(field.sharedItems ? { sharedItems: field.sharedItems.map((item) => ({ ...item })) } : {}),
			...(field.fieldGroup
				? {
						fieldGroup: {
							...field.fieldGroup,
							...(field.fieldGroup.range ? { range: { ...field.fieldGroup.range } } : {}),
							...(field.fieldGroup.discreteItems
								? { discreteItems: field.fieldGroup.discreteItems.map((item) => ({ ...item })) }
								: {}),
							...(field.fieldGroup.groupItems
								? { groupItems: field.fieldGroup.groupItems.map((item) => ({ ...item })) }
								: {}),
						},
					}
				: {}),
		})),
	}
}

export function clonePivotTableInfo(entry: PivotTableInfo): PivotTableInfo {
	return {
		...entry,
		...(entry.location ? { location: { ...entry.location } } : {}),
		...(entry.options ? { options: { ...entry.options } } : {}),
		...(entry.style ? { style: { ...entry.style } } : {}),
		fields: entry.fields.map((field) => ({
			...field,
			...(field.items ? { items: field.items.map((item) => ({ ...item })) } : {}),
		})),
		rowFields: entry.rowFields.map((field) => ({ ...field })),
		columnFields: entry.columnFields.map((field) => ({ ...field })),
		pageFields: entry.pageFields.map((field) => ({ ...field })),
		dataFields: entry.dataFields.map((field) => ({ ...field })),
		...(entry.rowItems
			? { rowItems: entry.rowItems.map((item) => clonePivotAxisItemInfo(item)) }
			: {}),
		...(entry.columnItems
			? { columnItems: entry.columnItems.map((item) => clonePivotAxisItemInfo(item)) }
			: {}),
		...(entry.formats
			? { formats: entry.formats.map((format) => clonePivotFormatInfo(format)) }
			: {}),
		...(entry.chartFormats
			? { chartFormats: entry.chartFormats.map((format) => clonePivotChartFormatInfo(format)) }
			: {}),
	}
}

function clonePivotAxisItemInfo(entry: PivotAxisItemInfo): PivotAxisItemInfo {
	return {
		...entry,
		fieldItems: entry.fieldItems.map((item) => ({ ...item })),
	}
}

function clonePivotFormatInfo(entry: PivotFormatInfo): PivotFormatInfo {
	return {
		...entry,
		...(entry.area ? { area: clonePivotAreaInfo(entry.area) } : {}),
	}
}

function clonePivotChartFormatInfo(entry: PivotChartFormatInfo): PivotChartFormatInfo {
	return {
		...entry,
		...(entry.area ? { area: clonePivotAreaInfo(entry.area) } : {}),
	}
}

function clonePivotAreaInfo(entry: PivotAreaInfo): PivotAreaInfo {
	return {
		...entry,
		...(entry.references
			? {
					references: entry.references.map((reference) => ({
						...reference,
						items: reference.items.map((item) => ({ ...item })),
					})),
				}
			: {}),
	}
}

export interface SlicerCacheInfo {
	readonly partPath: string
	readonly name?: string
	readonly sourceName?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
	readonly items?: readonly SlicerCacheItemInfo[]
}

export interface SlicerCacheItemInfo {
	readonly index: number
	readonly selected?: boolean
	readonly noData?: boolean
}

export interface SlicerInfo {
	readonly partPath: string
	readonly name?: string
	readonly cacheName?: string
	readonly caption?: string
}

export interface TimelineCacheInfo {
	readonly partPath: string
	readonly name?: string
	readonly sourceName?: string
	readonly pivotCacheId?: number
	readonly pivotTableNames: readonly string[]
	readonly state?: TimelineStateInfo
}

export interface TimelineInfo {
	readonly partPath: string
	readonly name?: string
	readonly cacheName?: string
	readonly caption?: string
}

export interface TimelineRangeInfo {
	readonly startDate: string
	readonly endDate: string
}

export interface TimelineStateInfo {
	readonly filterType?: string
	readonly filterId?: number
	readonly filterPivotName?: string
	readonly filterTabId?: number
	readonly lastRefreshVersion?: number
	readonly minimalRefreshVersion?: number
	readonly pivotCacheId?: number
	readonly singleRangeFilterState?: boolean
	readonly selection?: TimelineRangeInfo
	readonly bounds?: TimelineRangeInfo
}
