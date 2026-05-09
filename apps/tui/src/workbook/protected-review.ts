import { auditLossPolicy, type WorkbookInfo } from '@ascend/sdk'

export function protectedReviewReasons(info: WorkbookInfo): readonly string[] {
	const reasons: string[] = []
	if (info.externalReferenceCount > 0) reasons.push('External references are present.')
	if (info.pivotTableCount > 0)
		reasons.push('PivotTables are inspectable but not fully refreshable.')
	if (info.chartCount > 0)
		reasons.push('Charts are preserved and previewed before semantic editing.')
	for (const feature of auditLossPolicy(info.compatibility.features).blockedFeatures) {
		if (feature.tier === 'preserved' || feature.tier === 'unsupported') {
			reasons.push(`${feature.feature}: ${feature.tier}`)
		}
	}
	return reasons
}
