export type CompatibilityTier = 'exact' | 'normalized' | 'preserved' | 'unsupported'

export type CompatibilityStatus = 'clean' | 'has-preserved' | 'has-unsupported'

export interface FeatureReport {
	readonly feature: string
	readonly tier: CompatibilityTier
	readonly count: number
	readonly locations: readonly string[]
	readonly note?: string
}

export interface CompatibilityReport {
	readonly status: CompatibilityStatus
	readonly features: readonly FeatureReport[]
	readonly summary: {
		readonly exact: number
		readonly normalized: number
		readonly preserved: number
		readonly unsupported: number
	}
	readonly sourceFormat: string
	readonly sourceApp?: string
}

export function emptyReport(sourceFormat: string): CompatibilityReport {
	return {
		status: 'clean',
		features: [],
		summary: { exact: 0, normalized: 0, preserved: 0, unsupported: 0 },
		sourceFormat,
	}
}
