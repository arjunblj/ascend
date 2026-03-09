export type CorpusBenchmarkTier = 'smoke' | 'core' | 'extended' | 'stress'
export type CorpusAssertionClass =
	| 'exact-bytes'
	| 'preservation-only'
	| 'semantic-only'
	| 'semantic-plus-package'
export type CorpusRiskClass = 'low' | 'medium' | 'high'

export interface CorpusManifestEntry {
	readonly file: string
	readonly size_bytes: number
	readonly features: Record<string, boolean>
	readonly counts: Record<string, number>
	readonly source?: string
	readonly license?: string
	readonly vendorable?: boolean
	readonly benchmarkTier?: CorpusBenchmarkTier
	readonly assertionClass?: CorpusAssertionClass
	readonly riskClass?: CorpusRiskClass
	readonly featureTags?: readonly string[]
	readonly knownUnsupported?: readonly string[]
	readonly notes?: string
}

export interface NormalizedCorpusManifestEntry {
	readonly file: string
	readonly size_bytes: number
	readonly features: Record<string, boolean>
	readonly counts: Record<string, number>
	readonly source?: string
	readonly license?: string
	readonly vendorable: boolean
	readonly benchmarkTier: CorpusBenchmarkTier
	readonly assertionClass: CorpusAssertionClass
	readonly riskClass: CorpusRiskClass
	readonly featureTags: readonly string[]
	readonly knownUnsupported: readonly string[]
	readonly notes?: string
}

export interface CorpusSelection {
	readonly file?: string
	readonly tags?: readonly string[]
	readonly tiers?: readonly CorpusBenchmarkTier[]
	readonly risks?: readonly CorpusRiskClass[]
	readonly assertionClasses?: readonly CorpusAssertionClass[]
	readonly vendorableOnly?: boolean
}

export function normalizeManifest(
	entries: readonly CorpusManifestEntry[],
): readonly NormalizedCorpusManifestEntry[] {
	return entries.map(normalizeManifestEntry)
}

export function normalizeManifestEntry(entry: CorpusManifestEntry): NormalizedCorpusManifestEntry {
	const featureTags = entry.featureTags ?? deriveFeatureTags(entry)
	return {
		file: entry.file,
		size_bytes: entry.size_bytes,
		features: { ...entry.features },
		counts: { ...entry.counts },
		...(entry.source ? { source: entry.source } : {}),
		...(entry.license ? { license: entry.license } : {}),
		vendorable: entry.vendorable ?? false,
		benchmarkTier: entry.benchmarkTier ?? deriveBenchmarkTier(entry),
		assertionClass: entry.assertionClass ?? deriveAssertionClass(entry),
		riskClass: entry.riskClass ?? deriveRiskClass(entry),
		featureTags: [...new Set(featureTags)].sort((a, b) => a.localeCompare(b)),
		knownUnsupported: [...(entry.knownUnsupported ?? [])],
		...(entry.notes ? { notes: entry.notes } : {}),
	}
}

export function selectManifestEntries(
	entries: readonly NormalizedCorpusManifestEntry[],
	selection: CorpusSelection,
): readonly NormalizedCorpusManifestEntry[] {
	return entries.filter((entry) => matchesSelection(entry, selection))
}

export function matchesSelection(
	entry: NormalizedCorpusManifestEntry,
	selection: CorpusSelection,
): boolean {
	if (selection.file && entry.file !== selection.file) return false
	if (selection.vendorableOnly && !entry.vendorable) return false
	if (selection.tags && selection.tags.length > 0) {
		for (const tag of selection.tags) {
			if (!entry.featureTags.includes(tag)) return false
		}
	}
	if (
		selection.tiers &&
		selection.tiers.length > 0 &&
		!selection.tiers.includes(entry.benchmarkTier)
	) {
		return false
	}
	if (selection.risks && selection.risks.length > 0 && !selection.risks.includes(entry.riskClass)) {
		return false
	}
	if (
		selection.assertionClasses &&
		selection.assertionClasses.length > 0 &&
		!selection.assertionClasses.includes(entry.assertionClass)
	) {
		return false
	}
	return true
}

export function deriveFeatureTags(entry: CorpusManifestEntry): readonly string[] {
	const tags = new Set<string>()
	const featureTagMap: Record<string, string> = {
		macros: 'macro',
		charts: 'chart',
		pivot_tables: 'pivot',
		tables: 'table',
		drawings: 'drawing',
		comments: 'comment',
		threaded_comments: 'threaded-comment',
		conditional_formatting: 'conditional-formatting',
		data_validations: 'data-validation',
		merged_cells: 'merged-cells',
		hyperlinks: 'hyperlink',
		defined_names: 'defined-names',
		external_links: 'external-link',
		connections: 'connection',
		slicers: 'slicer',
		images_or_media: 'media',
		custom_xml: 'custom-xml',
		calc_chain: 'calc-chain',
	}
	for (const [feature, enabled] of Object.entries(entry.features)) {
		if (!enabled) continue
		tags.add(featureTagMap[feature] ?? feature.replaceAll('_', '-'))
	}
	if (entry.features.calc_chain) tags.add('formula-fidelity')
	if (entry.size_bytes >= 3_000_000) tags.add('large')
	else if (entry.size_bytes >= 250_000) tags.add('medium')
	else tags.add('small')
	return [...tags]
}

function deriveBenchmarkTier(entry: CorpusManifestEntry): CorpusBenchmarkTier {
	if (entry.size_bytes >= 4_000_000) return 'extended'
	if (entry.features.pivot_tables || entry.features.slicers || entry.features.macros)
		return 'extended'
	if (entry.features.charts || entry.features.tables || entry.features.drawings) return 'core'
	return 'smoke'
}

function deriveAssertionClass(entry: CorpusManifestEntry): CorpusAssertionClass {
	if (entry.features.pivot_tables || entry.features.slicers || entry.features.macros) {
		return 'semantic-plus-package'
	}
	if (entry.features.charts || entry.features.drawings || entry.features.custom_xml) {
		return 'preservation-only'
	}
	if (
		entry.features.conditional_formatting ||
		entry.features.data_validations ||
		entry.features.tables
	) {
		return 'semantic-plus-package'
	}
	return 'exact-bytes'
}

function deriveRiskClass(entry: CorpusManifestEntry): CorpusRiskClass {
	if (
		entry.features.pivot_tables ||
		entry.features.slicers ||
		entry.features.macros ||
		entry.features.external_links ||
		entry.features.connections
	) {
		return 'high'
	}
	if (
		entry.features.charts ||
		entry.features.drawings ||
		entry.features.conditional_formatting ||
		entry.features.data_validations ||
		entry.features.calc_chain ||
		entry.features.tables
	) {
		return 'medium'
	}
	return 'low'
}
