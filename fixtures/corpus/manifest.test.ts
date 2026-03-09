import { describe, expect, test } from 'bun:test'
import {
	type CorpusManifestEntry,
	matchesSelection,
	normalizeManifestEntry,
	selectManifestEntries,
} from './manifest.ts'

const PIVOT_ENTRY: CorpusManifestEntry = {
	file: 'pivot.xlsx',
	size_bytes: 5_000_000,
	features: {
		macros: false,
		charts: true,
		pivot_tables: true,
		tables: false,
		drawings: true,
		comments: false,
		threaded_comments: false,
		conditional_formatting: true,
		data_validations: false,
		merged_cells: false,
		hyperlinks: false,
		defined_names: true,
		external_links: false,
		connections: false,
		slicers: true,
		images_or_media: false,
		custom_xml: false,
		calc_chain: true,
	},
	counts: {
		worksheets: 4,
		charts: 3,
		tables: 0,
		drawings: 2,
		pivot_tables: 5,
		pivot_caches: 2,
		comments: 0,
	},
}

describe('normalizeManifestEntry', () => {
	test('derives tier, assertion class, risk, and tags', () => {
		const normalized = normalizeManifestEntry(PIVOT_ENTRY)
		expect(normalized.benchmarkTier).toBe('extended')
		expect(normalized.assertionClass).toBe('semantic-plus-package')
		expect(normalized.riskClass).toBe('high')
		expect(normalized.featureTags).toContain('pivot')
		expect(normalized.featureTags).toContain('slicer')
		expect(normalized.featureTags).toContain('large')
	})
})

describe('selectManifestEntries', () => {
	test('filters by tag and tier', () => {
		const entries = [
			normalizeManifestEntry(PIVOT_ENTRY),
			normalizeManifestEntry({
				file: 'small.xlsx',
				size_bytes: 13_000,
				features: {
					macros: false,
					charts: false,
					pivot_tables: false,
					tables: false,
					drawings: false,
					comments: false,
					threaded_comments: false,
					conditional_formatting: true,
					data_validations: false,
					merged_cells: false,
					hyperlinks: false,
					defined_names: false,
					external_links: false,
					connections: false,
					slicers: false,
					images_or_media: false,
					custom_xml: false,
					calc_chain: false,
				},
				counts: {
					worksheets: 1,
					charts: 0,
					tables: 0,
					drawings: 0,
					pivot_tables: 0,
					pivot_caches: 0,
					comments: 0,
				},
			}),
		]
		expect(selectManifestEntries(entries, { tags: ['pivot'] })).toHaveLength(1)
		expect(selectManifestEntries(entries, { tiers: ['smoke'] })).toHaveLength(1)
		expect(selectManifestEntries(entries, { risks: ['high'] })).toHaveLength(1)
		expect(selectManifestEntries(entries, { tags: ['conditional-formatting'] })).toHaveLength(2)
	})

	test('matchesSelection requires all requested tags', () => {
		const entry = normalizeManifestEntry(PIVOT_ENTRY)
		expect(matchesSelection(entry, { tags: ['pivot', 'slicer'] })).toBe(true)
		expect(matchesSelection(entry, { tags: ['pivot', 'macro'] })).toBe(false)
	})
})
