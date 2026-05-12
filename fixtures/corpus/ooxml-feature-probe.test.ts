import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../packages/io-xlsx/test/helpers.ts'
import { inspectOoxmlPackageFeatures } from './ooxml-feature-probe.ts'

describe('OOXML feature probe', () => {
	test('extracts pivot, slicer, and timeline relationship integrity independent of Ascend reader', () => {
		const bytes = makeXlsx({
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookProtection lockStructure="1"/>
  <pivotCaches><pivotCache cacheId="34" r:id="rIdPivotCache"/></pivotCaches>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotCache" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheDefinition" Target="/xl/pivotCache/./pivotCacheDefinition1.xml"/>
  <Relationship Id="rIdSlicerCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches\\slicerCache1.xml"/>
  <Relationship Id="rIdTimelineCache" Type="http://schemas.microsoft.com/office/2011/relationships/timelineCache" Target="/xl/timelineCaches/timelineCache1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': '<worksheet/>',
			'xl/worksheets/sheet2.xml': '<worksheet><sheetProtection sheet="1"/></worksheet>',
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
  <Relationship Id="rIdTimeline" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline1.xml"/>
</Relationships>`,
			'xl/pivotTables/pivotTable1.xml': `<?xml version="1.0"?>
<pivotTableDefinition name="PivotTable1" cacheId="34">
  <location ref="A3:D20"/>
</pivotTableDefinition>`,
			'xl/pivotCache/pivotCacheDefinition1.xml': `<?xml version="1.0"?>
<pivotCacheDefinition r:id="rIdRecords"/>`,
			'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`,
			'xl/pivotCache/pivotCacheRecords1.xml': '<pivotCacheRecords/>',
			'xl/slicerCaches/slicerCache1.xml': `<?xml version="1.0"?>
<slicerCacheDefinition name="Slicer_Product" sourceName="Product">
  <data><tabular pivotCacheId="34"/></data>
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
</slicerCacheDefinition>`,
			'xl/slicerCaches/_rels/slicerCache1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSlicer" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
</Relationships>`,
			'xl/slicers/slicer1.xml': `<?xml version="1.0"?>
<slicers><slicer name="Product" cache="Slicer_Product"/></slicers>`,
			'xl/timelineCaches/timelineCache1.xml': `<?xml version="1.0"?>
<timelineCacheDefinition name="Timeline_Date">
  <data><tabular pivotCacheId="34"/></data>
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
</timelineCacheDefinition>`,
			'xl/timelines/timeline1.xml': `<?xml version="1.0"?>
<timelines><timeline name="Date" cache="Timeline_Date"/></timelines>`,
			'customUI/customUI.xml': `<?xml version="1.0"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="Ribbon.OnLoad"/>`,
		})

		const probe = inspectOoxmlPackageFeatures(bytes)

		expect(probe.counts).toMatchObject({
			pivot_tables: 1,
			pivot_caches: 1,
			slicers: 1,
			slicer_caches: 1,
			timelines: 1,
			timeline_caches: 1,
			active_content: 1,
			workbook_protection: 1,
			sheet_protection: 1,
		})
		expect(probe.features).toMatchObject({
			pivot_tables: true,
			slicers: true,
			timelines: true,
			active_content: true,
			workbook_protection: true,
			sheet_protection: true,
			protection: true,
		})
		expect(probe.analytics.workbookPivotCaches).toEqual([{ cacheId: 34, relId: 'rIdPivotCache' }])
		expect(probe.analytics.pivotCacheRelationships[0]?.targetPartPath).toBe(
			'xl/pivotCache/pivotCacheDefinition1.xml',
		)
		expect(probe.analytics.pivotTableRelationships[0]?.targetPartPath).toBe(
			'xl/pivotTables/pivotTable1.xml',
		)
		expect(probe.analytics.pivotCaches[0]?.recordsPartPath).toBe(
			'xl/pivotCache/pivotCacheRecords1.xml',
		)
		expect(probe.analytics.slicerCacheRelationships[0]?.targetPartPath).toBe(
			'xl/slicerCaches/slicerCache1.xml',
		)
		expect(probe.analytics.slicerRelationships[0]?.targetPartPath).toBe('xl/slicers/slicer1.xml')
		expect(probe.analytics.timelineCacheRelationships[0]?.targetPartPath).toBe(
			'xl/timelineCaches/timelineCache1.xml',
		)
		expect(probe.analytics.timelineRelationships).toEqual([
			expect.objectContaining({
				sourcePartPath: 'xl/worksheets/sheet1.xml',
				id: 'rIdTimeline',
				targetPartPath: 'xl/timelines/timeline1.xml',
			}),
		])
		expect(probe.analytics.slicerCaches[0]).toMatchObject({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_Product',
			pivotCacheId: 34,
			pivotTableNames: ['PivotTable1'],
		})
		expect(probe.analytics.slicers[0]).toMatchObject({
			partPath: 'xl/slicers/slicer1.xml',
			name: 'Product',
			cacheName: 'Slicer_Product',
		})
		expect(probe.analytics.timelineCaches[0]?.name).toBe('Timeline_Date')
		expect(probe.analytics.timelines[0]?.cacheName).toBe('Timeline_Date')
	})

	test('extracts nonnumeric slicer and timeline analytics parts', () => {
		const bytes = makeXlsx({
			'xl/workbook.xml': '<workbook/>',
			'xl/worksheets/sheet1.xml': '<worksheet/>',
			'xl/tables/custom_sales.xml': '<table/>',
			'xl/tables/_rels/custom_sales.xml.rels': '<Relationships/>',
			'xl/tables/custom_sales.bin': 'not xml',
			'xl/slicerCaches/cache_region.xml': `<?xml version="1.0"?>
<slicerCacheDefinition name="Slicer_Region">
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
</slicerCacheDefinition>`,
			'xl/slicerCaches/_rels/cache_region.xml.rels': '<Relationships/>',
			'xl/slicerCaches/cache_region.bin': 'not xml',
			'xl/slicers/ui_region.xml': `<?xml version="1.0"?>
<slicers><slicer name="Region" cache="Slicer_Region"/></slicers>`,
			'xl/slicers/ui_region.bin': 'not xml',
			'xl/timelineCaches/cache_date.xml': `<?xml version="1.0"?>
<timelineCacheDefinition name="Timeline_Date">
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
</timelineCacheDefinition>`,
			'xl/timelineCaches/_rels/cache_date.xml.rels': '<Relationships/>',
			'xl/timelineCaches/cache_date.bin': 'not xml',
			'xl/timelines/ui_date.xml': `<?xml version="1.0"?>
<timelines><timeline name="Date" cache="Timeline_Date"/></timelines>`,
			'xl/timelines/ui_date.bin': 'not xml',
		})

		const probe = inspectOoxmlPackageFeatures(bytes)

		expect(probe.counts).toMatchObject({
			tables: 1,
			slicer_caches: 1,
			slicers: 1,
			timeline_caches: 1,
			timelines: 1,
		})
		expect(probe.analytics.slicerCaches[0]).toMatchObject({
			partPath: 'xl/slicerCaches/cache_region.xml',
			name: 'Slicer_Region',
			pivotTableNames: ['PivotTable1'],
		})
		expect(probe.analytics.slicers[0]).toMatchObject({
			partPath: 'xl/slicers/ui_region.xml',
			name: 'Region',
			cacheName: 'Slicer_Region',
		})
		expect(probe.analytics.timelineCaches[0]).toMatchObject({
			partPath: 'xl/timelineCaches/cache_date.xml',
			name: 'Timeline_Date',
			pivotTableNames: ['PivotTable1'],
		})
		expect(probe.analytics.timelines[0]).toMatchObject({
			partPath: 'xl/timelines/ui_date.xml',
			name: 'Date',
			cacheName: 'Timeline_Date',
		})
	})
})
