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
</Relationships>`,
			'xl/worksheets/sheet1.xml': '<worksheet/>',
			'xl/worksheets/sheet2.xml': '<worksheet><sheetProtection sheet="1"/></worksheet>',
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
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
			'xl/slicers/slicer1.xml': `<?xml version="1.0"?>
<slicers><slicer name="Product" cache="Slicer_Product"/></slicers>`,
			'xl/timelineCaches/timelineCache1.xml': `<?xml version="1.0"?>
<timelineCacheDefinition name="Timeline_Date">
  <data><tabular pivotCacheId="34"/></data>
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
</timelineCacheDefinition>`,
			'xl/timelines/timeline1.xml': `<?xml version="1.0"?>
<timelines><timeline name="Date" cache="Timeline_Date"/></timelines>`,
		})

		const probe = inspectOoxmlPackageFeatures(bytes)

		expect(probe.counts).toMatchObject({
			pivot_tables: 1,
			pivot_caches: 1,
			slicers: 1,
			slicer_caches: 1,
			timelines: 1,
			timeline_caches: 1,
			workbook_protection: 1,
			sheet_protection: 1,
		})
		expect(probe.features).toMatchObject({
			pivot_tables: true,
			slicers: true,
			timelines: true,
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
})
