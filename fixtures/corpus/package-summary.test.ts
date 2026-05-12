import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../packages/io-xlsx/test/helpers.ts'
import { summarizeOoxmlPackage } from './package-summary.ts'

describe('OOXML package summaries', () => {
	test('counts pivot definitions without relationship or records sidecars', () => {
		const bytes = makeXlsx({
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <pivotCaches><pivotCache cacheId="34" r:id="rIdPivotCache"/></pivotCaches>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': '<worksheet/>',
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`,
			'xl/pivotTables/pivotTable1.xml': '<pivotTableDefinition name="PivotTable1"/>',
			'xl/pivotTables/_rels/pivotTable1.xml.rels': '<Relationships/>',
			'xl/pivotCache/pivotCacheDefinition1.xml': '<pivotCacheDefinition r:id="rIdRecords"/>',
			'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`,
			'xl/pivotCache/pivotCacheRecords1.xml': '<pivotCacheRecords/>',
		})

		expect(summarizeOoxmlPackage(bytes).families).toMatchObject({
			pivotTables: 1,
			pivotCaches: 1,
		})
	})

	test('counts slicer and timeline data parts without relationship folders', () => {
		const bytes = makeXlsx({
			'xl/workbook.xml': '<workbook/>',
			'xl/worksheets/sheet1.xml': '<worksheet/>',
			'xl/slicerCaches/cache_region.xml': '<slicerCacheDefinition/>',
			'xl/slicerCaches/_rels/cache_region.xml.rels': '<Relationships/>',
			'xl/slicers/ui_region.xml': '<slicers/>',
			'xl/slicers/_rels/ui_region.xml.rels': '<Relationships/>',
			'xl/timelineCaches/cache_date.xml': '<timelineCacheDefinition/>',
			'xl/timelineCaches/_rels/cache_date.xml.rels': '<Relationships/>',
			'xl/timelines/ui_date.xml': '<timelines/>',
			'xl/timelines/_rels/ui_date.xml.rels': '<Relationships/>',
		})

		expect(summarizeOoxmlPackage(bytes).families).toMatchObject({
			slicerCaches: 1,
			slicers: 1,
			timelineCaches: 1,
			timelines: 1,
		})
	})
})
