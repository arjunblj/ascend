import { describe, expect, test } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { stringValue } from '@ascend/schema'
import { unzipSync } from 'fflate'
import { makeXlsx } from '../../test/helpers.ts'
import { inspectXlsxPackageGraph, type XlsxPackageGraph } from '../package-graph.ts'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
} from '../package-graph-fidelity.ts'
import { readXlsx } from '../reader/index.ts'
import { writeXlsx } from './index.ts'

const S0 = 0 as StyleId

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('slicer and timeline package fidelity', () => {
	test('preserves workbook cache relationships and sidecar parts after a dirty safe edit', () => {
		const sourceBytes = analyticsPackageWorkbook()
		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)

		expect(auditXlsxPackageGraphReadIntegrity(beforeGraph)).toEqual([])
		expect(analyticsPartIdentities(beforeGraph)).toEqual([
			{
				path: 'xl/slicerCaches/slicerCache7.xml',
				featureFamily: 'preservedSlicer',
				ownerScope: 'slicer',
				contentType: 'application/vnd.ms-excel.slicerCache+xml',
			},
			{
				path: 'xl/slicers/slicer7.xml',
				featureFamily: 'preservedSlicer',
				ownerScope: 'slicer',
				contentType: 'application/vnd.ms-excel.slicer+xml',
			},
			{
				path: 'xl/timelineCaches/timelineCache4.xml',
				featureFamily: 'preservedTimeline',
				ownerScope: 'timeline',
				contentType: 'application/vnd.ms-excel.timelineCache+xml',
			},
			{
				path: 'xl/timelines/timeline4.xml',
				featureFamily: 'preservedTimeline',
				ownerScope: 'timeline',
				contentType: 'application/vnd.ms-excel.timeline+xml',
			},
		])
		expect(analyticsRelationshipIdentities(beforeGraph)).toEqual([
			{
				sourcePartPath: 'xl/workbook.xml',
				relationshipPartPath: 'xl/_rels/workbook.xml.rels',
				id: 'rIdSlicerCacheOdd',
				rawTarget: 'slicerCaches/slicerCache7.xml',
				resolvedTarget: 'xl/slicerCaches/slicerCache7.xml',
				type: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
			},
			{
				sourcePartPath: 'xl/workbook.xml',
				relationshipPartPath: 'xl/_rels/workbook.xml.rels',
				id: 'rIdTimelineCacheAbsolute',
				rawTarget: '/xl/timelineCaches/timelineCache4.xml',
				resolvedTarget: 'xl/timelineCaches/timelineCache4.xml',
				type: 'http://schemas.microsoft.com/office/2011/relationships/timelineCache',
			},
			{
				sourcePartPath: 'xl/slicerCaches/slicerCache7.xml',
				relationshipPartPath: 'xl/slicerCaches/_rels/slicerCache7.xml.rels',
				id: 'rIdSlicerUi',
				rawTarget: '../slicers/slicer7.xml',
				resolvedTarget: 'xl/slicers/slicer7.xml',
				type: 'http://schemas.microsoft.com/office/2007/relationships/slicer',
			},
			{
				sourcePartPath: 'xl/timelineCaches/timelineCache4.xml',
				relationshipPartPath: 'xl/timelineCaches/_rels/timelineCache4.xml.rels',
				id: 'rIdTimelineUiFromCache',
				rawTarget: '/xl/timelines/timeline4.xml',
				resolvedTarget: 'xl/timelines/timeline4.xml',
				type: 'http://schemas.microsoft.com/office/2011/relationships/timeline',
			},
			{
				sourcePartPath: 'xl/worksheets/sheet2.xml',
				relationshipPartPath: 'xl/worksheets/_rels/sheet2.xml.rels',
				id: 'rIdTimelineUi',
				rawTarget: '../timelines/timeline4.xml',
				resolvedTarget: 'xl/timelines/timeline4.xml',
				type: 'http://schemas.microsoft.com/office/2011/relationships/timeline',
			},
		])

		const opened = readXlsx(sourceBytes)
		expectOk(opened)
		expect(opened.value.workbook.slicerCaches).toEqual([
			{
				partPath: 'xl/slicerCaches/slicerCache7.xml',
				name: 'Slicer_Region',
				sourceName: 'Region',
				pivotCacheId: 42,
				pivotTableNames: ['PivotTableMain'],
				items: [{ index: 0, selected: true }, { index: 1 }],
			},
		])
		expect(opened.value.workbook.slicers).toEqual([
			{
				partPath: 'xl/slicers/slicer7.xml',
				name: 'Region',
				cacheName: 'Slicer_Region',
				caption: 'Region',
			},
		])
		expect(opened.value.workbook.timelineCaches).toEqual([
			{
				partPath: 'xl/timelineCaches/timelineCache4.xml',
				name: 'Timeline_Order_Date',
				sourceName: 'Order Date',
				pivotCacheId: 42,
				pivotTableNames: ['PivotTableMain'],
				state: {
					filterId: 9,
					filterPivotName: 'PivotTableMain',
					filterType: 'dateRange',
					filterTabId: 2,
					pivotCacheId: 42,
					singleRangeFilterState: true,
					selection: {
						startDate: '2024-01-01T00:00:00',
						endDate: '2024-03-31T00:00:00',
					},
				},
			},
		])
		expect(opened.value.workbook.timelines).toEqual([
			{
				partPath: 'xl/timelines/timeline4.xml',
				name: 'Order_Date',
				cacheName: 'Timeline_Order_Date',
				caption: 'Order Date',
			},
		])

		const dataSheet = opened.value.workbook.sheets.find((sheet) => sheet.name === 'Data')
		if (!dataSheet) throw new Error('Data sheet was not parsed')
		dataSheet.cells.set(3, 0, {
			value: stringValue('safe dirty edit outside analytics package'),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: ['Data'],
			workbookMetaDirty: true,
		})
		expectOk(written)

		const afterGraph = inspectXlsxPackageGraph(written.value)
		expect(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph)).toEqual([])
		expect(auditXlsxPackageGraphBytePreservation(beforeGraph, sourceBytes, written.value)).toEqual(
			[],
		)
		expect(analyticsPartIdentities(afterGraph)).toEqual(analyticsPartIdentities(beforeGraph))
		expect(analyticsRelationshipIdentities(afterGraph)).toEqual(
			analyticsRelationshipIdentities(beforeGraph),
		)

		const beforeZip = unzipSync(sourceBytes)
		const afterZip = unzipSync(written.value)
		const workbookXml = decode(afterZip['xl/workbook.xml'])
		expect(workbookXml).toContain(
			'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"',
		)
		expect(workbookXml).toContain(
			'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"',
		)
		expect(workbookXml).toContain('xmlns:mx="urn:ascend:test-workbook-ext"')
		expect(workbookXml).toContain('<ext uri="{ASCEND-UNKNOWN-WORKBOOK-EXT}">')
		expect(workbookXml).toContain('<mx:keep attr="A&amp;B" nested="1"/>')
		expect(workbookXml).toContain('<x14:slicerCache r:id="rIdSlicerCacheOdd"/>')
		expect(workbookXml).toContain('<x15:timelineCacheRef r:id="rIdTimelineCacheAbsolute"/>')
		for (const partPath of [
			'xl/slicerCaches/slicerCache7.xml',
			'xl/slicerCaches/_rels/slicerCache7.xml.rels',
			'xl/slicers/slicer7.xml',
			'xl/timelineCaches/timelineCache4.xml',
			'xl/timelineCaches/_rels/timelineCache4.xml.rels',
			'xl/timelines/timeline4.xml',
		]) {
			expect(decode(afterZip[partPath]), partPath).toBe(decode(beforeZip[partPath]))
		}

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.slicerCaches).toEqual(opened.value.workbook.slicerCaches)
		expect(reopened.value.workbook.slicers).toEqual(opened.value.workbook.slicers)
		expect(reopened.value.workbook.timelineCaches).toEqual(opened.value.workbook.timelineCaches)
		expect(reopened.value.workbook.timelines).toEqual(opened.value.workbook.timelines)
	})

	test('preserves cache-to-UI package edges when slicer and timeline cache XML is edited', () => {
		const sourceBytes = analyticsPackageWorkbook()
		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
		const opened = readXlsx(sourceBytes)
		expectOk(opened)

		const slicerCache = opened.value.workbook.slicerCaches[0]
		const timelineCache = opened.value.workbook.timelineCaches[0]
		if (!slicerCache || !timelineCache) throw new Error('analytics caches were not parsed')
		opened.value.workbook.slicerCaches[0] = {
			...slicerCache,
			items: [{ index: 0 }, { index: 1, selected: true, noData: true }],
		}
		opened.value.workbook.timelineCaches[0] = {
			...timelineCache,
			state: {
				...timelineCache.state,
				singleRangeFilterState: true,
				selection: {
					startDate: '2024-04-01T00:00:00',
					endDate: '2024-06-30T00:00:00',
				},
			},
		}
		const dataSheet = opened.value.workbook.sheets.find((sheet) => sheet.name === 'Data')
		if (!dataSheet) throw new Error('Data sheet was not parsed')
		dataSheet.cells.set(3, 1, {
			value: stringValue('cache edit with dirty patch'),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: ['Data'],
			workbookMetaDirty: true,
		})
		expectOk(written)

		const afterGraph = inspectXlsxPackageGraph(written.value)
		expect(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph)).toEqual([])
		expect(analyticsPartIdentities(afterGraph)).toEqual(analyticsPartIdentities(beforeGraph))
		expect(analyticsRelationshipIdentities(afterGraph)).toEqual(
			analyticsRelationshipIdentities(beforeGraph),
		)

		const beforeZip = unzipSync(sourceBytes)
		const afterZip = unzipSync(written.value)
		const workbookXml = decode(afterZip['xl/workbook.xml'])
		expect(workbookXml).toContain('xmlns:mx="urn:ascend:test-workbook-ext"')
		expect(workbookXml).toContain('<ext uri="{ASCEND-UNKNOWN-WORKBOOK-EXT}">')
		expect(workbookXml).toContain('<mx:keep attr="A&amp;B" nested="1"/>')
		expect(workbookXml).toContain('<x14:slicerCache r:id="rIdSlicerCacheOdd"/>')
		expect(workbookXml).toContain('<x15:timelineCacheRef r:id="rIdTimelineCacheAbsolute"/>')
		const slicerCacheXml = decode(afterZip['xl/slicerCaches/slicerCache7.xml'])
		const timelineCacheXml = decode(afterZip['xl/timelineCaches/timelineCache4.xml'])
		expect(slicerCacheXml).toContain('<i x="0"/>')
		expect(slicerCacheXml).toContain('<i x="1" s="1" nd="1"/>')
		expect(timelineCacheXml).toContain(
			'<selection startDate="2024-04-01T00:00:00" endDate="2024-06-30T00:00:00"/>',
		)
		for (const partPath of ['xl/slicers/slicer7.xml', 'xl/timelines/timeline4.xml']) {
			expect(decode(afterZip[partPath]), partPath).toBe(decode(beforeZip[partPath]))
		}

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.slicerCaches[0]?.items).toEqual([
			{ index: 0 },
			{ index: 1, selected: true, noData: true },
		])
		expect(reopened.value.workbook.timelineCaches[0]?.state?.selection).toEqual({
			startDate: '2024-04-01T00:00:00',
			endDate: '2024-06-30T00:00:00',
		})
		expect(reopened.value.workbook.slicers).toEqual(opened.value.workbook.slicers)
		expect(reopened.value.workbook.timelines).toEqual(opened.value.workbook.timelines)
	})
})

function analyticsPartIdentities(graph: XlsxPackageGraph): readonly Record<string, unknown>[] {
	return graph.parts
		.filter(
			(part) =>
				part.featureFamily === 'preservedSlicer' || part.featureFamily === 'preservedTimeline',
		)
		.map((part) => ({
			path: part.path,
			featureFamily: part.featureFamily,
			ownerScope: part.ownerScope,
			contentType: part.contentType,
		}))
		.sort((left, right) => String(left.path).localeCompare(String(right.path)))
}

function analyticsRelationshipIdentities(
	graph: XlsxPackageGraph,
): readonly Record<string, unknown>[] {
	return graph.relationships
		.filter(
			(relationship) =>
				relationship.featureFamily === 'preservedSlicer' ||
				relationship.featureFamily === 'preservedTimeline',
		)
		.map((relationship) => ({
			sourcePartPath: relationship.sourcePartPath,
			relationshipPartPath: relationship.relationshipPartPath,
			id: relationship.id,
			type: relationship.type,
			rawTarget: relationship.rawTarget,
			resolvedTarget: relationship.resolvedTarget,
		}))
		.sort((left, right) =>
			`${left.relationshipPartPath}#${left.id}`.localeCompare(
				`${right.relationshipPartPath}#${right.id}`,
			),
		)
}

function decode(bytes: Uint8Array | undefined): string {
	if (!bytes) return ''
	return new TextDecoder().decode(bytes)
}

function analyticsPackageWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache7.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/slicers/slicer7.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCache4.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>
  <Override PartName="/xl/timelines/timeline4.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheetData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdSheetPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rIdPivotCacheMain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition3.xml"/>
  <Relationship Id="rIdSlicerCacheOdd" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache7.xml"/>
  <Relationship Id="rIdTimelineCacheAbsolute" Type="http://schemas.microsoft.com/office/2011/relationships/timelineCache" Target="/xl/timelineCaches/timelineCache4.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
  xmlns:mx="urn:ascend:test-workbook-ext">
  <pivotCaches><pivotCache cacheId="42" r:id="rIdPivotCacheMain"/></pivotCaches>
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rIdSheetData"/>
    <sheet name="Pivot" sheetId="2" r:id="rIdSheetPivot"/>
  </sheets>
  <extLst>
    <ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}">
      <x14:slicerCaches><x14:slicerCache r:id="rIdSlicerCacheOdd"/></x14:slicerCaches>
    </ext>
    <ext uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}">
      <x15:timelineCaches><x15:timelineCacheRef r:id="rIdTimelineCacheAbsolute"/></x15:timelineCaches>
    </ext>
    <ext uri="{ASCEND-UNKNOWN-WORKBOOK-EXT}">
      <mx:keep attr="A&amp;B" nested="1"/>
    </ext>
  </extLst>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Order Date</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>West</t></is></c><c r="B2"><v>45292</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>East</t></is></c><c r="B3"><v>45323</v></c></row>
  </sheetData>
</worksheet>`,
		'xl/worksheets/sheet2.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/worksheets/_rels/sheet2.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotTableMain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable3.xml"/>
  <Relationship Id="rIdTimelineUi" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline4.xml"/>
</Relationships>`,
		'xl/pivotTables/pivotTable3.xml': `<?xml version="1.0"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTableMain" cacheId="42">
  <location ref="A3:C8" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="2"><pivotField axis="axisRow"/><pivotField dataField="1"/></pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <dataFields count="1"><dataField fld="1" name="Count of Region" subtotal="count"/></dataFields>
</pivotTableDefinition>`,
		'xl/pivotCache/pivotCacheDefinition3.xml': `<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  r:id="rIdCacheRecordsMain" recordCount="2" refreshOnLoad="1" enableRefresh="1" saveData="1">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Data"/></cacheSource>
  <cacheFields count="2">
    <cacheField name="Region" databaseField="1"><sharedItems count="2"><s v="West"/><s v="East"/></sharedItems></cacheField>
    <cacheField name="Order Date" databaseField="1"><sharedItems containsDate="1" count="2"><d v="2024-01-01T00:00:00"/><d v="2024-02-01T00:00:00"/></sharedItems></cacheField>
  </cacheFields>
</pivotCacheDefinition>`,
		'xl/pivotCache/_rels/pivotCacheDefinition3.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdCacheRecordsMain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords3.xml"/>
</Relationships>`,
		'xl/pivotCache/pivotCacheRecords3.xml': `<?xml version="1.0"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
  <r><x v="0"/><d v="2024-01-01T00:00:00"/></r>
  <r><x v="1"/><d v="2024-02-01T00:00:00"/></r>
</pivotCacheRecords>`,
		'xl/slicerCaches/slicerCache7.xml': `<?xml version="1.0"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_Region" sourceName="Region">
  <pivotTables><pivotTable name="PivotTableMain"/></pivotTables>
  <data><tabular pivotCacheId="42"><items count="2"><i x="0" s="1"/><i x="1"/></items></tabular></data>
</slicerCacheDefinition>`,
		'xl/slicerCaches/_rels/slicerCache7.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSlicerUi" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer7.xml"/>
</Relationships>`,
		'xl/slicers/slicer7.xml': `<?xml version="1.0"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="Region" cache="Slicer_Region" caption="Region"/>
</slicers>`,
		'xl/timelineCaches/timelineCache4.xml': `<?xml version="1.0"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="Timeline_Order_Date" sourceName="Order Date">
  <data><tabular pivotCacheId="42"/></data>
  <pivotTables><pivotTable tabId="1" name="PivotTableMain"/></pivotTables>
  <state filterId="9" filterPivotName="PivotTableMain" filterType="dateRange" filterTabId="2" pivotCacheId="42" singleRangeFilterState="1">
    <selection startDate="2024-01-01T00:00:00" endDate="2024-03-31T00:00:00"/>
  </state>
</timelineCacheDefinition>`,
		'xl/timelineCaches/_rels/timelineCache4.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTimelineUiFromCache" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="/xl/timelines/timeline4.xml"/>
</Relationships>`,
		'xl/timelines/timeline4.xml': `<?xml version="1.0"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <timeline name="Order_Date" cache="Timeline_Order_Date" caption="Order Date"/>
</timelines>`,
	})
}
