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

describe('pivot package fidelity', () => {
	test('preserves pivot table and cache relationship graph after a dirty safe edit', () => {
		const sourceBytes = pivotPackageWorkbook()
		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
		expect(auditXlsxPackageGraphReadIntegrity(beforeGraph)).toEqual([])
		expect(pivotPartIdentities(beforeGraph)).toEqual([
			{
				path: 'xl/pivotCache/pivotCacheDefinition5.xml',
				featureFamily: 'preservedPivot',
				ownerScope: 'pivot',
				contentType:
					'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
			},
			{
				path: 'xl/pivotCache/pivotCacheDefinition9.xml',
				featureFamily: 'preservedPivot',
				ownerScope: 'pivot',
				contentType:
					'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
			},
			{
				path: 'xl/pivotCache/pivotCacheRecords5.xml',
				featureFamily: 'preservedPivot',
				ownerScope: 'pivot',
				contentType:
					'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
			},
			{
				path: 'xl/pivotCache/pivotCacheRecords9.xml',
				featureFamily: 'preservedPivot',
				ownerScope: 'pivot',
				contentType:
					'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
			},
			{
				path: 'xl/pivotTables/pivotTable1.xml',
				featureFamily: 'preservedPivot',
				ownerScope: 'pivot',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
			},
			{
				path: 'xl/pivotTables/pivotTable2.xml',
				featureFamily: 'preservedPivot',
				ownerScope: 'pivot',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
			},
		])
		expect(pivotRelationshipIdentities(beforeGraph)).toEqual([
			{
				sourcePartPath: 'xl/workbook.xml',
				relationshipPartPath: 'xl/_rels/workbook.xml.rels',
				id: 'rIdPivotCacheAlpha',
				rawTarget: 'pivotCache/pivotCacheDefinition5.xml',
				resolvedTarget: 'xl/pivotCache/pivotCacheDefinition5.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
			},
			{
				sourcePartPath: 'xl/workbook.xml',
				relationshipPartPath: 'xl/_rels/workbook.xml.rels',
				id: 'rIdPivotCacheOmega',
				rawTarget: 'pivotCache/pivotCacheDefinition9.xml',
				resolvedTarget: 'xl/pivotCache/pivotCacheDefinition9.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
			},
			{
				sourcePartPath: 'xl/pivotCache/pivotCacheDefinition5.xml',
				relationshipPartPath: 'xl/pivotCache/_rels/pivotCacheDefinition5.xml.rels',
				id: 'rIdCacheRecordsAbsolute',
				rawTarget: '/xl/pivotCache/pivotCacheRecords5.xml',
				resolvedTarget: 'xl/pivotCache/pivotCacheRecords5.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
			},
			{
				sourcePartPath: 'xl/pivotCache/pivotCacheDefinition9.xml',
				relationshipPartPath: 'xl/pivotCache/_rels/pivotCacheDefinition9.xml.rels',
				id: 'rIdCacheRecordsRelative',
				rawTarget: 'pivotCacheRecords9.xml',
				resolvedTarget: 'xl/pivotCache/pivotCacheRecords9.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
			},
			{
				sourcePartPath: 'xl/worksheets/sheet2.xml',
				relationshipPartPath: 'xl/worksheets/_rels/sheet2.xml.rels',
				id: 'rIdPivotTable17',
				rawTarget: '../pivotTables/pivotTable1.xml',
				resolvedTarget: 'xl/pivotTables/pivotTable1.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
			},
			{
				sourcePartPath: 'xl/worksheets/sheet2.xml',
				relationshipPartPath: 'xl/worksheets/_rels/sheet2.xml.rels',
				id: 'rIdPivotTable42',
				rawTarget: '../pivotTables/pivotTable2.xml',
				resolvedTarget: 'xl/pivotTables/pivotTable2.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
			},
		])

		const opened = readXlsx(sourceBytes, { pivotCacheRecordMaterializeLimit: 'all' })
		expectOk(opened)
		expect(
			opened.value.workbook.pivotTables.map((pivot) => ({
				name: pivot.name,
				partPath: pivot.partPath,
				sheetName: pivot.sheetName,
				cacheId: pivot.cacheId,
			})),
		).toEqual([
			{
				name: 'PivotTableA',
				partPath: 'xl/pivotTables/pivotTable1.xml',
				sheetName: 'Pivots',
				cacheId: 7,
			},
			{
				name: 'PivotTableB',
				partPath: 'xl/pivotTables/pivotTable2.xml',
				sheetName: 'Pivots',
				cacheId: 7,
			},
		])
		expect(
			opened.value.workbook.pivotCaches.map((cache) => ({
				cacheId: cache.cacheId,
				relId: cache.relId,
				partPath: cache.partPath,
				recordsPartPath: cache.recordsPartPath,
				parsedCount: cache.records?.parsedCount,
			})),
		).toEqual([
			{
				cacheId: 7,
				relId: 'rIdPivotCacheAlpha',
				partPath: 'xl/pivotCache/pivotCacheDefinition5.xml',
				recordsPartPath: 'xl/pivotCache/pivotCacheRecords5.xml',
				parsedCount: 2,
			},
			{
				cacheId: 31,
				relId: 'rIdPivotCacheOmega',
				partPath: 'xl/pivotCache/pivotCacheDefinition9.xml',
				recordsPartPath: 'xl/pivotCache/pivotCacheRecords9.xml',
				parsedCount: 1,
			},
		])

		const dataSheet = opened.value.workbook.sheets.find((sheet) => sheet.name === 'Data')
		if (!dataSheet) throw new Error('Data sheet was not parsed')
		dataSheet.cells.set(3, 0, {
			value: stringValue('safe dirty edit outside pivot package'),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)

		const afterGraph = inspectXlsxPackageGraph(written.value)
		expect(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph)).toEqual([])
		expect(auditXlsxPackageGraphBytePreservation(beforeGraph, sourceBytes, written.value)).toEqual(
			[],
		)
		expect(pivotPartIdentities(afterGraph)).toEqual(pivotPartIdentities(beforeGraph))
		expect(pivotRelationshipIdentities(afterGraph)).toEqual(
			pivotRelationshipIdentities(beforeGraph),
		)

		const beforeZip = unzipSync(sourceBytes)
		const afterZip = unzipSync(written.value)
		for (const partPath of [
			'xl/pivotTables/pivotTable1.xml',
			'xl/pivotTables/pivotTable2.xml',
			'xl/pivotCache/pivotCacheDefinition5.xml',
			'xl/pivotCache/pivotCacheDefinition9.xml',
			'xl/pivotCache/_rels/pivotCacheDefinition5.xml.rels',
			'xl/pivotCache/_rels/pivotCacheDefinition9.xml.rels',
			'xl/pivotCache/pivotCacheRecords5.xml',
			'xl/pivotCache/pivotCacheRecords9.xml',
		]) {
			expect(decode(afterZip[partPath]), partPath).toBe(decode(beforeZip[partPath]))
		}

		const reopened = readXlsx(written.value, { pivotCacheRecordMaterializeLimit: 'all' })
		expectOk(reopened)
		expect(
			reopened.value.workbook.pivotTables.map((pivot) => ({
				name: pivot.name,
				partPath: pivot.partPath,
				sheetName: pivot.sheetName,
				cacheId: pivot.cacheId,
			})),
		).toEqual(
			opened.value.workbook.pivotTables.map((pivot) => ({
				name: pivot.name,
				partPath: pivot.partPath,
				sheetName: pivot.sheetName,
				cacheId: pivot.cacheId,
			})),
		)
		expect(
			reopened.value.workbook.pivotCaches.map((cache) => ({
				cacheId: cache.cacheId,
				relId: cache.relId,
				partPath: cache.partPath,
				recordsPartPath: cache.recordsPartPath,
				parsedCount: cache.records?.parsedCount,
			})),
		).toEqual(
			opened.value.workbook.pivotCaches.map((cache) => ({
				cacheId: cache.cacheId,
				relId: cache.relId,
				partPath: cache.partPath,
				recordsPartPath: cache.recordsPartPath,
				parsedCount: cache.records?.parsedCount,
			})),
		)
	})

	test('keeps workbook pivot cache ids bound to the intended cache packages', () => {
		const sourceBytes = pivotPackageWorkbook({ secondPivotTableCacheId: 31 })
		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
		expect(auditXlsxPackageGraphReadIntegrity(beforeGraph)).toEqual([])

		const opened = readXlsx(sourceBytes, { pivotCacheRecordMaterializeLimit: 'all' })
		expectOk(opened)
		expect(pivotTableCacheBindings(opened.value.workbook)).toEqual([
			{
				name: 'PivotTableA',
				cacheId: 7,
				cachePartPath: 'xl/pivotCache/pivotCacheDefinition5.xml',
				recordsPartPath: 'xl/pivotCache/pivotCacheRecords5.xml',
				parsedCount: 2,
			},
			{
				name: 'PivotTableB',
				cacheId: 31,
				cachePartPath: 'xl/pivotCache/pivotCacheDefinition9.xml',
				recordsPartPath: 'xl/pivotCache/pivotCacheRecords9.xml',
				parsedCount: 1,
			},
		])

		const dataSheet = opened.value.workbook.sheets.find((sheet) => sheet.name === 'Data')
		if (!dataSheet) throw new Error('Data sheet was not parsed')
		dataSheet.cells.set(3, 1, {
			value: stringValue('dirty edit with split pivot caches'),
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
		expect(pivotRelationshipIdentities(afterGraph)).toEqual(
			pivotRelationshipIdentities(beforeGraph),
		)

		const afterZip = unzipSync(written.value)
		const workbookXml = decode(afterZip['xl/workbook.xml'])
		const workbookRelsXml = decode(afterZip['xl/_rels/workbook.xml.rels'])
		expect(workbookXml).toContain('<pivotCache cacheId="7" r:id="rIdPivotCacheAlpha"/>')
		expect(workbookXml).toContain('<pivotCache cacheId="31" r:id="rIdPivotCacheOmega"/>')
		expect(workbookRelsXml).toContain(
			'Id="rIdPivotCacheAlpha" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition5.xml"',
		)
		expect(workbookRelsXml).toContain(
			'Id="rIdPivotCacheOmega" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition9.xml"',
		)
		expect(decode(afterZip['xl/pivotTables/pivotTable1.xml'])).toContain(
			'name="PivotTableA" cacheId="7"',
		)
		expect(decode(afterZip['xl/pivotTables/pivotTable2.xml'])).toContain(
			'name="PivotTableB" cacheId="31"',
		)
		expect(decode(afterZip['xl/pivotCache/pivotCacheDefinition5.xml'])).toContain(
			'r:id="rIdCacheRecordsAbsolute"',
		)
		expect(decode(afterZip['xl/pivotCache/pivotCacheDefinition9.xml'])).toContain(
			'r:id="rIdCacheRecordsRelative"',
		)
		expect(decode(afterZip['xl/pivotCache/_rels/pivotCacheDefinition5.xml.rels'])).toContain(
			'Id="rIdCacheRecordsAbsolute" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="/xl/pivotCache/pivotCacheRecords5.xml"',
		)
		expect(decode(afterZip['xl/pivotCache/_rels/pivotCacheDefinition9.xml.rels'])).toContain(
			'Id="rIdCacheRecordsRelative" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords9.xml"',
		)

		const reopened = readXlsx(written.value, { pivotCacheRecordMaterializeLimit: 'all' })
		expectOk(reopened)
		expect(pivotTableCacheBindings(reopened.value.workbook)).toEqual(
			pivotTableCacheBindings(opened.value.workbook),
		)
	})
})

function pivotTableCacheBindings(workbook: {
	readonly pivotTables: readonly {
		readonly name?: string
		readonly cacheId?: number
	}[]
	readonly pivotCaches: readonly {
		readonly cacheId?: number
		readonly partPath: string
		readonly recordsPartPath?: string
		readonly records?: { readonly parsedCount: number }
	}[]
}): readonly Record<string, unknown>[] {
	const cacheById = new Map(workbook.pivotCaches.map((cache) => [cache.cacheId, cache]))
	return workbook.pivotTables.map((pivot) => {
		const cache = cacheById.get(pivot.cacheId)
		return {
			name: pivot.name,
			cacheId: pivot.cacheId,
			cachePartPath: cache?.partPath,
			recordsPartPath: cache?.recordsPartPath,
			parsedCount: cache?.records?.parsedCount,
		}
	})
}

function pivotPartIdentities(graph: XlsxPackageGraph): readonly Record<string, unknown>[] {
	return graph.parts
		.filter((part) => part.featureFamily === 'preservedPivot')
		.map((part) => ({
			path: part.path,
			featureFamily: part.featureFamily,
			ownerScope: part.ownerScope,
			contentType: part.contentType,
		}))
		.sort((left, right) => String(left.path).localeCompare(String(right.path)))
}

function pivotRelationshipIdentities(graph: XlsxPackageGraph): readonly Record<string, unknown>[] {
	return graph.relationships
		.filter((relationship) => relationship.featureFamily === 'preservedPivot')
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

function pivotPackageWorkbook(
	options: { readonly secondPivotTableCacheId?: number } = {},
): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition9.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords9.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheetData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdSheetPivots" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rIdPivotCacheAlpha" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition5.xml"/>
  <Relationship Id="rIdPivotCacheOmega" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition9.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <pivotCaches>
    <pivotCache cacheId="7" r:id="rIdPivotCacheAlpha"/>
    <pivotCache cacheId="31" r:id="rIdPivotCacheOmega"/>
  </pivotCaches>
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rIdSheetData"/>
    <sheet name="Pivots" sheetId="2" r:id="rIdSheetPivots"/>
  </sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>West</t></is></c><c r="B2"><v>10</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>East</t></is></c><c r="B3"><v>20</v></c></row>
  </sheetData>
</worksheet>`,
		'xl/worksheets/sheet2.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`,
		'xl/worksheets/_rels/sheet2.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotTable17" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
  <Relationship Id="rIdPivotTable42" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable2.xml"/>
</Relationships>`,
		'xl/pivotTables/pivotTable1.xml': pivotTableXml('PivotTableA', 7, 'A3:C8'),
		'xl/pivotTables/pivotTable2.xml': pivotTableXml(
			'PivotTableB',
			options.secondPivotTableCacheId ?? 7,
			'E3:G8',
		),
		'xl/pivotCache/pivotCacheDefinition5.xml': pivotCacheDefinitionXml(
			'rIdCacheRecordsAbsolute',
			2,
			'Data',
		),
		'xl/pivotCache/_rels/pivotCacheDefinition5.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdCacheRecordsAbsolute" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="/xl/pivotCache/pivotCacheRecords5.xml"/>
</Relationships>`,
		'xl/pivotCache/pivotCacheRecords5.xml': `<?xml version="1.0"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
  <r><x v="0"/><n v="10"/></r>
  <r><x v="1"/><n v="20"/></r>
</pivotCacheRecords>`,
		'xl/pivotCache/pivotCacheDefinition9.xml': pivotCacheDefinitionXml(
			'rIdCacheRecordsRelative',
			1,
			'Data',
		),
		'xl/pivotCache/_rels/pivotCacheDefinition9.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdCacheRecordsRelative" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords9.xml"/>
</Relationships>`,
		'xl/pivotCache/pivotCacheRecords9.xml': `<?xml version="1.0"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1">
  <r><x v="0"/><n v="30"/></r>
</pivotCacheRecords>`,
	})
}

function pivotTableXml(name: string, cacheId: number, ref: string): string {
	return `<?xml version="1.0"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${name}" cacheId="${cacheId}">
  <location ref="${ref}" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="2">
    <pivotField axis="axisRow" showAll="0"><items count="2"><item x="0"/><item x="1"/></items></pivotField>
    <pivotField dataField="1"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <dataFields count="1"><dataField fld="1" name="Sum of Amount" subtotal="sum"/></dataFields>
</pivotTableDefinition>`
}

function pivotCacheDefinitionXml(
	recordsRelId: string,
	recordCount: number,
	sourceSheet: string,
): string {
	return `<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  r:id="${recordsRelId}" recordCount="${recordCount}" refreshOnLoad="0" enableRefresh="1" saveData="1">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="${sourceSheet}"/></cacheSource>
  <cacheFields count="2">
    <cacheField name="Region" databaseField="1"><sharedItems count="2"><s v="West"/><s v="East"/></sharedItems></cacheField>
    <cacheField name="Amount" databaseField="1" numFmtId="0"><sharedItems containsNumber="1" count="2"><n v="10"/><n v="20"/></sharedItems></cacheField>
  </cacheFields>
</pivotCacheDefinition>`
}
