import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { unzipSync } from 'fflate'
import type { StyleId } from '../../packages/core/src/index.ts'
import { applyOperation } from '../../packages/engine/src/index.ts'
import {
	auditXlsxPackageGraphSafeEditIntegrity,
	inspectXlsxPackageGraph,
	readXlsx,
	writeXlsx,
	type XlsxPackageGraph,
	type XlsxPackageGraphFidelityIssue,
} from '../../packages/io-xlsx/src/index.ts'
import { stringValue } from '../../packages/schema/src/index.ts'

const S0 = 0 as StyleId
const TABLE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
const QUERY_TABLE_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml'

interface TableFixtureCase {
	readonly label: string
	readonly path: URL
	readonly expectedTables: number
}

const tableFixtures: readonly TableFixtureCase[] = [
	{
		label: 'Calamine absolute table relationship target',
		path: new URL('./calamine/table_with_absolute_paths.xlsx', import.meta.url),
		expectedTables: 1,
	},
	{
		label: 'ClosedXML workbook tables',
		path: new URL('./closedxml/Tables_UsingTables.xlsx', import.meta.url),
		expectedTables: 3,
	},
	{
		label: 'LibreOffice tableType table',
		path: new URL('./libreoffice/tdf167689_tableType.xlsx', import.meta.url),
		expectedTables: 1,
	},
	{
		label: 'POI structured references table',
		path: new URL('./poi/StructuredReferences.xlsx', import.meta.url),
		expectedTables: 1,
	},
]

describe('vendored table package graph fidelity', () => {
	for (const fixture of tableFixtures) {
		test(`${fixture.label} preserves table parts and relationships after safe edit`, () => {
			const sourceBytes = readFileSync(fixture.path)
			const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
			const beforeTableParts = tablePartIdentities(beforeGraph)
			const beforeTableRelationships = tableRelationshipIdentities(beforeGraph)
			const beforeTableOverrides = tableContentTypeOverrides(beforeGraph)
			expect(beforeTableParts).toHaveLength(fixture.expectedTables)
			expect(beforeTableOverrides).toHaveLength(fixture.expectedTables)

			const read = readXlsx(sourceBytes)
			expectOk(read)
			const tableSheet = read.value.workbook.sheets.find((sheet) => sheet.tables.length > 0)
			expect(tableSheet).toBeDefined()
			if (!tableSheet) return
			expect(read.value.workbook.sheets.reduce((sum, sheet) => sum + sheet.tables.length, 0)).toBe(
				fixture.expectedTables,
			)
			const tableNames = tableSheet.tables.map((table) => table.name)

			tableSheet.cells.set(199, 30, {
				value: stringValue('__ascend_table_package_safe_edit__'),
				formula: null,
				styleId: S0,
			})
			const written = writeXlsx(read.value.workbook, read.value.capsules, {
				dirtySheetNames: [tableSheet.name],
			})
			expectOk(written)

			const reopened = readXlsx(written.value)
			expectOk(reopened)
			expect(
				reopened.value.workbook.sheets.reduce((sum, sheet) => sum + sheet.tables.length, 0),
			).toBe(fixture.expectedTables)
			expect(
				reopened.value.workbook.sheets.flatMap((sheet) => sheet.tables.map((table) => table.name)),
			).toEqual(expect.arrayContaining(tableNames))

			const afterGraph = inspectXlsxPackageGraph(written.value)
			expectNoPackageGraphIssues(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph))
			expect(tablePartIdentities(afterGraph)).toEqual(beforeTableParts)
			expect(tableRelationshipIdentities(afterGraph)).toEqual(beforeTableRelationships)
			expect(tableContentTypeOverrides(afterGraph)).toEqual(beforeTableOverrides)
			expectTableBytesPreserved(beforeGraph, sourceBytes, written.value)
		})
	}

	test('LibreOffice query table sidecar preserves table-owned relationship and bytes', () => {
		const sourceBytes = readFileSync(
			new URL('./libreoffice/TableEmptyHeaders.xlsx', import.meta.url),
		)
		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
		const beforeQueryTableParts = queryTablePartIdentities(beforeGraph)
		const beforeTableRelationships = tableRelationshipIdentities(beforeGraph)
		const beforeQueryTableOverrides = queryTableContentTypeOverrides(beforeGraph)
		expect(beforeQueryTableParts).toEqual([
			expect.objectContaining({
				path: 'xl/queryTables/queryTable1.xml',
				contentType: QUERY_TABLE_CONTENT_TYPE,
				ownerScope: 'worksheet',
				sourceRelationshipPart: 'xl/tables/_rels/table1.xml.rels',
				sourceRelationshipId: 'rId1',
				sourceRelationshipRawTarget: '../queryTables/queryTable1.xml',
				preservationPolicy: 'preserve-exact',
			}),
		])
		expect(beforeQueryTableOverrides).toEqual([
			{
				partPath: 'xl/queryTables/queryTable1.xml',
				contentType: QUERY_TABLE_CONTENT_TYPE,
			},
		])

		const read = readXlsx(sourceBytes)
		expectOk(read)
		const sheet = read.value.workbook.sheets.find((entry) => entry.name === 'BTC')
		const table = sheet?.tables.find((entry) => entry.name === 'Bitcoin')
		expect(table?.queryTable).toMatchObject({
			relationshipId: 'rId1',
			partPath: 'xl/queryTables/queryTable1.xml',
			target: '../queryTables/queryTable1.xml',
		})
		if (!sheet) return

		sheet.cells.set(199, 30, {
			value: stringValue('__ascend_query_table_package_safe_edit__'),
			formula: null,
			styleId: S0,
		})
		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: [sheet.name],
		})
		expectOk(written)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedTable = reopened.value.workbook.sheets
			.find((entry) => entry.name === 'BTC')
			?.tables.find((entry) => entry.name === 'Bitcoin')
		expect(reopenedTable?.queryTable).toMatchObject({
			relationshipId: 'rId1',
			partPath: 'xl/queryTables/queryTable1.xml',
			target: '../queryTables/queryTable1.xml',
		})
		expect(reopenedTable?.columns.map((column) => column.queryTableFieldId)).toEqual([1, 2])

		const afterGraph = inspectXlsxPackageGraph(written.value)
		expectNoPackageGraphIssues(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph))
		expect(queryTablePartIdentities(afterGraph)).toEqual(beforeQueryTableParts)
		expect(tableRelationshipIdentities(afterGraph)).toEqual(beforeTableRelationships)
		expect(queryTableContentTypeOverrides(afterGraph)).toEqual(beforeQueryTableOverrides)
		expectFeatureBytesPreserved(beforeGraph, sourceBytes, written.value, 'preservedQueryTable')
	})

	test('LibreOffice query table sidecar keeps package identity through table resize', () => {
		const sourceBytes = readFileSync(
			new URL('./libreoffice/TableEmptyHeaders.xlsx', import.meta.url),
		)
		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
		const beforeTableParts = tablePartIdentities(beforeGraph)
		const beforeQueryTableParts = queryTablePartIdentities(beforeGraph)
		const beforeTableRelationships = tableRelationshipIdentities(beforeGraph)
		const beforeTableOverrides = tableContentTypeOverrides(beforeGraph)
		const beforeQueryTableOverrides = queryTableContentTypeOverrides(beforeGraph)

		const read = readXlsx(sourceBytes)
		expectOk(read)
		const resized = applyOperation(read.value.workbook, {
			op: 'resizeTable',
			table: 'Bitcoin',
			ref: 'A1:B17',
		})
		expectOk(resized)

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['BTC'],
		})
		expectOk(written)
		const writtenTableXml = decodeZipPart(written.value, 'xl/tables/table1.xml')
		expect(writtenTableXml).toContain('ref="A1:B17"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedTable = reopened.value.workbook.sheets
			.find((entry) => entry.name === 'BTC')
			?.tables.find((entry) => entry.name === 'Bitcoin')
		expect(reopenedTable).toMatchObject({
			ref: { start: { row: 0, col: 0 }, end: { row: 16, col: 1 } },
			autoFilter: { ref: 'A1:B17' },
			tableStyleInfo: { name: 'TableStyleMedium7' },
			queryTable: {
				relationshipId: 'rId1',
				partPath: 'xl/queryTables/queryTable1.xml',
				target: '../queryTables/queryTable1.xml',
			},
		})
		expect(reopenedTable?.columns.map((column) => column.queryTableFieldId)).toEqual([1, 2])

		const afterGraph = inspectXlsxPackageGraph(written.value)
		expectNoPackageGraphIssues(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph))
		expect(tablePartIdentities(afterGraph)).toEqual(beforeTableParts)
		expect(queryTablePartIdentities(afterGraph)).toEqual(beforeQueryTableParts)
		expect(tableRelationshipIdentities(afterGraph)).toEqual(beforeTableRelationships)
		expect(tableContentTypeOverrides(afterGraph)).toEqual(beforeTableOverrides)
		expect(queryTableContentTypeOverrides(afterGraph)).toEqual(beforeQueryTableOverrides)
		expectFeatureBytesPreserved(beforeGraph, sourceBytes, written.value, 'preservedQueryTable')
	})
})

function tablePartIdentities(graph: XlsxPackageGraph): readonly Record<string, unknown>[] {
	return graph.parts
		.filter((part) => part.featureFamily === 'preservedTable')
		.map((part) => ({
			path: part.path,
			contentType: part.contentType,
			contentTypeSource: part.contentTypeSource,
			ownerScope: part.ownerScope,
			sourceRelationshipPart: part.sourceRelationshipPart,
			sourceRelationshipId: part.sourceRelationshipId,
			sourceRelationshipType: part.sourceRelationshipType,
			sourceRelationshipRawType: part.sourceRelationshipRawType,
			sourceRelationshipRawTarget: part.sourceRelationshipRawTarget,
			sourceRelationshipResolvedTarget: part.sourceRelationshipResolvedTarget,
			preservationPolicy: part.preservationPolicy,
			bytePreservationExpected: part.bytePreservationExpected,
		}))
		.sort(compareJson)
}

function tableRelationshipIdentities(graph: XlsxPackageGraph): readonly Record<string, unknown>[] {
	return graph.relationships
		.filter(
			(relationship) =>
				relationship.featureFamily === 'preservedTable' ||
				relationship.featureFamily === 'preservedQueryTable' ||
				relationship.sourcePartPath.startsWith('xl/tables/'),
		)
		.map((relationship) => ({
			sourcePartPath: relationship.sourcePartPath,
			relationshipPartPath: relationship.relationshipPartPath,
			id: relationship.id,
			type: relationship.type,
			rawType: relationship.rawType,
			rawTarget: relationship.rawTarget,
			resolvedTarget: relationship.resolvedTarget,
			targetMode: relationship.targetMode,
			featureFamily: relationship.featureFamily,
		}))
		.sort(compareJson)
}

function tableContentTypeOverrides(graph: XlsxPackageGraph): readonly Record<string, unknown>[] {
	const tablePaths = new Set(
		graph.parts.filter((part) => part.featureFamily === 'preservedTable').map((part) => part.path),
	)
	return graph.contentTypeOverrides
		.filter((override) => tablePaths.has(override.partPath))
		.map((override) => ({
			partPath: override.partPath,
			contentType: override.contentType,
		}))
		.sort(compareJson)
}

function queryTablePartIdentities(graph: XlsxPackageGraph): readonly Record<string, unknown>[] {
	return graph.parts
		.filter((part) => part.featureFamily === 'preservedQueryTable')
		.map((part) => ({
			path: part.path,
			contentType: part.contentType,
			contentTypeSource: part.contentTypeSource,
			ownerScope: part.ownerScope,
			sourceRelationshipPart: part.sourceRelationshipPart,
			sourceRelationshipId: part.sourceRelationshipId,
			sourceRelationshipType: part.sourceRelationshipType,
			sourceRelationshipRawType: part.sourceRelationshipRawType,
			sourceRelationshipRawTarget: part.sourceRelationshipRawTarget,
			sourceRelationshipResolvedTarget: part.sourceRelationshipResolvedTarget,
			preservationPolicy: part.preservationPolicy,
			bytePreservationExpected: part.bytePreservationExpected,
		}))
		.sort(compareJson)
}

function queryTableContentTypeOverrides(
	graph: XlsxPackageGraph,
): readonly Record<string, unknown>[] {
	const queryTablePaths = new Set(
		graph.parts
			.filter((part) => part.featureFamily === 'preservedQueryTable')
			.map((part) => part.path),
	)
	return graph.contentTypeOverrides
		.filter((override) => queryTablePaths.has(override.partPath))
		.map((override) => ({
			partPath: override.partPath,
			contentType: override.contentType,
		}))
		.sort(compareJson)
}

function expectTableBytesPreserved(
	graph: XlsxPackageGraph,
	sourceBytes: Uint8Array,
	writtenBytes: Uint8Array,
): void {
	expectFeatureBytesPreserved(graph, sourceBytes, writtenBytes, 'preservedTable')
	for (const part of graph.parts.filter((entry) => entry.featureFamily === 'preservedTable')) {
		expect(part.contentType).toBe(TABLE_CONTENT_TYPE)
	}
}

function expectFeatureBytesPreserved(
	graph: XlsxPackageGraph,
	sourceBytes: Uint8Array,
	writtenBytes: Uint8Array,
	featureFamily: string,
): void {
	for (const part of graph.parts.filter((entry) => entry.featureFamily === featureFamily)) {
		const before = readZipPart(sourceBytes, part.path)
		const after = readZipPart(writtenBytes, part.path)
		expect(after).toEqual(before)
		expect(part.preservationPolicy).toBe('preserve-exact')
		expect(part.bytePreservationExpected).toBe(true)
	}
}

function readZipPart(bytes: Uint8Array, path: string): Uint8Array {
	const entry = unzipSync(bytes)[path]
	expect(entry).toBeDefined()
	if (!entry) throw new Error(`Missing ZIP part ${path}`)
	return entry
}

function decodeZipPart(bytes: Uint8Array, path: string): string {
	return new TextDecoder().decode(readZipPart(bytes, path))
}

function expectNoPackageGraphIssues(issues: readonly XlsxPackageGraphFidelityIssue[]): void {
	if (issues.length === 0) return
	throw new Error(issues.map(formatPackageGraphIssue).join('\n'))
}

function formatPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): string {
	return [
		issue.message,
		`severity=${issue.severity}`,
		issue.partPath ? `part=${issue.partPath}` : undefined,
		issue.sourcePartPath !== undefined
			? `source=${issue.sourcePartPath || '<package>'}`
			: undefined,
		issue.relationshipPartPath ? `rels=${issue.relationshipPartPath}` : undefined,
		issue.relationshipId ? `relId=${issue.relationshipId}` : undefined,
		issue.ownerScope ? `owner=${issue.ownerScope}` : undefined,
		issue.featureFamily ? `family=${issue.featureFamily}` : undefined,
		issue.suggestedAction ? `action=${issue.suggestedAction}` : undefined,
		issue.expected !== undefined ? `expected=${JSON.stringify(issue.expected)}` : undefined,
		issue.actual !== undefined ? `actual=${JSON.stringify(issue.actual)}` : undefined,
	]
		.filter((part): part is string => part !== undefined)
		.join(' | ')
}

function compareJson(left: Record<string, unknown>, right: Record<string, unknown>): number {
	return JSON.stringify(left).localeCompare(JSON.stringify(right))
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}
