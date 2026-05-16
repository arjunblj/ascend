import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
	inspectXlsxPackageGraph,
} from '@ascend/io-xlsx'
import {
	AscendWorkbook,
	type ExternalReferenceInfo,
	type WorkbookConnectionPartInfo,
	type WorkbookRefreshMetadataEntry,
} from '@ascend/sdk'

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

function externalReferenceContract(entry: ExternalReferenceInfo): Record<string, unknown> {
	return {
		partPath: entry.partPath,
		relId: entry.relId,
		externalBookRelId: entry.externalBookRelId,
		linkRelId: entry.linkRelId,
		linkRelationshipKind: entry.linkRelationshipKind,
		linkBindingStatus: entry.linkBindingStatus,
		target: entry.target,
		targetMode: entry.targetMode,
		sourceRelationshipPart: entry.sourceRelationshipPart,
		sourceRelationshipRawTarget: entry.sourceRelationshipRawTarget,
		linkRelationshipPart: entry.linkRelationshipPart,
		linkRelationshipRawTarget: entry.linkRelationshipRawTarget,
	}
}

function connectionContract(entry: WorkbookConnectionPartInfo): Record<string, unknown> {
	return {
		kind: entry.kind,
		partPath: entry.partPath,
		sheetName: entry.sheetName,
		name: entry.name,
		connectionId: entry.connectionId,
		refreshOnLoad: entry.refreshOnLoad,
		saveData: entry.saveData,
		refreshedVersion: entry.refreshedVersion,
	}
}

function refreshContract(entry: WorkbookRefreshMetadataEntry): Record<string, unknown> {
	return {
		kind: entry.kind,
		partPath: entry.partPath,
		state: entry.state,
		name: entry.name,
		sheetName: entry.sheetName,
		connectionId: entry.connectionId,
		refreshOnLoad: entry.refreshOnLoad,
		saveData: entry.saveData,
		refreshedVersion: entry.refreshedVersion,
		recommendedOps: entry.recommendedOps,
		warnings: entry.warnings,
	}
}

async function safeEdit(bytes: Uint8Array, sheet: string, ref: string): Promise<Uint8Array> {
	const workbook = await AscendWorkbook.open(bytes)
	const result = workbook.apply([{ op: 'setCells', sheet, updates: [{ ref, value: 'safe-edit' }] }])
	expect(result.errors).toEqual([])
	return workbook.toBytes()
}

function sortedContracts<T>(
	entries: readonly T[],
	contract: (entry: T) => Record<string, unknown>,
): readonly Record<string, unknown>[] {
	return entries
		.map(contract)
		.sort((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right)))
}

describe('external refresh corpus contract', () => {
	test('preserves public external workbook link identities through safe edits and reopen', async () => {
		const cases = [
			{
				path: '../xlsx/closedxml/Other_ExternalLinks_WorkbookWithExternalLink.xlsx',
				sheet: 'Sheet1',
				editRef: 'C10',
				expected: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					relId: 'rId2',
					externalBookRelId: 'rId1',
					linkRelId: 'rId1',
					linkRelationshipKind: 'externalLinkPath',
					linkBindingStatus: 'externalBookRelId',
					target: 'book1.xlsx',
					targetMode: 'External',
				},
			},
			{
				path: '../xlsx/libreoffice/MissingPathExternal.xlsx',
				sheet: '2er-Schichtplanung',
				editRef: 'Z10',
				expected: {
					partPath: 'xl/externalLinks/externalLink1.xml',
					relId: 'rId4',
					externalBookRelId: 'rId1',
					linkRelId: 'rId1',
					linkRelationshipKind: 'xlPathMissing',
					linkBindingStatus: 'externalBookRelId',
					target: 'Tabelle1',
					targetMode: 'External',
				},
			},
		] as const

		for (const entry of cases) {
			const source = loadFixture(entry.path)
			const beforeWorkbook = await AscendWorkbook.open(source)
			const before = beforeWorkbook
				.inspect()
				.externalReferenceDetails.map(externalReferenceContract)
			expect(before).toEqual([expect.objectContaining(entry.expected)])
			expect(
				beforeWorkbook
					.inspect()
					.compatibility.features.find((feature) => feature.feature === 'preservedExternalLink'),
			).toMatchObject({
				tier: 'preserved',
				count: 1,
				locations: ['xl/externalLinks/externalLink1.xml'],
			})

			const edited = await safeEdit(source, entry.sheet, entry.editRef)
			const reopened = await AscendWorkbook.open(edited)
			expect(reopened.inspect().externalReferenceDetails.map(externalReferenceContract)).toEqual(
				before,
			)
			expect(
				auditXlsxPackageGraphSafeEditIntegrity(
					inspectXlsxPackageGraph(source),
					inspectXlsxPackageGraph(edited),
				),
			).toEqual([])
			expect(
				auditXlsxPackageGraphBytePreservation(inspectXlsxPackageGraph(source), source, edited),
			).toEqual([])
		}
	})

	test('rewrites a public external link target only through explicit metadata operation', async () => {
		const source = loadFixture(
			'../xlsx/closedxml/Other_ExternalLinks_WorkbookWithExternalLink.xlsx',
		)
		const workbook = await AscendWorkbook.open(source)
		expect(workbook.inspect().externalReferenceDetails.map(externalReferenceContract)).toEqual([
			expect.objectContaining({
				partPath: 'xl/externalLinks/externalLink1.xml',
				linkRelId: 'rId1',
				linkBindingStatus: 'externalBookRelId',
				target: 'book1.xlsx',
			}),
		])

		const changed = workbook.apply(
			[
				{
					op: 'rewriteExternalLink',
					partPath: 'xl/externalLinks/externalLink1.xml',
					linkRelId: 'rId1',
					newTarget: 'linked/book2.xlsx',
				},
			],
			{ journal: true },
		)
		expect(changed.errors).toEqual([])
		expect(changed.journal).toMatchObject({
			supported: false,
			exact: false,
			issues: [
				expect.objectContaining({
					code: 'UNSUPPORTED_OPERATION',
					surface: 'package-parts',
				}),
			],
		})

		const edited = workbook.toBytes()
		expect(auditXlsxPackageGraphReadIntegrity(inspectXlsxPackageGraph(edited))).toEqual([])
		const reopened = await AscendWorkbook.open(edited)
		expect(reopened.inspect().externalReferenceDetails.map(externalReferenceContract)).toEqual([
			expect.objectContaining({
				partPath: 'xl/externalLinks/externalLink1.xml',
				linkRelId: 'rId1',
				linkBindingStatus: 'externalBookRelId',
				linkRelationshipRawTarget: 'linked/book2.xlsx',
				target: 'linked/book2.xlsx',
				targetMode: 'External',
			}),
		])
	})

	test('reports and preserves public query-table refresh surfaces without executing them', async () => {
		const source = loadFixture('../xlsx/libreoffice/queryTableExport.xlsx')
		const workbook = await AscendWorkbook.open(source)
		const beforeConnections = sortedContracts(workbook.connectionParts(), connectionContract)
		expect(beforeConnections).toEqual([
			{
				kind: 'connection',
				partPath: 'xl/connections.xml',
				name: 'conn_with_comma',
				connectionId: 2,
				saveData: true,
				refreshedVersion: 3,
			},
			{
				kind: 'connection',
				partPath: 'xl/connections.xml',
				name: 'conn_with_delim',
				connectionId: 1,
				saveData: true,
				refreshedVersion: 3,
			},
			{
				kind: 'queryTable',
				partPath: 'xl/queryTables/queryTable1.xml',
				sheetName: 'Sheet1',
				name: 'conn_with_delim',
				connectionId: 1,
			},
			{
				kind: 'queryTable',
				partPath: 'xl/queryTables/queryTable2.xml',
				sheetName: 'Sheet1',
				name: 'conn_with_comma',
				connectionId: 2,
			},
		])
		expect(workbook.refreshMetadata()).toMatchObject({
			refreshOnOpenCount: 0,
			notSavedCount: 0,
			unknownCount: 2,
		})
		expect(workbook.trustReport().findings).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					code: 'workbook.dataConnection',
					location: expect.objectContaining({
						partPath: 'xl/connections.xml',
						kind: 'connection',
						name: 'conn_with_delim',
						connectionId: 1,
					}),
				}),
				expect.objectContaining({
					code: 'workbook.dataConnection',
					location: expect.objectContaining({
						partPath: 'xl/queryTables/queryTable1.xml',
						kind: 'queryTable',
						sheet: 'Sheet1',
						name: 'conn_with_delim',
						connectionId: 1,
					}),
				}),
			]),
		)
		expect(sortedContracts(workbook.refreshMetadata().entries, refreshContract)).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					kind: 'queryTable',
					partPath: 'xl/queryTables/queryTable1.xml',
					state: 'unknown',
					connectionId: 1,
					recommendedOps: [
						expect.objectContaining({
							op: 'setConnectionRefresh',
							partPath: 'xl/queryTables/queryTable1.xml',
							refreshOnLoad: true,
							saveData: false,
						}),
					],
				}),
				expect.objectContaining({
					kind: 'workbookConnection',
					partPath: 'xl/connections.xml',
					state: 'cached',
					connectionId: 1,
					warnings: [
						'Workbook connection refresh metadata is inspectable and editable without executing the connection.',
					],
				}),
			]),
		)

		const edited = await safeEdit(source, 'Sheet1', 'Z99')
		const reopened = await AscendWorkbook.open(edited)
		expect(sortedContracts(reopened.connectionParts(), connectionContract)).toEqual(
			beforeConnections,
		)
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(edited),
				{ ignoredRelationshipFamilies: ['preservedDocumentProperties'] },
			),
		).toEqual([])
	})

	test('edits public table-owned query refresh metadata and reopens as cached', async () => {
		const source = loadFixture('../xlsx/libreoffice/TableEmptyHeaders.xlsx')
		const workbook = await AscendWorkbook.open(source)
		expect(workbook.refreshMetadata()).toMatchObject({
			refreshOnOpenCount: 5,
			notSavedCount: 1,
			unknownCount: 0,
		})
		expect(workbook.connectionParts()).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					kind: 'queryTable',
					partPath: 'xl/queryTables/queryTable1.xml',
					name: 'ExternalData_1',
					connectionId: 2,
					refreshOnLoad: true,
					saveData: false,
				}),
			]),
		)

		const changed = workbook.apply(
			[
				{
					op: 'setConnectionRefresh',
					partPath: 'xl/queryTables/queryTable1.xml',
					connectionId: 2,
					refreshOnLoad: false,
					saveData: true,
					refreshedVersion: 6,
				},
			],
			{ journal: true },
		)
		expect(changed.errors).toEqual([])
		expect(changed.journal).toMatchObject({
			supported: false,
			exact: false,
			issues: [
				expect.objectContaining({
					code: 'UNSUPPORTED_OPERATION',
					surface: 'package-parts',
				}),
			],
		})

		const edited = workbook.toBytes()
		expect(auditXlsxPackageGraphReadIntegrity(inspectXlsxPackageGraph(edited))).toEqual([])
		const reopened = await AscendWorkbook.open(edited)
		expect(reopened.connectionParts()).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					kind: 'queryTable',
					partPath: 'xl/queryTables/queryTable1.xml',
					name: 'ExternalData_1',
					connectionId: 2,
					refreshOnLoad: false,
					saveData: true,
					refreshedVersion: 6,
				}),
			]),
		)
		expect(reopened.refreshMetadata().entries).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					kind: 'queryTable',
					partPath: 'xl/queryTables/queryTable1.xml',
					state: 'cached',
					name: 'ExternalData_1',
					connectionId: 2,
					refreshOnLoad: false,
					saveData: true,
					refreshedVersion: 6,
					warnings: [
						'Query table refresh metadata is inspectable and editable without executing the query.',
					],
					recommendedOps: [],
				}),
			]),
		)
	})
})
