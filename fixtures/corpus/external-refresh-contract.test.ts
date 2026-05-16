import { afterEach, describe, expect, test } from 'bun:test'
import { existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
	inspectXlsxPackageGraph,
} from '@ascend/io-xlsx'
import {
	AscendWorkbook,
	commitAgentPlan,
	createAgentPlan,
	type ExternalReferenceInfo,
	type WorkbookConnectionPartInfo,
	type WorkbookRefreshMetadataEntry,
} from '@ascend/sdk'

const TEMP_DIR = join(tmpdir(), `ascend-external-refresh-contract-${process.pid}`)

afterEach(() => {
	if (existsSync(TEMP_DIR)) rmSync(TEMP_DIR, { recursive: true, force: true })
})

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

function definedContract(fields: Record<string, unknown>): Record<string, unknown> {
	return Object.fromEntries(Object.entries(fields).filter(([, value]) => value !== undefined))
}

function externalReferenceContract(entry: ExternalReferenceInfo): Record<string, unknown> {
	return definedContract({
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
	})
}

type ConnectionContractInfo = Pick<
	WorkbookConnectionPartInfo,
	| 'kind'
	| 'partPath'
	| 'sheetName'
	| 'name'
	| 'connectionId'
	| 'connectionType'
	| 'description'
	| 'deleted'
	| 'backgroundRefresh'
	| 'keepAlive'
	| 'refreshInterval'
	| 'refreshOnLoad'
	| 'saveData'
	| 'savePassword'
	| 'refreshedVersion'
	| 'sourceFile'
	| 'odcFile'
	| 'onlyUseConnectionFile'
	| 'command'
	| 'hasConnectionString'
>

function connectionContract(entry: ConnectionContractInfo): Record<string, unknown> {
	return definedContract({
		kind: entry.kind,
		partPath: entry.partPath,
		sheetName: entry.sheetName,
		name: entry.name,
		connectionId: entry.connectionId,
		connectionType: entry.connectionType,
		description: entry.description,
		deleted: entry.deleted,
		backgroundRefresh: entry.backgroundRefresh,
		keepAlive: entry.keepAlive,
		refreshInterval: entry.refreshInterval,
		refreshOnLoad: entry.refreshOnLoad,
		saveData: entry.saveData,
		savePassword: entry.savePassword,
		refreshedVersion: entry.refreshedVersion,
		sourceFile: entry.sourceFile,
		odcFile: entry.odcFile,
		onlyUseConnectionFile: entry.onlyUseConnectionFile,
		command: entry.command,
		hasConnectionString: entry.hasConnectionString,
	})
}

function refreshContract(entry: WorkbookRefreshMetadataEntry): Record<string, unknown> {
	return definedContract({
		kind: entry.kind,
		partPath: entry.partPath,
		state: entry.state,
		name: entry.name,
		sheetName: entry.sheetName,
		connectionId: entry.connectionId,
		connectionType: entry.connectionType,
		deleted: entry.deleted,
		backgroundRefresh: entry.backgroundRefresh,
		keepAlive: entry.keepAlive,
		refreshInterval: entry.refreshInterval,
		refreshOnLoad: entry.refreshOnLoad,
		saveData: entry.saveData,
		refreshedVersion: entry.refreshedVersion,
		sourceFile: entry.sourceFile,
		command: entry.command,
		hasConnectionString: entry.hasConnectionString,
		recommendedOps: entry.recommendedOps,
		warnings: entry.warnings,
	})
}

function tableContracts(workbook: AscendWorkbook): readonly Record<string, unknown>[] {
	return workbook.getWorkbookModel().sheets.flatMap((sheet) =>
		sheet.tables.map((table) => ({
			sheetName: sheet.name,
			name: table.name,
			ref: table.ref,
			tableType: table.tableType,
			queryTablePartPath: table.queryTable?.partPath,
		})),
	)
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
				connectionType: 6,
				deleted: true,
				backgroundRefresh: true,
				saveData: true,
				refreshedVersion: 3,
				sourceFile: 'C:\\data\\file2.csv',
			},
			{
				kind: 'connection',
				partPath: 'xl/connections.xml',
				name: 'conn_with_delim',
				connectionId: 1,
				connectionType: 6,
				deleted: true,
				backgroundRefresh: true,
				saveData: true,
				refreshedVersion: 3,
				sourceFile: 'C:\\data\\file1.csv',
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
					connectionType: 6,
					deleted: true,
					backgroundRefresh: true,
					sourceFile: 'C:\\data\\file1.csv',
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

	test('commit proof reports reopened public query-table connection metadata', async () => {
		const input = join(TEMP_DIR, 'libreoffice-table-owned-query.xlsx')
		const output = join(TEMP_DIR, 'libreoffice-table-owned-query-out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const source = loadFixture('../xlsx/libreoffice/TableEmptyHeaders.xlsx')
		await Bun.write(input, source)
		const ops = [
			{
				op: 'setConnectionRefresh' as const,
				partPath: 'xl/queryTables/queryTable1.xml',
				connectionId: 2,
				refreshOnLoad: false,
				saveData: true,
				refreshedVersion: 6,
			},
		]

		const plan = await createAgentPlan(input, ops)
		expect(
			plan.writePolicy.diagnostics.some((diagnostic) => diagnostic.severity === 'blocker'),
		).toBe(false)
		const committed = await commitAgentPlan(input, ops, {
			output,
			approvals: plan.approvals.map((approval) => approval.id),
		})

		expect(committed.postWrite.auditsPassed).toBe(true)
		expect(committed.postWrite.dataConnections).toMatchObject({
			total: 14,
			workbookConnections: 13,
			queryTables: 1,
			refreshOnOpen: 4,
			notSaved: 0,
			unknown: 0,
			cached: 10,
			preservationMode: 'preserve-exact',
			verification: 'reopened-output',
		})
		expect(committed.postWrite.dataConnections.partPaths).toEqual(
			expect.arrayContaining(['xl/connections.xml', 'xl/queryTables/queryTable1.xml']),
		)
		expect(committed.postWrite.dataConnections.names).toEqual(
			expect.arrayContaining(['Query - Bitcoin', 'ExternalData_1']),
		)
		expect(committed.postWrite.dataConnections.connectionIds).toEqual(
			expect.arrayContaining([1, 2, 13]),
		)
		expect(committed.postWrite.dataConnections.connections).toContainEqual(
			expect.objectContaining({
				kind: 'connection',
				partPath: 'xl/connections.xml',
				name: 'Query - Bitcoin',
				connectionId: 2,
				connectionType: 5,
				description: "Connection to the 'Bitcoin' query in the workbook.",
				backgroundRefresh: true,
				keepAlive: true,
				refreshInterval: 1,
				refreshOnLoad: true,
				command: 'SELECT * FROM [Bitcoin]',
				hasConnectionString: true,
				state: 'refresh-on-open',
			}),
		)
		expect(committed.postWrite.dataConnections.connections).toContainEqual(
			expect.objectContaining({
				kind: 'queryTable',
				partPath: 'xl/queryTables/queryTable1.xml',
				name: 'ExternalData_1',
				connectionId: 2,
				refreshOnLoad: false,
				saveData: true,
				refreshedVersion: 6,
				state: 'cached',
			}),
		)
		const reopened = await AscendWorkbook.open(new Uint8Array(readFileSync(output)))
		expect(sortedContracts(reopened.connectionParts(), connectionContract)).toEqual(
			sortedContracts(committed.postWrite.dataConnections.connections, connectionContract),
		)
	})

	test('fails closed on public query-table workbook with dangling thumbnail relationship', async () => {
		const input = join(TEMP_DIR, 'libreoffice-query-table-dangling-thumbnail.xlsx')
		const output = join(TEMP_DIR, 'libreoffice-query-table-dangling-thumbnail-out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const source = loadFixture('../xlsx/libreoffice/queryTableExport.xlsx')
		await Bun.write(input, source)
		const ops = [
			{ op: 'setCells' as const, sheet: 'Sheet1', updates: [{ ref: 'Z99', value: 'audit' }] },
		]

		const plan = await createAgentPlan(input, ops)
		expect(plan.writePolicy.diagnostics).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					code: 'pre-write-check-error',
					severity: 'blocker',
					message: expect.stringContaining('structural check error'),
				}),
			]),
		)
		expect(plan.writePolicy.diagnostics.map((diagnostic) => JSON.stringify(diagnostic))).toEqual(
			expect.arrayContaining([
				expect.stringContaining('docProps/thumbnail.jpeg'),
				expect.stringContaining('Package relationship _rels/.rels#rId2 resolves to missing target'),
			]),
		)
		await expect(
			commitAgentPlan(input, ops, {
				output,
				approvals: plan.approvals.map((approval) => approval.id),
			}),
		).rejects.toThrow('Commit blocked by write policy')
		expect(existsSync(output)).toBe(false)
		expect(Buffer.from(readFileSync(input)).equals(Buffer.from(source))).toBe(true)
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

	test('fails closed on public query-table column topology edits without dirtying output', async () => {
		const source = loadFixture('../xlsx/libreoffice/TableEmptyHeaders.xlsx')
		const workbook = await AscendWorkbook.open(source)
		const beforeConnections = sortedContracts(workbook.connectionParts(), connectionContract)
		const beforeTables = tableContracts(workbook)
		expect(beforeTables).toEqual([
			expect.objectContaining({
				sheetName: 'BTC',
				name: 'Bitcoin',
				tableType: 'queryTable',
				queryTablePartPath: 'xl/queryTables/queryTable1.xml',
			}),
		])

		const blocked = workbook.apply([{ op: 'resizeTable', table: 'Bitcoin', ref: 'A1:C16' }], {
			journal: true,
		})

		expect(blocked.errors.map((error) => error.message).join('\n')).toContain(
			'Cannot resize queryTable-backed table "Bitcoin"',
		)
		expect(blocked.journal).toMatchObject({
			supported: false,
			exact: false,
			inverseOps: [],
			issues: [
				expect.objectContaining({
					code: 'JOURNAL_UNAVAILABLE',
					surface: 'package-parts',
				}),
			],
		})
		expect(tableContracts(workbook)).toEqual(beforeTables)

		const saved = workbook.toBytes()
		const reopened = await AscendWorkbook.open(saved)
		expect(sortedContracts(reopened.connectionParts(), connectionContract)).toEqual(
			beforeConnections,
		)
		expect(tableContracts(reopened)).toEqual(beforeTables)
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(saved),
			),
		).toEqual([])
		expect(
			auditXlsxPackageGraphBytePreservation(inspectXlsxPackageGraph(source), source, saved),
		).toEqual([])
	})
})
