import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphSafeEditIntegrity,
	inspectXlsxPackageGraph,
} from '@ascend/io-xlsx'
import { type ActiveContentInfo, AscendWorkbook } from '@ascend/sdk'

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

function activeContentContract(entry: ActiveContentInfo): Record<string, unknown> {
	return {
		kind: entry.kind,
		partPath: entry.partPath,
		contentType: entry.contentType,
		anchor: entry.anchor,
		sheetName: entry.sheetName,
		sourcePartPath: entry.sourcePartPath,
		sourceRelationshipId: entry.sourceRelationshipId,
		relationshipCount: entry.relationshipCount,
		opaque: entry.opaque,
		executionPolicy: entry.executionPolicy,
		invalidationPolicy: entry.invalidationPolicy,
		resigningPolicy: entry.resigningPolicy,
		activeX: entry.activeX
			? {
					classId: entry.activeX.classId,
					persistence: entry.activeX.persistence,
					binaryRelationshipId: entry.activeX.binaryRelationshipId,
					binaryTarget: entry.activeX.binaryTarget,
				}
			: undefined,
		formControl: entry.formControl
			? {
					objectType: entry.formControl.objectType,
					macro: entry.formControl.macro,
					linkedCell: entry.formControl.linkedCell,
					listFillRange: entry.formControl.listFillRange,
					checked: entry.formControl.checked,
					dropLines: entry.formControl.dropLines,
				}
			: undefined,
		worksheetControl: entry.worksheetControl
			? {
					shapeId: entry.worksheetControl.shapeId,
					name: entry.worksheetControl.name,
					relationshipId: entry.worksheetControl.relationshipId,
					controlPrRelationshipId: entry.worksheetControl.controlPrRelationshipId,
					controlPrTarget: entry.worksheetControl.controlPrTarget,
					vmlShapeId: entry.worksheetControl.vmlShapeId,
					vmlImageTarget: entry.worksheetControl.vmlImageTarget,
				}
			: undefined,
		vbaProject: entry.vbaProject
			? {
					moduleCount: entry.vbaProject.moduleCount,
					modules: entry.vbaProject.modules,
					projectStreamPresent: entry.vbaProject.projectStreamPresent,
				}
			: undefined,
	}
}

function sortedActiveContracts(
	entries: readonly ActiveContentInfo[],
): readonly Record<string, unknown>[] {
	return entries
		.map(activeContentContract)
		.sort((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right)))
}

async function safeEdit(bytes: Uint8Array, sheet: string, ref: string): Promise<Uint8Array> {
	const workbook = await AscendWorkbook.open(bytes)
	const result = workbook.apply(
		[{ op: 'setCells', sheet, updates: [{ ref, value: 'safe-edit' }] }],
		{ journal: true },
	)
	expect(result.errors).toEqual([])
	expect(result.journal).toMatchObject({
		supported: true,
		exact: false,
		issues: [expect.objectContaining({ code: 'LOSSY_INVERSE', surface: 'package-parts' })],
	})
	return workbook.toBytes()
}

function activePackageIssues(
	issues: readonly { featureFamily?: unknown; partPath?: unknown; message?: unknown }[],
): readonly unknown[] {
	return issues.filter((issue) =>
		/Macro|Active|Control|Signature|Custom/i.test(
			String(issue.featureFamily ?? issue.partPath ?? issue.message ?? ''),
		),
	)
}

describe('active content corpus contract', () => {
	test('preserves a public Calamine VBA project as blocked opaque active content', async () => {
		const source = loadFixture('../xlsx/calamine/vba.xlsm')
		const workbook = await AscendWorkbook.open(source)
		const before = sortedActiveContracts(workbook.inspect().activeContent)
		expect(before).toEqual([
			expect.objectContaining({
				kind: 'vbaProject',
				partPath: 'xl/vbaProject.bin',
				opaque: true,
				executionPolicy: 'blocked',
				vbaProject: {
					moduleCount: 5,
					modules: [
						{ name: 'ThisWorkbook', kind: 'document' },
						{ name: 'Sheet1', kind: 'document' },
						{ name: 'Sheet2', kind: 'document' },
						{ name: 'Sheet3', kind: 'document' },
						{ name: 'testVBA', kind: 'standard' },
					],
					projectStreamPresent: true,
				},
			}),
		])
		expect(JSON.stringify(before)).not.toContain('Attribute VB_Name')
		expect(workbook.trustReport().findings).toContainEqual(
			expect.objectContaining({
				code: 'workbook.vbaProject',
				severity: 'blocked',
				message: 'Workbook contains VBA. Ascend preserves VBA but never executes it.',
			}),
		)

		const edited = await safeEdit(source, 'Sheet1', 'Z10')
		const reopened = await AscendWorkbook.open(edited)
		expect(sortedActiveContracts(reopened.inspect().activeContent)).toEqual(before)
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(edited),
			),
		).toEqual([])
		expect(
			activePackageIssues(
				auditXlsxPackageGraphBytePreservation(inspectXlsxPackageGraph(source), source, edited),
			),
		).toEqual([])
	})

	test('preserves public LibreOffice ActiveX control XML, binary, worksheet, and VML identity', async () => {
		const source = loadFixture('../xlsx/libreoffice/activex_checkbox.xlsx')
		const workbook = await AscendWorkbook.open(source)
		const before = sortedActiveContracts(workbook.inspect().activeContent)
		expect(before).toEqual([
			expect.objectContaining({
				kind: 'activeX',
				partPath: 'xl/activeX/activeX1.bin',
				sourcePartPath: 'xl/activeX/activeX1.xml',
				sourceRelationshipId: 'rId1',
			}),
			expect.objectContaining({
				kind: 'activeX',
				partPath: 'xl/activeX/activeX1.xml',
				sheetName: 'Sheet1',
				sourceRelationshipId: 'rId3',
				activeX: expect.objectContaining({
					classId: '{8BD21D40-EC42-11CE-9E0D-00AA006002F3}',
					persistence: 'persistStreamInit',
					binaryRelationshipId: 'rId1',
					binaryTarget: 'activeX1.bin',
				}),
				worksheetControl: expect.objectContaining({
					shapeId: 1025,
					name: 'CheckBox1343',
					relationshipId: 'rId3',
					controlPrRelationshipId: 'rId4',
					controlPrTarget: 'xl/media/image1.emf',
					vmlShapeId: 'CheckBox1343',
					vmlImageTarget: 'xl/media/image1.emf',
				}),
			}),
		])
		expect(
			workbook.trustReport().findings.filter((finding) => finding.code === 'workbook.activeX'),
		).toEqual([
			expect.objectContaining({ severity: 'blocked' }),
			expect.objectContaining({ severity: 'blocked' }),
		])

		const edited = await safeEdit(source, 'Sheet1', 'Z10')
		const reopened = await AscendWorkbook.open(edited)
		expect(sortedActiveContracts(reopened.inspect().activeContent)).toEqual(before)
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(edited),
			),
		).toEqual([])
		expect(
			activePackageIssues(
				auditXlsxPackageGraphBytePreservation(inspectXlsxPackageGraph(source), source, edited),
			),
		).toEqual([])
	})

	test('preserves public LibreOffice form controls while reporting non-executable bindings', async () => {
		const cases = [
			{
				path: '../xlsx/libreoffice/button-form-control.xlsx',
				expected: {
					partPath: 'xl/ctrlProps/ctrlProp1.xml',
					sheetName: 'Sheet1',
					sourceRelationshipId: 'rId3',
					formControl: expect.objectContaining({ objectType: 'Button' }),
					worksheetControl: expect.objectContaining({
						shapeId: 1025,
						name: 'Button 1',
					}),
				},
			},
			{
				path: '../xlsx/libreoffice/singlecontrol.xlsx',
				expected: {
					partPath: 'xl/ctrlProps/ctrlProp1.xml',
					sheetName: 'Sheet1',
					sourceRelationshipId: 'rId4',
					formControl: expect.objectContaining({
						objectType: 'CheckBox',
						checked: 'Checked',
					}),
					worksheetControl: expect.objectContaining({
						shapeId: 1026,
						name: 'Check Box 2',
					}),
				},
			},
			{
				path: '../xlsx/libreoffice/checkbox-form-control.xlsx',
				expected: {
					partPath: 'xl/ctrlProps/ctrlProp1.xml',
					sheetName: 'Sheet1',
					sourceRelationshipId: 'rId3',
					formControl: expect.objectContaining({
						objectType: 'CheckBox',
					}),
					worksheetControl: expect.objectContaining({
						shapeId: 1025,
						name: 'Check Box 1',
					}),
				},
			},
		] as const

		for (const entry of cases) {
			const source = loadFixture(entry.path)
			const workbook = await AscendWorkbook.open(source)
			const before = sortedActiveContracts(workbook.inspect().activeContent)
			expect(before).toEqual([expect.objectContaining({ kind: 'formControl', ...entry.expected })])
			expect(workbook.trustReport().findings).toContainEqual(
				expect.objectContaining({
					code: 'workbook.formControl',
					severity: 'warning',
					message: 'Workbook contains form controls that may bind to macros or workbook state.',
				}),
			)

			const edited = await safeEdit(source, 'Sheet1', 'Z10')
			const reopened = await AscendWorkbook.open(edited)
			expect(sortedActiveContracts(reopened.inspect().activeContent)).toEqual(before)
			expect(
				auditXlsxPackageGraphSafeEditIntegrity(
					inspectXlsxPackageGraph(source),
					inspectXlsxPackageGraph(edited),
				),
			).toEqual([])
			expect(
				activePackageIssues(
					auditXlsxPackageGraphBytePreservation(inspectXlsxPackageGraph(source), source, edited),
				),
			).toEqual([])
		}
	})
})
