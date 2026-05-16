import { afterEach, describe, expect, test } from 'bun:test'
import { existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { auditXlsxPackageGraphSafeEditIntegrity, inspectXlsxPackageGraph } from '@ascend/io-xlsx'
import { AscendWorkbook, commitAgentPlan, createAgentPlan } from '@ascend/sdk'

const TEMP_DIR = join(tmpdir(), `ascend-formula-binding-contract-${process.pid}`)

afterEach(() => {
	if (existsSync(TEMP_DIR)) rmSync(TEMP_DIR, { recursive: true, force: true })
})

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

function formulaBindingRefs(
	workbook: AscendWorkbook,
	sheetName: string,
	refs: readonly string[],
): readonly string[] {
	return refs
		.filter((ref) => workbook.sheet(sheetName)?.cell(ref)?.formulaBinding)
		.map((ref) => `${sheetName}!${ref}`)
}

function cellFormulaContract(
	workbook: AscendWorkbook,
	sheetName: string,
	refs: readonly string[],
): Record<string, unknown> {
	return Object.fromEntries(
		refs.map((ref) => [
			ref,
			{
				formula: workbook.formula(`${sheetName}!${ref}`)?.normalizedFormula ?? null,
				binding: workbook.sheet(sheetName)?.cell(ref)?.formulaBinding ?? null,
				value: workbook.sheet(sheetName)?.cell(ref)?.value ?? null,
			},
		]),
	)
}

function formulaBindingIntegrityIssues(workbook: AscendWorkbook): readonly unknown[] {
	return workbook.check().issues.filter((issue) => issue.rule === 'formula-binding-integrity')
}

describe('formula binding corpus contract', () => {
	test('commit proof reports missing public formula caches after save and reopen', async () => {
		const input = join(TEMP_DIR, 'closedxml-formulas-without-caches.xlsx')
		const output = join(TEMP_DIR, 'closedxml-formulas-without-caches-out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const source = loadFixture('../xlsx/closedxml/Misc_Formulas.xlsx')
		await Bun.write(input, source)
		const ops = [
			{ op: 'setCells' as const, sheet: 'Formulas', updates: [{ ref: 'H20', value: 'audit' }] },
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
		expect(committed.postWrite.formulaState).toMatchObject({
			calcChainState: 'present',
			calcChainParts: ['xl/calcChain.xml'],
			recalculationRequested: false,
			formulaCells: 12,
			cachedFormulaValues: 0,
			missingCachedFormulaValues: 12,
			formulaCacheState: 'all-missing',
			cachedValueKinds: [],
			missingCachedFormulaLocationSample: [
				'Formulas!C2',
				'Formulas!G2',
				'Formulas!C3',
				'Formulas!G3',
				'Formulas!C4',
				'Formulas!G4',
				'Formulas!B6',
				'Formulas!C6',
				'Formulas!A11',
				'Formulas!A12',
				'Formulas!A13',
				'Formulas!A14',
			],
		})
		expect(committed.postWrite.formulaState.warnings.join('\n')).toContain(
			'12 reopened formula cell(s) do not carry cached formula values',
		)
		const reopened = await AscendWorkbook.open(new Uint8Array(readFileSync(output)))
		expect(cellFormulaContract(reopened, 'Formulas', ['C2', 'G2', 'A14'])).toMatchObject({
			C2: { formula: 'A2+$B$2', value: { kind: 'empty' } },
			G2: { formula: 'IF(C2=F2,"Yes","No")', value: { kind: 'empty' } },
			A14: { formula: 'SUM(8:9)', value: { kind: 'empty' } },
		})
		expect(Buffer.from(readFileSync(input)).equals(Buffer.from(source))).toBe(true)
	})

	test('preserves public POI shared formulas through unrelated safe edits and reopen', async () => {
		const source = loadFixture('../xlsx/poi/shared_formulas.xlsx')
		const workbook = await AscendWorkbook.open(source)
		const refs = Array.from({ length: 40 }, (_, index) => `A${index + 2}`)
		const before = cellFormulaContract(workbook, 'Label', refs)
		expect(formulaBindingRefs(workbook, 'Label', refs)).toEqual(refs.map((ref) => `Label!${ref}`))
		expect(before.A2).toMatchObject({
			formula: 'B2',
			binding: { kind: 'shared', sharedIndex: '0', isMaster: true, ref: 'A2:A41' },
		})
		expect(before.A41).toMatchObject({
			formula: 'B41',
			binding: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A2' },
		})

		const changed = workbook.apply(
			[{ op: 'setCells', sheet: 'Label', updates: [{ ref: 'D1', value: 'safe-edit' }] }],
			{ journal: true },
		)
		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.issues.filter((issue) => issue.surface === 'shared-formulas')).toEqual(
			[],
		)

		const edited = workbook.toBytes()
		const reopened = await AscendWorkbook.open(edited)
		expect(formulaBindingIntegrityIssues(reopened)).toEqual([])
		expect(cellFormulaContract(reopened, 'Label', refs)).toEqual(before)
		expect(
			auditXlsxPackageGraphSafeEditIntegrity(
				inspectXlsxPackageGraph(source),
				inspectXlsxPackageGraph(edited),
			),
		).toEqual([])
	})

	test('materializes public Calamine multi-axis shared formulas when an edit intersects the group', async () => {
		const source = loadFixture('../xlsx/calamine/issue_565_multi_axis_shared.xlsx')
		const workbook = await AscendWorkbook.open(source)
		const refs = ['B1', 'C1', 'D1', 'B2', 'C2', 'D2']
		expect(formulaBindingRefs(workbook, 'Sheet1', refs)).toEqual(refs.map((ref) => `Sheet1!${ref}`))

		const changed = workbook.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'C2', value: 99 }] }],
			{ journal: true },
		)
		expect(changed.errors).toEqual([])
		expect(changed.affectedCells).toEqual(refs)
		expect(changed.journal?.exact).toBe(false)
		expect(
			changed.journal?.issues
				.filter((issue) => issue.surface === 'shared-formulas')
				.map((issue) => issue.reason),
		).toEqual(refs.map(() => 'formula-binding-metadata'))

		const reopened = await AscendWorkbook.open(workbook.toBytes())
		expect(formulaBindingIntegrityIssues(reopened)).toEqual([])
		expect(formulaBindingRefs(reopened, 'Sheet1', refs)).toEqual([])
		expect(cellFormulaContract(reopened, 'Sheet1', refs)).toMatchObject({
			B1: { formula: 'B1' },
			C1: { formula: 'C1' },
			D1: { formula: 'D1' },
			B2: { formula: 'B2' },
			C2: { formula: null, value: { kind: 'number', value: 99 } },
			D2: { formula: 'D2' },
		})
	})

	test('fails closed on destructive public ClosedXML legacy array edits without dropping metadata', async () => {
		const source = loadFixture('../xlsx/closedxml/Other_Formulas_ArrayFormula.xlsx')
		const workbook = await AscendWorkbook.open(source)
		const before = cellFormulaContract(workbook, 'Sheet1', ['A1', 'B2'])
		expect(before.A1).toMatchObject({
			formula: '1+2',
			binding: { kind: 'array', ref: 'A1:B2' },
		})
		expect(formulaBindingRefs(workbook, 'Sheet1', ['A1', 'B2'])).toEqual(['Sheet1!A1'])

		const changed = workbook.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B2', value: 99 }] }],
			{ journal: true },
		)
		expect(changed.errors).toEqual([
			expect.objectContaining({ code: 'VALIDATION_ERROR', refs: ['B2', 'A1:B2'] }),
		])
		expect(changed.affectedCells).toEqual([])
		expect(changed.journal).toMatchObject({
			supported: false,
			exact: false,
			issues: [expect.objectContaining({ code: 'JOURNAL_UNAVAILABLE' })],
		})

		const reopened = await AscendWorkbook.open(workbook.toBytes())
		expect(formulaBindingIntegrityIssues(reopened)).toEqual([])
		expect(cellFormulaContract(reopened, 'Sheet1', ['A1', 'B2'])).toEqual(before)
	})

	test('reports public ClosedXML data-table member edits as binding-lossy and reopens cleanly', async () => {
		const source = loadFixture('../xlsx/closedxml/Other_Formulas_DataTableFormula-Excel-Input.xlsx')
		const workbook = await AscendWorkbook.open(source)
		expect(formulaBindingRefs(workbook, '1D Row', ['C4', 'C6'])).toEqual(['1D Row!C4'])
		expect(workbook.sheet('1D Row')?.cell('C4')?.formulaBinding).toMatchObject({
			kind: 'dataTable',
			ref: 'C4:C8',
		})

		const changed = workbook.apply(
			[{ op: 'setCells', sheet: '1D Row', updates: [{ ref: 'C6', value: 99 }] }],
			{ journal: true },
		)
		expect(changed.errors).toEqual([])
		expect([...changed.affectedCells].sort()).toEqual(['C4', 'C6'])
		expect(changed.journal?.exact).toBe(false)
		expect(changed.journal?.issues.filter((issue) => issue.surface === 'data-tables')).toEqual([
			expect.objectContaining({
				surface: 'data-tables',
				reason: 'formula-binding-metadata',
				refs: ['1D Row!C4'],
			}),
		])

		const reopened = await AscendWorkbook.open(workbook.toBytes())
		expect(formulaBindingIntegrityIssues(reopened)).toEqual([])
		expect(formulaBindingRefs(reopened, '1D Row', ['C4', 'C6'])).toEqual([])
		expect(cellFormulaContract(reopened, '1D Row', ['C4', 'C6'])).toMatchObject({
			C4: { formula: null },
			C6: { formula: null, value: { kind: 'number', value: 99 } },
		})
	})
})
