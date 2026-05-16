import { afterEach, describe, expect, test } from 'bun:test'
import { existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendWorkbook, commitAgentPlan, createAgentPlan } from '@ascend/sdk'

const TEMP_DIR = join(tmpdir(), `ascend-data-validation-contract-${process.pid}`)

afterEach(() => {
	if (existsSync(TEMP_DIR)) rmSync(TEMP_DIR, { recursive: true, force: true })
})

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

describe('data validation corpus contract', () => {
	test('reports public LibreOffice data validations after safe edit and reopen', async () => {
		const input = join(TEMP_DIR, 'libreoffice-data-validations.xlsx')
		const output = join(TEMP_DIR, 'libreoffice-data-validations-out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const sourceBytes = loadFixture('../xlsx/libreoffice/dataValidity.xlsx')
		await Bun.write(input, sourceBytes)
		const ops = [
			{ op: 'setCells' as const, sheet: 'Foglio1', updates: [{ ref: 'B1', value: 'audit' }] },
		]

		const plan = await createAgentPlan(input, ops)
		const committed = await commitAgentPlan(input, ops, {
			output,
			approvals: plan.approvals.map((approval) => approval.id),
		})

		expect(committed.postWrite.auditsPassed).toBe(true)
		expect(committed.postWrite.dataValidations).toMatchObject({
			total: 3,
			formulaBacked: 3,
			listValidations: 0,
			x14Validations: 0,
			types: ['decimal', 'whole', 'custom'],
			preservationMode: 'generated',
			verification: 'reopened-output',
		})
		expect(committed.postWrite.dataValidations.ranges).toEqual(
			expect.arrayContaining(['Foglio1!C3:C7', 'Foglio2!C4:G8', 'Foglio3!C3:C7']),
		)
		expect(committed.postWrite.dataValidations.validations).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					sheetName: 'Foglio1',
					sqref: 'C3:C7',
					location: 'Foglio1!C3:C7',
					type: 'decimal',
					operator: 'greaterThan',
					formula1: '3.14',
					allowBlank: true,
				}),
				expect.objectContaining({
					sheetName: 'Foglio2',
					sqref: 'C4:G8',
					location: 'Foglio2!C4:G8',
					type: 'whole',
					formula1: '1',
					formula2: '10',
				}),
				expect.objectContaining({
					sheetName: 'Foglio3',
					sqref: 'C3:C7',
					location: 'Foglio3!C3:C7',
					type: 'custom',
					formula1: 'ISTEXT(C3)',
				}),
			]),
		)
		const reopened = await AscendWorkbook.open(new Uint8Array(readFileSync(output)))
		expect(reopened.inspect().dataValidationCount).toBe(3)
		expect(Buffer.from(readFileSync(input)).equals(Buffer.from(sourceBytes))).toBe(true)
	})
})
