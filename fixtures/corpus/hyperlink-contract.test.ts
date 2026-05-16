import { afterEach, describe, expect, test } from 'bun:test'
import { existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { commitAgentPlan, createAgentPlan } from '../../packages/sdk/src/agent-workflow.ts'
import { AscendWorkbook } from '../../packages/sdk/src/workbook.ts'

const TEMP_DIR = join(tmpdir(), `ascend-hyperlink-contract-${process.pid}`)

afterEach(() => {
	if (existsSync(TEMP_DIR)) rmSync(TEMP_DIR, { recursive: true, force: true })
})

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

describe('hyperlink corpus contract', () => {
	test('reports public ClosedXML hyperlinks after safe edit and reopen', async () => {
		const input = join(TEMP_DIR, 'closedxml-hyperlinks.xlsx')
		const output = join(TEMP_DIR, 'closedxml-hyperlinks-out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const sourceBytes = loadFixture('../xlsx/closedxml/Misc_Hyperlinks.xlsx')
		await Bun.write(input, sourceBytes)
		const ops = [
			{ op: 'setCells' as const, sheet: 'Second Sheet', updates: [{ ref: 'B2', value: 'audit' }] },
		]

		const plan = await createAgentPlan(input, ops)
		const committed = await commitAgentPlan(input, ops, {
			output,
			approvals: plan.approvals.map((approval) => approval.id),
		})

		expect(committed.postWrite.auditsPassed).toBe(true)
		expect(committed.postWrite.hyperlinks).toMatchObject({
			total: 10,
			externalTargets: 5,
			internalLocations: 5,
			displayed: 5,
			withTooltips: 2,
			locations: [
				'Hyperlinks!A1',
				'Hyperlinks!A2',
				'Hyperlinks!A3',
				'Hyperlinks!A4',
				'Hyperlinks!A5',
				'Hyperlinks!A6',
				'Hyperlinks!A7',
				'Hyperlinks!A8',
				'Hyperlinks!A9',
				'Hyperlinks!A11',
			],
			preservationMode: 'generated',
			verification: 'reopened-output',
		})
		expect(committed.postWrite.hyperlinks.links).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					sheetName: 'Hyperlinks',
					ref: 'A1',
					location: 'Hyperlinks!A1',
					target: 'http://www.yahoo.com',
				}),
				expect.objectContaining({
					sheetName: 'Hyperlinks',
					ref: 'A7',
					location: 'Hyperlinks!A7',
					internalLocation: "'Second Sheet'!A1",
					display: 'Link to an address in another worksheet',
				}),
				expect.objectContaining({
					sheetName: 'Hyperlinks',
					ref: 'A8',
					location: 'Hyperlinks!A8',
					internalLocation: 'Hyperlinks!B1:C2',
					display: 'Link to a range in this worksheet',
					tooltip: 'SquareBox',
				}),
			]),
		)
		const reopened = await AscendWorkbook.open(new Uint8Array(readFileSync(output)))
		expect(reopened.sheet('Hyperlinks')?.getHyperlinks()).toHaveLength(10)
		expect(Buffer.from(readFileSync(input)).equals(Buffer.from(sourceBytes))).toBe(true)
	})
})
