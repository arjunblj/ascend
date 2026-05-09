import { afterEach, describe, expect, test } from 'bun:test'
import { existsSync, mkdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createZip, encode } from '../../io-xlsx/src/writer/zip.ts'
import { AscendWorkbook, commitAgentPlan, createAgentPlan } from './index.ts'

const TEMP_DIR = join(tmpdir(), `ascend-agent-workflow-${process.pid}`)

afterEach(() => {
	if (existsSync(TEMP_DIR)) rmSync(TEMP_DIR, { recursive: true, force: true })
})

describe('agent workflow loss audit', () => {
	test('plans report blocked preserved features and commits require explicit allow-loss', async () => {
		const input = join(TEMP_DIR, 'preserved.xlsx')
		const output = join(TEMP_DIR, 'out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makePreservedCustomXlsx())
		const ops = [{ op: 'setCells' as const, sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }]

		const plan = await createAgentPlan(input, ops)
		expect(plan.lossAudit.ok).toBe(false)
		expect(plan.lossAudit.blockedFeatures[0]?.feature).toBe('preservedOther')
		expect(plan.trace.kind).toBe('plan')
		expect(plan.trace.traceDigest).toMatch(/^[a-f0-9]{64}$/)
		expect(plan.trace.phases.find((phase) => phase.phase === 'loss-audit')?.status).toBe('blocked')
		expect(plan.modelOutput.blocked).toBe(true)
		expect(plan.modelOutput.nextActions.join('\n')).toContain('allowLoss')

		await expect(commitAgentPlan(input, ops, { output })).rejects.toThrow(
			'Workbook contains preserved or unsupported features',
		)

		const committed = await commitAgentPlan(input, ops, {
			output,
			allowLoss: ['preservedOther'],
		})
		expect(committed.lossAudit.ok).toBe(true)
		expect(committed.outputSha256).toMatch(/^[a-f0-9]{64}$/)
		expect(committed.trace.kind).toBe('commit')
		expect(committed.trace.outputSha256).toBe(committed.outputSha256)
		expect(committed.modelOutput.blocked).toBe(false)
		expect(committed.modelOutput.digests.traceDigest).toBe(committed.trace.traceDigest)
	})

	test('clean workbooks commit without allow-loss', async () => {
		const input = join(TEMP_DIR, 'clean.xlsx')
		const output = join(TEMP_DIR, 'clean-out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		await wb.save(input)

		const committed = await commitAgentPlan(
			input,
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] }],
			{ output },
		)
		expect(committed.lossAudit.ok).toBe(true)
		expect(committed.trace.artifacts.map((artifact) => artifact.name)).toContain('apply')
		expect(committed.modelOutput.counts.operations).toBe(1)
	})
})

function makePreservedCustomXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/custom/custom1.xml" ContentType="application/custom+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/custom/custom1.xml': encode('<custom>preserve me</custom>'),
			}),
		),
	)
}
