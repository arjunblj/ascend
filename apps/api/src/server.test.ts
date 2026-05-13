import { afterAll, describe, expect, test } from 'bun:test'
import { unlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendWorkbook } from '@ascend/sdk'
import { createZip, encode } from '../../../packages/io-xlsx/src/writer/zip.ts'
import { createServer } from './server.ts'

const TEMP_FILE = join(
	tmpdir(),
	`ascend-api-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
)
const OUTPUT_FILE = join(
	tmpdir(),
	`ascend-api-out-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
)
const MACRO_FILE = join(
	tmpdir(),
	`ascend-api-macro-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
)
const MACRO_OUTPUT_FILE = join(
	tmpdir(),
	`ascend-api-macro-out-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsm`,
)

let server: ReturnType<typeof createServer> | undefined

interface ApiEnvelope {
	readonly ok: boolean
	readonly data?: {
		readonly approvals?: readonly { readonly id: string }[]
	}
	readonly error?: {
		readonly message?: string
	}
}

afterAll(async () => {
	server?.stop(true)
	await unlink(TEMP_FILE).catch(() => {})
	await unlink(OUTPUT_FILE).catch(() => {})
	await unlink(MACRO_FILE).catch(() => {})
	await unlink(MACRO_OUTPUT_FILE).catch(() => {})
})

async function postJson(
	path: string,
	body: unknown,
): Promise<{ status: number; body: ApiEnvelope }> {
	server ??= createServer({ port: 0 })
	const response = await fetch(`http://127.0.0.1:${server.port}${path}`, {
		method: 'POST',
		headers: { 'Content-Type': 'application/json' },
		body: JSON.stringify(body),
	})
	return { status: response.status, body: (await response.json()) as ApiEnvelope }
}

describe('Ascend API server', () => {
	test('plan and commit require exact approval ids', async () => {
		const workbook = AscendWorkbook.create()
		workbook.apply([{ op: 'addSheet', name: 'Scratch' }])
		await workbook.save(TEMP_FILE)
		const ops = [{ op: 'deleteSheet', sheet: 'Scratch' }]

		const plan = await postJson('/plan', { file: TEMP_FILE, ops })
		expect(plan.status).toBe(200)
		expect(plan.body.ok).toBe(true)
		const approvalId = plan.body.data.approvals[0].id
		expect(approvalId).toBe('op:0:deletesheet')

		const aliasCommit = await postJson('/commit', {
			file: TEMP_FILE,
			ops,
			output: OUTPUT_FILE,
			approvals: ['deleteSheet'],
		})
		expect(aliasCommit.status).toBe(400)
		expect(aliasCommit.body.ok).toBe(false)
		expect(aliasCommit.body.error.message).toBe('Commit requires explicit approval')

		const exactCommit = await postJson('/commit', {
			file: TEMP_FILE,
			ops,
			output: OUTPUT_FILE,
			approvals: [approvalId],
		})
		expect(exactCommit.status).toBe(200)
		expect(exactCommit.body.ok).toBe(true)
		expect(exactCommit.body.data.approvals[0].id).toBe(approvalId)
	})

	test('commit requires exact approval ids for preserved lossy features', async () => {
		await Bun.write(MACRO_FILE, signedMacroWorkbook())
		const ops = [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] }]
		const plan = await postJson('/plan', { file: MACRO_FILE, ops })
		expect(plan.status).toBe(200)
		expect(plan.body.ok).toBe(true)
		const approvalIds = plan.body.data.approvals.map((approval) => approval.id)
		expect(approvalIds).toEqual(
			expect.arrayContaining([
				expect.stringMatching(/^loss:preservedmacro:preserved:/),
				expect.stringMatching(/^loss:preservedsignature:preserved:/),
			]),
		)

		const aliasCommit = await postJson('/commit', {
			file: MACRO_FILE,
			ops,
			output: MACRO_OUTPUT_FILE,
			approvals: ['preservedMacro', 'preservedSignature'],
		})
		expect(aliasCommit.status).toBe(400)
		expect(aliasCommit.body.ok).toBe(false)
		expect(aliasCommit.body.error.message).toBe('Commit requires explicit approval')

		const exactCommit = await postJson('/commit', {
			file: MACRO_FILE,
			ops,
			output: MACRO_OUTPUT_FILE,
			approvals: approvalIds,
		})
		expect(exactCommit.status).toBe(200)
		expect(exactCommit.body.ok).toBe(true)
		expect(exactCommit.body.data.approvals.map((approval) => approval.id)).toEqual(approvalIds)
	})
})

function signedMacroWorkbook(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>`),
				'xl/_rels/vbaProject.bin.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVbaSignature" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/vbaProject.bin': encode('macro-bytes'),
				'xl/vbaProjectSignature.bin': encode('signature-bytes'),
			}),
		),
	)
}
