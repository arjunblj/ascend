import { afterAll, describe, expect, test } from 'bun:test'
import { unlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendWorkbook } from '@ascend/sdk'
import { createServer } from './server.ts'

const TEMP_FILE = join(
	tmpdir(),
	`ascend-api-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
)
const OUTPUT_FILE = join(
	tmpdir(),
	`ascend-api-out-${Date.now()}-${Math.random().toString(16).slice(2)}.xlsx`,
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
})
