import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./practical-latency-contracts.ts', import.meta.url))

describe('practical latency contracts benchmark', () => {
	test('public-tracked preset separates tracked fixtures from generated edit input', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--input-preset',
				'public-tracked',
				'--contract',
				'edit-verify',
				'--repeat',
				'1',
				'--warmup',
				'0',
				'--dry-run',
				'--json',
			],
			{ cwd: process.cwd(), stderr: 'pipe', stdout: 'pipe' },
		)
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])
		expect(exitCode, stderr).toBe(0)
		const payload = JSON.parse(stdout) as {
			readonly args?: {
				readonly inputPreset?: string
				readonly inputFile?: string
				readonly tableInputFile?: string
				readonly editInputFile?: string
			}
			readonly inputs?: readonly {
				readonly role?: string
				readonly kind?: string
				readonly releaseClaimable?: boolean
			}[]
			readonly results?: readonly { readonly id?: string; readonly command?: string }[]
		}
		expect(payload.args?.inputPreset).toBe('public-tracked')
		expect(payload.args?.inputFile).toBe('fixtures/xlsx/calamine/issue_174.xlsx')
		expect(payload.args?.tableInputFile).toBe('fixtures/xlsx/calamine/table-multiple.xlsx')
		expect(payload.args?.editInputFile).toBeUndefined()
		expect(payload.inputs?.find((input) => input.role === 'first-view')).toMatchObject({
			kind: 'tracked-file',
			releaseClaimable: true,
		})
		expect(payload.inputs?.find((input) => input.role === 'edit')).toMatchObject({
			kind: 'generated',
			releaseClaimable: true,
		})
		expect(payload.inputs?.find((input) => input.role === 'table-inspection')).toMatchObject({
			kind: 'tracked-file',
			releaseClaimable: true,
		})
		for (const result of payload.results ?? []) {
			if (result.id === 'workflow-commit' || result.id === 'post-write-breakdown') {
				expect(result.command).not.toContain('--input-file')
			}
		}
	})
})
