import { afterAll, describe, expect, test } from 'bun:test'
import { existsSync, unlinkSync } from 'node:fs'
import { AscendWorkbook } from '@ascend/sdk'

const CLI = new URL('./index.ts', import.meta.url).pathname
const TEST_FILE = 'test-visuals.xlsx'

function run(...args: string[]): Promise<{ stdout: string; stderr: string; exitCode: number }> {
	return new Promise((resolve) => {
		const proc = Bun.spawn(['bun', 'run', CLI, ...args], {
			stdout: 'pipe',
			stderr: 'pipe',
			cwd: import.meta.dir,
		})

		proc.exited.then(async (exitCode) => {
			const stdout = await new Response(proc.stdout).text()
			const stderr = await new Response(proc.stderr).text()
			resolve({ stdout: stdout.trim(), stderr: stderr.trim(), exitCode })
		})
	})
}

afterAll(() => {
	const path = `${import.meta.dir}/${TEST_FILE}`
	if (existsSync(path)) unlinkSync(path)
})

describe('inspect visuals detail', () => {
	test('inspect --detail visuals --json returns workbook visual inventory', async () => {
		const wb = AscendWorkbook.create()
		await wb.save(`${import.meta.dir}/${TEST_FILE}`)

		const { exitCode, stdout } = await run('inspect', TEST_FILE, '--detail', 'visuals', '--json')

		expect(exitCode).toBe(0)
		const parsed = JSON.parse(stdout)
		expect(parsed.data.load.mode).toBe('full')
		expect(parsed.data.sheetImageCount).toBe(0)
		expect(parsed.data.sheets[0]).toMatchObject({
			sheet: 'Sheet1',
			hasDrawing: false,
			hasLegacyDrawing: false,
			imageCount: 0,
		})
	})

	test('inspect --detail visuals shows drawing object links in text output', async () => {
		const fixture = new URL(
			'../../../fixtures/xlsx/libreoffice/textbox-hyperlink.xlsx',
			import.meta.url,
		).pathname
		const { exitCode, stdout } = await run('inspect', fixture, '--detail', 'visuals')

		expect(exitCode).toBe(0)
		expect(stdout).toContain('Drawing Object Links')
		expect(stdout).toContain('TextBox 1')
		expect(stdout).toContain('https://www.google.com/')
		expect(stdout).toContain('Sheet drawing objects')
	})
})
