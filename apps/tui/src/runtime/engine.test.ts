import { describe, expect, test } from 'bun:test'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { AscendWorkbook, type CellStyle, parseRange } from '@ascend/sdk'
import stripAnsi from 'strip-ansi'
import { visibleLength } from '../render/ansi-text.ts'
import { WorkbookTuiEngine } from './engine.ts'

describe('WorkbookTuiEngine', () => {
	test('edits a cell through the headless trace path', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 24, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		const result = await engine.runHeadless(
			[
				{ kind: 'key', key: 'text', text: '4' },
				{ kind: 'key', key: 'text', text: '2' },
				{ kind: 'key', key: 'Enter' },
			],
			{ size: { rows: 24, cols: 80 }, includeFrames: true },
		)
		expect(result.state.dirty).toBe(true)
		expect(result.state.selection.active.row).toBe(1)
		expect(result.frames.at(-1)?.lines.join('\n')).toContain('42')
	})

	test('supports Excel-like navigation and command mode', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 24, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'ArrowRight' })
		await engine.dispatch({ kind: 'key', key: ':' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'p' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'e' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'r' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'f' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		expect(engine.state().selection.active.col).toBe(1)
		expect(engine.state().message).toContain('perf')
	})

	test('terminal calibration is available as a headless command surface', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		const result = await engine.dispatch({ kind: 'command', command: 'calibrate' })
		expect(result.handled).toBe(true)
		expect(engine.state().workspace.focusedRegion).toBe('inspector')
		expect(engine.state().message).toContain('Terminal calibration')
		const frame = engine.render({ rows: 18, cols: 90 }).lines.join('\n')
		expect(frame).toContain('Terminal Calibration')
		expect(frame).toContain('Keyboard:')
		expect(frame).toContain('Shortcut:')
	})

	test('Ctrl+Arrow navigates to used sheet boundaries', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 12, cols: 60 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '1\t2\t3\n4\t5\t6' })
		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+ArrowRight' })
		expect(engine.state().selection.active).toEqual({ row: 0, col: 2 })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+ArrowDown' })
		expect(engine.state().selection.active).toEqual({ row: 1, col: 2 })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+ArrowLeft' })
		expect(engine.state().selection.active).toEqual({ row: 1, col: 0 })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+ArrowUp' })
		expect(engine.state().selection.active).toEqual({ row: 0, col: 0 })
	})

	test('does not exit wq when an unnamed workbook cannot be saved', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 24, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		const result = await engine.dispatch({ kind: 'command', command: 'wq' })
		expect(result.shouldExit).toBe(false)
		expect(engine.state().dirty).toBe(true)
		expect(engine.state().message).toContain('Save As')
	})

	test('guards dirty workbooks before quit, new, or open replacements', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-'))
		try {
			const path = join(dir, 'other.xlsx')
			await AscendWorkbook.create().save(path)
			const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'key', key: 'text', text: 'dirty' })
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			const quit = await engine.dispatch({ kind: 'command', command: 'quit' })
			expect(quit.shouldExit).toBeFalsy()
			expect(engine.state().message).toContain('Unsaved changes')
			const open = await engine.dispatch({ kind: 'command', command: `open ${path}` })
			expect(open.handled).toBe(false)
			expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('dirty')
			const forced = await engine.dispatch({ kind: 'command', command: `open! ${path}` })
			expect(forced.handled).toBe(true)
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('new! resets selection, dialogs, hydration cache, and undo history', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'old' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.render({ rows: 18, cols: 90 })
		await engine.dispatch({ kind: 'command', command: 'format' })
		expect(engine.state().activeDialog?.id).toBe('format-cells')
		await engine.dispatch({ kind: 'command', command: 'new!' })
		expect(engine.state().activeDialog).toBeUndefined()
		expect(engine.state().selection.active).toEqual({ row: 0, col: 0 })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).not.toContain('old')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.state().message).toContain('Nothing to undo')
	})

	test('starts on the requested sheet', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-'))
		try {
			const path = join(dir, 'book.xlsx')
			const workbook = AscendWorkbook.create()
			workbook.applyAndRecalc([{ op: 'addSheet', name: 'Data' }])
			await workbook.save(path)
			const engine = await WorkbookTuiEngine.create({
				path,
				sheet: 'Data',
				size: { rows: 24, cols: 80 },
			})
			expect(engine.state().sheetName).toBe('Data')
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('opens row-limited workbook previews as read-only', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-'))
		try {
			const path = join(dir, 'preview.xlsx')
			const workbook = AscendWorkbook.create()
			workbook.applyAndRecalc([
				{
					op: 'setCells',
					sheet: 'Sheet1',
					updates: [
						{ ref: 'A1', value: 'visible' },
						{ ref: 'A2', value: 'capped' },
					],
				},
			])
			await workbook.save(path)
			const engine = await WorkbookTuiEngine.create({
				path,
				loadOptions: { mode: 'values', maxRows: 1 },
				size: { rows: 18, cols: 90 },
			})
			const document = engine.state().workspace.documents[0]
			expect(document?.readOnly).toBe(true)
			expect(document?.info?.load.isPartial).toBe(true)
			expect(engine.state().message).toContain('first 1 rows')
			expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('visible')

			const edit = await engine.dispatch({ kind: 'key', key: 'text', text: 'x' })
			expect(edit.handled).toBe(false)
			expect(engine.state().dirty).toBe(false)
			expect(engine.state().message).toContain('read-only')
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('open command loads a workbook path inside the TUI session', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-'))
		try {
			const path = join(dir, 'openable.xlsx')
			const workbook = AscendWorkbook.create()
			workbook.applyAndRecalc([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'opened' }] },
			])
			await workbook.save(path)
			const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
			const result = await engine.dispatch({ kind: 'command', command: `open ${path}` })
			expect(result.handled).toBe(true)
			expect(engine.state().message).toContain('Opened openable.xlsx')
			expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('opened')
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('save-as writes the active workbook and rebases the document path', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-'))
		try {
			const path = join(dir, 'saved.xlsx')
			const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'key', key: 'text', text: 'saved-as' })
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			const result = await engine.dispatch({ kind: 'command', command: `save-as ${path}` })
			expect(result.handled).toBe(true)
			expect(engine.state().dirty).toBe(false)
			expect(engine.state().workspace.documents[0]?.path).toBe(path)
			const reopened = await AscendWorkbook.open(path)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({
				kind: 'string',
				value: 'saved-as',
			})
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('recent workbook store backs File backstage across sessions', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-recent-engine-'))
		try {
			const storePath = join(dir, 'recent.sqlite')
			const workbookPath = join(dir, 'remember-me.xlsx')
			const engine = await WorkbookTuiEngine.create({
				size: { rows: 12, cols: 80 },
				recentStorePath: storePath,
			})
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'command', command: `save-as ${workbookPath}` })

			const reopened = await WorkbookTuiEngine.create({
				size: { rows: 12, cols: 80 },
				recentStorePath: storePath,
			})
			expect(reopened.state().workspace.fileHub.entries[0]).toMatchObject({
				label: 'remember-me.xlsx',
				path: workbookPath,
			})
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('export command writes CSV and JSON without rebasing the active workbook path', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-'))
		try {
			const csvPath = join(dir, 'export.csv')
			const jsonPath = join(dir, 'export.json')
			const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'key', key: 'text', text: 'Name\tValue\nAlpha\t7' })
			const csv = await engine.dispatch({ kind: 'command', command: `export ${csvPath}` })
			expect(csv.handled).toBe(true)
			expect(await Bun.file(csvPath).text()).toContain('Alpha,7')
			const json = await engine.dispatch({
				kind: 'command',
				command: `export ${JSON.stringify({ path: jsonPath, format: 'json' })}`,
			})
			expect(json.handled).toBe(true)
			expect(await Bun.file(jsonPath).text()).toContain('"sheets"')
			expect(engine.state().workspace.documents[0]?.path).toBeNull()
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('keeps ANSI-styled frame lines at the requested visible width', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 12, cols: 60 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '1' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		const frame = engine.render({ rows: 12, cols: 60 })
		expect(frame.lines.every((line) => visibleLength(line) === 60)).toBe(true)
	})

	test('renders and hit-tests the cursor against the data grid geometry', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 12, cols: 60 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		expect(engine.render({ rows: 12, cols: 60 }).cursor).toEqual({
			row: 5,
			col: 8,
			visible: true,
		})
		await engine.dispatch({ kind: 'mouse', action: 'press', row: 5, col: 19 })
		expect(engine.state().selection.active).toEqual({ row: 0, col: 1 })
		expect(engine.render({ rows: 12, cols: 60 }).cursor).toMatchObject({ row: 5, col: 19 })
	})

	test('context menu is navigable and executes native Excel-core actions', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 96 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'Shift+F10' })
		expect(engine.state().workspace.focusedRegion).toBe('contextMenu')
		expect(engine.state().contextMenu?.address).toBe('A1')
		expect(stripAnsi(engine.render({ rows: 16, cols: 96 }).lines.join('\n'))).toContain(
			'Context: cell A1',
		)

		await engine.dispatch({ kind: 'key', key: 'ArrowDown' })
		await engine.dispatch({ kind: 'key', key: 'ArrowDown' })
		await engine.dispatch({ kind: 'key', key: 'ArrowDown' })
		await engine.dispatch({ kind: 'key', key: 'ArrowDown' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })

		expect(engine.state().contextMenu).toBeUndefined()
		expect(engine.state().activeDialog?.id).toBe('format-cells')
		expect(engine.state().workspace.focusedRegion).toBe('dialog')
	})

	test('right click selects a cell and opens its context menu', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 96 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'mouse', action: 'press', row: 5, col: 19, button: 2 })

		expect(engine.state().selection.active).toEqual({ row: 0, col: 1 })
		expect(engine.state().contextMenu?.address).toBe('B1')
	})

	test('status bar summarizes selected numeric ranges', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 150 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '1\t2' })
		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		await engine.dispatch({ kind: 'key', key: 'Shift+ArrowRight' })
		const frame = engine.render({ rows: 18, cols: 150 }).lines.join('\n')
		expect(frame).toContain('Count 2')
		expect(frame).toContain('Average 1.5')
		expect(frame).toContain('Numerical Count 2')
		expect(frame).toContain('Sum 3')
		expect(frame).toContain('Min 1')
		expect(frame).toContain('Max 2')
	})

	test('ready Enter moves down and F2 edits existing cell contents', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 100 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '42' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		expect(engine.state().selection.active).toEqual({ row: 1, col: 0 })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		expect(engine.state().selection.active).toEqual({ row: 2, col: 0 })
		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		await engine.dispatch({ kind: 'key', key: 'F2' })
		expect(engine.state().mode).toBe('entering')
		expect(engine.state().editBuffer).toBe('42')
	})

	test('formula bar shows active cell content in ready mode', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 100 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '=1+1' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		expect(engine.render({ rows: 18, cols: 100 }).lines.join('\n')).toContain('=1+1')
	})

	test('File backstage keyboard actions are functional', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 100 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+O' })
		await engine.dispatch({ kind: 'key', key: 'ArrowRight' })
		expect(engine.state().workspace.fileHub.section).toBe('open')
		await engine.dispatch({ kind: 'key', key: 'ArrowDown' })
		expect(engine.state().workspace.fileHub.selectedIndex).toBe(1)
		await engine.dispatch({ kind: 'key', key: 'o' })
		expect(engine.state().mode).toBe('command')
		expect(engine.state().commandBuffer).toBe('open ')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		expect(engine.state().workspace.fileHub.visible).toBe(false)
	})

	test('common Excel navigation chords move and extend selection', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 100 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '1\t2\n3\t4' })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+End' })
		expect(engine.state().selection.active).toEqual({ row: 1, col: 1 })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Home' })
		expect(engine.state().selection.active).toEqual({ row: 0, col: 0 })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Shift+ArrowRight' })
		expect(engine.state().selection.kind).toBe('range')
		expect(engine.state().selection.active).toEqual({ row: 0, col: 1 })
	})

	test('formula bar and grid expose Excel selection landmarks', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 100 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'Shift+ArrowRight' })
		const frame = stripAnsi(engine.render({ rows: 18, cols: 100 }).lines.join('\n'))
		expect(frame).toContain('A1:B1')
		expect(frame).toContain('{        }')
		expect(frame).toContain('[        ]')
	})

	test('records renderer telemetry samples', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 100 } })
		engine.recordRendererStats({
			frameBuildMs: 0,
			frameDiffMs: 1.5,
			encodeMs: 0.5,
			writeMs: 0.25,
			changedCells: 20,
			bytesOut: 120,
			droppedFrames: 0,
			fps: 60,
		})
		const latest = engine.state().telemetry.at(-1)
		expect(latest?.diffMs).toBe(1.5)
		expect(latest?.encodeMs).toBe(0.5)
		expect(latest?.ptyWriteMs).toBe(0.25)
		expect(latest?.bytesWritten).toBe(120)
		expect(latest?.fps).toBe(60)
	})

	test('range edit writes only the active cell', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'Shift+ArrowRight' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '7' })
		const result = await engine.dispatch({ kind: 'key', key: 'Enter' })
		expect(result.handled).toBe(true)
		expect(engine.state().message).not.toContain('Invalid')
	})

	test('literal entry replaces an existing formula', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		for (const text of '=1+1') await engine.dispatch({ kind: 'key', key: 'text', text })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '4' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '2' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		const frame = engine.render({ rows: 16, cols: 80 })
		expect(frame.lines.join('\n')).toContain('42')
		expect(frame.lines.join('\n')).not.toContain('#')
	})

	test('invalid formula entry remains editable and keeps the error visible', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		for (const text of '=INVALID((') await engine.dispatch({ kind: 'key', key: 'text', text })
		const result = await engine.dispatch({ kind: 'key', key: 'Enter' })
		expect(result.handled).toBe(false)
		expect(engine.state().mode).toBe('editing')
		expect(engine.state().selection.active).toEqual({ row: 0, col: 0 })
		expect(engine.state().editBuffer).toBe('=INVALID((')
		expect(engine.state().message).toContain('Failed to parse')
	})

	test('F2 loads existing formula text for editing', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		for (const text of '=1+1') await engine.dispatch({ kind: 'key', key: 'text', text })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		await engine.dispatch({ kind: 'key', key: 'F2' })
		expect(engine.state().editBuffer).toBe('=1+1')
	})

	test('formula point mode inserts a navigated reference while preserving the edit target', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '2' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'goto B1' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '=' })
		await engine.dispatch({ kind: 'key', key: 'ArrowLeft' })
		expect(engine.state().mode).toBe('point')
		expect(engine.state().editBuffer).toBe('=A1')
		for (const text of '*2') await engine.dispatch({ kind: 'key', key: 'text', text })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'goto B1' })
		expect(engine.render({ rows: 16, cols: 80 }).lines.join('\n')).toContain('4')
	})

	test('formula point cancel restores the edit target selection', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'command', command: 'goto B1' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '=' })
		await engine.dispatch({ kind: 'key', key: 'ArrowLeft' })
		expect(engine.state().selection.active).toEqual({ row: 0, col: 0 })
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		expect(engine.state().selection.active).toEqual({ row: 0, col: 1 })
	})

	test('formula editing supports cursor insertion and F4 reference cycling', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '3' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'goto B1' })
		for (const text of '=A1*2') await engine.dispatch({ kind: 'key', key: 'text', text })
		await engine.dispatch({ kind: 'key', key: 'Home' })
		await engine.dispatch({ kind: 'key', key: 'ArrowRight' })
		await engine.dispatch({ kind: 'key', key: 'F4' })
		expect(engine.state().editBuffer).toBe('=$A$1*2')
		expect(engine.state().editCursor).toBe(5)
		await engine.dispatch({ kind: 'key', key: 'End' })
		await engine.dispatch({ kind: 'key', key: 'ArrowLeft' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '+' })
		expect(engine.state().editBuffer).toBe('=$A$1*+2')
		await engine.dispatch({ kind: 'key', key: 'Backspace' })
		expect(engine.state().editBuffer).toBe('=$A$1*2')
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		expect(engine.render({ rows: 16, cols: 80 }).lines.join('\n')).toContain('6')
	})

	test('registered dialog commands accept JSON input and journal operations', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		const result = await engine.dispatch({
			kind: 'command',
			command: 'format {"bold":true,"numberFormat":"0.00"}',
		})
		expect(result.handled).toBe(true)
		expect(engine.state().message).toContain('Committed 1 operation')
	})

	test('registered dialog commands report invalid JSON', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		const result = await engine.dispatch({ kind: 'command', command: 'format {"bold":' })
		expect(result.handled).toBe(false)
		expect(engine.state().message).toContain('Invalid command JSON')
	})

	test('registered commands honor focus context', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 16, cols: 80 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+O' })
		const result = await engine.dispatch({ kind: 'command', command: 'format' })
		expect(result.handled).toBe(false)
		expect(engine.state().message).toContain('File hub')
		expect(engine.state().activeDialog).toBeUndefined()
	})

	test('file help, keytips, and info routes are visible without shell knowledge', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 100 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'F1' })
		expect(engine.render({ rows: 18, cols: 100 }).lines.join('\n')).toContain('Ascend TUI Help')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		await engine.dispatch({ kind: 'key', key: 'F10' })
		expect(engine.render({ rows: 18, cols: 100 }).lines.join('\n')).toContain('KeyTips')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		await engine.dispatch({ kind: 'command', command: 'status' })
		expect(engine.render({ rows: 18, cols: 100 }).lines.join('\n')).toContain('Workbook Health')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		await engine.dispatch({ kind: 'command', command: 'info' })
		const frame = engine.render({ rows: 18, cols: 100 }).lines.join('\n')
		expect(frame).toContain('Info')
		expect(frame).toContain('Workbook info')
	})

	test('F10 keytips route to File and ribbon command search', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 100 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'F10' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'h' })
		expect(engine.state().mode).toBe('command')
		expect(engine.state().commandBuffer).toBe('home.')
		expect(stripAnsi(engine.render({ rows: 18, cols: 100 }).lines.join('\n'))).toContain('Copy')

		await engine.dispatch({ kind: 'key', key: 'Escape' })
		await engine.dispatch({ kind: 'key', key: 'F10' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'f' })
		expect(engine.state().workspace.fileHub.visible).toBe(true)
		expect(engine.state().workspace.focusedRegion).toBe('fileHub')
	})

	test('command palette Enter runs the highlighted fuzzy match', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 100 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: ':' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'form' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		expect(engine.state().activeDialog?.id).toBe('format-cells')
	})

	test('save-copy writes another file without rebasing the active workbook path', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-'))
		try {
			const original = join(dir, 'original.xlsx')
			const copy = join(dir, 'copy.xlsx')
			const workbook = AscendWorkbook.create()
			await workbook.save(original)
			const engine = await WorkbookTuiEngine.create({
				size: { rows: 18, cols: 100 },
				path: original,
			})
			await engine.dispatch({ kind: 'key', key: 'text', text: 'changed' })
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			const result = await engine.dispatch({ kind: 'command', command: `save-copy ${copy}` })
			expect(result.handled).toBe(true)
			const doc = engine.state().workspace.documents[0]
			expect(doc?.path).toBe(original)
			expect(engine.state().message).toContain('Saved copy')
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('registered dialog commands open terminal dialog surfaces without JSON', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		const result = await engine.dispatch({ kind: 'command', command: 'format' })
		expect(result.handled).toBe(true)
		expect(engine.state().activeDialog?.id).toBe('format-cells')
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('Format Cells')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		expect(engine.state().activeDialog).toBeUndefined()
	})

	test('editable dialog fields apply through the operation path', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'command', command: 'format' })

		for (const text of '0.00') await engine.dispatch({ kind: 'key', key: 'text', text })
		await engine.dispatch({ kind: 'key', key: 'Tab' })
		await engine.dispatch({ kind: 'key', key: 'text', text: ' ' })

		const dialog = engine.state().activeDialog
		expect(dialog?.fields[0]?.value).toBe('0.00')
		expect(dialog?.fields[1]?.value).toBe('true')
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('0.00')

		const result = await engine.dispatch({ kind: 'key', key: 'Enter' })
		expect(result.handled).toBe(true)
		expect(engine.state().activeDialog).toBeUndefined()
		expect(engine.state().message).toContain('Committed 1 operation')
	})

	test('default Format Cells dialog does not dirty the workbook', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-'))
		try {
			const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({
				kind: 'command',
				command: `save-as ${join(dir, 'format-noop.xlsx')}`,
			})
			expect(engine.state().dirty).toBe(false)
			await engine.dispatch({ kind: 'command', command: 'format' })
			const result = await engine.dispatch({ kind: 'key', key: 'Enter' })
			expect(result.handled).toBe(true)
			expect(engine.state().message).toBe('No changes.')
			expect(engine.state().dirty).toBe(false)
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('editable dialog select fields cycle with arrow keys and space', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'command', command: 'format' })

		await engine.dispatch({ kind: 'key', key: 'Tab' })
		await engine.dispatch({ kind: 'key', key: 'Tab' })
		await engine.dispatch({ kind: 'key', key: 'Tab' })
		expect(engine.state().activeDialog?.fields[3]?.name).toBe('horizontal')

		await engine.dispatch({ kind: 'key', key: 'text', text: ' ' })
		expect(engine.state().activeDialog?.fields[3]?.value).toBe('general')
		await engine.dispatch({ kind: 'key', key: 'ArrowRight' })
		expect(engine.state().activeDialog?.fields[3]?.value).toBe('left')
		await engine.dispatch({ kind: 'key', key: 'ArrowLeft' })
		expect(engine.state().activeDialog?.fields[3]?.value).toBe('general')
	})

	test('Excel shortcut keys route to native TUI actions', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+1' })
		expect(engine.state().activeDialog?.id).toBe('format-cells')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Shift+L' })
		expect(engine.state().message).toContain('Committed 1 operation')
		await engine.dispatch({ kind: 'key', key: 'Alt+ArrowDown' })
		expect(engine.state().message).toContain('Committed 1 operation')
		await engine.dispatch({ kind: 'key', key: 'F5' })
		expect(engine.state().mode).toBe('command')
		expect(engine.state().commandBuffer).toBe('goto ')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '/' })
		expect(engine.state().activeDialog?.id).toBe('find-replace')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		await engine.dispatch({ kind: 'key', key: 'F9' })
		expect(engine.state().message).toContain('Recalculation queued')
	})

	test('typed command fallbacks route to the same native actions as shortcuts', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '1' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '2' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'autosum' })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('3')

		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		const copy = await engine.dispatch({ kind: 'command', command: 'copy' })
		expect(copy.handled).toBe(true)
		await engine.dispatch({ kind: 'command', command: 'goto B1' })
		const paste = await engine.dispatch({ kind: 'command', command: 'paste' })
		expect(paste.handled).toBe(true)
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('1')

		await engine.dispatch({ kind: 'command', command: 'goto B2' })
		const freeze = await engine.dispatch({ kind: 'command', command: 'freeze' })
		expect(freeze.handled).toBe(true)
		expect(engine.state().message).toContain('Frozen panes')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.state().message).toContain('Undid view.freeze')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Y' })
		expect(engine.state().message).toContain('Redid view.freeze')
	})

	test('table and comment commands apply native workbook operations', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'Name\tValue\nA\t1\nB\t2' })

		const table = await engine.dispatch({
			kind: 'command',
			command: 'table create {"ref":"A1:B3","name":"Revenue","hasHeaders":true}',
		})
		expect(table.handled).toBe(true)
		expect(engine.state().message).toContain('Committed 1 operation')

		await engine.dispatch({ kind: 'command', command: 'goto B2' })
		const comment = await engine.dispatch({
			kind: 'command',
			command: 'comment {"text":"Review this","author":"Ada"}',
		})
		expect(comment.handled).toBe(true)
		expect(engine.state().message).toContain('Committed 1 operation')
	})

	test('undo reverses comment edits', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({
			kind: 'command',
			command: 'comment {"ref":"A1","text":"Review this","author":"Ada"}',
		})
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.state().message).toContain('Undid review.comment')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Y' })
		expect(engine.state().message).toContain('Redid review.comment')
	})

	test('undo reverses table creation enough to allow recreating the same table', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'Name\tValue\nA\t1\nB\t2' })
		await engine.dispatch({
			kind: 'command',
			command: 'table create {"ref":"A1:B3","name":"Revenue","hasHeaders":true}',
		})
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.state().message).toContain('Undid insert.table')
		const recreate = await engine.dispatch({
			kind: 'command',
			command: 'table create {"ref":"A1:B3","name":"Revenue","hasHeaders":true}',
		})
		expect(recreate.handled).toBe(true)
		expect(engine.state().message).toContain('Committed 1 operation')
	})

	test('undo and redo cover filter and print metadata commands', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'command', command: 'filter {"range":"A1:C10"}' })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.state().message).toContain('Undid data.filter')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Y' })
		expect(engine.state().message).toContain('Redid data.filter')

		await engine.dispatch({
			kind: 'command',
			command: 'print {"range":"A1:D20","orientation":"landscape","fitToWidth":1,"fitToHeight":0}',
		})
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.state().message).toContain('Undid file.printPreview')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Y' })
		expect(engine.state().message).toContain('Redid file.printPreview')
	})

	test('undo and redo cover validation and conditional formatting dialogs', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({
			kind: 'command',
			command:
				'validate {"range":"A1:A10","rule":{"type":"whole","operator":"between","formula1":"1","formula2":"10"}}',
		})
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.state().message).toContain('Undid data.validation')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Y' })
		expect(engine.state().message).toContain('Redid data.validation')

		await engine.dispatch({
			kind: 'command',
			command:
				'conditional-format {"range":"A1:A10","rule":{"type":"expression","formula":"A1>5","priority":1}}',
		})
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.state().message).toContain('Undid data.conditionalFormatting')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Y' })
		expect(engine.state().message).toContain('Redid data.conditionalFormatting')
	})

	test('table and comment shortcuts open editable dialogs', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+T' })
		expect(engine.state().activeDialog?.id).toBe('create-table')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		await engine.dispatch({ kind: 'key', key: 'Shift+F2' })
		expect(engine.state().activeDialog?.id).toBe('comment')
	})

	test('chart and pivot commands open object workflow dialogs', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'command', command: 'chart' })
		expect(engine.state().activeDialog?.id).toBe('chart-wizard')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		await engine.dispatch({ kind: 'command', command: 'pivot' })
		expect(engine.state().activeDialog?.id).toBe('pivot-fields')
	})

	test('object inspector lists workbook visual and pivot inventory', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		const result = await engine.dispatch({ kind: 'command', command: 'objects' })
		expect(result.handled).toBe(true)
		expect(engine.state().workspace.focusedRegion).toBe('inspector')
		expect(engine.state().message).toContain('Objects:')
		const frame = engine.render({ rows: 18, cols: 90 }).lines.join('\n')
		expect(frame).toContain('Object Inspector')
		expect(frame).toContain('No charts or pivot tables')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		expect(engine.state().workspace.focusedRegion).toBe('grid')
		expect(engine.state().inspectorLines).toHaveLength(0)
	})

	test('formula trace inspector lists precedents and dependents for the active cell', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 100 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '2' })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'goto B1' })
		for (const text of '=A1*2') await engine.dispatch({ kind: 'key', key: 'text', text })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		const sourceTrace = await engine.dispatch({ kind: 'command', command: 'trace precedents' })
		expect(sourceTrace.handled).toBe(true)
		expect(engine.state().workspace.focusedRegion).toBe('inspector')
		expect(engine.state().message).toContain('dependents')
		expect(engine.render({ rows: 18, cols: 100 }).lines.join('\n')).toContain('Sheet1!B1')
		await engine.dispatch({ kind: 'key', key: 'Escape' })
		expect(engine.state().workspace.focusedRegion).toBe('grid')

		await engine.dispatch({ kind: 'command', command: 'goto B1' })
		const formulaTrace = await engine.dispatch({
			kind: 'command',
			command: 'trace precedents {"maxDepth":1}',
		})
		expect(formulaTrace.handled).toBe(true)
		const frame = engine.render({ rows: 18, cols: 100 }).lines.join('\n')
		expect(frame).toContain('Trace Sheet1!B1')
		expect(frame).toContain('Formula =A1*2')
		expect(frame).toContain('Sheet1!A1')
	})

	test('print preview command applies print area and page setup operations', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		const result = await engine.dispatch({
			kind: 'command',
			command: 'print {"range":"A1:D20","orientation":"landscape","fitToWidth":1,"fitToHeight":0}',
		})
		expect(result.handled).toBe(true)
		expect(engine.state().message).toContain('Committed 2 operations')
	})

	test('find command selects the first matching value in the requested range', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'Alpha\tBeta\nGamma\tDelta' })
		const result = await engine.dispatch({
			kind: 'command',
			command: 'find {"range":"A1:B2","findText":"delta","action":"find","lookIn":"values"}',
		})
		expect(result.handled).toBe(true)
		expect(engine.state().selection.active).toEqual({ row: 1, col: 1 })
		expect(engine.state().message).toContain('B2')
	})

	test('replace all updates values through journaled operations', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'North\tNorthwest\nSouth\tnorth' })
		const result = await engine.dispatch({
			kind: 'command',
			command:
				'replace {"range":"A1:B2","findText":"North","replaceText":"East","action":"replaceAll","lookIn":"values"}',
		})
		expect(result.handled).toBe(true)
		expect(engine.state().message).toContain('Replaced 3 matches')
		const frame = engine.render({ rows: 18, cols: 90 }).lines.join('\n')
		expect(frame).toContain('East')
		expect(frame).toContain('Eastwest')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.state().message).toContain('Undid home.findReplace')
		await engine.dispatch({ kind: 'command', command: 'goto B1' })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('Northwest')
	})

	test('find-replace can rewrite formulas without flattening them to values', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '1\n2\n3' })
		await engine.dispatch({ kind: 'command', command: 'goto B1' })
		for (const text of '=A1+A2') await engine.dispatch({ kind: 'key', key: 'text', text })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		const result = await engine.dispatch({
			kind: 'command',
			command:
				'replace {"range":"B1","findText":"A2","replaceText":"A3","action":"replace","lookIn":"formulas"}',
		})
		expect(result.handled).toBe(true)
		await engine.dispatch({ kind: 'command', command: 'show formulas' })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('=A1+A3')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('=A1+A2')
	})

	test('find-replace shortcut opens an editable dialog', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+F' })
		expect(engine.state().activeDialog?.id).toBe('find-replace')
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('Find and Replace')
	})

	test('show formulas toggles grid visualization from values to formula text', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		for (const text of '=1+1') await engine.dispatch({ kind: 'key', key: 'text', text })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('2')
		await engine.dispatch({ kind: 'command', command: 'show formulas' })
		const sourceFrame = engine.render({ rows: 18, cols: 90 }).lines.join('\n')
		expect(engine.state().showFormulas).toBe(true)
		expect(sourceFrame).toContain('=1+1')
	})

	test('undo and redo replay journaled cell edits', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		for (const text of '42') await engine.dispatch({ kind: 'key', key: 'text', text })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('42')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		expect(engine.state().message).toContain('Undid home.edit')
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).not.toContain('42')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Y' })
		expect(engine.state().message).toContain('Redid home.edit')
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('42')
	})

	test('undo and redo restore exact cell formatting preimages', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-'))
		try {
			const formattedPath = join(dir, 'formatted.xlsx')
			const undonePath = join(dir, 'undone.xlsx')
			const redonePath = join(dir, 'redone.xlsx')
			const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'key', key: 'text', text: '42' })
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			await engine.dispatch({ kind: 'command', command: 'goto A1' })
			await engine.dispatch({
				kind: 'command',
				command: 'format {"italic":true}',
			})
			await engine.dispatch({
				kind: 'command',
				command: 'format {"bold":true,"numberFormat":"0.00"}',
			})
			await engine.dispatch({ kind: 'command', command: `save-as ${formattedPath}` })
			const formatted = await AscendWorkbook.open(formattedPath)
			expect(cellStyle(formatted, 'Sheet1', 'A1')?.font?.bold).toBe(true)
			expect(cellStyle(formatted, 'Sheet1', 'A1')?.font?.italic).toBe(true)
			expect(cellStyle(formatted, 'Sheet1', 'A1')?.numberFormat).toBe('0.00')

			await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
			await engine.dispatch({ kind: 'command', command: `save-as ${undonePath}` })
			const undone = await AscendWorkbook.open(undonePath)
			expect(cellStyle(undone, 'Sheet1', 'A1')?.font?.italic).toBe(true)
			expect(cellStyle(undone, 'Sheet1', 'A1')?.font?.bold).not.toBe(true)
			expect(cellStyle(undone, 'Sheet1', 'A1')?.numberFormat).not.toBe('0.00')

			await engine.dispatch({ kind: 'key', key: 'Ctrl+Y' })
			await engine.dispatch({ kind: 'command', command: `save-as ${redonePath}` })
			const redone = await AscendWorkbook.open(redonePath)
			expect(cellStyle(redone, 'Sheet1', 'A1')?.font?.bold).toBe(true)
			expect(cellStyle(redone, 'Sheet1', 'A1')?.font?.italic).toBe(true)
			expect(cellStyle(redone, 'Sheet1', 'A1')?.numberFormat).toBe('0.00')
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('cell edit undo preserves preexisting custom formatting', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'ascend-tui-'))
		try {
			const path = join(dir, 'styled-undo.xlsx')
			const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
			await engine.dispatch({ kind: 'command', command: 'new' })
			await engine.dispatch({ kind: 'key', key: 'text', text: '10' })
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			await engine.dispatch({ kind: 'command', command: 'goto A1' })
			await engine.dispatch({
				kind: 'command',
				command: 'format {"italic":true,"numberFormat":"0.0"}',
			})
			await engine.dispatch({ kind: 'key', key: 'text', text: '20' })
			await engine.dispatch({ kind: 'key', key: 'Enter' })
			await engine.dispatch({ kind: 'command', command: 'goto A1' })
			await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
			await engine.dispatch({ kind: 'command', command: `save-as ${path}` })
			const reopened = await AscendWorkbook.open(path)
			expect(reopened.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 10 })
			expect(cellStyle(reopened, 'Sheet1', 'A1')?.font?.italic).toBe(true)
			expect(cellStyle(reopened, 'Sheet1', 'A1')?.numberFormat).toBe('0.0')
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('bracketed TSV paste writes a grid instead of entering edit mode', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		const result = await engine.dispatch({
			kind: 'key',
			key: 'text',
			text: 'Name\tValue\nAlpha\t=1+1\nBeta\t3\n',
		})
		expect(result.handled).toBe(true)
		expect(engine.state().mode).toBe('ready')
		expect(engine.state().message).toContain('Pasted 6 cells')
		const valuesFrame = engine.render({ rows: 18, cols: 90 }).lines.join('\n')
		expect(valuesFrame).toContain('Alpha')
		expect(valuesFrame).toContain('Beta')
		expect(valuesFrame).toContain('3')
		await engine.dispatch({ kind: 'command', command: 'show formulas' })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('=1+1')
	})

	test('ragged TSV paste only clears cells that are actually pasted', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'command', command: 'goto B2' })
		for (const text of '=1+1') await engine.dispatch({ kind: 'key', key: 'text', text })
		await engine.dispatch({ kind: 'key', key: 'Enter' })
		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'A\tB\nC' })
		await engine.dispatch({ kind: 'command', command: 'goto B2' })
		await engine.dispatch({ kind: 'command', command: 'show formulas' })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('=1+1')
	})

	test('fill shortcuts are undoable through copyRange preimages', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: '1\n9' })
		await engine.dispatch({ kind: 'command', command: 'goto A1:A2' })
		await engine.dispatch({ kind: 'key', key: 'Ctrl+D' })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('1')
		await engine.dispatch({ kind: 'key', key: 'Ctrl+Z' })
		await engine.dispatch({ kind: 'command', command: 'goto A2' })
		expect(engine.render({ rows: 18, cols: 90 }).lines.join('\n')).toContain('9')
	})

	test('Ctrl+C and Ctrl+V round-trip selected cells as TSV', async () => {
		const engine = await WorkbookTuiEngine.create({ size: { rows: 18, cols: 90 } })
		await engine.dispatch({ kind: 'command', command: 'new' })
		await engine.dispatch({ kind: 'key', key: 'text', text: 'A\tB' })
		await engine.dispatch({ kind: 'command', command: 'goto A1' })
		await engine.dispatch({ kind: 'key', key: 'Shift+ArrowRight' })

		const copy = await engine.dispatch({ kind: 'key', key: 'Ctrl+C' })
		expect(copy.handled).toBe(true)
		expect(engine.state().message).toContain('Copied 2 cells')

		await engine.dispatch({ kind: 'command', command: 'goto A3' })
		const paste = await engine.dispatch({ kind: 'key', key: 'Ctrl+V' })
		expect(paste.handled).toBe(true)
		const frame = engine.render({ rows: 18, cols: 90 }).lines.join('\n')
		expect(frame).toContain('A')
		expect(frame).toContain('B')
		expect(engine.state().message).toContain('Pasted 2 cells')
	})
})

function cellStyle(
	workbook: AscendWorkbook,
	sheetName: string,
	ref: string,
): CellStyle | undefined {
	const parsed = parseRange(ref)
	const model = workbook.getWorkbookModel()
	const cell = model.getSheet(sheetName)?.cells.get(parsed.start.row, parsed.start.col)
	if (!cell) return undefined
	return model.styles.get(cell.styleId)
}
