import { describe, expect, test } from 'bun:test'
import { fitAnsi, visibleLength } from './ansi-text.ts'
import { diffFrames } from './diff.ts'
import { createFrame } from './frame.ts'
import { OpenTuiRenderer } from './opentui-renderer.ts'
import { runRendererBakeoff } from './renderer-bakeoff.ts'

describe('render primitives', () => {
	test('diffFrames emits a full redraw for the first frame', () => {
		const frame = createFrame({ rows: 2, cols: 8 }, ['alpha', 'beta'])
		const patch = diffFrames(null, frame)
		expect(patch.fullRedraw).toBe(true)
		expect(patch.lines).toHaveLength(2)
	})

	test('diffFrames emits only changed rows for same-size frames', () => {
		const previous = createFrame({ rows: 3, cols: 8 }, ['one', 'two', 'three'])
		const next = createFrame({ rows: 3, cols: 8 }, ['one', 'TWO', 'three'])
		const patch = diffFrames(previous, next)
		expect(patch.fullRedraw).toBe(false)
		expect(patch.lines).toEqual([{ row: 2, text: 'TWO     ' }])
	})

	test('ANSI fitting preserves escape sequences while matching visible width', () => {
		const styled = '\x1b[31mabcdef\x1b[0m'
		const fitted = fitAnsi(styled, 4)
		expect(fitted).toContain('\x1b[31m')
		expect(fitted).toContain('\x1b[')
		expect(visibleLength(fitted)).toBe(4)
	})

	test('renderer bakeoff records ANSI and OpenTUI candidates', async () => {
		const frames = [
			createFrame({ rows: 2, cols: 8 }, ['alpha', 'beta']),
			createFrame({ rows: 2, cols: 8 }, ['alpha', 'BETA']),
		]
		const results = await runRendererBakeoff(frames)
		expect(results.find((result) => result.candidate === 'ansi')).toMatchObject({
			status: 'passed',
			frames: 2,
		})
		expect(results.find((result) => result.candidate === 'opentui')).toMatchObject({
			status: 'passed',
			frames: 2,
		})
	})

	test('OpenTUI renderer can draw a logical frame in test mode', async () => {
		const renderer = new OpenTuiRenderer()
		await renderer.init({
			isTty: false,
			color: 'truecolor',
			unicode: true,
			mouse: false,
			bracketedPaste: false,
			hyperlinks: false,
			graphics: 'off',
			keyboardProtocol: 'legacy',
			profile: 'legacy',
		})
		const stats = await renderer.draw(createFrame({ rows: 2, cols: 8 }, ['alpha', 'beta']))
		expect(stats.bytesOut).toBeGreaterThan(0)
		await renderer.shutdown()
	})
})
