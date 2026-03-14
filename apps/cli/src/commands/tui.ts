import type { CellValue } from '@ascend/schema'
import { EMPTY } from '@ascend/schema'
import type { CompactCellInfo } from '@ascend/sdk'
import { AscendWorkbook, formatDisplayCellValue, indexToColumn, ops, parseA1 } from '@ascend/sdk'

export const usage = `Usage: ascend tui <file> [flags]

  Open a workbook in an interactive terminal spreadsheet.

Arguments:
  <file>              Path to the workbook file

Flags:
  --sheet <name>      Start on a specific sheet

Navigation:
  Arrow keys / hjkl   Move cursor
  Tab / Shift+Tab     Move right / left
  PgUp / PgDn         Scroll page
  g                   Go to cell (enter address)

Editing:
  Enter / F2          Edit cell
  Escape              Cancel edit
  Delete              Clear cell

Other:
  :                   Command mode (:w :q :wq :sheet Name)
  q                   Quit
`

const CSI = '\x1b['
const BOX_V = '│'

const C = {
	reset: `${CSI}0m`,
	bold: `${CSI}1m`,
	hdrBg: `${CSI}48;5;236m`,
	hdrFg: `${CSI}38;5;252m`,
	actBg: `${CSI}48;5;17m`,
	actFg: `${CSI}38;5;15m`,
	gridFg: `${CSI}38;5;240m`,
	errFg: `${CSI}38;5;196m`,
	numFg: `${CSI}38;5;114m`,
	strFg: `${CSI}38;5;252m`,
	boolFg: `${CSI}38;5;215m`,
	emptyFg: `${CSI}38;5;242m`,
	stBg: `${CSI}48;5;236m`,
	stFg: `${CSI}38;5;252m`,
	fbBg: `${CSI}48;5;234m`,
	fbFg: `${CSI}38;5;252m`,
	tabAct: `${CSI}48;5;24m${CSI}38;5;15m`,
	tabOff: `${CSI}48;5;236m${CSI}38;5;248m`,
} as const

interface State {
	wb: AscendWorkbook
	filePath: string
	si: number
	cr: number
	cc: number
	sr: number
	sc: number
	vr: number
	vc: number
	cw: number[]
	editing: boolean
	ebuf: string
	ecur: number
	cmdMode: boolean
	cbuf: string
	msg: string
	dirty: boolean
	viewCache: Map<string, CompactCellInfo>
	viewCacheDirty: boolean
}

function termSize(): [number, number] {
	return [process.stdout.rows ?? 24, process.stdout.columns ?? 80]
}

function colorOf(v: CellValue): string {
	switch (v.kind) {
		case 'number':
		case 'date':
			return C.numFg
		case 'error':
			return C.errFg
		case 'boolean':
			return C.boolFg
		case 'empty':
			return C.emptyFg
		default:
			return C.strFg
	}
}

function trunc(s: string, w: number): string {
	return s.length <= w ? s : `${s.slice(0, w - 1)}…`
}

function pad(text: string, w: number, right: boolean): string {
	const t = trunc(text, w)
	const p = w - t.length
	return p <= 0 ? t : right ? ' '.repeat(p) + t : t + ' '.repeat(p)
}

function sheetName(s: State): string {
	return s.wb.inspect().sheets[s.si]?.name ?? 'Sheet1'
}

function computeColWidths(s: State): number[] {
	const [, cols] = termSize()
	const rhw = 6
	const avail = cols - rhw - 1
	const defaultW = 10
	const ws: number[] = []
	let used = 0
	for (let c = s.sc; used < avail; c++) {
		let w = defaultW
		if (used + w + 1 > avail) {
			w = avail - used - 1
			if (w < 4) break
		}
		ws.push(w)
		used += w + 1
		if (c > s.sc + 50) break
	}
	return ws
}

function refreshViewCache(s: State): void {
	if (!s.viewCacheDirty && s.viewCache.size > 0) return
	s.viewCache.clear()
	const name = sheetName(s)
	const startCol = indexToColumn(s.sc)
	const endCol = indexToColumn(s.sc + s.vc - 1)
	const startRow = s.sr + 1
	const endRow = s.sr + s.vr
	const range = `${startCol}${startRow}:${endCol}${endRow}`
	const data = s.wb.readRangeCompact(name, range, { includeRefs: true })
	if (data) {
		for (const cell of data.cells) {
			if (cell.ref) s.viewCache.set(cell.ref, cell)
		}
	}
	s.viewCacheDirty = false
}

function getCachedValue(s: State, row: number, col: number): CellValue {
	const ref = `${indexToColumn(col)}${row + 1}`
	const cached = s.viewCache.get(ref)
	return cached?.value ?? EMPTY
}

function renderFormulaBar(s: State): string {
	const [, cols] = termSize()
	const addr = `${indexToColumn(s.cc)}${s.cr + 1}`
	let content = ''
	if (s.editing) {
		content = s.ebuf
	} else {
		const ref = `${sheetName(s)}!${indexToColumn(s.cc)}${s.cr + 1}`
		const info = s.wb.formula(ref)
		if (info?.formula) {
			content = `=${info.formula}`
		} else {
			content = formatDisplayCellValue(getCachedValue(s, s.cr, s.cc))
		}
	}
	const aw = 8
	const cw = cols - aw - 3
	return `${C.fbBg}${C.bold}${C.fbFg} ${addr.padEnd(aw)}${C.reset}${C.gridFg}${BOX_V}${C.reset}${C.fbBg}${C.fbFg} ${trunc(content, cw).padEnd(cw)} ${C.reset}`
}

function renderGrid(s: State): string {
	const [rows] = termSize()
	s.vr = rows - 3
	s.cw = computeColWidths(s)
	s.vc = s.cw.length
	const rhw = 6
	const lines: string[] = []

	refreshViewCache(s)

	const hdr: string[] = [`${C.hdrBg}${C.hdrFg}${' '.repeat(rhw)}${C.reset}`]
	for (let ci = 0; ci < s.cw.length; ci++) {
		const col = s.sc + ci
		const w = s.cw[ci] ?? 10
		const lbl = indexToColumn(col)
		const isAct = col === s.cc
		const bg = isAct ? C.actBg : C.hdrBg
		const fg = isAct ? C.actFg : C.hdrFg
		hdr.push(`${C.gridFg}${BOX_V}${C.reset}${bg}${fg}${C.bold}${pad(lbl, w, false)}${C.reset}`)
	}
	lines.push(hdr.join(''))

	for (let ri = 0; ri < s.vr; ri++) {
		const row = s.sr + ri
		const rStr = String(row + 1)
		const isActRow = row === s.cr
		const rb = isActRow ? C.actBg : C.hdrBg
		const rf = isActRow ? C.actFg : C.hdrFg
		const parts: string[] = [`${rb}${rf}${C.bold}${rStr.padStart(rhw)}${C.reset}`]

		for (let ci = 0; ci < s.cw.length; ci++) {
			const col = s.sc + ci
			const w = s.cw[ci] ?? 10
			const isAct = row === s.cr && col === s.cc
			const value = getCachedValue(s, row, col)
			const txt = formatDisplayCellValue(value)
			const rAlign = value.kind === 'number' || value.kind === 'date'

			if (isAct) {
				const display = s.editing ? pad(s.ebuf, w, false) : pad(txt, w, rAlign)
				parts.push(`${C.gridFg}${BOX_V}${C.reset}${C.actBg}${C.actFg}${C.bold}${display}${C.reset}`)
			} else {
				parts.push(`${C.gridFg}${BOX_V}${C.reset}${colorOf(value)}${pad(txt, w, rAlign)}${C.reset}`)
			}
		}
		lines.push(parts.join(''))
	}
	return lines.join('\n')
}

function renderTabs(s: State): string {
	const info = s.wb.inspect()
	return info.sheets
		.map((sh, i) => `${i === s.si ? C.tabAct : C.tabOff} ${trunc(sh.name, 15)} ${C.reset}`)
		.join(' ')
}

function renderStatus(s: State): string {
	const [, cols] = termSize()
	let left = ''
	if (s.cmdMode) left = `:${s.cbuf}`
	else if (s.msg) left = s.msg
	else left = `${sheetName(s)} | ${indexToColumn(s.cc)}${s.cr + 1}${s.dirty ? ' [modified]' : ''}`
	const mode = s.editing ? ' EDIT ' : s.cmdMode ? ' CMD ' : ' NORMAL '
	const p = cols - left.length - mode.length - 2
	return `${C.stBg}${C.stFg} ${left}${' '.repeat(Math.max(0, p))}${mode} ${C.reset}`
}

function draw(s: State): void {
	const [rows] = termSize()
	const buf: string[] = [`${CSI}H`]
	buf.push(renderFormulaBar(s))
	buf.push(`${CSI}2;1H`)
	buf.push(renderGrid(s))
	buf.push(`${CSI}${rows - 1};1H`)
	buf.push(renderTabs(s))
	buf.push(`${CSI}${rows};1H`)
	buf.push(renderStatus(s))
	process.stdout.write(buf.join(''))
}

function ensureVisible(s: State): void {
	let scrolled = false
	if (s.cr < s.sr) {
		s.sr = s.cr
		scrolled = true
	}
	if (s.cr >= s.sr + s.vr) {
		s.sr = s.cr - s.vr + 1
		scrolled = true
	}
	if (s.cc < s.sc) {
		s.sc = s.cc
		scrolled = true
	}
	if (s.cc >= s.sc + s.vc) {
		s.sc = s.cc - s.vc + 1
		scrolled = true
	}
	if (scrolled) s.viewCacheDirty = true
}

function beginEdit(s: State): void {
	const ref = `${sheetName(s)}!${indexToColumn(s.cc)}${s.cr + 1}`
	const info = s.wb.formula(ref)
	if (info?.formula) {
		s.ebuf = `=${info.formula}`
	} else {
		const v = getCachedValue(s, s.cr, s.cc)
		s.ebuf = v.kind === 'empty' ? '' : formatDisplayCellValue(v)
	}
	s.editing = true
	s.ecur = s.ebuf.length
}

function commitEdit(s: State): void {
	if (!s.editing) return
	s.editing = false
	const text = s.ebuf.trim()
	const sheet = sheetName(s)
	const ref = `${indexToColumn(s.cc)}${s.cr + 1}`
	if (text === '') {
		s.wb.applyAndRecalc([{ op: 'clearRange', sheet, range: ref, what: 'all' }])
	} else if (text.startsWith('=')) {
		s.wb.applyAndRecalc([ops.setFormula(sheet, ref, text.slice(1))])
	} else {
		s.wb.applyAndRecalc([ops.setCell(sheet, ref, parseInput(text))])
	}
	s.dirty = true
	s.viewCacheDirty = true
	s.msg = ''
}

function parseInput(t: string): string | number | boolean {
	if (t.toUpperCase() === 'TRUE') return true
	if (t.toUpperCase() === 'FALSE') return false
	const n = Number(t)
	if (!Number.isNaN(n) && t.trim() !== '') return n
	return t
}

function handleCmd(s: State): void {
	const cmd = s.cbuf.trim()
	s.cmdMode = false
	s.cbuf = ''

	if (cmd === 'q' || cmd === 'quit') {
		teardown()
		process.exit(0)
	}
	if (cmd === 'w' || cmd === 'write' || cmd === 'save') {
		s.wb.save(s.filePath).then(() => {
			s.dirty = false
			s.msg = `Saved ${s.filePath}`
			draw(s)
		})
		return
	}
	if (cmd === 'wq') {
		s.wb.save(s.filePath).then(() => {
			teardown()
			process.exit(0)
		})
		return
	}
	if (cmd.startsWith('sheet ')) {
		const name = cmd.slice(6).trim()
		const info = s.wb.inspect()
		const idx = info.sheets.findIndex((sh) => sh.name.toLowerCase() === name.toLowerCase())
		if (idx >= 0) {
			s.si = idx
			s.cr = 0
			s.cc = 0
			s.sr = 0
			s.sc = 0
			s.viewCacheDirty = true
			s.msg = `Switched to ${info.sheets[idx]?.name}`
		} else {
			s.msg = `Sheet not found: ${name}`
		}
		return
	}
	s.msg = `Unknown command: ${cmd}`
}

function onNormal(s: State, key: Buffer): void {
	const str = key.toString('utf8')
	const b0 = key[0] ?? 0

	if (str === '\x1b[A' || str === 'k') {
		s.cr = Math.max(0, s.cr - 1)
	} else if (str === '\x1b[B' || str === 'j') {
		s.cr++
	} else if (str === '\x1b[C' || str === 'l' || str === '\t') {
		s.cc++
	} else if (str === '\x1b[D' || str === 'h' || str === '\x1b[Z') {
		s.cc = Math.max(0, s.cc - 1)
	} else if (str === '\x1b[5~') {
		s.cr = Math.max(0, s.cr - s.vr)
	} else if (str === '\x1b[6~') {
		s.cr += s.vr
	} else if (str === '\r') {
		beginEdit(s)
	} else if (str === 'q') {
		if (s.dirty) {
			s.msg = 'Unsaved changes. :wq to save+quit, :q to force quit.'
			return
		}
		teardown()
		process.exit(0)
	} else if (str === ':') {
		s.cmdMode = true
		s.cbuf = ''
	} else if (str === '\x7f' || str === '\x1b[3~') {
		const sheet = sheetName(s)
		const ref = `${indexToColumn(s.cc)}${s.cr + 1}`
		s.wb.applyAndRecalc([{ op: 'clearRange', sheet, range: ref, what: 'all' }])
		s.dirty = true
		s.viewCacheDirty = true
	} else if (str === 'g') {
		s.msg = 'Go to: (type address + Enter)'
		s.cmdMode = true
		s.cbuf = 'goto '
	} else if (b0 >= 0x20 && b0 < 0x7f && str.length === 1) {
		s.ebuf = str
		s.editing = true
		s.ecur = 1
	}
}

function onEdit(s: State, key: Buffer): void {
	const str = key.toString('utf8')
	if (str === '\r') {
		commitEdit(s)
		s.cr++
	} else if (str === '\t') {
		commitEdit(s)
		s.cc++
	} else if (str === '\x1b') {
		s.editing = false
		s.ebuf = ''
		s.msg = ''
	} else if (str === '\x7f' || str === '\b') {
		if (s.ecur > 0) {
			s.ebuf = s.ebuf.slice(0, s.ecur - 1) + s.ebuf.slice(s.ecur)
			s.ecur--
		}
	} else if (str === '\x1b[D') {
		s.ecur = Math.max(0, s.ecur - 1)
	} else if (str === '\x1b[C') {
		s.ecur = Math.min(s.ebuf.length, s.ecur + 1)
	} else if (str === '\x01') {
		s.ecur = 0
	} else if (str === '\x05') {
		s.ecur = s.ebuf.length
	} else if (str.length === 1 && str.charCodeAt(0) >= 0x20) {
		s.ebuf = s.ebuf.slice(0, s.ecur) + str + s.ebuf.slice(s.ecur)
		s.ecur++
	}
}

function onCmd(s: State, key: Buffer): void {
	const str = key.toString('utf8')
	if (str === '\r') {
		if (s.cbuf.startsWith('goto ')) {
			const addr = s.cbuf.slice(5).trim().toUpperCase()
			try {
				const parsed = parseA1(addr)
				s.cr = parsed.row
				s.cc = parsed.col
				s.msg = ''
			} catch {
				s.msg = `Invalid address: ${addr}`
			}
		} else {
			handleCmd(s)
		}
		s.cmdMode = false
		s.cbuf = ''
	} else if (str === '\x1b') {
		s.cmdMode = false
		s.cbuf = ''
		s.msg = ''
	} else if (str === '\x7f' || str === '\b') {
		s.cbuf = s.cbuf.slice(0, -1)
		if (s.cbuf === '') {
			s.cmdMode = false
			s.msg = ''
		}
	} else if (str.length === 1 && str.charCodeAt(0) >= 0x20) {
		s.cbuf += str
	}
}

function onMouse(s: State, seq: string): void {
	// biome-ignore lint/suspicious/noControlCharactersInRegex: ANSI escape sequence for mouse events
	const m = seq.match(/\x1b\[<(\d+);(\d+);(\d+)([mM])/)
	if (!m) return
	const btn = Number.parseInt(m[1] ?? '0', 10)
	const tr = Number.parseInt(m[3] ?? '0', 10) - 1
	const tc = Number.parseInt(m[2] ?? '0', 10) - 1

	if (btn === 64) {
		s.sr = Math.max(0, s.sr - 3)
		s.viewCacheDirty = true
		return
	}
	if (btn === 65) {
		s.sr += 3
		s.viewCacheDirty = true
		return
	}
	if (btn === 0) {
		const gr = tr - 1
		const rhw = 6
		if (gr < 0 || gr >= s.vr || tc < rhw) return
		let off = rhw
		for (let ci = 0; ci < s.cw.length; ci++) {
			const w = (s.cw[ci] ?? 10) + 1
			if (tc >= off && tc < off + w) {
				s.cr = s.sr + gr
				s.cc = s.sc + ci
				return
			}
			off += w
		}
	}
}

function teardown(): void {
	process.stdout.write(`${CSI}?25h${CSI}?1000l${CSI}?1003l${CSI}?1006l${CSI}2J${CSI}H`)
	if (process.stdin.isTTY) process.stdin.setRawMode(false)
}

export async function tuiCommand(args: string[], flags: Map<string, string>): Promise<number> {
	const file = args[0]
	if (!file) {
		console.error('Usage: ascend tui <file>')
		return 1
	}

	const wb = await AscendWorkbook.open(file)

	const reqSheet = flags.get('sheet')
	let si = 0
	if (reqSheet) {
		const info = wb.inspect()
		const idx = info.sheets.findIndex((sh) => sh.name.toLowerCase() === reqSheet.toLowerCase())
		if (idx >= 0) si = idx
	}

	const s: State = {
		wb,
		filePath: file,
		si,
		cr: 0,
		cc: 0,
		sr: 0,
		sc: 0,
		vr: 20,
		vc: 8,
		cw: [],
		editing: false,
		ebuf: '',
		ecur: 0,
		cmdMode: false,
		cbuf: '',
		msg: '',
		dirty: false,
		viewCache: new Map(),
		viewCacheDirty: true,
	}

	if (!process.stdin.isTTY) {
		console.error('tui requires an interactive terminal')
		return 1
	}

	process.stdin.setRawMode(true)
	process.stdin.resume()
	process.stdout.write(`${CSI}?25l${CSI}2J${CSI}H`)
	process.stdout.write(`${CSI}?1000h${CSI}?1006h`)

	draw(s)

	process.on('SIGINT', () => {
		teardown()
		process.exit(0)
	})
	process.on('SIGWINCH', () => {
		s.viewCacheDirty = true
		process.stdout.write(`${CSI}2J`)
		draw(s)
	})

	process.stdin.on('data', (data: Buffer) => {
		if (data.toString('utf8') === '\x03') {
			teardown()
			process.exit(0)
		}
		if (s.editing) onEdit(s, data)
		else if (s.cmdMode) onCmd(s, data)
		else {
			const str = data.toString('utf8')
			if (str.startsWith('\x1b[<')) onMouse(s, str)
			else onNormal(s, data)
		}
		ensureVisible(s)
		draw(s)
	})

	return new Promise<number>(() => {})
}
