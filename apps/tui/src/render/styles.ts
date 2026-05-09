const CSI = '\x1b['

export const ANSI = {
	reset: `${CSI}0m`,
	bold: `${CSI}1m`,
	inverse: `${CSI}7m`,
	hideCursor: `${CSI}?25l`,
	showCursor: `${CSI}?25h`,
	clear: `${CSI}2J`,
	home: `${CSI}H`,
	altScreen: `${CSI}?1049h`,
	mainScreen: `${CSI}?1049l`,
	enableMouse: `${CSI}?1000h${CSI}?1002h${CSI}?1006h`,
	disableMouse: `${CSI}?1000l${CSI}?1002l${CSI}?1006l`,
	enableBracketedPaste: `${CSI}?2004h`,
	disableBracketedPaste: `${CSI}?2004l`,
} as const

export const THEME = {
	reset: ANSI.reset,
	bold: ANSI.bold,
	canvas: `${CSI}48;5;234m${CSI}38;5;252m`,
	ribbon: `${CSI}48;5;236m${CSI}38;5;252m`,
	header: `${CSI}48;5;236m${CSI}38;5;252m`,
	active: `${CSI}48;5;23m${CSI}38;5;15m`,
	selection: `${CSI}48;5;24m${CSI}38;5;15m`,
	grid: `${CSI}38;5;240m`,
	status: `${CSI}48;5;236m${CSI}38;5;252m`,
	formula: `${CSI}48;5;234m${CSI}38;5;252m`,
	number: `${CSI}38;5;114m`,
	text: `${CSI}38;5;252m`,
	bool: `${CSI}38;5;215m`,
	error: `${CSI}38;5;196m`,
	muted: `${CSI}38;5;245m`,
	warn: `${CSI}38;5;220m`,
	success: `${CSI}38;5;78m`,
} as const

export function moveTo(row: number, col: number): string {
	return `${CSI}${row};${col}H`
}

export function colorForKind(kind: string): string {
	switch (kind) {
		case 'number':
		case 'date':
			return THEME.number
		case 'boolean':
			return THEME.bool
		case 'error':
			return THEME.error
		case 'empty':
			return THEME.muted
		default:
			return THEME.text
	}
}
