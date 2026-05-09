import type { TerminalCapabilities } from '../runtime/types.ts'
import { defaultShortcutCheck } from './shortcut-check.ts'

export function describeTerminalProfile(capabilities: TerminalCapabilities): string {
	return `${capabilities.profile} / ${capabilities.color} / ${capabilities.graphics}`
}

export interface TerminalCalibrationReport {
	readonly profile: string
	readonly lines: readonly string[]
	readonly warnings: readonly string[]
}

export function buildTerminalCalibrationReport(
	capabilities: TerminalCapabilities,
): TerminalCalibrationReport {
	const shortcutCheck = defaultShortcutCheck(capabilities)
	const warnings = calibrationWarnings(capabilities)
	return {
		profile: describeTerminalProfile(capabilities),
		warnings,
		lines: [
			'Terminal Calibration',
			`Profile: ${describeTerminalProfile(capabilities)}`,
			`TTY: ${capabilities.isTty ? 'yes' : 'no'}  Keyboard: ${capabilities.keyboardProtocol}`,
			`Mouse: ${capabilities.mouse ? 'yes' : 'no'}  Paste: ${capabilities.bracketedPaste ? 'bracketed' : 'plain'}`,
			`Color: ${capabilities.color}  Unicode: ${capabilities.unicode ? 'yes' : 'no'}  Graphics: ${capabilities.graphics}`,
			`Links: ${capabilities.hyperlinks ? 'OSC 8' : 'off'}  Density: compact grid`,
			'Keyboard profiles: PC Ctrl+Arrow/Home/End and Mac Cmd+Arrow/Fn fallbacks map to the same spreadsheet actions.',
			'Mac fallbacks: Option+Left/Right switches sheets, Ctrl+G opens Go To, Ctrl+P opens command search.',
			'ASCII mode: grid markers use [active], {selected}, F filter, A1/D1 sort, c comment, d validation, ! invalid, RO protected.',
			...shortcutCheck.notes.map((note) => `Shortcut: ${note}`),
			...(warnings.length > 0 ? warnings.map((warning) => `Warning: ${warning}`) : ['Ready.']),
		],
	}
}

function calibrationWarnings(capabilities: TerminalCapabilities): readonly string[] {
	const warnings: string[] = []
	if (!capabilities.isTty) warnings.push('Non-TTY mode renders headless frames only.')
	if (capabilities.keyboardProtocol === 'legacy') {
		warnings.push('Some Excel key chords need visible command or palette fallbacks.')
	}
	if (!capabilities.mouse) warnings.push('Mouse selection is unavailable in this terminal.')
	if (capabilities.color === 'none') warnings.push('Color styling is disabled.')
	if (!capabilities.bracketedPaste)
		warnings.push('Large paste input may arrive as ordinary text chunks.')
	return warnings
}
