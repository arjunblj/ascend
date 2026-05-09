import type { TerminalCapabilities } from '../runtime/types.ts'

export interface ShortcutCheckResult {
	readonly profile: string
	readonly notes: readonly string[]
}

export function defaultShortcutCheck(capabilities?: TerminalCapabilities): ShortcutCheckResult {
	const profile = capabilities?.profile ?? 'modern'
	const notes = [
		'Most Excel navigation keys are enabled.',
		'Terminal-reserved keys keep visible command fallbacks.',
	]
	if (capabilities?.keyboardProtocol === 'kitty' || capabilities?.keyboardProtocol === 'csi-u') {
		notes.push('Enhanced keyboard protocol improves modified key fidelity.')
	}
	if (capabilities && !capabilities.bracketedPaste) {
		notes.push('Paste uses plain text fallback instead of bracketed paste.')
	}
	return {
		profile,
		notes,
	}
}
