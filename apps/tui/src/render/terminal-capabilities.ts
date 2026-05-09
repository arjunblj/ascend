import type { TerminalCapabilities } from '../runtime/types.ts'

export function detectTerminalCapabilities(): TerminalCapabilities {
	const env = process.env
	const term = env.TERM ?? ''
	const program = env.TERM_PROGRAM ?? ''
	const isTty = Boolean(process.stdout.isTTY)
	const noColor = 'NO_COLOR' in env || term === 'dumb'
	const truecolor = env.COLORTERM === 'truecolor' || env.COLORTERM === '24bit'
	const tmux = Boolean(env.TMUX) || term.includes('screen')
	const profile = tmux
		? 'ssh-tmux-limited'
		: process.platform === 'darwin' && program === 'Apple_Terminal'
			? 'mac-terminal'
			: process.platform === 'win32'
				? 'windows-terminal'
				: truecolor
					? 'modern'
					: 'legacy'
	const graphics =
		tmux || !isTty
			? 'unicode'
			: program === 'iTerm.app'
				? 'iterm'
				: term.includes('kitty') || program.includes('WezTerm') || program.includes('Ghostty')
					? 'kitty'
					: 'unicode'
	return {
		isTty,
		color: noColor ? 'none' : truecolor ? 'truecolor' : '256',
		unicode: term !== 'dumb',
		mouse: isTty,
		bracketedPaste: isTty,
		hyperlinks: isTty && !tmux,
		graphics,
		keyboardProtocol: graphics === 'kitty' ? 'kitty' : 'legacy',
		profile,
	}
}
