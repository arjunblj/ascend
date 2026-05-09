import type { TerminalCapabilities } from '../runtime/types.ts'

export function describeGraphicsFallback(capabilities: TerminalCapabilities): string {
	switch (capabilities.graphics) {
		case 'kitty':
			return 'Kitty graphics with Unicode fallback'
		case 'iterm':
			return 'iTerm2 inline images with Unicode fallback'
		case 'sixel':
			return 'Sixel graphics with Unicode fallback'
		case 'unicode':
			return 'Unicode/Braille/block graphics'
		case 'off':
			return 'ASCII-only graphics'
	}
}
