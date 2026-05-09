import type { InputEvent } from '../runtime/types.ts'

export function isMouseEvent(event: InputEvent): event is Extract<InputEvent, { kind: 'mouse' }> {
	return event.kind === 'mouse'
}
