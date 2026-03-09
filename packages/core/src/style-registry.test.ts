import { describe, expect, test } from 'bun:test'
import { StyleRegistry } from './style-registry.ts'

describe('StyleRegistry', () => {
	test('stores frozen style copies', () => {
		const registry = new StyleRegistry()
		const style = { font: { bold: true }, numberFormat: '0.0%' }
		const styleId = registry.register(style)
		style.font.bold = false
		style.numberFormat = 'General'

		const stored = registry.get(styleId)
		expect(stored?.font?.bold).toBe(true)
		expect(stored?.numberFormat).toBe('0.0%')
		expect(stored && Object.isFrozen(stored)).toBe(true)
		expect(stored?.font && Object.isFrozen(stored.font)).toBe(true)
	})

	test('clone shares until register diverges', () => {
		const registry = new StyleRegistry()
		const baseId = registry.register({ numberFormat: '0.0%' })
		const clone = registry.clone()
		const cloneId = clone.register({ numberFormat: '0.00%' })

		expect(registry.size).toBe(2)
		expect(clone.size).toBe(3)
		expect(registry.get(baseId)?.numberFormat).toBe('0.0%')
		expect(clone.get(cloneId)?.numberFormat).toBe('0.00%')
		expect(registry.get(cloneId)).toBeUndefined()
	})
})
