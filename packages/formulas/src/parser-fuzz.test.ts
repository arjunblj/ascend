import { describe, expect, test } from 'bun:test'
import { parseFormula } from './parser.ts'

function rng(seed: number): () => number {
	let state = seed >>> 0
	return () => {
		state = (Math.imul(1664525, state) + 1013904223) >>> 0
		return state / 0x1_0000_0000
	}
}

function randomFormula(random: () => number): string {
	const parts = [
		'A1',
		'B2',
		'C$3',
		'$D$4',
		'Sheet1!A1',
		"'My Sheet'!B2",
		'SUM(',
		'IF(',
		'VLOOKUP(',
		'INDEX(',
		'MATCH(',
		'1',
		'2.5',
		'-3',
		'"hello"',
		'TRUE',
		'FALSE',
		'+',
		'-',
		'*',
		'/',
		'^',
		'&',
		'=',
		'<>',
		'<',
		'>',
		'<=',
		'>=',
		',',
		':',
		'(',
		')',
		'{',
		'}',
		';',
		' ',
		'#REF!',
		'#N/A',
		'A1:B10',
		'1:1',
		'A:Z',
		'@',
		'#',
	]
	const len = 1 + Math.floor(random() * 8)
	let formula = ''
	for (let i = 0; i < len; i++) {
		formula += parts[Math.floor(random() * parts.length)]
	}
	return formula
}

describe('formula parser fuzz', () => {
	test('random formulas never throw (only return ok or error)', () => {
		const random = rng(0xdeadbeef)
		let threw = false
		for (let i = 0; i < 500; i++) {
			const formula = randomFormula(random)
			try {
				parseFormula(formula)
			} catch {
				threw = true
			}
		}
		expect(threw).toBe(false)
	})

	test('adversarial deeply nested parens', () => {
		const depth = 200
		const formula = `${'('.repeat(depth)}1${')'.repeat(depth)}`
		let threw = false
		try {
			parseFormula(formula)
		} catch {
			threw = true
		}
		expect(threw).toBe(false)
	})

	test('very long formula string', () => {
		const formula = Array.from({ length: 500 }, (_, i) => `A${i + 1}`).join('+')
		let threw = false
		try {
			parseFormula(formula)
		} catch {
			threw = true
		}
		expect(threw).toBe(false)
	})
})
