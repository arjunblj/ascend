import { describe, expect, test } from 'bun:test'

import { hashLegacyProtectionPassword, legacyProtectionPasswordMatches } from './protection.ts'

describe('Excel legacy protection passwords', () => {
	test('hashes worksheet protection passwords using the OOXML legacy algorithm', () => {
		expect(hashLegacyProtectionPassword('password')).toBe('83AF')
		expect(hashLegacyProtectionPassword('test')).toBe('CBEB')
		expect(hashLegacyProtectionPassword('')).toBe('CE4B')
	})

	test('matches legacy protection hashes case-insensitively', () => {
		expect(legacyProtectionPasswordMatches('password', '83af')).toBe(true)
		expect(legacyProtectionPasswordMatches('wrong', '83AF')).toBe(false)
	})
})
