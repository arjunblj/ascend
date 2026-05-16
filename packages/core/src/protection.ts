export function hashLegacyProtectionPassword(password: string): string {
	let hash = 0
	for (let index = 0; index < password.length; index++) {
		let value = password.charCodeAt(index) << (index + 1)
		const rotated = value >> 15
		value &= 0x7fff
		hash ^= value | rotated
	}
	hash ^= password.length
	hash ^= 0xce4b
	return (hash & 0xffff).toString(16).toUpperCase().padStart(4, '0')
}

export function legacyProtectionPasswordMatches(password: string, hash: string): boolean {
	return hashLegacyProtectionPassword(password).toUpperCase() === hash.toUpperCase()
}
