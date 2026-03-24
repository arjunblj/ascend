export function levenshtein(a: string, b: string): number {
	if (a === b) return 0
	if (a.length === 0) return b.length
	if (b.length === 0) return a.length
	const m = a.length
	const n = b.length
	const prev = new Array<number>(n + 1)
	const curr = new Array<number>(n + 1)
	for (let j = 0; j <= n; j++) prev[j] = j
	for (let i = 1; i <= m; i++) {
		curr[0] = i
		for (let j = 1; j <= n; j++) {
			curr[j] =
				a[i - 1] === b[j - 1]
					? (prev[j - 1] ?? 0)
					: 1 + Math.min(prev[j] ?? i, curr[j - 1] ?? j, prev[j - 1] ?? Math.max(i, j))
		}
		for (let j = 0; j <= n; j++) prev[j] = curr[j] ?? 0
	}
	return prev[n] ?? 0
}
