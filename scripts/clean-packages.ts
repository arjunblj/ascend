import { rm } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const packageNames = ['schema', 'core', 'formulas', 'engine', 'io-xlsx', 'io-csv', 'verify', 'sdk']
const repoRoot = fileURLToPath(new URL('..', import.meta.url))

for (const name of packageNames) {
	await rm(join(repoRoot, 'packages', name, 'dist'), { recursive: true, force: true })
	await rm(join(repoRoot, 'packages', name, 'tsconfig.tsbuildinfo'), { force: true })
}
