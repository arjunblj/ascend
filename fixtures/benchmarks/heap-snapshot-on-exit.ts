import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname } from 'node:path'

const writtenSnapshots = new Set<string>()

export function writeHeapSnapshot(target: string | undefined): void {
	if (!target) return
	if (writtenSnapshots.has(target)) return
	writtenSnapshots.add(target)
	mkdirSync(dirname(target), { recursive: true })
	writeFileSync(target, JSON.stringify(Bun.generateHeapSnapshot()))
	console.error(`heap snapshot: ${target}`)
}

export function writeHeapSnapshotFromEnv(): void {
	writeHeapSnapshot(process.env.ASCEND_HEAP_SNAPSHOT)
}

process.once('beforeExit', async () => {
	writeHeapSnapshotFromEnv()
})
