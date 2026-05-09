import type { FileHubEntry } from '../runtime/types.ts'

export interface RecoverySnapshot {
	readonly id: string
	readonly workbookPath: string
	readonly snapshotPath: string
	readonly createdAt: number
}

export class RecoveryIndex {
	private readonly snapshots: RecoverySnapshot[] = []

	add(snapshot: RecoverySnapshot): void {
		this.snapshots.push(snapshot)
	}

	entries(): readonly FileHubEntry[] {
		return this.snapshots.map((snapshot) => ({
			label: snapshot.workbookPath.split(/[\\/]/).pop() ?? snapshot.workbookPath,
			path: snapshot.snapshotPath,
			detail: `Recovered ${new Date(snapshot.createdAt).toLocaleString()}`,
		}))
	}
}
