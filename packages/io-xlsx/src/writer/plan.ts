import { encode } from './zip.ts'

export type WritePartOwner =
	| { readonly kind: 'package' }
	| { readonly kind: 'workbook' }
	| { readonly kind: 'sheet'; readonly sheetName: string }

export type WritePartOrigin = 'generated' | 'preserved-inline' | 'preserved-source' | 'capsule'

export interface WritePartDescriptor {
	readonly path: string
	readonly owner: WritePartOwner
	readonly origin: WritePartOrigin
	readonly contentType?: string
	/** When set, part is built via streaming; not in parts map. */
	readonly streamingBuild?: (onChunk: (chunk: string) => void) => void
}

export interface WritePlanResult {
	readonly parts: ReadonlyMap<string, Uint8Array>
	readonly descriptors: readonly WritePartDescriptor[]
	readonly extraOverrides: readonly { partPath: string; contentType: string }[]
	readonly skippedCapsulePaths: ReadonlySet<string>
}

export interface WritePlanSummary {
	readonly totalParts: number
	readonly byOrigin: Readonly<Record<WritePartOrigin, number>>
	readonly byOwnerKind: Readonly<Record<WritePartOwner['kind'], number>>
	readonly sheetPartCounts: Readonly<Record<string, number>>
}

export class WritePlanBuilder {
	private readonly parts = new Map<string, Uint8Array>()
	private readonly descriptors: WritePartDescriptor[] = []
	private readonly extraOverrides: Array<{ partPath: string; contentType: string }> = []
	private readonly skippedCapsulePaths = new Set<string>()

	constructor(private readonly includeParts = true) {}

	putXml(path: string, xml: string, descriptor: Omit<WritePartDescriptor, 'path'>): void {
		if (this.includeParts) {
			this.parts.set(path, encode(xml))
		}
		this.record(path, descriptor)
	}

	putBytes(path: string, bytes: Uint8Array, descriptor: Omit<WritePartDescriptor, 'path'>): void {
		if (this.includeParts) {
			this.parts.set(path, bytes)
		}
		this.record(path, descriptor)
	}

	recordOnly(path: string, descriptor: Omit<WritePartDescriptor, 'path'>): void {
		this.record(path, descriptor)
	}

	putStreamingSheet(
		path: string,
		descriptor: Omit<WritePartDescriptor, 'path'>,
		build: (onChunk: (chunk: string) => void) => void,
	): void {
		if (this.includeParts) {
			this.descriptors.push({ path, ...descriptor, streamingBuild: build })
		} else {
			this.record(path, descriptor)
		}
	}

	addOverride(partPath: string, contentType: string): void {
		this.extraOverrides.push({ partPath, contentType })
	}

	skipCapsulePath(path: string): void {
		this.skippedCapsulePaths.add(path)
	}

	isCapsulePathSkipped(path: string): boolean {
		return this.skippedCapsulePaths.has(path)
	}

	build(): WritePlanResult {
		return {
			parts: this.parts,
			descriptors: this.descriptors,
			extraOverrides: this.extraOverrides,
			skippedCapsulePaths: this.skippedCapsulePaths,
		}
	}

	private record(path: string, descriptor: Omit<WritePartDescriptor, 'path'>): void {
		this.descriptors.push({ path, ...descriptor })
	}
}

export function summarizeWritePlan(plan: WritePlanResult): WritePlanSummary {
	const byOrigin: Record<WritePartOrigin, number> = {
		generated: 0,
		'preserved-inline': 0,
		'preserved-source': 0,
		capsule: 0,
	}
	const byOwnerKind: Record<WritePartOwner['kind'], number> = {
		package: 0,
		workbook: 0,
		sheet: 0,
	}
	const sheetPartCounts: Record<string, number> = {}
	for (const descriptor of plan.descriptors) {
		byOrigin[descriptor.origin] += 1
		byOwnerKind[descriptor.owner.kind] += 1
		if (descriptor.owner.kind === 'sheet') {
			sheetPartCounts[descriptor.owner.sheetName] =
				(sheetPartCounts[descriptor.owner.sheetName] ?? 0) + 1
		}
	}
	return {
		totalParts: plan.descriptors.length,
		byOrigin,
		byOwnerKind,
		sheetPartCounts,
	}
}
