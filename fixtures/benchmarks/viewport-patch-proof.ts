import { AscendSession, AscendWorkbook, type Operation } from '../../packages/sdk/src/index.ts'

export interface ViewportPatchProofCaseResult {
	readonly name: string
	readonly expected: string
	readonly observed: string
	readonly passed: boolean
	readonly patchBytes?: number
	readonly changedRefs?: readonly string[]
	readonly invalidationReason?: string
	readonly requiredAction?: string
}

export interface ViewportPatchProofResult {
	readonly generatedAt: string
	readonly cases: readonly ViewportPatchProofCaseResult[]
	readonly passed: boolean
	readonly totalPatchBytes: number
}

const BASE_REQUEST = {
	sheet: 'Sheet1',
	topRow: 0,
	leftCol: 0,
	rowCount: 1,
	colCount: 1,
} as const

export async function runViewportPatchProof(): Promise<ViewportPatchProofResult> {
	const cases = [
		await retainedPatchCase(),
		await skippedTokenCase(),
		await invalidTokenCase(),
		await crossSessionTokenCase(),
		await expiredHistoryCase(),
		await projectionChangeCase(),
		await metadataInvalidationCase(),
	]
	return {
		generatedAt: new Date().toISOString(),
		cases,
		passed: cases.every((entry) => entry.passed),
		totalPatchBytes: cases.reduce((sum, entry) => sum + (entry.patchBytes ?? 0), 0),
	}
}

export function viewportPatchProofMarkdown(result: ViewportPatchProofResult): string {
	return [
		'# Viewport Patch Proof',
		'',
		`Generated: ${result.generatedAt}`,
		'Boundary: this proves bounded single-session viewport patch retention and explicit refresh reasons. It is not CRDT collaboration, multi-writer convergence, or unbounded event sourcing.',
		'',
		'| Case | Expected | Observed | Passed | Patch bytes | Changed refs | Invalidation reason | Required action |',
		'| --- | --- | --- | --- | ---: | --- | --- | --- |',
		...result.cases.map(markdownRow),
		'',
		`All passed: ${result.passed}`,
		`Total patch bytes: ${result.totalPatchBytes}`,
	].join('\n')
}

async function retainedPatchCase(): Promise<ViewportPatchProofCaseResult> {
	const session = await openSeedSession()
	try {
		const base = session.readViewport(BASE_REQUEST)
		await applyExact(session, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
		])
		const patch = session.readViewportPatchResult({
			...BASE_REQUEST,
			changedSince: base.changeToken,
		})
		const refs = patch.patch?.changedCells.map((cell) => cell.ref) ?? []
		return patchResult('retained-patch', 'patch:A1', refs.includes('A1'), patch.patch)
	} finally {
		session.close()
	}
}

async function skippedTokenCase(): Promise<ViewportPatchProofCaseResult> {
	const session = await openSeedSession()
	try {
		const base = session.readViewport(BASE_REQUEST)
		await applyExact(session, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
		])
		const first = session.readViewport({ ...BASE_REQUEST, changedSince: base.changeToken })
		await applyExact(session, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 3 }] },
		])
		const skipped = session.readViewportPatchResult({
			...BASE_REQUEST,
			changedSince: base.changeToken,
		})
		const refs = skipped.patch?.changedCells.map((cell) => `${cell.ref}=${cell.flatValue}`) ?? []
		return {
			...patchResult(
				'skipped-token-retained',
				'patch from older retained token',
				first.patch !== undefined && refs.includes('A1=3'),
				skipped.patch,
			),
			changedRefs: refs,
		}
	} finally {
		session.close()
	}
}

async function invalidTokenCase(): Promise<ViewportPatchProofCaseResult> {
	const session = await openSeedSession()
	try {
		const result = session.readViewportPatchResult({
			...BASE_REQUEST,
			changedSince: 'not-a-token',
		})
		return invalidationResult('invalid-token', 'base-token-invalid', result.patchInvalidation)
	} finally {
		session.close()
	}
}

async function crossSessionTokenCase(): Promise<ViewportPatchProofCaseResult> {
	const seed = await seedBytes()
	const first = await AscendSession.open(seed, { mode: 'interactive' })
	const second = await AscendSession.open(seed, { mode: 'interactive' })
	try {
		const base = first.readViewport(BASE_REQUEST)
		const result = second.readViewport({
			...BASE_REQUEST,
			changedSince: base.changeToken,
		})
		return invalidationResult(
			'cross-session-token',
			'base-snapshot-missing',
			result.patchInvalidation,
		)
	} finally {
		first.close()
		second.close()
	}
}

async function expiredHistoryCase(): Promise<ViewportPatchProofCaseResult> {
	const session = await openSeedSession()
	try {
		const base = session.readViewport(BASE_REQUEST)
		for (let index = 0; index < 130; index++) {
			await applyExact(session, [
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: index + 2 }] },
			])
		}
		const result = session.readViewportPatchResult({
			...BASE_REQUEST,
			changedSince: base.changeToken,
		})
		return invalidationResult('expired-history', 'base-token-expired', result.patchInvalidation)
	} finally {
		session.close()
	}
}

async function projectionChangeCase(): Promise<ViewportPatchProofCaseResult> {
	const session = await openSeedSession()
	try {
		const base = session.readViewport(BASE_REQUEST)
		await applyExact(session, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
		])
		const result = session.readViewportPatchResult({
			...BASE_REQUEST,
			rowCount: 2,
			changedSince: base.changeToken,
		})
		return invalidationResult(
			'projection-change',
			'base-snapshot-missing',
			result.patchInvalidation,
		)
	} finally {
		session.close()
	}
}

async function metadataInvalidationCase(): Promise<ViewportPatchProofCaseResult> {
	const session = await openSeedSession()
	try {
		const base = session.readViewport(BASE_REQUEST)
		await applyExact(session, [{ op: 'setComment', sheet: 'Sheet1', ref: 'A1', text: 'review' }])
		const result = session.readViewportPatchResult({
			...BASE_REQUEST,
			changedSince: base.changeToken,
		})
		return invalidationResult(
			'metadata-invalidation',
			'viewport-invalidated',
			result.patchInvalidation,
		)
	} finally {
		session.close()
	}
}

async function openSeedSession(): Promise<AscendSession> {
	return AscendSession.open(await seedBytes(), { mode: 'interactive' })
}

async function seedBytes(): Promise<Uint8Array> {
	const wb = AscendWorkbook.create()
	applyWorkbookExact(wb, [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
	return wb.toBytes()
}

async function applyExact(session: AscendSession, ops: readonly Operation[]): Promise<void> {
	const result = await session.apply(ops)
	if (result.apply.errors.length > 0)
		throw new Error(result.apply.errors.map((error) => error.message).join('\n'))
}

function applyWorkbookExact(workbook: AscendWorkbook, ops: readonly Operation[]): void {
	const result = workbook.apply(ops)
	if (result.errors.length > 0)
		throw new Error(result.errors.map((error) => error.message).join('\n'))
}

function patchResult(
	name: string,
	expected: string,
	passed: boolean,
	patch: NonNullable<ReturnType<AscendSession['readViewportPatchResult']>['patch']> | null,
): ViewportPatchProofCaseResult {
	const changedRefs = patch?.changedCells.map((cell) => cell.ref) ?? []
	return {
		name,
		expected,
		observed: patch ? `patch:${changedRefs.join(',') || 'empty'}` : 'no-patch',
		passed,
		...(patch ? { patchBytes: patch.byteLength, changedRefs } : {}),
	}
}

function invalidationResult(
	name: string,
	expected: string,
	invalidation: ReturnType<AscendSession['readViewportPatchResult']>['patchInvalidation'],
): ViewportPatchProofCaseResult {
	return {
		name,
		expected,
		observed: invalidation?.reason ?? 'no-invalidation',
		passed: invalidation?.reason === expected,
		invalidationReason: invalidation?.reason,
		requiredAction: invalidation?.requiredAction,
	}
}

function markdownRow(row: ViewportPatchProofCaseResult): string {
	return [
		row.name,
		row.expected,
		row.observed,
		String(row.passed),
		String(row.patchBytes ?? 0),
		row.changedRefs?.join(', ') ?? 'n/a',
		row.invalidationReason ?? 'n/a',
		row.requiredAction ?? 'n/a',
	]
		.map((cell) => ` ${cell} `)
		.join('|')
		.replace(/^/, '|')
		.replace(/$/, '|')
}

if (import.meta.main) {
	const result = await runViewportPatchProof()
	if (process.argv.includes('--json')) console.log(JSON.stringify(result, null, 2))
	else {
		console.log(viewportPatchProofMarkdown(result))
		console.error(`Generated viewport patch proof over ${result.cases.length} cases.`)
	}
}
