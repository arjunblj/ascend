import { randomUUID } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { type AscendError, AscendException, ascendError, type Operation } from '@ascend/schema'
import {
	type AgentCommitOptions,
	type AgentCommitResult,
	commitAgentPlanFromWorkbook,
	type PreparedAgentPlan,
	sha256Bytes,
} from './agent-workflow.ts'
import type { PathMutationResult } from './types.ts'
import type { AscendWorkbook } from './workbook.ts'

export interface PreparedPlanHandle {
	readonly file: string
	readonly inputSha256: string
	readonly planDigest: string
	readonly operationCount: number
	readonly pathMutations?: PathMutationResult
	commit(options?: AgentCommitOptions): Promise<AgentCommitResult>
}

export interface PreparedPlanStoreOptions {
	readonly preparedPlanMaxHandles?: number
	readonly preparedPlanTtlMs?: number
	readonly now?: () => number
}

export interface PreparedPathMutationPlanHandleOptions {
	readonly file: string
	readonly inputSha256: string
	readonly planDigest: string
	readonly operationCount: number
	readonly workbook: AscendWorkbook
	readonly ops: readonly Operation[]
	readonly sourceBytes: Uint8Array
	readonly preparedCheck: ReturnType<AscendWorkbook['check']>
	readonly pathMutations?: PathMutationResult
}

export interface PreparedPlanMetadata {
	readonly id: string
	readonly file: string
	readonly inputSha256: string
	readonly planDigest: string
	readonly operationCount: number
	readonly expiresAt: string
	readonly ttlMs: number
}

interface PreparedPlanRecord {
	readonly handle: PreparedPlanHandle
	readonly expiresAtMs: number
}

export type PreparedPlanUnavailableReason = 'expired' | 'evicted' | 'already-used' | 'unknown'

export type PreparedPlanTakeResult =
	| { readonly ok: true; readonly handle: PreparedPlanHandle }
	| { readonly ok: false; readonly error: AscendError }

const DEFAULT_PREPARED_PLAN_MAX_HANDLES = 64
const DEFAULT_PREPARED_PLAN_TTL_MS = 5 * 60 * 1000

export class PreparedPlanStore {
	private readonly handles = new Map<string, PreparedPlanRecord>()
	private readonly unavailableReasons = new Map<string, PreparedPlanUnavailableReason>()
	private readonly maxHandles: number
	private readonly ttlMs: number
	private readonly now: () => number

	constructor(options: PreparedPlanStoreOptions = {}) {
		this.maxHandles = positiveIntegerOption(
			options.preparedPlanMaxHandles,
			DEFAULT_PREPARED_PLAN_MAX_HANDLES,
		)
		this.ttlMs = positiveIntegerOption(options.preparedPlanTtlMs, DEFAULT_PREPARED_PLAN_TTL_MS)
		this.now = options.now ?? Date.now
	}

	add(handle: PreparedPlanHandle): PreparedPlanMetadata {
		this.pruneExpired()
		while (this.handles.size >= this.maxHandles) {
			const oldest = this.handles.keys().next().value
			if (oldest === undefined) break
			this.handles.delete(oldest)
			this.rememberUnavailable(oldest, 'evicted')
		}
		const id = randomUUID()
		const expiresAtMs = this.now() + this.ttlMs
		this.handles.set(id, { handle, expiresAtMs })
		return {
			id,
			file: handle.file,
			inputSha256: handle.inputSha256,
			planDigest: handle.planDigest,
			operationCount: handle.operationCount,
			expiresAt: new Date(expiresAtMs).toISOString(),
			ttlMs: this.ttlMs,
		}
	}

	take(id: string): PreparedPlanTakeResult {
		const record = this.handles.get(id)
		if (!record) {
			this.pruneExpired()
			return { ok: false, error: preparedPlanUnavailableError(id, this.unavailableReason(id)) }
		}
		if (record.expiresAtMs <= this.now()) {
			this.handles.delete(id)
			this.rememberUnavailable(id, 'expired')
			return { ok: false, error: preparedPlanUnavailableError(id, 'expired') }
		}
		this.handles.delete(id)
		this.rememberUnavailable(id, 'already-used')
		return { ok: true, handle: record.handle }
	}

	private pruneExpired(): void {
		const now = this.now()
		for (const [id, record] of this.handles) {
			if (record.expiresAtMs <= now) {
				this.handles.delete(id)
				this.rememberUnavailable(id, 'expired')
			}
		}
	}

	private unavailableReason(id: string): PreparedPlanUnavailableReason {
		return this.unavailableReasons.get(id) ?? 'unknown'
	}

	private rememberUnavailable(id: string, reason: PreparedPlanUnavailableReason): void {
		this.unavailableReasons.set(id, reason)
		while (this.unavailableReasons.size > this.maxHandles * 2) {
			const oldest = this.unavailableReasons.keys().next().value
			if (oldest === undefined) break
			this.unavailableReasons.delete(oldest)
		}
	}
}

export function preparedPlanUnavailableError(
	planHandle: string,
	reason: PreparedPlanUnavailableReason,
): AscendError {
	const messages: Record<PreparedPlanUnavailableReason, string> = {
		expired: 'Prepared plan handle expired',
		evicted: 'Prepared plan handle was evicted',
		'already-used': 'Prepared plan handle has already been used',
		unknown: 'Prepared plan handle was not found',
	}
	return ascendError('VALIDATION_ERROR', messages[reason], {
		details: {
			rule: 'prepared-plan-handle-unavailable',
			reason,
			planHandle,
		},
		suggestedFix: 'Re-run ascend plan with prepare=true before committing.',
	})
}

export function preparedPlanHandle(prepared: PreparedAgentPlan): PreparedPlanHandle {
	const commit = prepared.commit
	return {
		file: prepared.file,
		inputSha256: prepared.inputSha256,
		planDigest: prepared.planDigest,
		operationCount: prepared.operationCount,
		commit: (options) => commit(options),
	}
}

export function preparedPathMutationPlanHandle(
	prepared: PreparedPathMutationPlanHandleOptions,
): PreparedPlanHandle {
	let committed = false
	return {
		file: prepared.file,
		inputSha256: prepared.inputSha256,
		planDigest: prepared.planDigest,
		operationCount: prepared.operationCount,
		...(prepared.pathMutations !== undefined ? { pathMutations: prepared.pathMutations } : {}),
		commit: async (options = {}) => {
			if (committed) {
				throw new AscendException(
					ascendError('VALIDATION_ERROR', 'Prepared agent plan has already been committed', {
						suggestedFix: 'Create a fresh prepared plan before committing another output.',
					}),
				)
			}
			const current = await readWorkbookFileBytes(prepared.file)
			const currentSha256 = sha256Bytes(current)
			if (currentSha256 !== prepared.inputSha256) {
				throw new AscendException(
					ascendError('VALIDATION_ERROR', 'Input workbook changed after agent plan was prepared', {
						details: {
							expected: prepared.inputSha256,
							actual: currentSha256,
							planDigest: prepared.planDigest,
						},
						suggestedFix: 'Re-run ascend plan and commit with the new input workbook.',
					}),
				)
			}
			const result = await commitAgentPlanFromWorkbook(
				prepared.file,
				prepared.inputSha256,
				prepared.workbook,
				prepared.ops,
				{
					...options,
					expectSha256: options.expectSha256 ?? prepared.inputSha256,
				},
				{ sourceBytes: prepared.sourceBytes, preparedCheck: prepared.preparedCheck },
			)
			committed = true
			return result
		},
	}
}

export function withPreparedPlanHandle<T extends object>(
	result: T,
	preparedPlan: PreparedPlanMetadata | undefined,
): T | (T & { readonly preparedPlan: PreparedPlanMetadata }) {
	return preparedPlan ? { ...result, preparedPlan } : result
}

async function readWorkbookFileBytes(file: string): Promise<Uint8Array> {
	if (typeof Bun !== 'undefined') return Bun.file(file).bytes()
	return new Uint8Array(await readFile(file))
}

function positiveIntegerOption(value: number | undefined, fallback: number): number {
	if (value === undefined || !Number.isFinite(value)) return fallback
	return Math.max(1, Math.floor(value))
}
