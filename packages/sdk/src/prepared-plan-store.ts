import { randomUUID } from 'node:crypto'
import { type AscendError, ascendError } from '@ascend/schema'
import type { AgentCommitOptions, AgentCommitResult, PreparedAgentPlan } from './agent-workflow.ts'
import type { PathMutationResult } from './types.ts'

export interface PreparedPlanHandle {
	readonly file: string
	readonly inputSha256: string
	readonly planDigest: string
	readonly operationCount: number
	readonly pathMutations?: PathMutationResult
	commit(options: AgentCommitOptions): Promise<AgentCommitResult>
}

export interface PreparedPlanStoreOptions {
	readonly preparedPlanMaxHandles?: number
	readonly preparedPlanTtlMs?: number
	readonly now?: () => number
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

export function withPreparedPlanHandle<T extends object>(
	result: T,
	preparedPlan: PreparedPlanMetadata | undefined,
): T | (T & { readonly preparedPlan: PreparedPlanMetadata }) {
	return preparedPlan ? { ...result, preparedPlan } : result
}

function positiveIntegerOption(value: number | undefined, fallback: number): number {
	if (value === undefined || !Number.isFinite(value)) return fallback
	return Math.max(1, Math.floor(value))
}
