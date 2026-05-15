import { describe, expect, test } from 'bun:test'
import { mkdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
	AscendWorkbook,
	createAgentPlanFromWorkbook,
	preparedPathMutationPlanHandle,
	resolveOperationInputForWorkbook,
	resolveOperationInputShape,
	sha256Bytes,
	withPathMutationResult,
} from './index.ts'

const TEMP_DIR = join(tmpdir(), `ascend-operation-input-${process.pid}`)

const apiSource = (input: {
	readonly ops?: readonly unknown[] | null
	readonly mutations?: readonly unknown[] | null
}) => ({
	hasOpsKey: input.ops !== undefined,
	ops: input.ops ?? null,
	hasMutationsKey: input.mutations !== undefined,
	mutations: input.mutations ?? null,
	operationSchemaSuggestedFix: 'Call /operations for canonical operation schemas and examples.',
})

describe('operation input helpers', () => {
	test('rejects mixed ops and path mutations', () => {
		const result = resolveOperationInputShape(
			apiSource({
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
				mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 1 }],
			}),
		)

		expect(result.ok).toBe(false)
		if (!result.ok) {
			expect(result.error.code).toBe('VALIDATION_ERROR')
			expect(result.error.message).toBe('Provide either ops or mutations, not both')
		}
	})

	test('parses canonical ops with transport-specific repair text', () => {
		const result = resolveOperationInputShape(apiSource({ ops: [{ op: 'missingOperation' }] }))

		expect(result.ok).toBe(false)
		if (!result.ok) {
			expect(result.error.details?.suggestedFix).toBeUndefined()
			expect(result.error.suggestedFix).toBe(
				'Call /operations for canonical operation schemas and examples.',
			)
			expect(result.error.details).toEqual(
				expect.objectContaining({
					issueCount: 1,
				}),
			)
		}
	})

	test('accepts explicit empty ops as a no-op batch', () => {
		const result = resolveOperationInputShape(apiSource({ ops: [] }))

		expect(result).toEqual({ ok: true, ops: [] })
	})

	test('accepts explicit empty path mutations as a no-op batch', () => {
		const wb = AscendWorkbook.create()
		const result = resolveOperationInputForWorkbook(wb, apiSource({ mutations: [] }))

		expect(result.ok).toBe(true)
		if (result.ok) {
			expect(result.ops).toEqual([])
			expect(result.pathMutations).toMatchObject({
				mutationCount: 0,
				issueCount: 0,
				issues: [],
				replayable: true,
			})
		}
	})

	test('reports malformed path mutations with structured issue details', () => {
		const result = resolveOperationInputShape(apiSource({ mutations: [{ path: 123 }] }))

		expect(result.ok).toBe(false)
		if (!result.ok) {
			expect(result.error.code).toBe('VALIDATION_ERROR')
			expect(result.error.details).toEqual(
				expect.objectContaining({
					issueCount: 1,
					issues: ['mutations[0]: Mutation path must be a string or string array.'],
				}),
			)
		}
	})

	test('compiles replayable path mutations and preserves compiler metadata', () => {
		const wb = AscendWorkbook.create()
		const result = resolveOperationInputForWorkbook(
			wb,
			apiSource({ mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 42 }] }),
		)

		expect(result.ok).toBe(true)
		if (result.ok) {
			expect(result.ops).toEqual([
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 42 }] },
			])
			expect(result.pathMutations?.replayable).toBe(true)
		}
	})

	test('rejects non-replayable path mutation batches without hiding partial ops', () => {
		const wb = AscendWorkbook.create()
		const result = resolveOperationInputForWorkbook(
			wb,
			apiSource({ mutations: [{ path: '/sheets/Missing/cells/A1/value', value: 1 }] }),
		)

		expect(result.ok).toBe(false)
		if (!result.ok) {
			expect(result.error.code).toBe('VALIDATION_ERROR')
			expect(result.error.details).toEqual(
				expect.objectContaining({
					compiledOps: [],
					supportedPathShapes: expect.any(Array),
				}),
			)
		}
	})

	test('attaches path mutation metadata only when present', () => {
		const result = { ok: true as const }
		expect(withPathMutationResult(result, undefined)).toBe(result)

		const wb = AscendWorkbook.create()
		const compiled = resolveOperationInputForWorkbook(
			wb,
			apiSource({ mutations: [{ path: '/sheets/Sheet1/cells/A1/value', value: 1 }] }),
		)
		expect(compiled.ok).toBe(true)
		if (compiled.ok) {
			expect(withPathMutationResult(result, compiled.pathMutations)).toEqual({
				ok: true,
				pathMutations: compiled.pathMutations,
			})
		}
	})

	test('prepared path mutation handles default the hash guard and remain one-shot', async () => {
		mkdirSync(TEMP_DIR, { recursive: true })
		const input = join(TEMP_DIR, 'prepared-path-mutation.xlsx')
		const output = join(TEMP_DIR, 'prepared-path-mutation-out.xlsx')
		const secondOutput = join(TEMP_DIR, 'prepared-path-mutation-second.xlsx')
		const seed = AscendWorkbook.create()
		await seed.save(input)

		const opened = await AscendWorkbook.openSourceBytes(input)
		const pathMutations = opened.workbook.compilePathMutations([
			{ path: '/sheets/Sheet1/cells/A1/value', value: 99 },
		])
		expect(pathMutations.replayable).toBe(true)
		const inputSha256 = sha256Bytes(opened.sourceBytes)
		const plan = await createAgentPlanFromWorkbook(
			input,
			inputSha256,
			opened.workbook,
			pathMutations.ops,
		)
		const handle = preparedPathMutationPlanHandle({
			file: input,
			inputSha256,
			planDigest: plan.planDigest,
			operationCount: plan.operationCount,
			workbook: opened.workbook,
			ops: pathMutations.ops,
			sourceBytes: opened.sourceBytes,
			pathMutations,
		})

		const committed = await handle.commit({ output })
		expect(committed.trace.phases.find((phase) => phase.phase === 'hash-guard')?.summary).toBe(
			'Input hash matched expected SHA-256.',
		)
		expect(committed.postWrite.valid).toBe(true)

		await expect(handle.commit({ output: secondOutput })).rejects.toThrow(
			'Prepared agent plan has already been committed',
		)
	})
})
