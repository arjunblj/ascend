import { type AscendError, ascendError, type Operation } from '@ascend/schema'
import { operationValidationDetails, parseOperations } from './ops.ts'
import { SUPPORTED_PATH_MUTATION_SHAPES } from './path-mutations.ts'
import type { PathMutation, PathMutationResult } from './types.ts'

export type ResolvedOperationInput =
	| {
			readonly ok: true
			readonly ops: readonly Operation[]
			readonly pathMutations?: PathMutationResult
	  }
	| { readonly ok: false; readonly error: AscendError }

export interface OperationInputSource {
	readonly hasOpsKey: boolean
	readonly ops: readonly unknown[] | null
	readonly hasMutationsKey: boolean
	readonly mutations: readonly unknown[] | null
	readonly operationSchemaSuggestedFix: string
}

export type OperationInputShape =
	| { readonly ok: true; readonly ops: readonly Operation[] }
	| { readonly ok: true; readonly mutations: readonly PathMutation[] }
	| { readonly ok: false; readonly error: AscendError }

export interface PathMutationCompiler {
	compilePathMutations(mutations: readonly PathMutation[]): PathMutationResult
}

export function resolveOperationInputForWorkbook(
	workbook: PathMutationCompiler,
	input: OperationInputSource,
): ResolvedOperationInput {
	const shape = resolveOperationInputShape(input)
	if (!shape.ok || !('mutations' in shape)) return shape
	return compilePathMutationInput(workbook, shape.mutations)
}

export function resolveOperationInputShape(input: OperationInputSource): OperationInputShape {
	if (input.hasOpsKey && input.hasMutationsKey) {
		return {
			ok: false,
			error: ascendError('VALIDATION_ERROR', 'Provide either ops or mutations, not both', {
				retryStrategy: 'modified',
				suggestedFix: 'Send canonical operations in ops or path-addressed mutations in mutations.',
			}),
		}
	}
	const hasOps = input.ops !== null && input.ops.length > 0
	const hasMutations = input.mutations !== null && input.mutations.length > 0
	if (!hasOps && !hasMutations) {
		return {
			ok: false,
			error: ascendError('VALIDATION_ERROR', 'Missing or invalid ops or mutations', {
				retryStrategy: 'modified',
				suggestedFix:
					'Send non-empty ops, or send mutations like {"path":"/sheets/Sheet1/cells/A1/value","value":123}.',
			}),
		}
	}
	if (hasOps) {
		const parsed = parseOperations(input.ops)
		if (!parsed.ok) {
			return {
				ok: false,
				error: ascendError('VALIDATION_ERROR', parsed.error, {
					details: operationValidationDetails(parsed),
					retryStrategy: 'modified',
					suggestedFix: input.operationSchemaSuggestedFix,
				}),
			}
		}
		return { ok: true, ops: parsed.value }
	}

	const parsedMutations = parsePathMutationBody(input.mutations ?? [])
	if (!parsedMutations.ok) return parsedMutations
	return { ok: true, mutations: parsedMutations.mutations }
}

function parsePathMutationBody(
	value: readonly unknown[],
):
	| { readonly ok: true; readonly mutations: readonly PathMutation[] }
	| { readonly ok: false; readonly error: AscendError } {
	const mutations: PathMutation[] = []
	for (const [index, entry] of value.entries()) {
		if (entry === null || typeof entry !== 'object' || Array.isArray(entry)) {
			return {
				ok: false,
				error: pathMutationShapeError(index, 'Mutation must be an object with path and value.'),
			}
		}
		const path = (entry as Record<string, unknown>).path
		if (
			typeof path !== 'string' &&
			(!Array.isArray(path) || !path.every((segment) => typeof segment === 'string'))
		) {
			return {
				ok: false,
				error: pathMutationShapeError(index, 'Mutation path must be a string or string array.'),
			}
		}
		mutations.push({
			path,
			...(Object.hasOwn(entry, 'value') ? { value: (entry as Record<string, unknown>).value } : {}),
		})
	}
	return { ok: true, mutations }
}

export function compilePathMutationInput(
	workbook: PathMutationCompiler,
	mutations: readonly PathMutation[],
): ResolvedOperationInput {
	const compiled = workbook.compilePathMutations(mutations)
	if (!compiled.replayable) return { ok: false, error: pathMutationCompileError(compiled) }
	return { ok: true, ops: compiled.ops, pathMutations: compiled }
}

function pathMutationShapeError(index: number, message: string): AscendError {
	return ascendError('VALIDATION_ERROR', message, {
		details: {
			issueCount: 1,
			issues: [`mutations[${index}]: ${message}`],
			issueDetails: [
				{ code: 'invalid_path_mutation', mutationIndex: index, path: `mutations[${index}]` },
			],
		},
		retryStrategy: 'modified',
		suggestedFix: 'Use mutations shaped like {"path":"/sheets/Sheet1/cells/A1/value","value":123}.',
	})
}

function pathMutationCompileError(result: PathMutationResult): AscendError {
	return ascendError('VALIDATION_ERROR', 'Path mutation compilation failed', {
		details: {
			mutationCount: result.mutationCount,
			issueCount: result.issueCount,
			issues: result.issues.map((issue) => issue.message),
			issueDetails: result.issues,
			compiledOps: result.ops,
			supportedPathShapes: SUPPORTED_PATH_MUTATION_SHAPES,
		},
		retryStrategy: 'modified',
		suggestedFix:
			'Use supported paths such as /sheets/{sheet}/cells/{A1}/value, /sheets/{sheet}/cells/{A1}/formula, /sheets/{sheet}/ranges/{A1:B2}/clear, /tables/{table}/rows/append, or /names/{name}/ref.',
	})
}

export function withPathMutationResult<T extends object>(
	result: T,
	compiled: PathMutationResult | undefined,
): T | (T & { readonly pathMutations: PathMutationResult }) {
	return compiled ? { ...result, pathMutations: compiled } : result
}
