import type { Workbook, WorkbookConnectionPartInfo } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'

type SetConnectionRefreshOp = Extract<Operation, { op: 'setConnectionRefresh' }>

export function handleSetConnectionRefresh(
	workbook: Workbook,
	op: SetConnectionRefreshOp,
): Result<PatchResult> {
	if (!hasConnectionSelector(op)) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'setConnectionRefresh requires partPath, name, connectionId, or sheet',
				{
					suggestedFix:
						'Use inspect --detail connections to choose a workbook connection or query table.',
				},
			),
		)
	}
	if (!hasConnectionRefreshUpdate(op)) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'setConnectionRefresh requires refreshOnLoad, saveData, backgroundRefresh, keepAlive, refreshInterval, or refreshedVersion',
				{
					suggestedFix:
						'Set refreshOnLoad=true for refresh-on-open, saveData=false when cached external data should not be trusted, or edit workbook connection scheduling metadata.',
				},
			),
		)
	}
	const updateValidation = validateConnectionRefreshUpdateValues(op)
	if (updateValidation) return err(updateValidation)

	const matches = resolveConnectionRefreshMatches(workbook, op)
	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching connection refresh metadata found', {
				suggestedFix:
					'Inspect connectionParts and provide a matching partPath, name, connectionId, or sheet.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				`setConnectionRefresh matched ${matches.length} connection parts`,
				{
					suggestedFix:
						'Provide a more specific selector such as partPath plus name or connectionId.',
				},
			),
		)
	}

	const match = matches[0]
	if (!match) return err(ascendError('VALIDATION_ERROR', 'No matching connection metadata found'))
	if (hasWorkbookConnectionOnlyUpdate(op) && match.kind !== 'connection') {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'setConnectionRefresh backgroundRefresh, keepAlive, and refreshInterval only apply to workbook connection parts',
				{
					suggestedFix:
						'Choose a workbook connection in xl/connections.xml, or omit workbook-connection scheduling fields for query tables.',
				},
			),
		)
	}
	const index = workbook.connectionParts.indexOf(match)
	if (index < 0)
		return err(ascendError('VALIDATION_ERROR', 'No matching connection metadata found'))
	const updated: WorkbookConnectionPartInfo = {
		...match,
		...(op.refreshOnLoad !== undefined ? { refreshOnLoad: op.refreshOnLoad } : {}),
		...(op.saveData !== undefined ? { saveData: op.saveData } : {}),
		...(op.backgroundRefresh !== undefined ? { backgroundRefresh: op.backgroundRefresh } : {}),
		...(op.keepAlive !== undefined ? { keepAlive: op.keepAlive } : {}),
		...(op.refreshInterval !== undefined ? { refreshInterval: op.refreshInterval } : {}),
		...(op.refreshedVersion !== undefined ? { refreshedVersion: op.refreshedVersion } : {}),
	}
	workbook.connectionParts.splice(index, 1, updated)

	const warnings = refreshWarnings(updated)
	return ok(
		patch(
			[],
			updated.sheetName ? [updated.sheetName] : [],
			false,
			warnings.length ? warnings : undefined,
		),
	)
}

function hasConnectionSelector(op: SetConnectionRefreshOp): boolean {
	return (
		op.partPath !== undefined ||
		op.name !== undefined ||
		op.connectionId !== undefined ||
		op.sheet !== undefined
	)
}

function hasConnectionRefreshUpdate(op: SetConnectionRefreshOp): boolean {
	return (
		op.refreshOnLoad !== undefined ||
		op.saveData !== undefined ||
		op.backgroundRefresh !== undefined ||
		op.keepAlive !== undefined ||
		op.refreshInterval !== undefined ||
		op.refreshedVersion !== undefined
	)
}

function validateConnectionRefreshUpdateValues(op: SetConnectionRefreshOp) {
	for (const field of ['refreshOnLoad', 'saveData', 'backgroundRefresh', 'keepAlive'] as const) {
		if (op[field] !== undefined && typeof op[field] !== 'boolean') {
			return ascendError('VALIDATION_ERROR', `setConnectionRefresh ${field} must be boolean`, {
				suggestedFix: `Set ${field}=true or ${field}=false.`,
			})
		}
	}
	if (
		op.refreshInterval !== undefined &&
		(!Number.isInteger(op.refreshInterval) || op.refreshInterval < 0)
	) {
		return ascendError(
			'VALIDATION_ERROR',
			'setConnectionRefresh refreshInterval must be a non-negative integer',
			{
				suggestedFix:
					'Use the non-negative refresh interval in minutes expected by Excel metadata.',
			},
		)
	}
	if (
		op.refreshedVersion !== undefined &&
		(!Number.isInteger(op.refreshedVersion) || op.refreshedVersion < 0)
	) {
		return ascendError(
			'VALIDATION_ERROR',
			'setConnectionRefresh refreshedVersion must be a non-negative integer',
			{
				suggestedFix: 'Use the non-negative refreshedVersion value expected by Excel metadata.',
			},
		)
	}
	return null
}

function hasWorkbookConnectionOnlyUpdate(op: SetConnectionRefreshOp): boolean {
	return (
		op.backgroundRefresh !== undefined ||
		op.keepAlive !== undefined ||
		op.refreshInterval !== undefined
	)
}

function resolveConnectionRefreshMatches(
	workbook: Workbook,
	op: SetConnectionRefreshOp,
): WorkbookConnectionPartInfo[] {
	return workbook.connectionParts.filter((part) => {
		if (part.kind === 'powerQueryMashup') return false
		if (op.partPath !== undefined && part.partPath !== op.partPath) return false
		if (op.name !== undefined && part.name !== op.name) return false
		if (op.connectionId !== undefined && part.connectionId !== op.connectionId) return false
		if (op.sheet !== undefined && !sameOptionalSheetName(part.sheetName, op.sheet)) return false
		return true
	})
}

function sameOptionalSheetName(left: string | undefined, right: string): boolean {
	return left !== undefined && left.toLowerCase() === right.toLowerCase()
}

function refreshWarnings(part: WorkbookConnectionPartInfo) {
	const warnings = []
	if (part.refreshOnLoad) {
		warnings.push(
			ascendError(
				'VALIDATION_ERROR',
				'Connection is marked refresh-on-open; external data may change when Excel opens the workbook.',
				{
					details: {
						kind: part.kind,
						partPath: part.partPath,
						name: part.name,
						connectionId: part.connectionId,
						sheetName: part.sheetName,
					},
					suggestedFix:
						'Open in Excel or another connection-aware engine to refresh query output before treating saved cells as current.',
				},
			),
		)
	}
	if (part.saveData === false) {
		warnings.push(
			ascendError(
				'VALIDATION_ERROR',
				'Connection cache data is not saved; refresh is required before cached external output can be trusted.',
				{
					details: {
						kind: part.kind,
						partPath: part.partPath,
						name: part.name,
						connectionId: part.connectionId,
						sheetName: part.sheetName,
					},
					suggestedFix: 'Refresh the connection in Excel or a query-aware engine.',
				},
			),
		)
	}
	return warnings
}
