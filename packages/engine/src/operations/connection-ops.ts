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
				'setConnectionRefresh requires refreshOnLoad, saveData, or refreshedVersion',
				{
					suggestedFix:
						'Set refreshOnLoad=true for refresh-on-open or saveData=false when cached external data should not be trusted.',
				},
			),
		)
	}

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
	const index = workbook.connectionParts.indexOf(match)
	if (index < 0)
		return err(ascendError('VALIDATION_ERROR', 'No matching connection metadata found'))
	const updated: WorkbookConnectionPartInfo = {
		...match,
		...(op.refreshOnLoad !== undefined ? { refreshOnLoad: op.refreshOnLoad } : {}),
		...(op.saveData !== undefined ? { saveData: op.saveData } : {}),
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
		op.refreshOnLoad !== undefined || op.saveData !== undefined || op.refreshedVersion !== undefined
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
		if (op.sheet !== undefined && part.sheetName !== op.sheet) return false
		return true
	})
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
