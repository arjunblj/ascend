import type { Workbook, WorkbookProperties, WorkbookView } from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'

type RewriteExternalLinkOp = Extract<Operation, { op: 'rewriteExternalLink' }>
type SetWorkbookPropertiesOp = Extract<Operation, { op: 'setWorkbookProperties' }>
type SetWorkbookViewOp = Extract<Operation, { op: 'setWorkbookView' }>
type SetCalcSettingsOp = Extract<Operation, { op: 'setCalcSettings' }>

const WORKBOOK_PROPERTY_KEYS = [
	'codeName',
	'defaultThemeVersion',
	'filterPrivacy',
	'date1904',
] as const

export function handleSetWorkbookProperties(
	workbook: Workbook,
	op: SetWorkbookPropertiesOp,
): Result<PatchResult> {
	const mode = op.mode ?? 'merge'
	if (mode !== 'merge' && mode !== 'replace') {
		return err(
			ascendError('VALIDATION_ERROR', 'setWorkbookProperties mode must be merge or replace', {
				suggestedFix: 'Use mode="merge" to update selected properties or mode="replace".',
			}),
		)
	}

	const validated = validateWorkbookProperties(op.properties)
	if (!validated.ok) return validated

	const next: Record<string, string | number | boolean> =
		mode === 'replace' ? {} : { ...workbook.workbookProperties }
	for (const key of WORKBOOK_PROPERTY_KEYS) {
		if (!(key in op.properties)) continue
		const value = op.properties[key]
		if (value === null || value === undefined) {
			delete next[key]
		} else {
			next[key] = value
		}
	}

	const oldDateSystem = workbook.calcSettings.dateSystem
	const date1904 =
		'date1904' in op.properties
			? op.properties.date1904 === true
			: mode === 'replace'
				? false
				: workbook.workbookProperties.date1904 === true
	if (mode === 'replace' || 'date1904' in op.properties) {
		workbook.calcSettings = {
			...workbook.calcSettings,
			dateSystem: date1904 ? '1904' : '1900',
		}
	}
	workbook.workbookProperties = next as WorkbookProperties

	return ok(patch([], [], oldDateSystem !== workbook.calcSettings.dateSystem))
}

export function handleSetWorkbookView(
	workbook: Workbook,
	op: SetWorkbookViewOp,
): Result<PatchResult> {
	const index = op.index ?? 0
	if (!Number.isInteger(index) || index < 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'setWorkbookView index must be a non-negative integer', {
				suggestedFix: 'Use index=0 for the primary workbook view.',
			}),
		)
	}
	if (op.view === null) {
		if (index >= workbook.workbookViews.length) {
			return err(
				ascendError('VALIDATION_ERROR', 'setWorkbookView cannot delete a missing view', {
					suggestedFix: 'Inspect workbook views and choose an existing view index.',
				}),
			)
		}
		workbook.workbookViews.splice(index, 1)
		return ok(patch([], [], false))
	}

	const mode = op.mode ?? 'merge'
	if (mode !== 'merge' && mode !== 'replace') {
		return err(
			ascendError('VALIDATION_ERROR', 'setWorkbookView mode must be merge or replace', {
				suggestedFix: 'Use mode="merge" to update selected view fields or mode="replace".',
			}),
		)
	}
	const validated = validateWorkbookView(op.view)
	if (!validated.ok) return validated
	if (index > workbook.workbookViews.length) {
		return err(
			ascendError('VALIDATION_ERROR', 'setWorkbookView index cannot skip view slots', {
				suggestedFix: `Use index=${workbook.workbookViews.length} to append the next view.`,
			}),
		)
	}

	const current = workbook.workbookViews[index]
	const next: Record<string, string | number> = mode === 'replace' || !current ? {} : { ...current }
	for (const key of ['activeTab', 'firstSheet', 'visibility', 'tabRatio'] as const) {
		if (!(key in op.view)) continue
		const value = op.view[key]
		if (value === null || value === undefined) {
			delete next[key]
		} else {
			next[key] = value
		}
	}
	if (index === workbook.workbookViews.length) workbook.workbookViews.push(next as WorkbookView)
	else workbook.workbookViews[index] = next as WorkbookView

	return ok(patch([], [], false))
}

export function handleSetCalcSettings(
	workbook: Workbook,
	op: SetCalcSettingsOp,
): Result<PatchResult> {
	const validated = validateCalcSettings(op.settings)
	if (!validated.ok) return validated
	const old = workbook.calcSettings
	const next = { ...old }
	if (op.settings.calcMode !== undefined) next.calcMode = op.settings.calcMode
	if (op.settings.fullCalcOnLoad !== undefined) next.fullCalcOnLoad = op.settings.fullCalcOnLoad
	assignOptionalCalcSetting(next, 'calcCompleted', op.settings.calcCompleted)
	assignOptionalCalcSetting(next, 'calcOnSave', op.settings.calcOnSave)
	assignOptionalCalcSetting(next, 'forceFullCalc', op.settings.forceFullCalc)
	assignOptionalCalcSetting(next, 'calcId', op.settings.calcId)
	if (op.settings.dateSystem !== undefined) next.dateSystem = op.settings.dateSystem
	if (op.settings.iterativeCalc === null) {
		next.iterativeCalc = { enabled: false, maxIterations: 100, maxChange: 0.001 }
	} else if (op.settings.iterativeCalc) {
		next.iterativeCalc = { ...old.iterativeCalc, ...op.settings.iterativeCalc }
	}
	workbook.calcSettings = next
	if (op.settings.dateSystem !== undefined) {
		workbook.workbookProperties = {
			...workbook.workbookProperties,
			date1904: op.settings.dateSystem === '1904',
		}
	}
	const recalcRequired =
		old.dateSystem !== workbook.calcSettings.dateSystem ||
		old.iterativeCalc.enabled !== workbook.calcSettings.iterativeCalc.enabled ||
		old.iterativeCalc.maxIterations !== workbook.calcSettings.iterativeCalc.maxIterations ||
		old.iterativeCalc.maxChange !== workbook.calcSettings.iterativeCalc.maxChange
	return ok(patch([], [], recalcRequired))
}

function assignOptionalCalcSetting<
	K extends 'calcCompleted' | 'calcOnSave' | 'forceFullCalc' | 'calcId',
>(
	target: Partial<Record<K, boolean | number>>,
	key: K,
	value: boolean | number | null | undefined,
): void {
	if (value === undefined) return
	if (value === null) delete target[key]
	else target[key] = value
}

export function handleRewriteExternalLink(
	workbook: Workbook,
	op: RewriteExternalLinkOp,
): Result<PatchResult> {
	if (
		op.partPath === undefined &&
		op.relId === undefined &&
		op.linkRelId === undefined &&
		op.target === undefined
	) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'rewriteExternalLink requires partPath, relId, linkRelId, or target',
				{
					suggestedFix:
						'Use inspect --detail external-refs and select a stable external link identifier.',
				},
			),
		)
	}
	if (op.newTarget.trim() === '') {
		return err(
			ascendError('VALIDATION_ERROR', 'newTarget must be a non-empty external workbook target', {
				suggestedFix: 'Provide the replacement workbook path or URL for the external link target.',
			}),
		)
	}

	const matches = workbook.externalReferenceDetails
		.map((entry, index) => ({ entry, index }))
		.filter(({ entry }) => {
			if (op.partPath !== undefined && entry.partPath !== op.partPath) return false
			if (op.relId !== undefined && entry.relId !== op.relId) return false
			if (op.linkRelId !== undefined && entry.linkRelId !== op.linkRelId) return false
			if (op.target !== undefined && entry.target !== op.target) return false
			return true
		})

	if (matches.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching external link found for rewriteExternalLink', {
				suggestedFix:
					'Inspect workbook externalReferenceDetails and provide a matching partPath, relId, linkRelId, or target.',
			}),
		)
	}
	if (matches.length > 1) {
		return err(
			ascendError('VALIDATION_ERROR', `rewriteExternalLink matched ${matches.length} links`, {
				suggestedFix:
					'Provide a more specific selector such as partPath plus linkRelId before committing.',
			}),
		)
	}

	const match = matches[0]
	if (!match) {
		return err(
			ascendError('VALIDATION_ERROR', 'No matching external link found for rewriteExternalLink'),
		)
	}
	workbook.externalReferenceDetails[match.index] = {
		...match.entry,
		target: op.newTarget,
		...(op.targetMode !== undefined ? { targetMode: op.targetMode } : {}),
	}

	return ok(patch([], [], false))
}

function validateWorkbookProperties(
	properties: SetWorkbookPropertiesOp['properties'],
): Result<undefined> {
	if (!properties || typeof properties !== 'object' || Array.isArray(properties)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setWorkbookProperties requires a properties object', {
				suggestedFix: 'Provide properties such as { "codeName": "Model", "filterPrivacy": true }.',
			}),
		)
	}
	if (
		properties.codeName !== undefined &&
		properties.codeName !== null &&
		properties.codeName.trim() === ''
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'workbook property codeName must be non-empty', {
				suggestedFix: 'Use null to clear codeName, or provide a non-empty workbook code name.',
			}),
		)
	}
	if (
		properties.defaultThemeVersion !== undefined &&
		properties.defaultThemeVersion !== null &&
		(!Number.isInteger(properties.defaultThemeVersion) || properties.defaultThemeVersion < 0)
	) {
		return err(
			ascendError(
				'VALIDATION_ERROR',
				'workbook property defaultThemeVersion must be a non-negative integer',
				{
					suggestedFix: 'Use a non-negative integer theme version, or null to clear it.',
				},
			),
		)
	}
	return ok(undefined)
}

function validateWorkbookView(view: NonNullable<SetWorkbookViewOp['view']>): Result<undefined> {
	for (const key of ['activeTab', 'firstSheet', 'tabRatio'] as const) {
		const value = view[key]
		if (value === undefined || value === null) continue
		if (!Number.isInteger(value) || value < 0) {
			return err(
				ascendError('VALIDATION_ERROR', `workbook view ${key} must be a non-negative integer`, {
					suggestedFix: `Use a non-negative integer for ${key}, or null to clear it.`,
				}),
			)
		}
	}
	return ok(undefined)
}

function validateCalcSettings(settings: SetCalcSettingsOp['settings']): Result<undefined> {
	if (!settings || typeof settings !== 'object' || Array.isArray(settings)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setCalcSettings requires a settings object', {
				suggestedFix: 'Provide settings such as { "calcMode": "manual" }.',
			}),
		)
	}
	if (
		settings.calcId !== undefined &&
		settings.calcId !== null &&
		(!Number.isInteger(settings.calcId) || settings.calcId < 0)
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'calcId must be a non-negative integer', {
				suggestedFix: 'Use a non-negative integer calcId, or null to clear it.',
			}),
		)
	}
	const iterative = settings.iterativeCalc
	if (iterative && iterative.maxIterations !== undefined) {
		if (!Number.isInteger(iterative.maxIterations) || iterative.maxIterations < 1) {
			return err(
				ascendError('VALIDATION_ERROR', 'iterativeCalc.maxIterations must be positive', {
					suggestedFix: 'Use a positive integer maxIterations value.',
				}),
			)
		}
	}
	if (iterative && iterative.maxChange !== undefined && iterative.maxChange < 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'iterativeCalc.maxChange must be non-negative', {
				suggestedFix: 'Use a non-negative maxChange convergence threshold.',
			}),
		)
	}
	return ok(undefined)
}
