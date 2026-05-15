import type {
	Workbook,
	WorkbookCoreDocumentProperties,
	WorkbookCustomDocumentProperty,
	WorkbookDocumentProperties,
	WorkbookDocumentPropertyAppValue,
	WorkbookProperties,
	WorkbookThemeColor,
	WorkbookView,
} from '@ascend/core'
import type { Operation, Result } from '@ascend/schema'
import { ascendError, err, ok } from '@ascend/schema'
import { type PatchResult, patch } from './helpers.ts'

type RewriteExternalLinkOp = Extract<Operation, { op: 'rewriteExternalLink' }>
type SetWorkbookPropertiesOp = Extract<Operation, { op: 'setWorkbookProperties' }>
type SetDocumentPropertiesOp = Extract<Operation, { op: 'setDocumentProperties' }>
type SetWorkbookViewOp = Extract<Operation, { op: 'setWorkbookView' }>
type SetCalcSettingsOp = Extract<Operation, { op: 'setCalcSettings' }>
type SetThemeOp = Extract<Operation, { op: 'setTheme' }>
type MutableThemeMetadata = {
	colorCount: number
	name?: string
	colorSchemeName?: string
	majorFontLatin?: string
	minorFontLatin?: string
}

const WORKBOOK_PROPERTY_KEYS = [
	'codeName',
	'defaultThemeVersion',
	'filterPrivacy',
	'date1904',
] as const

function externalLinkSelectorRelId(entry: {
	readonly linkRelId?: string
	readonly externalBookRelId?: string
}): string | undefined {
	return entry.linkRelId ?? entry.externalBookRelId
}

const CORE_DOCUMENT_PROPERTY_KEYS = [
	'title',
	'subject',
	'creator',
	'keywords',
	'description',
	'lastModifiedBy',
	'revision',
	'created',
	'modified',
	'category',
	'contentStatus',
	'language',
	'identifier',
	'version',
] as const
const THEME_COLOR_SLOTS = new Set([
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
])

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

export function handleSetDocumentProperties(
	workbook: Workbook,
	op: SetDocumentPropertiesOp,
): Result<PatchResult> {
	const mode = op.mode ?? 'merge'
	if (mode !== 'merge' && mode !== 'replace') {
		return err(
			ascendError('VALIDATION_ERROR', 'setDocumentProperties mode must be merge or replace', {
				suggestedFix: 'Use mode="merge" for targeted docProps edits or mode="replace".',
			}),
		)
	}
	const validated = validateDocumentProperties(op.properties)
	if (!validated.ok) return validated

	const next: {
		core?: MutableCoreDocumentProperties
		app?: Record<string, WorkbookDocumentPropertyAppValue>
		custom?: WorkbookCustomDocumentProperty[]
	} = mode === 'replace' ? {} : cloneMutableDocumentProperties(workbook.documentProperties)

	if ('core' in op.properties) {
		if (op.properties.core === null) {
			delete next.core
		} else if (op.properties.core) {
			const core = mode === 'replace' || !next.core ? {} : { ...next.core }
			for (const key of CORE_DOCUMENT_PROPERTY_KEYS) {
				if (!(key in op.properties.core)) continue
				const value = op.properties.core[key]
				if (value === null || value === undefined) delete core[key]
				else core[key] = value
			}
			if (Object.keys(core).length > 0) next.core = core
			else delete next.core
		}
	}
	if ('app' in op.properties) {
		if (op.properties.app === null) {
			delete next.app
		} else if (op.properties.app) {
			const app = mode === 'replace' || !next.app ? {} : { ...next.app }
			for (const [key, value] of Object.entries(op.properties.app)) {
				if (value === null || value === undefined) delete app[key]
				else app[key] = Array.isArray(value) ? [...value] : value
			}
			if (Object.keys(app).length > 0) next.app = app
			else delete next.app
		}
	}
	if ('custom' in op.properties) {
		if (op.properties.custom === null) delete next.custom
		else if (op.properties.custom) {
			next.custom = op.properties.custom.map((property) => ({ ...property }))
		}
	}
	workbook.documentProperties = next as WorkbookDocumentProperties
	return ok(patch([], [], false))
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

export function handleSetTheme(workbook: Workbook, op: SetThemeOp): Result<PatchResult> {
	if (
		op.themeName === undefined &&
		op.colorSchemeName === undefined &&
		op.majorFontLatin === undefined &&
		op.minorFontLatin === undefined &&
		op.themeColors === undefined
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'setTheme requires theme metadata or themeColors', {
				suggestedFix:
					'Provide themeName, colorSchemeName, majorFontLatin, minorFontLatin, or themeColors.',
			}),
		)
	}

	const validated = validateThemeOperation(op)
	if (!validated.ok) return validated

	const metadata: MutableThemeMetadata = { ...workbook.themeMetadata }
	assignThemeField(metadata, 'name', op.themeName)
	assignThemeField(metadata, 'colorSchemeName', op.colorSchemeName)
	assignThemeField(metadata, 'majorFontLatin', op.majorFontLatin)
	assignThemeField(metadata, 'minorFontLatin', op.minorFontLatin)
	workbook.themeMetadata = metadata
	if (op.themeColors !== undefined) {
		const merged = mergeThemeColors(workbook.themeColors, op.themeColors)
		workbook.themeColors.splice(0, workbook.themeColors.length, ...merged)
		workbook.themeMetadata = { ...workbook.themeMetadata, colorCount: merged.length }
	}

	return ok(patch([], [], false))
}

function assignThemeField(
	target: MutableThemeMetadata,
	key: 'name' | 'colorSchemeName' | 'majorFontLatin' | 'minorFontLatin',
	value: string | undefined,
): void {
	if (value === undefined) return
	target[key] = value
}

function mergeThemeColors(
	existing: readonly WorkbookThemeColor[],
	updates: NonNullable<SetThemeOp['themeColors']>,
): WorkbookThemeColor[] {
	const bySlot = new Map(existing.map((color) => [color.slot, { ...color }]))
	for (const update of updates) {
		bySlot.set(update.slot, {
			slot: update.slot,
			...(update.rgb !== undefined ? { rgb: update.rgb.toUpperCase() } : {}),
			...(update.systemColor !== undefined ? { systemColor: update.systemColor } : {}),
			...(update.lastColor !== undefined ? { lastColor: update.lastColor.toUpperCase() } : {}),
		})
	}
	return [...bySlot.values()]
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
			if (op.linkRelId !== undefined && externalLinkSelectorRelId(entry) !== op.linkRelId)
				return false
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
		typeof properties.codeName !== 'string'
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'workbook property codeName must be a string or null', {
				suggestedFix: 'Use null to clear codeName, or provide a workbook code name string.',
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
	for (const [field, value] of [
		['filterPrivacy', properties.filterPrivacy],
		['date1904', properties.date1904],
	] as const) {
		if (value !== undefined && value !== null && typeof value !== 'boolean') {
			return err(
				ascendError('VALIDATION_ERROR', `workbook property ${field} must be a boolean or null`, {
					suggestedFix: `Use true, false, or null for ${field}.`,
				}),
			)
		}
	}
	return ok(undefined)
}

function validateDocumentProperties(
	properties: SetDocumentPropertiesOp['properties'],
): Result<undefined> {
	if (!properties || typeof properties !== 'object' || Array.isArray(properties)) {
		return err(
			ascendError('VALIDATION_ERROR', 'setDocumentProperties requires a properties object', {
				suggestedFix: 'Provide properties such as { "core": { "title": "Forecast" } }.',
			}),
		)
	}
	if (properties.core !== undefined && properties.core !== null) {
		if (typeof properties.core !== 'object' || Array.isArray(properties.core)) {
			return err(
				ascendError('VALIDATION_ERROR', 'document core properties must be an object or null'),
			)
		}
		for (const [key, value] of Object.entries(properties.core)) {
			if (value !== null && value !== undefined && typeof value !== 'string') {
				return err(
					ascendError('VALIDATION_ERROR', `document core property ${key} must be a string or null`),
				)
			}
		}
	}
	if (properties.app !== undefined && properties.app !== null) {
		if (typeof properties.app !== 'object' || Array.isArray(properties.app)) {
			return err(
				ascendError('VALIDATION_ERROR', 'document app properties must be an object or null'),
			)
		}
		for (const [key, value] of Object.entries(properties.app)) {
			if (value !== null && !isDocumentPropertyAppValue(value)) {
				return err(
					ascendError(
						'VALIDATION_ERROR',
						`document app property ${key} must be a scalar, scalar array, or null`,
					),
				)
			}
		}
	}
	if (properties.custom !== undefined && properties.custom !== null) {
		if (!Array.isArray(properties.custom)) {
			return err(
				ascendError('VALIDATION_ERROR', 'document custom properties must be an array or null'),
			)
		}
		for (const [index, property] of properties.custom.entries()) {
			if (!property || typeof property !== 'object' || Array.isArray(property)) {
				return err(
					ascendError('VALIDATION_ERROR', `document custom property ${index} must be an object`),
				)
			}
			if (typeof property.name !== 'string' || property.name.trim() === '') {
				return err(
					ascendError('VALIDATION_ERROR', `document custom property ${index} requires a name`),
				)
			}
			if (
				typeof property.value !== 'string' &&
				typeof property.value !== 'boolean' &&
				!(typeof property.value === 'number' && Number.isFinite(property.value))
			) {
				return err(
					ascendError(
						'VALIDATION_ERROR',
						`document custom property ${property.name} value must be a scalar`,
					),
				)
			}
		}
	}
	return ok(undefined)
}

function isDocumentPropertyAppValue(value: unknown): boolean {
	return (
		typeof value === 'string' ||
		typeof value === 'boolean' ||
		(typeof value === 'number' && Number.isFinite(value)) ||
		isScalarArray(value)
	)
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

function cloneMutableDocumentProperties(properties: WorkbookDocumentProperties): {
	core?: MutableCoreDocumentProperties
	app?: Record<string, WorkbookDocumentPropertyAppValue>
	custom?: WorkbookCustomDocumentProperty[]
} {
	return {
		...(properties.core ? { core: { ...properties.core } } : {}),
		...(properties.app
			? {
					app: Object.fromEntries(
						Object.entries(properties.app).map(([key, value]) => [
							key,
							Array.isArray(value) ? [...value] : value,
						]),
					),
				}
			: {}),
		...(properties.custom
			? { custom: properties.custom.map((property) => ({ ...property })) }
			: {}),
	}
}

function isScalarArray(value: unknown): value is readonly (string | number | boolean)[] {
	return (
		Array.isArray(value) &&
		value.every(
			(entry) =>
				typeof entry === 'string' ||
				typeof entry === 'boolean' ||
				(typeof entry === 'number' && Number.isFinite(entry)),
		)
	)
}

type MutableCoreDocumentProperties = {
	-readonly [K in keyof WorkbookCoreDocumentProperties]?: WorkbookCoreDocumentProperties[K]
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
		settings.calcMode !== undefined &&
		settings.calcMode !== 'auto' &&
		settings.calcMode !== 'manual' &&
		settings.calcMode !== 'autoNoTable'
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'calcMode must be auto, manual, or autoNoTable', {
				suggestedFix: 'Use one of Excel calc modes: auto, manual, or autoNoTable.',
			}),
		)
	}
	for (const [field, value] of [
		['fullCalcOnLoad', settings.fullCalcOnLoad],
		['calcCompleted', settings.calcCompleted],
		['calcOnSave', settings.calcOnSave],
		['forceFullCalc', settings.forceFullCalc],
	] as const) {
		if (value !== undefined && value !== null && typeof value !== 'boolean') {
			return err(
				ascendError('VALIDATION_ERROR', `${field} must be a boolean or null`, {
					suggestedFix: `Use true, false, or null for ${field}.`,
				}),
			)
		}
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
	if (
		settings.dateSystem !== undefined &&
		settings.dateSystem !== '1900' &&
		settings.dateSystem !== '1904'
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'dateSystem must be 1900 or 1904', {
				suggestedFix: 'Use dateSystem="1900" or dateSystem="1904".',
			}),
		)
	}
	const iterative = settings.iterativeCalc
	if (
		iterative !== undefined &&
		iterative !== null &&
		(typeof iterative !== 'object' || Array.isArray(iterative))
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'iterativeCalc must be an object or null', {
				suggestedFix: 'Use null to reset iterative calculation or provide iterativeCalc fields.',
			}),
		)
	}
	if (iterative && iterative.enabled !== undefined && typeof iterative.enabled !== 'boolean') {
		return err(
			ascendError('VALIDATION_ERROR', 'iterativeCalc.enabled must be a boolean', {
				suggestedFix: 'Use true or false for iterativeCalc.enabled.',
			}),
		)
	}
	if (iterative && iterative.maxIterations !== undefined) {
		if (!Number.isInteger(iterative.maxIterations) || iterative.maxIterations < 1) {
			return err(
				ascendError('VALIDATION_ERROR', 'iterativeCalc.maxIterations must be positive', {
					suggestedFix: 'Use a positive integer maxIterations value.',
				}),
			)
		}
	}
	if (
		iterative &&
		iterative.maxChange !== undefined &&
		(!Number.isFinite(iterative.maxChange) || iterative.maxChange < 0)
	) {
		return err(
			ascendError('VALIDATION_ERROR', 'iterativeCalc.maxChange must be non-negative', {
				suggestedFix: 'Use a non-negative maxChange convergence threshold.',
			}),
		)
	}
	return ok(undefined)
}

function validateThemeOperation(op: SetThemeOp): Result<undefined> {
	for (const [field, value] of [
		['themeName', op.themeName],
		['colorSchemeName', op.colorSchemeName],
		['majorFontLatin', op.majorFontLatin],
		['minorFontLatin', op.minorFontLatin],
	] as const) {
		if (value === undefined || value === null) continue
		if (value.trim() === '') {
			return err(
				ascendError('VALIDATION_ERROR', `${field} must be non-empty when provided`, {
					suggestedFix: `Omit ${field} to leave it unchanged, or provide a non-empty string.`,
				}),
			)
		}
	}
	if (op.themeColors === undefined) return ok(undefined)
	if (op.themeColors.length === 0) {
		return err(
			ascendError('VALIDATION_ERROR', 'themeColors must contain at least one color update', {
				suggestedFix: 'Provide one or more theme color entries with slot and rgb or systemColor.',
			}),
		)
	}
	const seen = new Set<string>()
	for (const color of op.themeColors) {
		if (!THEME_COLOR_SLOTS.has(color.slot)) {
			return err(
				ascendError('VALIDATION_ERROR', `Unsupported theme color slot "${color.slot}"`, {
					suggestedFix: `Use one of ${[...THEME_COLOR_SLOTS].join(', ')}.`,
				}),
			)
		}
		if (seen.has(color.slot)) {
			return err(
				ascendError('VALIDATION_ERROR', `Duplicate theme color slot "${color.slot}"`, {
					suggestedFix: 'Provide each theme color slot at most once.',
				}),
			)
		}
		seen.add(color.slot)
		if (color.rgb === undefined && color.systemColor === undefined) {
			return err(
				ascendError('VALIDATION_ERROR', 'Theme color requires rgb or systemColor', {
					suggestedFix: 'Use rgb for fixed colors or systemColor plus optional lastColor.',
				}),
			)
		}
		if (color.rgb !== undefined && !/^[0-9A-Fa-f]{6}$/.test(color.rgb)) {
			return err(
				ascendError('VALIDATION_ERROR', 'Theme color rgb must be 6 hex digits', {
					suggestedFix: 'Use values such as 4F81BD without a leading #.',
				}),
			)
		}
		if (color.lastColor !== undefined && !/^[0-9A-Fa-f]{6}$/.test(color.lastColor)) {
			return err(
				ascendError('VALIDATION_ERROR', 'Theme color lastColor must be 6 hex digits', {
					suggestedFix: 'Use values such as FFFFFF without a leading #.',
				}),
			)
		}
	}
	return ok(undefined)
}
