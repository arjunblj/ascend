import {
	inspectXlsxPackageGraph,
	type XlsxPackageGraph,
	type XlsxPackageGraphPart,
} from '@ascend/io-xlsx'
import { WorkbookDocument, type WorkbookLoadOptions } from './session.ts'
import type {
	WorkbookTrustFinding,
	WorkbookTrustFindingSeverity,
	WorkbookTrustReport,
} from './workbook-trust.ts'

export type WorkbookOpenIntent = 'risk-inventory' | 'read-values' | 'formula-analysis' | 'edit-plan'

export type WorkbookOpenFeatureCategory =
	| 'active-content'
	| 'security'
	| 'analytics'
	| 'visual'
	| 'formula'
	| 'metadata'
	| 'package'
	| 'worksheet'
	| 'unknown'

export type WorkbookOpenCostClass = 'tiny' | 'small' | 'medium' | 'large'

export interface WorkbookOpenFeatureSignal {
	readonly featureFamily: string
	readonly category: WorkbookOpenFeatureCategory
	readonly count: number
	readonly sampleParts: readonly string[]
}

export interface WorkbookOpenPlan {
	readonly intent: WorkbookOpenIntent
	readonly recommendedMode: NonNullable<WorkbookLoadOptions['mode']>
	readonly recommendedLoadOptions: Pick<WorkbookLoadOptions, 'mode' | 'richMetadata'>
	readonly richMetadataRecommended: boolean
	readonly reviewBeforeHydration: boolean
	readonly costClass: WorkbookOpenCostClass
	readonly partCount: number
	readonly worksheetPartCount: number
	readonly relationshipCount: number
	readonly formulaSignal: boolean
	readonly featureSignals: readonly WorkbookOpenFeatureSignal[]
	readonly riskFeatures: readonly WorkbookOpenFeatureSignal[]
	readonly reasons: readonly string[]
}

export type InteractiveOpenStepId =
	| 'plan-open'
	| 'open-preview'
	| 'review-trust'
	| 'promote-editable'

export interface InteractiveOpenStep {
	readonly id: InteractiveOpenStepId
	readonly title: string
	readonly recommended: boolean
}

export interface InteractiveOpenTrustSummary {
	readonly severity: WorkbookTrustFindingSeverity | 'none'
	readonly title: string
	readonly explanation: string
	readonly recommendedAction: string
	readonly findingCount: number
	readonly emittedFindingCount: number
	readonly topFindings: readonly WorkbookTrustFinding[]
}

export interface InteractiveOpenPlan {
	readonly plan: WorkbookOpenPlan
	readonly trustReport: WorkbookTrustReport
	readonly trustSummary: InteractiveOpenTrustSummary
	readonly recommendedMode: WorkbookOpenPlan['recommendedMode']
	readonly recommendedLoadOptions: WorkbookOpenPlan['recommendedLoadOptions']
	readonly previewLoadOptions: WorkbookOpenPlan['recommendedLoadOptions']
	readonly editableLoadOptions: Pick<WorkbookLoadOptions, 'mode' | 'richMetadata'>
	readonly reviewBeforeEdit: boolean
	readonly steps: readonly InteractiveOpenStep[]
}

export interface InspectWorkbookOpenPlanOptions {
	readonly intent?: WorkbookOpenIntent
	readonly samplePartLimit?: number
	readonly password?: string
}

export interface PlanInteractiveOpenOptions extends InspectWorkbookOpenPlanOptions {
	readonly maxTrustFindings?: number
}

export function inspectWorkbookOpenPlan(
	bytes: Uint8Array,
	options: InspectWorkbookOpenPlanOptions = {},
): WorkbookOpenPlan {
	return planWorkbookOpenFromGraph(inspectXlsxPackageGraph(bytes, passwordGraphOptions(options)), {
		...options,
		byteLength: bytes.byteLength,
	})
}

export async function planInteractiveOpen(
	bytes: Uint8Array,
	options: PlanInteractiveOpenOptions = {},
): Promise<InteractiveOpenPlan> {
	const packageGraph = inspectXlsxPackageGraph(bytes, passwordGraphOptions(options))
	const plan = planWorkbookOpen(packageGraph, options)
	const document = await WorkbookDocument.open(bytes, {
		...plan.recommendedLoadOptions,
		...(options.password !== undefined ? { password: options.password } : {}),
	})
	const trustReport = await document.trustReport({
		packageGraph,
		...(options.maxTrustFindings !== undefined ? { maxFindings: options.maxTrustFindings } : {}),
	})
	const trustSummary = summarizeInteractiveOpenTrust(plan, trustReport)
	const reviewBeforeEdit =
		plan.reviewBeforeHydration ||
		trustSummary.severity === 'blocked' ||
		trustSummary.severity === 'warning'
	return {
		plan,
		trustReport,
		trustSummary,
		recommendedMode: plan.recommendedMode,
		recommendedLoadOptions: plan.recommendedLoadOptions,
		previewLoadOptions: plan.recommendedLoadOptions,
		editableLoadOptions: {
			mode: 'full',
			...(plan.richMetadataRecommended ? { richMetadata: true } : {}),
		},
		reviewBeforeEdit,
		steps: interactiveOpenSteps(reviewBeforeEdit),
	}
}

function passwordGraphOptions(options: Pick<InspectWorkbookOpenPlanOptions, 'password'>): {
	readonly password?: string
} {
	return options.password === undefined ? {} : { password: options.password }
}

interface PlanWorkbookOpenOptions extends InspectWorkbookOpenPlanOptions {
	readonly byteLength?: number
}

export function planWorkbookOpen(
	packageGraph: XlsxPackageGraph,
	options: InspectWorkbookOpenPlanOptions = {},
): WorkbookOpenPlan {
	return planWorkbookOpenFromGraph(packageGraph, options)
}

function planWorkbookOpenFromGraph(
	packageGraph: XlsxPackageGraph,
	options: PlanWorkbookOpenOptions,
): WorkbookOpenPlan {
	const intent = options.intent ?? 'edit-plan'
	const samplePartLimit = options.samplePartLimit ?? 3
	const featureSignals = summarizeFeatureSignals(packageGraph.parts, samplePartLimit)
	const signalByFamily = new Map(featureSignals.map((signal) => [signal.featureFamily, signal]))
	const riskFeatures = featureSignals.filter((signal) => isRiskSignal(signal))
	const worksheetPartCount = countFeature(packageGraph.parts, 'worksheet')
	const hasActiveOrSecurityRisk = riskFeatures.some(
		(signal) => signal.category === 'active-content' || signal.category === 'security',
	)
	const hasUnknownReviewRisk = riskFeatures.some((signal) => signal.category === 'unknown')
	const hasRichMetadata =
		featureSignals.some((signal) => signal.category === 'analytics') ||
		featureSignals.some((signal) => signal.category === 'visual') ||
		hasFeature(signalByFamily, 'preservedMetadata') ||
		hasFeature(signalByFamily, 'preservedComments') ||
		hasFeature(signalByFamily, 'preservedThreadedComments')
	const formulaSignal = hasFeature(signalByFamily, 'preservedCalcChain')
	const richMetadataRecommended = hasRichMetadata && intent !== 'risk-inventory'
	const reviewBeforeHydration =
		intent !== 'risk-inventory' && (hasActiveOrSecurityRisk || hasUnknownReviewRisk)
	const recommendedMode = chooseRecommendedMode({
		intent,
		reviewBeforeHydration,
		richMetadataRecommended,
		formulaSignal,
	})
	const recommendedLoadOptions: WorkbookOpenPlan['recommendedLoadOptions'] = {
		mode: recommendedMode,
		...(richMetadataRecommended ? { richMetadata: true } : {}),
	}
	return {
		intent,
		recommendedMode,
		recommendedLoadOptions,
		richMetadataRecommended,
		reviewBeforeHydration,
		costClass: classifyCost(packageGraph, worksheetPartCount, options.byteLength),
		partCount: packageGraph.parts.length,
		worksheetPartCount,
		relationshipCount: packageGraph.relationships.length,
		formulaSignal,
		featureSignals,
		riskFeatures,
		reasons: buildReasons({
			intent,
			recommendedMode,
			richMetadataRecommended,
			reviewBeforeHydration,
			formulaSignal,
			riskFeatures,
			worksheetPartCount,
		}),
	}
}

function summarizeInteractiveOpenTrust(
	plan: WorkbookOpenPlan,
	report: WorkbookTrustReport,
): InteractiveOpenTrustSummary {
	const severity = highestTrustSeverity(report)
	const activeOrExternal =
		report.summary.byCategory['active-content'] + report.summary.byCategory['external-content']
	if (severity === 'blocked') {
		return {
			severity,
			title: 'Review required',
			explanation: 'The workbook contains active content that Ascend preserves but never executes.',
			recommendedAction: plan.reviewBeforeHydration
				? 'Open metadata-only preview, review the trust findings, and promote to editable only after the source is trusted.'
				: 'Open a read-only preview, review the trust findings, and promote to editable only after the source is trusted.',
			findingCount: report.summary.findingCount,
			emittedFindingCount: report.summary.emittedFindingCount,
			topFindings: report.findings.slice(0, 5),
		}
	}
	if (severity === 'warning') {
		return {
			severity,
			title: 'Review recommended',
			explanation:
				activeOrExternal > 0
					? 'The workbook references external or active workbook content that should be reviewed before editing.'
					: 'The workbook has trust findings that should be reviewed before editing.',
			recommendedAction: plan.reviewBeforeHydration
				? 'Open metadata-only first, review findings, then promote to editable.'
				: 'Open a read-only preview first and review findings before saving changes.',
			findingCount: report.summary.findingCount,
			emittedFindingCount: report.summary.emittedFindingCount,
			topFindings: report.findings.slice(0, 5),
		}
	}
	if (severity === 'info') {
		return {
			severity,
			title: 'Trust notes available',
			explanation: 'The workbook has informational trust notes but no blocking findings.',
			recommendedAction: 'Open the preview using the recommended load mode.',
			findingCount: report.summary.findingCount,
			emittedFindingCount: report.summary.emittedFindingCount,
			topFindings: report.findings.slice(0, 5),
		}
	}
	return {
		severity: 'none',
		title: 'No trust findings',
		explanation: 'No trust findings were emitted for the recommended preview load.',
		recommendedAction: 'Open the preview using the recommended load mode.',
		findingCount: report.summary.findingCount,
		emittedFindingCount: report.summary.emittedFindingCount,
		topFindings: [],
	}
}

function highestTrustSeverity(report: WorkbookTrustReport): WorkbookTrustFindingSeverity | 'none' {
	if (report.summary.bySeverity.blocked > 0) return 'blocked'
	if (report.summary.bySeverity.warning > 0) return 'warning'
	if (report.summary.bySeverity.info > 0) return 'info'
	return 'none'
}

function interactiveOpenSteps(reviewBeforeEdit: boolean): readonly InteractiveOpenStep[] {
	return [
		{ id: 'plan-open', title: 'Plan open', recommended: true },
		{ id: 'open-preview', title: 'Open preview', recommended: true },
		{ id: 'review-trust', title: 'Review trust findings', recommended: reviewBeforeEdit },
		{ id: 'promote-editable', title: 'Promote editable', recommended: !reviewBeforeEdit },
	]
}

function summarizeFeatureSignals(
	parts: readonly XlsxPackageGraphPart[],
	samplePartLimit: number,
): WorkbookOpenFeatureSignal[] {
	const groups = new Map<string, string[]>()
	for (const part of parts) {
		const existing = groups.get(part.featureFamily)
		if (existing) existing.push(part.path)
		else groups.set(part.featureFamily, [part.path])
	}
	return [...groups.entries()]
		.map(([featureFamily, paths]) => ({
			featureFamily,
			category: categorizeFeature(featureFamily),
			count: paths.length,
			sampleParts: paths.slice(0, samplePartLimit),
		}))
		.sort(
			(left, right) =>
				categoryRank(left.category) - categoryRank(right.category) ||
				left.featureFamily.localeCompare(right.featureFamily),
		)
}

function chooseRecommendedMode(options: {
	readonly intent: WorkbookOpenIntent
	readonly reviewBeforeHydration: boolean
	readonly richMetadataRecommended: boolean
	readonly formulaSignal: boolean
}): NonNullable<WorkbookLoadOptions['mode']> {
	if (options.intent === 'risk-inventory' || options.reviewBeforeHydration) {
		return 'metadata-only'
	}
	if (options.intent === 'read-values') return 'values'
	if (
		options.intent === 'formula-analysis' ||
		options.richMetadataRecommended ||
		options.formulaSignal
	) {
		return 'formula'
	}
	return 'full'
}

function buildReasons(options: {
	readonly intent: WorkbookOpenIntent
	readonly recommendedMode: NonNullable<WorkbookLoadOptions['mode']>
	readonly richMetadataRecommended: boolean
	readonly reviewBeforeHydration: boolean
	readonly formulaSignal: boolean
	readonly riskFeatures: readonly WorkbookOpenFeatureSignal[]
	readonly worksheetPartCount: number
}): string[] {
	const reasons = [`Intent '${options.intent}' maps to '${options.recommendedMode}' mode.`]
	if (options.reviewBeforeHydration) {
		reasons.push(
			'Active, security, signature, or unknown package features should be inventoried before hydration or edit planning.',
		)
	}
	if (options.richMetadataRecommended) {
		reasons.push(
			'Pivot, slicer, timeline, visual, comment, or metadata sidecars are present; preserve their inventory while planning.',
		)
	}
	if (options.formulaSignal) {
		reasons.push('A calculation-chain package part signals formula/dependency content.')
	}
	if (options.worksheetPartCount === 0) {
		reasons.push(
			'No worksheet package part was found; metadata inspection is the safest first read.',
		)
	}
	if (options.riskFeatures.length > 0) {
		reasons.push(
			`Risk families: ${options.riskFeatures.map((feature) => feature.featureFamily).join(', ')}.`,
		)
	}
	return reasons
}

function classifyCost(
	graph: XlsxPackageGraph,
	worksheetPartCount: number,
	byteLength: number | undefined,
): WorkbookOpenCostClass {
	const partCount = graph.parts.length
	const relationshipCount = graph.relationships.length
	const bytes = byteLength ?? 0
	if (bytes > 25_000_000 || partCount > 250 || worksheetPartCount > 64) return 'large'
	if (bytes > 5_000_000 || partCount > 80 || worksheetPartCount > 24 || relationshipCount > 300) {
		return 'medium'
	}
	if (bytes > 1_000_000 || partCount > 20 || worksheetPartCount > 4 || relationshipCount > 80) {
		return 'small'
	}
	return 'tiny'
}

function countFeature(parts: readonly XlsxPackageGraphPart[], featureFamily: string): number {
	return parts.filter((part) => part.featureFamily === featureFamily).length
}

function hasFeature(
	signals: ReadonlyMap<string, WorkbookOpenFeatureSignal>,
	featureFamily: string,
): boolean {
	return signals.has(featureFamily)
}

function isRiskSignal(signal: WorkbookOpenFeatureSignal): boolean {
	return (
		signal.category === 'active-content' ||
		signal.category === 'security' ||
		signal.category === 'unknown' ||
		signal.featureFamily === 'preservedSignature'
	)
}

function categorizeFeature(featureFamily: string): WorkbookOpenFeatureCategory {
	switch (featureFamily) {
		case 'preservedMacro':
		case 'preservedMacroSheet':
		case 'preservedActiveX':
		case 'preservedControl':
		case 'preservedCustomUi':
		case 'preservedEmbedding':
			return 'active-content'
		case 'preservedSignature':
		case 'preservedVendorSecurity':
			return 'security'
		case 'preservedPivot':
		case 'preservedSlicer':
		case 'preservedTimeline':
		case 'preservedTable':
		case 'preservedConnection':
		case 'preservedExternalLink':
		case 'preservedPowerQuery':
		case 'preservedQueryTable':
		case 'preservedDataModel':
			return 'analytics'
		case 'preservedChart':
		case 'preservedChartSheet':
		case 'preservedChartStyle':
		case 'preservedChartColor':
		case 'preservedDrawing':
		case 'preservedMedia':
		case 'preservedVml':
			return 'visual'
		case 'preservedCalcChain':
			return 'formula'
		case 'preservedMetadata':
		case 'preservedComments':
		case 'preservedThreadedComments':
		case 'preservedCustomXml':
		case 'preservedDocumentProperties':
		case 'preservedRevision':
		case 'preservedWorksheetSidecar':
		case 'preservedPrinterSettings':
		case 'preservedStyles':
		case 'preservedTheme':
			return 'metadata'
		case 'packageContentTypes':
		case 'packageRelationships':
		case 'workbook':
		case 'sharedStrings':
			return 'package'
		case 'worksheet':
			return 'worksheet'
		default:
			return 'unknown'
	}
}

function categoryRank(category: WorkbookOpenFeatureCategory): number {
	switch (category) {
		case 'active-content':
			return 0
		case 'security':
			return 1
		case 'unknown':
			return 2
		case 'analytics':
			return 3
		case 'visual':
			return 4
		case 'formula':
			return 5
		case 'metadata':
			return 6
		case 'worksheet':
			return 7
		case 'package':
			return 8
	}
}
