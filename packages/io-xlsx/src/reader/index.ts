import type { CellStyle, StyleId } from '@ascend/core'
import { Workbook } from '@ascend/core'
import type {
	AscendError,
	CompatibilityReport,
	CompatibilityStatus,
	CompatibilityTier,
	FeatureReport,
	Result,
} from '@ascend/schema'
import { ascendError, emptyReport, err, ok } from '@ascend/schema'
import { parseContentTypes } from './content-types.ts'
import {
	getRelsPath,
	parseRelationships,
	REL_OFFICE_DOC,
	REL_SHARED_STRINGS,
	REL_STYLES,
	resolvePath,
} from './relationships.ts'
import { parseSharedStrings } from './shared-strings.ts'
import { parseSheet } from './sheet.ts'
import { parseStyles } from './styles.ts'
import { parseWorkbookXml } from './workbook.ts'
import { extractZip } from './zip.ts'

export interface ReadXlsxResult {
	readonly workbook: Workbook
	readonly report: CompatibilityReport
}

const decoder = new TextDecoder('utf-8')

export function readXlsx(bytes: Uint8Array): Result<ReadXlsxResult, AscendError> {
	let parts: Map<string, Uint8Array>
	try {
		parts = extractZip(bytes)
	} catch (e) {
		return err(
			ascendError('CORRUPT_FILE', `Invalid ZIP: ${e instanceof Error ? e.message : 'unknown'}`),
		)
	}

	const contentTypesXml = readPart(parts, '[Content_Types].xml')
	if (!contentTypesXml) {
		return err(ascendError('CORRUPT_FILE', 'Missing [Content_Types].xml'))
	}
	const contentTypes = parseContentTypes(contentTypesXml)

	const rootRelsXml = readPart(parts, '_rels/.rels')
	if (!rootRelsXml) {
		return err(ascendError('CORRUPT_FILE', 'Missing _rels/.rels'))
	}
	const rootRels = parseRelationships(rootRelsXml)

	const docRel = rootRels.find((r) => r.type === REL_OFFICE_DOC)
	if (!docRel) {
		return err(ascendError('CORRUPT_FILE', 'No officeDocument relationship found'))
	}
	const workbookPath = docRel.target.replace(/^\//, '')

	const wbXml = readPart(parts, workbookPath)
	if (!wbXml) {
		return err(ascendError('CORRUPT_FILE', `Missing workbook: ${workbookPath}`))
	}
	const wbInfo = parseWorkbookXml(wbXml)

	const wbRelsPath = getRelsPath(workbookPath)
	const wbRelsXml = readPart(parts, wbRelsPath)
	const wbRels = wbRelsXml ? parseRelationships(wbRelsXml) : []
	const relMap = new Map(wbRels.map((r) => [r.id, r]))

	const ssPart = wbRels.find((r) => r.type === REL_SHARED_STRINGS)
	const ssPath = ssPart ? resolvePath(workbookPath, ssPart.target) : undefined
	const ssXml = ssPath ? readPart(parts, ssPath) : undefined
	const sharedStrings = ssXml ? parseSharedStrings(ssXml) : []

	const stylesPart = wbRels.find((r) => r.type === REL_STYLES)
	const stylesPath = stylesPart ? resolvePath(workbookPath, stylesPart.target) : undefined
	const stylesXml = stylesPath ? readPart(parts, stylesPath) : undefined
	const parsedStyles = stylesXml
		? parseStyles(stylesXml)
		: { cellStyles: [{}], isDateFormat: [false] }

	const workbook = new Workbook()
	workbook.calcSettings = wbInfo.calcSettings

	const styleIds = registerStyles(workbook, parsedStyles.cellStyles)

	for (const entry of wbInfo.sheets) {
		const rel = relMap.get(entry.rId)
		if (!rel) continue

		const sheetPath = resolvePath(workbookPath, rel.target)
		const sheetXml = readPart(parts, sheetPath)
		if (!sheetXml) continue

		const sheet = parseSheet(entry.name, sheetXml, {
			sharedStrings,
			styleIds,
			isDateFormat: parsedStyles.isDateFormat,
		})
		sheet.state = entry.state
		workbook.sheets.push(sheet)
	}

	for (const dn of wbInfo.definedNames) {
		workbook.definedNames.set(dn.name, dn.formula)
	}

	const report = buildReport(contentTypes)

	return ok({ workbook, report })
}

function readPart(parts: Map<string, Uint8Array>, path: string): string | undefined {
	const data = parts.get(path)
	if (!data) return undefined
	return decoder.decode(data)
}

function registerStyles(workbook: Workbook, cellStyles: CellStyle[]): StyleId[] {
	return cellStyles.map((style) => workbook.styles.register(style))
}

function buildReport(contentTypes: {
	overrides: ReadonlyMap<string, string>
}): CompatibilityReport {
	const features: FeatureReport[] = []

	const unsupportedTypes: [string, string][] = [
		['chart', 'chart+xml'],
		['pivotTable', 'pivotTable+xml'],
		['drawing', 'drawing+xml'],
		['vbaProject', 'vbaProject'],
	]

	for (const [feature, pattern] of unsupportedTypes) {
		const locations: string[] = []
		for (const [path, ct] of contentTypes.overrides) {
			if (ct.includes(pattern)) locations.push(path)
		}
		if (locations.length > 0) {
			features.push({
				feature,
				tier: 'unsupported' as CompatibilityTier,
				count: locations.length,
				locations,
			})
		}
	}

	if (features.length === 0) return emptyReport('xlsx')

	const summary = { exact: 0, normalized: 0, preserved: 0, unsupported: 0 }
	for (const f of features) {
		summary[f.tier] += f.count
	}

	const status: CompatibilityStatus =
		summary.unsupported > 0 ? 'has-unsupported' : summary.preserved > 0 ? 'has-preserved' : 'clean'

	return {
		status,
		features,
		summary,
		sourceFormat: 'xlsx',
	}
}
