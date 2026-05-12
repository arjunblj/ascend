import { describe, expect, test } from 'bun:test'
import type { StyleId } from '../../packages/core/src/index.ts'
import { applyOperation } from '../../packages/engine/src/index.ts'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
	extractZip,
	inspectXlsxPackageGraph,
	readXlsx,
	writeXlsx,
	type XlsxPackageGraph,
	type XlsxPackageGraphFidelityIssue,
} from '../../packages/io-xlsx/src/index.ts'
import { makeXlsx } from '../../packages/io-xlsx/test/helpers.ts'
import { numberValue } from '../../packages/schema/src/index.ts'

const S0 = 0 as StyleId
const EXTERNAL_LINK_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'

describe('external link package graph fidelity', () => {
	test('inventories relationship-id bindings and keeps symbolic formulas through safe cell edits', () => {
		const sourceBytes = externalLinkPackageFixture()
		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
		expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(beforeGraph))
		expect(externalLinkPartIdentities(beforeGraph)).toEqual([
			expect.objectContaining({
				path: 'xl/externalLinks/externalLink1.xml',
				ownerScope: 'external-link',
				featureFamily: 'preservedExternalLink',
				preservationPolicy: 'preserve-exact',
			}),
			expect.objectContaining({
				path: 'xl/externalLinks/externalLink2.xml',
				ownerScope: 'external-link',
				featureFamily: 'preservedExternalLink',
				preservationPolicy: 'preserve-exact',
			}),
			expect.objectContaining({
				path: 'xl/externalLinks/externalLink3.xml',
				ownerScope: 'external-link',
				featureFamily: 'preservedExternalLink',
				preservationPolicy: 'preserve-exact',
			}),
			expect.objectContaining({
				path: 'xl/externalLinks/externalLink4.xml',
				ownerScope: 'external-link',
				featureFamily: 'preservedExternalLink',
				preservationPolicy: 'preserve-exact',
			}),
		])
		expect(externalLinkContentTypeOverrides(beforeGraph)).toEqual([
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				contentType: EXTERNAL_LINK_CONTENT_TYPE,
			},
			{
				partPath: 'xl/externalLinks/externalLink2.xml',
				contentType: EXTERNAL_LINK_CONTENT_TYPE,
			},
			{
				partPath: 'xl/externalLinks/externalLink3.xml',
				contentType: EXTERNAL_LINK_CONTENT_TYPE,
			},
			{
				partPath: 'xl/externalLinks/externalLink4.xml',
				contentType: EXTERNAL_LINK_CONTENT_TYPE,
			},
		])

		const read = readXlsx(sourceBytes)
		expectOk(read)
		expect(read.value.workbook.externalReferences).toEqual([
			'xl/externalLinks/externalLink1.xml',
			'xl/externalLinks/externalLink2.xml',
			'xl/externalLinks/externalLink3.xml',
			'xl/externalLinks/externalLink4.xml',
		])
		expect(read.value.workbook.externalReferenceDetails).toEqual([
			expect.objectContaining({
				partPath: 'xl/externalLinks/externalLink1.xml',
				relId: 'rIdExternal1',
				externalBookRelId: 'rIdChosen',
				linkRelId: 'rIdChosen',
				linkRelationshipKind: 'xlStartup',
				linkBindingStatus: 'externalBookRelId',
				target: 'personal.xlsb',
				targetMode: 'External',
			}),
			expect.objectContaining({
				partPath: 'xl/externalLinks/externalLink2.xml',
				relId: 'rIdExternal2',
				externalBookRelId: 'rIdMissing',
				linkRelId: 'rIdLibrary',
				linkRelationshipKind: 'xlLibrary',
				linkBindingStatus: 'fallbackPathRelationship',
				target: 'library.xlsx',
				targetMode: 'External',
			}),
			expect.objectContaining({
				partPath: 'xl/externalLinks/externalLink3.xml',
				relId: 'rIdExternal3',
				linkRelId: 'rIdPathMissing',
				linkRelationshipKind: 'xlPathMissing',
				linkBindingStatus: 'fallbackPathRelationship',
				target: 'missing-source.xlsx',
				targetMode: 'External',
			}),
			expect.objectContaining({
				partPath: 'xl/externalLinks/externalLink4.xml',
				relId: 'rIdExternal4',
				externalBookRelId: 'rIdAbsent',
				linkBindingStatus: 'missingPathRelationship',
			}),
		])
		expect(read.value.workbook.definedNames.get('ExternalTotal')).toBe('[1]Sheet1!$A$1')
		expect(read.value.workbook.sheets[0]?.cells.get(0, 1)?.formula).toBe(
			'[1]Sheet1!$A$1+ExternalTotal',
		)

		read.value.workbook.sheets[0]?.cells.set(2, 0, {
			value: numberValue(42),
			formula: null,
			styleId: S0,
		})
		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.definedNames.get('ExternalTotal')).toBe('[1]Sheet1!$A$1')
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 1)?.formula).toBe(
			'[1]Sheet1!$A$1+ExternalTotal',
		)
		expect(reopened.value.workbook.externalReferenceDetails).toEqual(
			read.value.workbook.externalReferenceDetails,
		)

		const afterGraph = inspectXlsxPackageGraph(written.value)
		expectNoPackageGraphIssues(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph))
		expectNoPackageGraphIssues(
			auditXlsxPackageGraphBytePreservation(beforeGraph, sourceBytes, written.value),
		)
		expect(externalLinkPartIdentities(afterGraph)).toEqual(externalLinkPartIdentities(beforeGraph))
	})

	test('rewrites only the externalBook-bound relationship target for a selected link', () => {
		const read = readXlsx(externalLinkPackageFixture())
		expectOk(read)

		const rewritten = applyOperation(read.value.workbook, {
			op: 'rewriteExternalLink',
			partPath: 'xl/externalLinks/externalLink1.xml',
			linkRelId: 'rIdChosen',
			newTarget: '../sources/reforecast.xlsx',
		})
		expectOk(rewritten)
		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			workbookMetaDirty: true,
		})
		expectOk(written)

		const rels = decodeZipPart(written.value, 'xl/externalLinks/_rels/externalLink1.xml.rels')
		expect(rels).toContain('Id="rIdOther"')
		expect(rels).toContain('Target="../sources/wrong.xlsx"')
		expect(rels).toContain('Id="rIdChosen"')
		expect(rels).toContain('Target="../sources/reforecast.xlsx"')
		expect(decodeZipPart(written.value, 'xl/externalLinks/externalLink1.xml')).toContain(
			'r:id="rIdChosen"',
		)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.externalReferenceDetails[0]).toMatchObject({
			partPath: 'xl/externalLinks/externalLink1.xml',
			externalBookRelId: 'rIdChosen',
			linkRelId: 'rIdChosen',
			linkBindingStatus: 'externalBookRelId',
			target: '../sources/reforecast.xlsx',
			targetMode: 'External',
		})
	})
})

function externalLinkPackageFixture(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="${EXTERNAL_LINK_CONTENT_TYPE}"/>
  <Override PartName="/xl/externalLinks/externalLink2.xml" ContentType="${EXTERNAL_LINK_CONTENT_TYPE}"/>
  <Override PartName="/xl/externalLinks/externalLink3.xml" ContentType="${EXTERNAL_LINK_CONTENT_TYPE}"/>
  <Override PartName="/xl/externalLinks/externalLink4.xml" ContentType="${EXTERNAL_LINK_CONTENT_TYPE}"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdExternal1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
  <Relationship Id="rIdExternal2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink2.xml"/>
  <Relationship Id="rIdExternal3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink3.xml"/>
  <Relationship Id="rIdExternal4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink4.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
  <definedNames>
    <definedName name="ExternalTotal">[1]Sheet1!$A$1</definedName>
  </definedNames>
  <externalReferences>
    <externalReference r:id="rIdExternal1"/>
    <externalReference r:id="rIdExternal2"/>
    <externalReference r:id="rIdExternal3"/>
    <externalReference r:id="rIdExternal4"/>
  </externalReferences>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>10</v></c>
      <c r="B1"><f>[1]Sheet1!$A$1+ExternalTotal</f><v>15</v></c>
    </row>
  </sheetData>
</worksheet>`,
		'xl/externalLinks/externalLink1.xml': externalLinkXml('rIdChosen'),
		'xl/externalLinks/_rels/externalLink1.xml.rels': relationshipsXml(`
  <Relationship Id="rIdOther" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="../sources/wrong.xlsx" TargetMode="External"/>
  <Relationship Id="rIdChosen" Type="http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup" Target="personal.xlsb" TargetMode="External"/>
`),
		'xl/externalLinks/externalLink2.xml': externalLinkXml('rIdMissing'),
		'xl/externalLinks/_rels/externalLink2.xml.rels': relationshipsXml(`
  <Relationship Id="rIdLibrary" Type="http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary" Target="library.xlsx" TargetMode="External"/>
  <Relationship Id="rIdFallback2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="fallback.xlsx" TargetMode="External"/>
`),
		'xl/externalLinks/externalLink3.xml': externalLinkXml(),
		'xl/externalLinks/_rels/externalLink3.xml.rels': relationshipsXml(`
  <Relationship Id="rIdPathMissing" Type="http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing" Target="missing-source.xlsx" TargetMode="External"/>
`),
		'xl/externalLinks/externalLink4.xml': externalLinkXml('rIdAbsent'),
		'xl/externalLinks/_rels/externalLink4.xml.rels': relationshipsXml(''),
	})
}

function externalLinkXml(externalBookRelId?: string): string {
	return `<?xml version="1.0"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook${externalBookRelId ? ` r:id="${externalBookRelId}"` : ''}>
    <sheetNames><sheetName val="Sheet1"/></sheetNames>
  </externalBook>
</externalLink>`
}

function relationshipsXml(inner: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${inner}</Relationships>`
}

function externalLinkPartIdentities(graph: XlsxPackageGraph): readonly unknown[] {
	return graph.parts
		.filter(
			(part) => part.featureFamily === 'preservedExternalLink' && !part.path.endsWith('.rels'),
		)
		.map((part) => ({
			path: part.path,
			contentType: part.contentType,
			contentTypeSource: part.contentTypeSource,
			ownerScope: part.ownerScope,
			sourceRelationshipPart: part.sourceRelationshipPart,
			sourceRelationshipId: part.sourceRelationshipId,
			sourceRelationshipType: part.sourceRelationshipType,
			sourceRelationshipRawTarget: part.sourceRelationshipRawTarget,
			sourceRelationshipResolvedTarget: part.sourceRelationshipResolvedTarget,
			featureFamily: part.featureFamily,
			preservationPolicy: part.preservationPolicy,
			bytePreservationExpected: part.bytePreservationExpected,
		}))
}

function externalLinkContentTypeOverrides(
	graph: XlsxPackageGraph,
): readonly XlsxPackageGraph['contentTypeOverrides'][number][] {
	return graph.contentTypeOverrides.filter((override) =>
		override.partPath.startsWith('xl/externalLinks/externalLink'),
	)
}

function decodeZipPart(bytes: Uint8Array, partPath: string): string {
	return new TextDecoder().decode(extractZip(bytes).readBytes(partPath) ?? new Uint8Array())
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

function expectNoPackageGraphIssues(issues: readonly XlsxPackageGraphFidelityIssue[]): void {
	expect(
		issues.map((issue) => ({
			code: issue.code,
			message: issue.message,
			partPath: issue.partPath,
			relationshipPartPath: issue.relationshipPartPath,
			relationshipId: issue.relationshipId,
			expected: issue.expected,
			actual: issue.actual,
		})),
	).toEqual([])
}
