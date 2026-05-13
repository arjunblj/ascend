import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { strToU8, unzipSync, zipSync } from 'fflate'
import type { StyleId } from '../../packages/core/src/index.ts'
import { applyOperation } from '../../packages/engine/src/index.ts'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
	inspectXlsxPackageGraph,
	readXlsx,
	writeXlsx,
	type XlsxPackageGraph,
	type XlsxPackageGraphFidelityIssue,
} from '../../packages/io-xlsx/src/index.ts'
import { stringValue } from '../../packages/schema/src/index.ts'

const S0 = 0 as StyleId
const COMMENT_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
const VML_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
const THREADED_COMMENT_CONTENT_TYPE = 'application/vnd.ms-excel.threadedcomments+xml'
const PERSON_CONTENT_TYPE = 'application/vnd.ms-excel.person+xml'

const legacyCommentFixtures = [
	{
		label: 'ClosedXML legacy comments',
		path: new URL('./closedxml/Comments_AddingComments.xlsx', import.meta.url),
		sheetName: 'Comments',
		expectedComments: 13,
		expectedCommentParts: 11,
		expectedVmlParts: 11,
	},
	{
		label: 'LibreOffice legacy comment',
		path: new URL('./libreoffice/tdf117287_comment.xlsx', import.meta.url),
		sheetName: 'Tabelle1',
		expectedComments: 1,
		expectedCommentParts: 1,
		expectedVmlParts: 1,
	},
] as const

describe('comment package graph fidelity', () => {
	for (const fixture of legacyCommentFixtures) {
		test(`${fixture.label} preserves comment sidecars after safe dirty edit`, () => {
			const sourceBytes = readFileSync(fixture.path)
			const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
			expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(beforeGraph))
			const beforeCommentParts = featurePartIdentities(beforeGraph, 'preservedComments')
			const beforeVmlParts = featurePartIdentities(beforeGraph, 'preservedVml')
			const beforeRelationships = commentRelationshipIdentities(beforeGraph)
			const beforeOverrides = featureContentTypeOverrides(beforeGraph, [
				'preservedComments',
				'preservedVml',
			])
			expect(beforeCommentParts).toHaveLength(fixture.expectedCommentParts)
			expect(beforeVmlParts).toHaveLength(fixture.expectedVmlParts)
			expect(beforeCommentParts.every((part) => part.contentType === COMMENT_CONTENT_TYPE)).toBe(
				true,
			)
			expect(beforeVmlParts.every((part) => part.contentType === VML_CONTENT_TYPE)).toBe(true)

			const opened = readXlsx(sourceBytes)
			expectOk(opened)
			const sheet = opened.value.workbook.sheets.find((entry) => entry.name === fixture.sheetName)
			expect(sheet).toBeDefined()
			if (!sheet) return
			expect(sheet.comments.size).toBe(fixture.expectedComments)

			sheet.cells.set(199, 30, {
				value: stringValue('__ascend_comment_package_safe_edit__'),
				formula: null,
				styleId: S0,
			})
			const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
				dirtySheetNames: [sheet.name],
			})
			expectOk(written)

			const reopened = readXlsx(written.value)
			expectOk(reopened)
			const reopenedSheet = reopened.value.workbook.sheets.find(
				(entry) => entry.name === fixture.sheetName,
			)
			expect(reopenedSheet?.comments.size).toBe(fixture.expectedComments)

			const afterGraph = inspectXlsxPackageGraph(written.value)
			expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(afterGraph))
			expectNoPackageGraphIssues(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph))
			expectFeatureBytesPreserved(beforeGraph, sourceBytes, written.value, 'preservedComments')
			expectFeatureBytesPreserved(beforeGraph, sourceBytes, written.value, 'preservedVml')
			expect(featurePartIdentities(afterGraph, 'preservedComments')).toEqual(beforeCommentParts)
			expect(featurePartIdentities(afterGraph, 'preservedVml')).toEqual(beforeVmlParts)
			expect(commentRelationshipIdentities(afterGraph)).toEqual(beforeRelationships)
			expect(
				featureContentTypeOverrides(afterGraph, ['preservedComments', 'preservedVml']),
			).toEqual(beforeOverrides)
		})
	}

	test('LibreOffice legacy comment text edits preserve VML layout and package identity', () => {
		const fixture = legacyCommentFixtures[1]
		const sourceBytes = readFileSync(fixture.path)
		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
		expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(beforeGraph))
		const beforeCommentParts = featurePartIdentities(beforeGraph, 'preservedComments')
		const beforeVmlParts = featurePartIdentities(beforeGraph, 'preservedVml')
		const beforeRelationships = commentRelationshipIdentities(beforeGraph)
		const beforeOverrides = featureContentTypeOverrides(beforeGraph, [
			'preservedComments',
			'preservedVml',
		])
		const commentPartPath = singleFeaturePartPath(beforeGraph, 'preservedComments')
		const vmlPartPath = singleFeaturePartPath(beforeGraph, 'preservedVml')
		const sourceCommentXml = decodeZipPart(sourceBytes, commentPartPath)
		const sourceVmlBytes = readZipPart(sourceBytes, vmlPartPath)

		const opened = readXlsx(sourceBytes)
		expectOk(opened)
		const sheet = opened.value.workbook.sheets.find((entry) => entry.name === fixture.sheetName)
		expect(sheet).toBeDefined()
		if (!sheet) return
		const original = sheet.comments.get('C9')
		expect(original).toMatchObject({
			text: 'visible comment',
			author: 'LO',
			legacyDrawing: expect.objectContaining({
				shapeId: '_x0000_s1025',
				row: 8,
				column: 2,
			}),
		})
		if (!original) return
		sheet.comments.set('C9', {
			...original,
			text: 'Visible comment updated by Ascend',
		})

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: [sheet.name],
		})
		expectOk(written)

		const writtenCommentXml = decodeZipPart(written.value, commentPartPath)
		expect(writtenCommentXml).toContain('Visible comment updated by Ascend')
		expect(writtenCommentXml).not.toEqual(sourceCommentXml)
		expect(readZipPart(written.value, vmlPartPath)).toEqual(sourceVmlBytes)

		const afterGraph = inspectXlsxPackageGraph(written.value)
		expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(afterGraph))
		expectNoPackageGraphIssues(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph))
		expect(featurePartIdentities(afterGraph, 'preservedComments')).toEqual(beforeCommentParts)
		expect(featurePartIdentities(afterGraph, 'preservedVml')).toEqual(beforeVmlParts)
		expect(commentRelationshipIdentities(afterGraph)).toEqual(beforeRelationships)
		expect(featureContentTypeOverrides(afterGraph, ['preservedComments', 'preservedVml'])).toEqual(
			beforeOverrides,
		)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedComment = reopened.value.workbook.getSheet(fixture.sheetName)?.comments.get('C9')
		expect(reopenedComment).toMatchObject({
			text: 'Visible comment updated by Ascend',
			author: 'LO',
			legacyDrawing: original.legacyDrawing,
		})
	})

	test('row and column structural edits rewrite comment sidecars coherently', () => {
		const legacyFixture = legacyCommentFixtures[1]
		const legacySourceBytes = readFileSync(legacyFixture.path)
		const legacyBeforeGraph = inspectXlsxPackageGraph(legacySourceBytes)
		expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(legacyBeforeGraph))
		const legacyBeforeCommentParts = featurePartIdentities(legacyBeforeGraph, 'preservedComments')
		const legacyBeforeVmlParts = featurePartIdentities(legacyBeforeGraph, 'preservedVml')
		const legacyBeforeRelationships = commentRelationshipIdentities(legacyBeforeGraph)
		const legacyBeforeOverrides = featureContentTypeOverrides(legacyBeforeGraph, [
			'preservedComments',
			'preservedVml',
		])
		const legacyCommentPartPath = singleFeaturePartPath(legacyBeforeGraph, 'preservedComments')
		const legacyVmlPartPath = singleFeaturePartPath(legacyBeforeGraph, 'preservedVml')
		const legacySourceEntries = unzipSync(legacySourceBytes)

		const legacyOpened = readXlsx(legacySourceBytes)
		expectOk(legacyOpened)
		expectOk(
			applyOperation(legacyOpened.value.workbook, {
				op: 'insertRows',
				sheet: legacyFixture.sheetName,
				at: 0,
				count: 1,
			}),
		)
		expectOk(
			applyOperation(legacyOpened.value.workbook, {
				op: 'insertCols',
				sheet: legacyFixture.sheetName,
				at: 1,
				count: 1,
			}),
		)
		const legacySheet = legacyOpened.value.workbook.getSheet(legacyFixture.sheetName)
		expect(legacySheet).toBeDefined()
		if (!legacySheet) return
		expect([...legacySheet.comments.keys()]).toEqual(['D10'])
		for (const [ref, comment] of legacySheet.comments) {
			const drawing = comment.legacyDrawing
			expect(drawing).toBeDefined()
			if (!drawing) return
			expect(`${String.fromCharCode(65 + drawing.column)}${drawing.row + 1}`).toBe(ref)
		}

		const legacyWritten = writeXlsx(legacyOpened.value.workbook, legacyOpened.value.capsules, {
			dirtySheetNames: [legacyFixture.sheetName],
		})
		expectOk(legacyWritten)
		const legacyWrittenEntries = unzipSync(legacyWritten.value)
		expect(legacyWrittenEntries[legacyCommentPartPath]).not.toEqual(
			legacySourceEntries[legacyCommentPartPath],
		)
		expect(legacyWrittenEntries[legacyVmlPartPath]).not.toEqual(
			legacySourceEntries[legacyVmlPartPath],
		)
		const legacyCommentXml = decodeZipPart(legacyWritten.value, legacyCommentPartPath)
		expect(legacyCommentXml).toContain('ref="D10"')
		expect(legacyCommentXml).not.toContain('ref="C9"')
		const legacyVmlXml = decodeZipPart(legacyWritten.value, legacyVmlPartPath)
		expect(legacyVmlXml).toContain('<x:Row>9</x:Row><x:Column>3</x:Column>')

		const legacyAfterGraph = inspectXlsxPackageGraph(legacyWritten.value)
		expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(legacyAfterGraph))
		expectNoPackageGraphIssues(
			auditXlsxPackageGraphSafeEditIntegrity(legacyBeforeGraph, legacyAfterGraph),
		)
		expect(featurePartIdentities(legacyAfterGraph, 'preservedComments')).toEqual(
			legacyBeforeCommentParts,
		)
		expect(featurePartIdentities(legacyAfterGraph, 'preservedVml')).toEqual(legacyBeforeVmlParts)
		expect(commentRelationshipIdentities(legacyAfterGraph)).toEqual(legacyBeforeRelationships)
		expect(
			featureContentTypeOverrides(legacyAfterGraph, ['preservedComments', 'preservedVml']),
		).toEqual(legacyBeforeOverrides)

		const legacyReopened = readXlsx(legacyWritten.value)
		expectOk(legacyReopened)
		const legacyReopenedSheet = legacyReopened.value.workbook.getSheet(legacyFixture.sheetName)
		expect([...(legacyReopenedSheet?.comments.keys() ?? [])]).toEqual(['D10'])
		expect(legacyReopenedSheet?.comments.get('D10')?.legacyDrawing).toMatchObject({
			row: 9,
			column: 3,
		})

		const threadedSourceBytes = threadedCommentWorkbook()
		const threadedSourceEntries = unzipSync(threadedSourceBytes)
		const threadedBeforeGraph = inspectXlsxPackageGraph(threadedSourceBytes)
		expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(threadedBeforeGraph))
		const threadedBeforeParts = featurePartIdentities(
			threadedBeforeGraph,
			'preservedThreadedComments',
		)
		const threadedBeforeRelationships = commentRelationshipIdentities(threadedBeforeGraph)
		const threadedBeforeOverrides = featureContentTypeOverrides(threadedBeforeGraph, [
			'preservedThreadedComments',
		])

		const threadedOpened = readXlsx(threadedSourceBytes)
		expectOk(threadedOpened)
		expectOk(
			applyOperation(threadedOpened.value.workbook, {
				op: 'insertRows',
				sheet: 'Sheet1',
				at: 0,
				count: 1,
			}),
		)
		expectOk(
			applyOperation(threadedOpened.value.workbook, {
				op: 'insertCols',
				sheet: 'Sheet1',
				at: 0,
				count: 1,
			}),
		)
		const threadedSheet = threadedOpened.value.workbook.getSheet('Sheet1')
		expect(threadedSheet?.threadedComments.map((comment) => comment.ref)).toEqual(['B2', 'B2'])

		const threadedWritten = writeXlsx(
			threadedOpened.value.workbook,
			threadedOpened.value.capsules,
			{ dirtySheetNames: ['Sheet1'] },
		)
		expectOk(threadedWritten)
		const threadedWrittenEntries = unzipSync(threadedWritten.value)
		expect(threadedWrittenEntries['xl/threadedComments/threadedComment1.xml']).not.toEqual(
			threadedSourceEntries['xl/threadedComments/threadedComment1.xml'],
		)
		expect(threadedWrittenEntries['xl/persons/person.xml']).toEqual(
			threadedSourceEntries['xl/persons/person.xml'],
		)
		const threadedCommentXml = decodeZipPart(
			threadedWritten.value,
			'xl/threadedComments/threadedComment1.xml',
		)
		expect(threadedCommentXml).toContain('ref="B2"')
		expect(threadedCommentXml).not.toContain('ref="A1"')

		const threadedAfterGraph = inspectXlsxPackageGraph(threadedWritten.value)
		expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(threadedAfterGraph))
		expectNoPackageGraphIssues(
			auditXlsxPackageGraphSafeEditIntegrity(threadedBeforeGraph, threadedAfterGraph),
		)
		expect(featurePartIdentities(threadedAfterGraph, 'preservedThreadedComments')).toEqual(
			threadedBeforeParts,
		)
		expect(commentRelationshipIdentities(threadedAfterGraph)).toEqual(threadedBeforeRelationships)
		expect(featureContentTypeOverrides(threadedAfterGraph, ['preservedThreadedComments'])).toEqual(
			threadedBeforeOverrides,
		)

		const threadedReopened = readXlsx(threadedWritten.value)
		expectOk(threadedReopened)
		expect(
			threadedReopened.value.workbook
				.getSheet('Sheet1')
				?.threadedComments.map(({ ref, text, id, parentId, personId, author, done }) => ({
					ref,
					text,
					id,
					parentId,
					personId,
					author,
					done,
				})),
		).toEqual([
			{
				ref: 'B2',
				text: 'Please review',
				id: 'tc1',
				parentId: undefined,
				personId: '0',
				author: 'Ada Lovelace',
				done: undefined,
			},
			{
				ref: 'B2',
				text: 'Reviewed',
				id: 'tc2',
				parentId: 'tc1',
				personId: '1',
				author: 'Grace Hopper',
				done: true,
			},
		])
	})

	test('synthetic threaded comments preserve thread and person sidecars after safe dirty edit', () => {
		const sourceBytes = threadedCommentWorkbook()
		const sourceEntries = unzipSync(sourceBytes)
		const beforeGraph = inspectXlsxPackageGraph(sourceBytes)
		expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(beforeGraph))
		const beforeThreadedParts = featurePartIdentities(beforeGraph, 'preservedThreadedComments')
		const beforeRelationships = commentRelationshipIdentities(beforeGraph)
		const beforeOverrides = featureContentTypeOverrides(beforeGraph, ['preservedThreadedComments'])
		expect(beforeThreadedParts).toEqual([
			expect.objectContaining({
				path: 'xl/persons/person.xml',
				contentType: PERSON_CONTENT_TYPE,
			}),
			expect.objectContaining({
				path: 'xl/threadedComments/threadedComment1.xml',
				contentType: THREADED_COMMENT_CONTENT_TYPE,
			}),
		])

		const opened = readXlsx(sourceBytes)
		expectOk(opened)
		const sheet = opened.value.workbook.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		expect(sheet.threadedComments).toEqual([
			{
				ref: 'A1',
				text: 'Please review',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				id: 'tc1',
				personId: '0',
				author: 'Ada Lovelace',
				dateTime: '2024-01-01T00:00:00.000',
			},
			{
				ref: 'A1',
				text: 'Reviewed',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				id: 'tc2',
				parentId: 'tc1',
				personId: '1',
				author: 'Grace Hopper',
				dateTime: '2024-01-02T00:00:00.000',
				done: true,
			},
		])

		sheet.cells.set(9, 0, {
			value: stringValue('__ascend_threaded_comment_package_safe_edit__'),
			formula: null,
			styleId: S0,
		})
		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: [sheet.name],
		})
		expectOk(written)

		const writtenEntries = unzipSync(written.value)
		expect(writtenEntries['xl/threadedComments/threadedComment1.xml']).toEqual(
			sourceEntries['xl/threadedComments/threadedComment1.xml'],
		)
		expect(writtenEntries['xl/persons/person.xml']).toEqual(sourceEntries['xl/persons/person.xml'])

		const afterGraph = inspectXlsxPackageGraph(written.value)
		expectNoPackageGraphIssues(auditXlsxPackageGraphReadIntegrity(afterGraph))
		expectNoPackageGraphIssues(auditXlsxPackageGraphSafeEditIntegrity(beforeGraph, afterGraph))
		expectNoPackageGraphIssues(
			auditXlsxPackageGraphBytePreservation(beforeGraph, sourceBytes, written.value),
		)
		expect(featurePartIdentities(afterGraph, 'preservedThreadedComments')).toEqual(
			beforeThreadedParts,
		)
		expect(commentRelationshipIdentities(afterGraph)).toEqual(beforeRelationships)
		expect(featureContentTypeOverrides(afterGraph, ['preservedThreadedComments'])).toEqual(
			beforeOverrides,
		)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedSheet = reopened.value.workbook.getSheet('Sheet1')
		expect(reopenedSheet?.threadedComments).toEqual(sheet.threadedComments)
		expect(reopenedSheet?.cells.get(9, 0)?.value).toEqual({
			kind: 'string',
			value: '__ascend_threaded_comment_package_safe_edit__',
		})
	})
})

function threadedCommentWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
  <Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPackage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdPerson" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdThreaded" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
</Relationships>`,
		'xl/persons/person.xml': `<?xml version="1.0"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="0" displayName="Ada Lovelace"/>
  <person id="1" displayName="Grace Hopper"/>
</personList>`,
		'xl/threadedComments/threadedComment1.xml': `<?xml version="1.0"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="0" id="tc1" dT="2024-01-01T00:00:00.000">
    <text>Please review</text>
  </threadedComment>
  <threadedComment ref="A1" personId="1" id="tc2" parentId="tc1" dT="2024-01-02T00:00:00.000" done="1">
    <text>Reviewed</text>
  </threadedComment>
</ThreadedComments>`,
	})
}

function makeXlsx(parts: Record<string, string>): Uint8Array {
	const entries: Record<string, Uint8Array> = {}
	for (const [path, content] of Object.entries(parts)) entries[path] = strToU8(content)
	return zipSync(entries)
}

function featurePartIdentities(
	graph: XlsxPackageGraph,
	featureFamily: string,
): readonly Record<string, unknown>[] {
	return graph.parts
		.filter((part) => part.featureFamily === featureFamily)
		.map((part) => ({
			path: part.path,
			contentType: part.contentType,
			contentTypeSource: part.contentTypeSource,
			ownerScope: part.ownerScope,
			sourceRelationshipPart: part.sourceRelationshipPart,
			sourceRelationshipId: part.sourceRelationshipId,
			sourceRelationshipType: part.sourceRelationshipType,
			sourceRelationshipRawType: part.sourceRelationshipRawType,
			sourceRelationshipRawTarget: part.sourceRelationshipRawTarget,
			sourceRelationshipResolvedTarget: part.sourceRelationshipResolvedTarget,
			preservationPolicy: part.preservationPolicy,
			bytePreservationExpected: part.bytePreservationExpected,
		}))
		.sort(compareJson)
}

function commentRelationshipIdentities(
	graph: XlsxPackageGraph,
): readonly Record<string, unknown>[] {
	return graph.relationships
		.filter((relationship) =>
			['preservedComments', 'preservedVml', 'preservedThreadedComments'].includes(
				relationship.featureFamily ?? '',
			),
		)
		.map((relationship) => ({
			sourcePartPath: relationship.sourcePartPath,
			relationshipPartPath: relationship.relationshipPartPath,
			id: relationship.id,
			type: relationship.type,
			rawType: relationship.rawType,
			rawTarget: relationship.rawTarget,
			resolvedTarget: relationship.resolvedTarget,
			targetMode: relationship.targetMode,
			featureFamily: relationship.featureFamily,
		}))
		.sort(compareJson)
}

function featureContentTypeOverrides(
	graph: XlsxPackageGraph,
	featureFamilies: readonly string[],
): readonly Record<string, unknown>[] {
	const paths = new Set(
		graph.parts
			.filter((part) => featureFamilies.includes(part.featureFamily ?? ''))
			.map((part) => part.path),
	)
	return graph.contentTypeOverrides
		.filter((override) => paths.has(override.partPath))
		.map((override) => ({
			partPath: override.partPath,
			contentType: override.contentType,
		}))
		.sort(compareJson)
}

function expectFeatureBytesPreserved(
	graph: XlsxPackageGraph,
	sourceBytes: Uint8Array,
	writtenBytes: Uint8Array,
	featureFamily: string,
): void {
	for (const part of graph.parts.filter((entry) => entry.featureFamily === featureFamily)) {
		const before = readZipPart(sourceBytes, part.path)
		const after = readZipPart(writtenBytes, part.path)
		expect(after).toEqual(before)
		expect(part.preservationPolicy).toBe('preserve-exact')
		expect(part.bytePreservationExpected).toBe(true)
	}
}

function singleFeaturePartPath(graph: XlsxPackageGraph, featureFamily: string): string {
	const paths = graph.parts
		.filter((entry) => entry.featureFamily === featureFamily)
		.map((entry) => entry.path)
	expect(paths).toHaveLength(1)
	const [path] = paths
	if (!path) throw new Error(`Missing ${featureFamily} package part`)
	return path
}

function decodeZipPart(bytes: Uint8Array, path: string): string {
	return new TextDecoder().decode(readZipPart(bytes, path))
}

function readZipPart(bytes: Uint8Array, path: string): Uint8Array {
	const entry = unzipSync(bytes)[path]
	expect(entry).toBeDefined()
	if (!entry) throw new Error(`Missing ZIP part ${path}`)
	return entry
}

function expectNoPackageGraphIssues(issues: readonly XlsxPackageGraphFidelityIssue[]): void {
	if (issues.length === 0) return
	throw new Error(issues.map(formatPackageGraphIssue).join('\n'))
}

function formatPackageGraphIssue(issue: XlsxPackageGraphFidelityIssue): string {
	return [
		issue.message,
		`severity=${issue.severity}`,
		issue.partPath ? `part=${issue.partPath}` : undefined,
		issue.sourcePartPath !== undefined
			? `source=${issue.sourcePartPath || '<package>'}`
			: undefined,
		issue.relationshipPartPath ? `rels=${issue.relationshipPartPath}` : undefined,
		issue.relationshipId ? `relId=${issue.relationshipId}` : undefined,
		issue.ownerScope ? `owner=${issue.ownerScope}` : undefined,
		issue.featureFamily ? `family=${issue.featureFamily}` : undefined,
		issue.suggestedAction ? `action=${issue.suggestedAction}` : undefined,
		issue.expected !== undefined ? `expected=${JSON.stringify(issue.expected)}` : undefined,
		issue.actual !== undefined ? `actual=${JSON.stringify(issue.actual)}` : undefined,
	]
		.filter((part): part is string => part !== undefined)
		.join(' | ')
}

function compareJson(left: Record<string, unknown>, right: Record<string, unknown>): number {
	return JSON.stringify(left).localeCompare(JSON.stringify(right))
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}
