import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../test/helpers.ts'
import { inspectXlsxPackageGraph } from './package-graph.ts'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
} from './package-graph-fidelity.ts'

describe('XLSX package graph fidelity audits', () => {
	test('reports unclassified parts and missing internal relationship targets', () => {
		const graph = inspectXlsxPackageGraph(
			makeXlsx({
				'[Content_Types].xml': contentTypesXml(`
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/odd/package.bin" ContentType="application/octet-stream"/>
`),
				'_rels/.rels': relationshipsXml(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
				'xl/workbook.xml': '<workbook/>',
				'xl/_rels/workbook.xml.rels': relationshipsXml(`
  <Relationship Id="rIdMissing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/missing.xml"/>
`),
				'odd/package.bin': 'opaque',
			}),
		)

		const issues = auditXlsxPackageGraphReadIntegrity(graph)
		expect(issues.map((issue) => issue.code)).toContain('package_feature_classification')
		expect(issues.map((issue) => issue.code)).toContain('package_relationship_target')
		expect(issues.find((issue) => issue.code === 'package_relationship_target')).toMatchObject({
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			relationshipId: 'rIdMissing',
		})
	})

	test('compares relationship identity with raw strict dialect and package scope', () => {
		const before = inspectXlsxPackageGraph(
			makeXlsx({
				...baseWorkbookParts(),
				'_rels/.rels': relationshipsXml(`
  <Relationship Id="rIdOffice" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
			}),
		)
		const after = inspectXlsxPackageGraph(
			makeXlsx({
				...baseWorkbookParts(),
				'_rels/.rels': relationshipsXml(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
			}),
		)

		const issues = auditXlsxPackageGraphSafeEditIntegrity(before, after)
		expect(issues.map((issue) => issue.code)).toContain('package_preserved_relationship_identity')
		expect(
			issues.find((issue) => issue.code === 'package_preserved_relationship_identity'),
		).toMatchObject({
			relationshipPartPath: '_rels/.rels',
			relationshipId: 'rIdOffice',
		})
	})

	test('detects byte drift for preserve-exact sidecars', () => {
		const beforeBytes = makeXlsx({
			...baseWorkbookParts(),
			...chartSidecarParts('<c:chartSpace><c:chart/></c:chartSpace>'),
		})
		const afterBytes = makeXlsx({
			...baseWorkbookParts(),
			...chartSidecarParts('<c:chartSpace><c:edited/></c:chartSpace>'),
		})
		const before = inspectXlsxPackageGraph(beforeBytes)

		const issues = auditXlsxPackageGraphBytePreservation(before, beforeBytes, afterBytes)
		expect(issues).toContainEqual(
			expect.objectContaining({
				code: 'package_preserved_part_bytes',
				partPath: 'xl/charts/chart1.xml',
			}),
		)
	})
})

function baseWorkbookParts(): Record<string, string> {
	return {
		'[Content_Types].xml': contentTypesXml(`
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
`),
		'_rels/.rels': relationshipsXml(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
		'xl/workbook.xml': '<workbook/>',
		'xl/_rels/workbook.xml.rels': relationshipsXml(`
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
`),
		'xl/worksheets/sheet1.xml': '<worksheet/>',
	}
}

function chartSidecarParts(chartXml: string): Record<string, string> {
	return {
		'xl/worksheets/_rels/sheet1.xml.rels': relationshipsXml(`
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
`),
		'xl/drawings/drawing1.xml': '<xdr:wsDr/>',
		'xl/drawings/_rels/drawing1.xml.rels': relationshipsXml(`
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
`),
		'xl/charts/chart1.xml': chartXml,
	}
}

function contentTypesXml(inner: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${inner}
</Types>`
}

function relationshipsXml(inner: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}
</Relationships>`
}
