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
			severity: 'error',
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			relationshipId: 'rIdMissing',
			featureFamily: 'worksheet',
			suggestedAction: expect.stringContaining('relationship target'),
		})
		expect(issues.find((issue) => issue.code === 'package_feature_classification')).toMatchObject({
			severity: 'warning',
			partPath: 'odd/package.bin',
			ownerScope: 'unknown',
			featureFamily: 'preservedOther',
		})
	})

	test('reports relationship sidecars whose source part is missing', () => {
		const graph = inspectXlsxPackageGraph(
			makeXlsx({
				'[Content_Types].xml': contentTypesXml(`
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
`),
				'_rels/.rels': relationshipsXml(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
				'xl/workbook.xml': '<workbook/>',
				'xl/media/image1.png': 'image-bytes',
				'xl/drawings/_rels/missingDrawing.xml.rels': relationshipsXml(`
  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
`),
			}),
		)

		expect(auditXlsxPackageGraphReadIntegrity(graph)).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_source',
				severity: 'error',
				sourcePartPath: 'xl/drawings/missingDrawing.xml',
				relationshipPartPath: 'xl/drawings/_rels/missingDrawing.xml.rels',
				relationshipId: 'rIdImage',
				featureFamily: 'preservedMedia',
				suggestedAction: expect.stringContaining('orphan relationship sidecar'),
			}),
		)
	})

	test('reports empty relationship sidecars whose source part is missing', () => {
		const graph = inspectXlsxPackageGraph(
			makeXlsx({
				'[Content_Types].xml': contentTypesXml(`
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
`),
				'_rels/.rels': relationshipsXml(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
				'xl/workbook.xml': '<workbook/>',
				'xl/externalLinks/_rels/externalLink9.xml.rels': relationshipsXml(''),
			}),
		)

		expect(auditXlsxPackageGraphReadIntegrity(graph)).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_source',
				severity: 'error',
				sourcePartPath: 'xl/externalLinks/externalLink9.xml',
				relationshipPartPath: 'xl/externalLinks/_rels/externalLink9.xml.rels',
				ownerScope: 'relationship-part',
				suggestedAction: expect.stringContaining('orphan relationship sidecar'),
			}),
		)
	})

	test('reports orphan comment relationship sidecars by their missing comment source parts', () => {
		const graph = inspectXlsxPackageGraph(
			makeXlsx({
				'[Content_Types].xml': contentTypesXml(`
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/drawings/vmlDrawing1.vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/xml"/>
`),
				'_rels/.rels': relationshipsXml(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
				'xl/workbook.xml': '<workbook/>',
				'xl/drawings/vmlDrawing1.vml': '<xml/>',
				'xl/metadata.xml': '<metadata/>',
				'xl/_rels/comments1.xml.rels': relationshipsXml(`
  <Relationship Id="rIdCommentVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="drawings/vmlDrawing1.vml"/>
`),
				'xl/threadedComments/_rels/threadedComment1.xml.rels': relationshipsXml(`
  <Relationship Id="rIdThreadedMeta" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="../metadata.xml"/>
`),
			}),
		)

		const issues = auditXlsxPackageGraphReadIntegrity(graph)
		expect(issues).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_source',
				severity: 'error',
				sourcePartPath: 'xl/comments1.xml',
				relationshipPartPath: 'xl/_rels/comments1.xml.rels',
				relationshipId: 'rIdCommentVml',
				featureFamily: 'preservedVml',
				suggestedAction: expect.stringContaining('orphan relationship sidecar'),
			}),
		)
		expect(issues).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_source',
				severity: 'error',
				sourcePartPath: 'xl/threadedComments/threadedComment1.xml',
				relationshipPartPath: 'xl/threadedComments/_rels/threadedComment1.xml.rels',
				relationshipId: 'rIdThreadedMeta',
				suggestedAction: expect.stringContaining('orphan relationship sidecar'),
			}),
		)
	})

	test('reports orphan chart sidecar relationship parts by their missing chart source parts', () => {
		const graph = inspectXlsxPackageGraph(
			makeXlsx({
				'[Content_Types].xml': contentTypesXml(`
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>
  <Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>
`),
				'_rels/.rels': relationshipsXml(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
				'xl/workbook.xml': '<workbook/>',
				'xl/charts/style1.xml': '<cs:chartStyle/>',
				'xl/charts/colors1.xml': '<cs:colors/>',
				'xl/charts/_rels/chart99.xml.rels': relationshipsXml(`
  <Relationship Id="rIdStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
  <Relationship Id="rIdColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
`),
			}),
		)

		const issues = auditXlsxPackageGraphReadIntegrity(graph)
		expect(issues).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_source',
				severity: 'error',
				sourcePartPath: 'xl/charts/chart99.xml',
				relationshipPartPath: 'xl/charts/_rels/chart99.xml.rels',
				relationshipId: 'rIdStyle',
				featureFamily: 'preservedChartStyle',
				suggestedAction: expect.stringContaining('orphan relationship sidecar'),
			}),
		)
		expect(issues).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_source',
				severity: 'error',
				sourcePartPath: 'xl/charts/chart99.xml',
				relationshipPartPath: 'xl/charts/_rels/chart99.xml.rels',
				relationshipId: 'rIdColors',
				featureFamily: 'preservedChartColor',
				suggestedAction: expect.stringContaining('orphan relationship sidecar'),
			}),
		)
	})

	test('accepts empty relationship sidecars whose source part exists', () => {
		const graph = inspectXlsxPackageGraph(
			makeXlsx({
				'[Content_Types].xml': contentTypesXml(`
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
`),
				'_rels/.rels': relationshipsXml(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
				'xl/workbook.xml': '<workbook/>',
				'xl/externalLinks/externalLink1.xml': '<externalLink/>',
				'xl/externalLinks/_rels/externalLink1.xml.rels': relationshipsXml(''),
			}),
		)

		expect(
			auditXlsxPackageGraphReadIntegrity(graph).filter(
				(issue) =>
					issue.code === 'package_relationship_source' &&
					issue.relationshipPartPath === 'xl/externalLinks/_rels/externalLink1.xml.rels',
			),
		).toHaveLength(0)
	})

	test('reports stale content type overrides whose target part is missing', () => {
		const graph = inspectXlsxPackageGraph(
			makeXlsx({
				'[Content_Types].xml': contentTypesXml(`
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
`),
				'_rels/.rels': relationshipsXml(`
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
`),
				'xl/workbook.xml': '<workbook/>',
			}),
		)

		expect(auditXlsxPackageGraphReadIntegrity(graph)).toContainEqual(
			expect.objectContaining({
				code: 'package_content_type_override_target',
				severity: 'error',
				partPath: 'xl/comments1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
				expected: 'xl/comments1.xml',
				actual: undefined,
				suggestedAction: expect.stringContaining('stale content type override'),
			}),
		)
	})

	test('reports content type overrides that disagree with resolved package part type', () => {
		const graph = {
			parts: [
				{
					path: 'xl/threadedComments/threadedComment1.xml',
					contentType: 'application/vnd.ms-excel.threadedcomments+xml',
					contentTypeSource: 'override',
					ownerScope: 'worksheet',
					featureFamily: 'preservedThreadedComments',
					preservationPolicy: 'preserve-exact',
					bytePreservationExpected: true,
				},
			],
			relationships: [],
			contentTypeDefaults: [],
			contentTypeOverrides: [
				{
					partPath: 'xl/threadedComments/threadedComment1.xml',
					contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
				},
			],
		} as const

		expect(auditXlsxPackageGraphReadIntegrity(graph)).toContainEqual(
			expect.objectContaining({
				code: 'package_content_type_override_mismatch',
				severity: 'error',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				featureFamily: 'preservedThreadedComments',
				ownerScope: 'worksheet',
				expected: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
				actual: 'application/vnd.ms-excel.threadedcomments+xml',
				suggestedAction: expect.stringContaining('content type override'),
			}),
		)
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
			severity: 'error',
			sourcePartPath: '',
			relationshipPartPath: '_rels/.rels',
			relationshipId: 'rIdOffice',
			featureFamily: 'workbook',
			suggestedAction: expect.stringContaining('raw type'),
		})
	})

	test('uses OPC source, id, target, and TargetMode fields for adversarial relationship drift', () => {
		const before = inspectXlsxPackageGraph(
			makeXlsx({
				...baseWorkbookParts(),
				...adversarialRelationshipParts({
					imageTarget: '../media/image%201.png',
					hyperlinkTarget: 'https://example.com/original',
					absoluteTarget: '/xl/charts/chart1.xml',
				}),
			}),
		)
		const after = inspectXlsxPackageGraph(
			makeXlsx({
				...baseWorkbookParts(),
				...adversarialRelationshipParts({
					imageTarget: '../media/image2.png',
					hyperlinkTarget: 'https://example.com/changed',
					absoluteTarget: '../charts/chart1.xml',
				}),
			}),
		)

		const issues = auditXlsxPackageGraphSafeEditIntegrity(before, after)
		expect(issues).toContainEqual(
			expect.objectContaining({
				code: 'package_preserved_relationship_identity',
				sourcePartPath: 'xl/drawings/drawing1.xml',
				relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
				relationshipId: 'rId20',
				featureFamily: 'preservedMedia',
				expected: expect.objectContaining({
					rawTarget: '../media/image%201.png',
					resolvedTarget: 'xl/media/image 1.png',
				}),
				actual: expect.objectContaining({
					rawTarget: '../media/image2.png',
					resolvedTarget: 'xl/media/image2.png',
				}),
			}),
		)
		expect(issues).toContainEqual(
			expect.objectContaining({
				code: 'package_preserved_relationship_identity',
				relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
				relationshipId: 'rId7',
				featureFamily: 'preservedOther',
				expected: expect.objectContaining({
					rawTarget: 'https://example.com/original',
					targetMode: 'External',
				}),
				actual: expect.objectContaining({
					rawTarget: 'https://example.com/changed',
					targetMode: 'External',
				}),
			}),
		)
		expect(issues).toContainEqual(
			expect.objectContaining({
				code: 'package_preserved_relationship_identity',
				relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
				relationshipId: 'rId100',
				featureFamily: 'preservedChart',
				expected: expect.objectContaining({
					rawTarget: '/xl/charts/chart1.xml',
					resolvedTarget: 'xl/charts/chart1.xml',
				}),
				actual: expect.objectContaining({
					rawTarget: '../charts/chart1.xml',
					resolvedTarget: 'xl/charts/chart1.xml',
				}),
			}),
		)
	})

	test('classifies external workbook path relationships as external-link metadata', () => {
		const graph = inspectXlsxPackageGraph(
			makeXlsx({
				...baseWorkbookParts({
					extraContentTypes: `
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
`,
				}),
				'xl/_rels/workbook.xml.rels': relationshipsXml(`
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
`),
				'xl/workbook.xml': '<workbook/>',
				'xl/externalLinks/externalLink1.xml':
					'<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
				'xl/externalLinks/_rels/externalLink1.xml.rels': relationshipsXml(`
  <Relationship Id="rIdPath" Type="http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing" Target="missing.xlsx" TargetMode="External"/>
`),
			}),
		)

		expect(graph.relationships).toContainEqual(
			expect.objectContaining({
				relationshipPartPath: 'xl/externalLinks/_rels/externalLink1.xml.rels',
				id: 'rIdPath',
				targetMode: 'External',
				featureFamily: 'preservedExternalLink',
			}),
		)
		expect(auditXlsxPackageGraphReadIntegrity(graph)).not.toContainEqual(
			expect.objectContaining({
				code: 'package_feature_classification',
				partPath: 'missing.xlsx',
			}),
		)
	})

	test('treats default-covered content type override removal as package graph drift', () => {
		const before = inspectXlsxPackageGraph(
			makeXlsx({
				...baseWorkbookParts({
					extraContentTypes: `
  <Override PartName="/xl/charts/style1.xml" ContentType="application/xml"/>
`,
				}),
				'xl/charts/style1.xml': '<cs:chartStyle/>',
			}),
		)
		const after = inspectXlsxPackageGraph(
			makeXlsx({
				...baseWorkbookParts(),
				'xl/charts/style1.xml': '<cs:chartStyle/>',
			}),
		)

		expect(auditXlsxPackageGraphSafeEditIntegrity(before, after)).toContainEqual(
			expect.objectContaining({
				code: 'package_content_type_override',
				partPath: 'xl/charts/style1.xml',
				featureFamily: 'preservedChartStyle',
				expected: {
					partPath: 'xl/charts/style1.xml',
					contentType: 'application/xml',
				},
				actual: undefined,
			}),
		)
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

function baseWorkbookParts(
	options: { readonly extraContentTypes?: string } = {},
): Record<string, string> {
	return {
		'[Content_Types].xml': contentTypesXml(`
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
${options.extraContentTypes ?? ''}
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

function adversarialRelationshipParts(options: {
	readonly imageTarget: string
	readonly hyperlinkTarget: string
	readonly absoluteTarget: string
}): Record<string, string> {
	return {
		'xl/worksheets/_rels/sheet1.xml.rels': relationshipsXml(`
  <Relationship Id="rId42" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${options.absoluteTarget}"/>
`),
		'xl/drawings/drawing1.xml': '<xdr:wsDr/>',
		'xl/drawings/_rels/drawing1.xml.rels': relationshipsXml(`
  <Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${options.imageTarget}"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${options.hyperlinkTarget}" TargetMode="External"/>
  <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
`),
		'xl/media/image 1.png': 'image-one',
		'xl/media/image2.png': 'image-two',
		'xl/charts/chart1.xml': '<c:chartSpace/>',
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
