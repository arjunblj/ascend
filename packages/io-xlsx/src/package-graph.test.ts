import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../test/helpers.ts'
import { inspectXlsxPackageGraph } from './package-graph.ts'

describe('XLSX package graph', () => {
	test('normalizes content types, relationship identity, owners, and feature families', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>
  <Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>
  <Override PartName="/xl/xmlMaps.xml" ContentType="application/xml"/>
  <Override PartName="/xl/customProperty1.bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty"/>
  <Override PartName="/xl/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>
  <Override PartName="/xl/model/item.data" ContentType="application/vnd.ms-excel.model"/>
  <Override PartName="/xl/customData/item1.data" ContentType="application/vnd.ms-excel.customData"/>
  <Override PartName="/xl/data/connectionsPayload.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>
  <Override PartName="/xl/revisions/revisionHeaders.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': '<workbook/>',
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdXmlMaps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps" Target="xmlMaps.xml"/>
  <Relationship Id="rIdCustomProperty" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty" Target="customProperty1.bin"/>
  <Relationship Id="rIdDiagramData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>
  <Relationship Id="rIdDataModel" Type="http://schemas.microsoft.com/office/2011/relationships/model" Target="model/item.data"/>
  <Relationship Id="rIdPowerQuery" Type="http://schemas.microsoft.com/office/2014/relationships/powerQueryMashup" Target="customData/item1.data"/>
  <Relationship Id="rIdConnections" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="data/connectionsPayload.xml"/>
  <Relationship Id="rIdRevisionHeaders" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders" Target="revisions/revisionHeaders.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': '<worksheet/>',
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
			'xl/drawings/drawing1.xml': '<xdr:wsDr/>',
			'xl/drawings/_rels/drawing1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image%201.png"/>
</Relationships>`,
			'xl/charts/chart1.xml': '<c:chartSpace/>',
			'xl/charts/style1.xml': '<cs:chartStyle/>',
			'xl/charts/colors1.xml': '<cs:colors/>',
			'xl/xmlMaps.xml': '<xmlMaps/>',
			'xl/customProperty1.bin': 'custom-property-bytes',
			'xl/diagrams/data1.xml': '<dgm:dataModel/>',
			'xl/model/item.data': 'data-model-bytes',
			'xl/customData/item1.data': 'power-query-bytes',
			'xl/data/connectionsPayload.xml': '<connections/>',
			'xl/revisions/revisionHeaders.xml': '<headers/>',
			'xl/media/image 1.png': 'not-really-a-png',
		})

		const graph = inspectXlsxPackageGraph(bytes)
		expect(graph.contentTypeDefaults).toContainEqual({
			extension: 'png',
			contentType: 'image/png',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: '',
			relationshipPartPath: '_rels/.rels',
			id: 'rIdOffice',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
			rawType: 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument',
			rawTarget: 'xl/workbook.xml',
			resolvedTarget: 'xl/workbook.xml',
			featureFamily: 'workbook',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			id: 'rIdConnections',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections',
			rawTarget: 'data/connectionsPayload.xml',
			resolvedTarget: 'xl/data/connectionsPayload.xml',
			featureFamily: 'preservedConnection',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/drawings/drawing1.xml',
			relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
			id: 'rIdImage',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
			rawTarget: '../media/image%201.png',
			resolvedTarget: 'xl/media/image 1.png',
			featureFamily: 'preservedMedia',
		})
		expect(graph.parts.find((part) => part.path === 'xl/workbook.xml')).toMatchObject({
			contentTypeSource: 'override',
			ownerScope: 'workbook',
			sourceRelationshipId: 'rIdOffice',
			sourceRelationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
			sourceRelationshipRawType:
				'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument',
			featureFamily: 'workbook',
		})
		expect(graph.parts.find((part) => part.path === 'xl/worksheets/sheet1.xml')).toMatchObject({
			ownerScope: 'worksheet',
			sourceRelationshipId: 'rIdSheet',
			featureFamily: 'worksheet',
		})
		expect(graph.parts.find((part) => part.path === 'xl/charts/style1.xml')).toMatchObject({
			featureFamily: 'preservedChartStyle',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/charts/colors1.xml')).toMatchObject({
			featureFamily: 'preservedChartColor',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/media/image 1.png')).toMatchObject({
			contentType: 'image/png',
			contentTypeSource: 'default',
			ownerScope: 'drawing',
			featureFamily: 'preservedMedia',
		})
		expect(graph.parts.find((part) => part.path === 'xl/xmlMaps.xml')).toMatchObject({
			sourceRelationshipId: 'rIdXmlMaps',
			featureFamily: 'preservedCustomXml',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/customProperty1.bin')).toMatchObject({
			sourceRelationshipId: 'rIdCustomProperty',
			featureFamily: 'preservedMetadata',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/diagrams/data1.xml')).toMatchObject({
			sourceRelationshipId: 'rIdDiagramData',
			featureFamily: 'preservedDrawing',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/model/item.data')).toMatchObject({
			sourceRelationshipId: 'rIdDataModel',
			featureFamily: 'preservedDataModel',
			preservationPolicy: 'inspect-only',
			bytePreservationExpected: true,
		})
		expect(graph.parts.find((part) => part.path === 'xl/customData/item1.data')).toMatchObject({
			sourceRelationshipId: 'rIdPowerQuery',
			featureFamily: 'preservedPowerQuery',
			preservationPolicy: 'inspect-only',
			bytePreservationExpected: true,
		})
		expect(
			graph.parts.find((part) => part.path === 'xl/data/connectionsPayload.xml'),
		).toMatchObject({
			sourceRelationshipId: 'rIdConnections',
			featureFamily: 'preservedConnection',
			preservationPolicy: 'preserve-exact',
		})
		expect(
			graph.parts.find((part) => part.path === 'xl/revisions/revisionHeaders.xml'),
		).toMatchObject({
			sourceRelationshipId: 'rIdRevisionHeaders',
			featureFamily: 'preservedRevision',
			preservationPolicy: 'inspect-only',
			bytePreservationExpected: true,
		})
	})

	test('classifies vendor security and orphan worksheet sidecar parts without preservedOther fallback', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="xen" ContentType="application/octet-stream"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rDellEncryptedDoc" Type="http://schemas.dell.com/ddp/2016/relationships/xenFile" Target="ddp/ddpfile.xen"/>
  <Relationship Id="rDellMetadataFile" Type="http://schemas.dell.com/ddp/2016/relationships/metadataFile" Target="docProps/metadata.xml"/>
</Relationships>`,
			'xl/workbook.xml': '<workbook/>',
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': '<worksheet/>',
			'xl/worksheets/sheet1_formatted.xml': '<worksheet/>',
			'ddp/ddpfile.xen': 'opaque-vendor-security',
			'docProps/metadata.xml': '<items/>',
		})

		const graph = inspectXlsxPackageGraph(bytes)
		expect(graph.parts.find((part) => part.path === 'ddp/ddpfile.xen')).toMatchObject({
			ownerScope: 'security',
			sourceRelationshipId: 'rDellEncryptedDoc',
			featureFamily: 'preservedVendorSecurity',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'docProps/metadata.xml')).toMatchObject({
			ownerScope: 'security',
			sourceRelationshipId: 'rDellMetadataFile',
			featureFamily: 'preservedVendorSecurity',
			preservationPolicy: 'preserve-exact',
		})
		expect(
			graph.parts.find((part) => part.path === 'xl/worksheets/sheet1_formatted.xml'),
		).toMatchObject({
			ownerScope: 'unknown',
			featureFamily: 'preservedWorksheetSidecar',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.some((part) => part.featureFamily === 'preservedOther')).toBe(false)
	})

	test('reads content type attributes with XML whitespace, quotes, and entities', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension = 'rels' ContentType = 'application/vnd.openxmlformats-package.relationships+xml'/>
  <Default Extension = 'bin' ContentType = 'application/octet-stream'/>
  <Override NotPartName="/ignored.xml" PartName = '/xl/workbook.xml' ContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/>
  <Override PartName = '/xl/media/a&amp;b.bin' ContentType = 'application/vnd.example.opaque'/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': '<workbook/>',
			'xl/media/a&b.bin': 'opaque',
		})

		const graph = inspectXlsxPackageGraph(bytes)

		expect(graph.contentTypeDefaults).toContainEqual({
			extension: 'bin',
			contentType: 'application/octet-stream',
		})
		expect(graph.parts.find((part) => part.path === 'xl/workbook.xml')).toMatchObject({
			contentTypeSource: 'override',
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
		})
		expect(graph.parts.find((part) => part.path === 'xl/media/a&b.bin')).toMatchObject({
			contentTypeSource: 'override',
			contentType: 'application/vnd.example.opaque',
		})
	})
})
