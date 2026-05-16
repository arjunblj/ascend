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
  <Override PartName="/xl/links/bookLink.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
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
  <Relationship Id="rIdPowerQueryOpaque" Type="http://schemas.microsoft.com/office/2014/relationships/powerQueryMashup" Target="opaque-mashup.bin"/>
  <Relationship Id="rIdDataModelOpaque" Type="http://schemas.microsoft.com/office/2011/relationships/model" Target="opaque-model.bin"/>
  <Relationship Id="rIdCalcChainOpaque" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="opaque-calc.bin"/>
  <Relationship Id="rIdSheetMetadataOpaque" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata" Target="opaque-sheet-metadata.bin"/>
  <Relationship Id="rIdCustomXmlOpaque" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="opaque-custom.xml"/>
  <Relationship Id="rIdVbaOpaque" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="opaque-vba.bin"/>
  <Relationship Id="rIdVbaSignatureOpaque" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="opaque-signature.bin"/>
  <Relationship Id="rIdConnections" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="data/connectionsPayload.xml"/>
  <Relationship Id="rIdExternalLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="links/bookLink.xml"/>
  <Relationship Id="rIdRevisionHeaders" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders" Target="revisions/revisionHeaders.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': '<worksheet/>',
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rIdHyperlink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/report" TargetMode="External"/>
  <Relationship Id="rIdControlProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="https://example.invalid/control.xml" TargetMode="External"/>
  <Relationship Id="rIdControlPropsOpaque" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="opaque-control.bin"/>
  <Relationship Id="rIdActiveX" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="https://example.invalid/control.ocx" TargetMode="External"/>
</Relationships>`,
			'xl/drawings/drawing1.xml': '<xdr:wsDr/>',
			'xl/drawings/_rels/drawing1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image%201.png"/>
  <Relationship Id="rIdLinkedImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.invalid/logo.png" TargetMode="External"/>
  <Relationship Id="rIdOleObject" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="file:///C:/objects/report.bin" TargetMode="External"/>
  <Relationship Id="rIdOleObjectOpaque" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="opaque-ole.bin"/>
  <Relationship Id="rIdOpaqueChartStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style-package.bin"/>
  <Relationship Id="rIdOpaqueChartColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="color-package.bin"/>
</Relationships>`,
			'xl/charts/chart1.xml': '<c:chartSpace/>',
			'xl/charts/style1.xml': '<cs:chartStyle/>',
			'xl/charts/colors1.xml': '<cs:colors/>',
			'xl/xmlMaps.xml': '<xmlMaps/>',
			'xl/customProperty1.bin': 'custom-property-bytes',
			'xl/diagrams/data1.xml': '<dgm:dataModel/>',
			'xl/model/item.data': 'data-model-bytes',
			'xl/customData/item1.data': 'power-query-bytes',
			'xl/opaque-calc.bin': 'calc-chain-bytes',
			'xl/opaque-sheet-metadata.bin': 'metadata-bytes',
			'xl/opaque-custom.xml': '<custom/>',
			'xl/opaque-vba.bin': 'vba-bytes',
			'xl/opaque-signature.bin': 'signature-bytes',
			'xl/data/connectionsPayload.xml': '<connections/>',
			'xl/links/bookLink.xml': '<externalLink/>',
			'xl/revisions/revisionHeaders.xml': '<headers/>',
			'xl/worksheets/opaque-control.bin': 'control-bytes',
			'xl/drawings/opaque-ole.bin': 'ole-bytes',
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
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			id: 'rIdExternalLink',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
			rawTarget: 'links/bookLink.xml',
			resolvedTarget: 'xl/links/bookLink.xml',
			featureFamily: 'preservedExternalLink',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			id: 'rIdPowerQueryOpaque',
			type: 'http://schemas.microsoft.com/office/2014/relationships/powerQueryMashup',
			rawTarget: 'opaque-mashup.bin',
			resolvedTarget: 'xl/opaque-mashup.bin',
			featureFamily: 'preservedPowerQuery',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			id: 'rIdDataModelOpaque',
			type: 'http://schemas.microsoft.com/office/2011/relationships/model',
			rawTarget: 'opaque-model.bin',
			resolvedTarget: 'xl/opaque-model.bin',
			featureFamily: 'preservedDataModel',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			id: 'rIdCalcChainOpaque',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
			rawTarget: 'opaque-calc.bin',
			resolvedTarget: 'xl/opaque-calc.bin',
			featureFamily: 'preservedCalcChain',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			id: 'rIdSheetMetadataOpaque',
			type: 'http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata',
			rawTarget: 'opaque-sheet-metadata.bin',
			resolvedTarget: 'xl/opaque-sheet-metadata.bin',
			featureFamily: 'preservedMetadata',
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
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/drawings/drawing1.xml',
			relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
			id: 'rIdLinkedImage',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
			rawTarget: 'https://example.invalid/logo.png',
			targetMode: 'External',
			featureFamily: 'preservedMedia',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/worksheets/sheet1.xml',
			relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
			id: 'rIdHyperlink',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
			rawTarget: 'https://example.invalid/report',
			targetMode: 'External',
			featureFamily: 'preservedHyperlink',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/worksheets/sheet1.xml',
			relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
			id: 'rIdControlProps',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp',
			rawTarget: 'https://example.invalid/control.xml',
			targetMode: 'External',
			featureFamily: 'preservedControl',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/worksheets/sheet1.xml',
			relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
			id: 'rIdActiveX',
			type: 'http://schemas.microsoft.com/office/2006/relationships/activeXControl',
			rawTarget: 'https://example.invalid/control.ocx',
			targetMode: 'External',
			featureFamily: 'preservedActiveX',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/drawings/drawing1.xml',
			relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
			id: 'rIdOleObject',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject',
			rawTarget: 'file:///C:/objects/report.bin',
			targetMode: 'External',
			featureFamily: 'preservedEmbedding',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/drawings/drawing1.xml',
			relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
			id: 'rIdOpaqueChartStyle',
			type: 'http://schemas.microsoft.com/office/2011/relationships/chartStyle',
			rawTarget: 'style-package.bin',
			resolvedTarget: 'xl/drawings/style-package.bin',
			featureFamily: 'preservedChartStyle',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/drawings/drawing1.xml',
			relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
			id: 'rIdOpaqueChartColors',
			type: 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle',
			rawTarget: 'color-package.bin',
			resolvedTarget: 'xl/drawings/color-package.bin',
			featureFamily: 'preservedChartColor',
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
		expect(graph.parts.find((part) => part.path === 'xl/opaque-calc.bin')).toMatchObject({
			ownerScope: 'metadata',
			sourceRelationshipId: 'rIdCalcChainOpaque',
			featureFamily: 'preservedCalcChain',
			preservationPolicy: 'discard-on-recalc',
			bytePreservationExpected: false,
		})
		expect(graph.parts.find((part) => part.path === 'xl/opaque-sheet-metadata.bin')).toMatchObject({
			ownerScope: 'metadata',
			sourceRelationshipId: 'rIdSheetMetadataOpaque',
			featureFamily: 'preservedMetadata',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/opaque-custom.xml')).toMatchObject({
			ownerScope: 'custom-xml',
			sourceRelationshipId: 'rIdCustomXmlOpaque',
			featureFamily: 'preservedCustomXml',
		})
		expect(graph.parts.find((part) => part.path === 'xl/opaque-vba.bin')).toMatchObject({
			ownerScope: 'active-content',
			sourceRelationshipId: 'rIdVbaOpaque',
			featureFamily: 'preservedMacro',
		})
		expect(graph.parts.find((part) => part.path === 'xl/opaque-signature.bin')).toMatchObject({
			ownerScope: 'security',
			sourceRelationshipId: 'rIdVbaSignatureOpaque',
			featureFamily: 'preservedSignature',
		})
		expect(
			graph.parts.find((part) => part.path === 'xl/worksheets/opaque-control.bin'),
		).toMatchObject({
			ownerScope: 'active-content',
			sourceRelationshipId: 'rIdControlPropsOpaque',
			featureFamily: 'preservedControl',
		})
		expect(graph.parts.find((part) => part.path === 'xl/drawings/opaque-ole.bin')).toMatchObject({
			ownerScope: 'active-content',
			sourceRelationshipId: 'rIdOleObjectOpaque',
			featureFamily: 'preservedEmbedding',
		})
		expect(
			graph.parts.find((part) => part.path === 'xl/data/connectionsPayload.xml'),
		).toMatchObject({
			sourceRelationshipId: 'rIdConnections',
			featureFamily: 'preservedConnection',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/links/bookLink.xml')).toMatchObject({
			ownerScope: 'external-link',
			sourceRelationshipId: 'rIdExternalLink',
			featureFamily: 'preservedExternalLink',
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
