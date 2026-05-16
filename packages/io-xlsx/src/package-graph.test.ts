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
  <Relationship Id="rIdOpaqueCoreProps" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="package/services/metadata/core-properties/source.psmdcp"/>
  <Relationship Id="rIdOpaqueThumbnail" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="package/services/metadata/thumbnail.bin"/>
  <Relationship Id="rIdOpaqueAppProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="package/services/app.bin"/>
  <Relationship Id="rIdOpaqueCustomProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="package/services/custom.bin"/>
  <Relationship Id="rIdOpaqueSignatureOrigin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="package/signatures/origin.sigs"/>
  <Relationship Id="rIdOpaqueCustomUi" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="package/ui/custom-ui.bin"/>
</Relationships>`,
			'package/services/metadata/core-properties/source.psmdcp': '<core/>',
			'package/services/metadata/thumbnail.bin': 'thumbnail-bytes',
			'package/services/app.bin': '<app/>',
			'package/services/custom.bin': '<custom/>',
			'package/signatures/origin.sigs': '',
			'package/signatures/_rels/origin.sigs.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOpaqueSignature" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="signature-package.bin"/>
</Relationships>`,
			'package/signatures/signature-package.bin': '<Signature/>',
			'package/ui/custom-ui.bin': '<customUI/>',
			'xl/workbook.xml': '<workbook/>',
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdXmlMaps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps" Target="xmlMaps.xml"/>
  <Relationship Id="rIdXmlMapsOpaque" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps" Target="opaque-xmlmaps.bin"/>
  <Relationship Id="rIdCustomProperty" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty" Target="customProperty1.bin"/>
  <Relationship Id="rIdCustomPropertyOpaque" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty" Target="opaque-custom-property.bin"/>
  <Relationship Id="rIdDiagramData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>
  <Relationship Id="rIdDiagramDataOpaque" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="opaque-diagram.bin"/>
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
  <Relationship Id="rIdRevisionHeadersOpaque" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders" Target="opaque-revision.bin"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': '<worksheet/>',
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rIdQueryTableOpaque" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="opaque-query.bin"/>
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
			'xl/opaque-xmlmaps.bin': 'xml-map-bytes',
			'xl/customProperty1.bin': 'custom-property-bytes',
			'xl/opaque-custom-property.bin': 'custom-property-opaque-bytes',
			'xl/diagrams/data1.xml': '<dgm:dataModel/>',
			'xl/opaque-diagram.bin': 'diagram-bytes',
			'xl/model/item.data': 'data-model-bytes',
			'xl/customData/item1.data': 'power-query-bytes',
			'xl/opaque-mashup.bin': 'opaque-power-query-bytes',
			'xl/opaque-model.bin': 'opaque-data-model-bytes',
			'xl/opaque-calc.bin': 'calc-chain-bytes',
			'xl/opaque-sheet-metadata.bin': 'metadata-bytes',
			'xl/opaque-custom.xml': '<custom/>',
			'xl/opaque-vba.bin': 'vba-bytes',
			'xl/opaque-signature.bin': 'signature-bytes',
			'xl/data/connectionsPayload.xml': '<connections/>',
			'xl/links/bookLink.xml': '<externalLink/>',
			'xl/revisions/revisionHeaders.xml': '<headers/>',
			'xl/opaque-revision.bin': 'revision-bytes',
			'xl/worksheets/opaque-query.bin': 'query-table-bytes',
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
			sourcePartPath: '',
			relationshipPartPath: '_rels/.rels',
			id: 'rIdOpaqueCoreProps',
			type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
			rawTarget: 'package/services/metadata/core-properties/source.psmdcp',
			resolvedTarget: 'package/services/metadata/core-properties/source.psmdcp',
			featureFamily: 'preservedDocumentProperties',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: '',
			relationshipPartPath: '_rels/.rels',
			id: 'rIdOpaqueAppProps',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
			rawTarget: 'package/services/app.bin',
			resolvedTarget: 'package/services/app.bin',
			featureFamily: 'preservedDocumentProperties',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: '',
			relationshipPartPath: '_rels/.rels',
			id: 'rIdOpaqueThumbnail',
			type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail',
			rawTarget: 'package/services/metadata/thumbnail.bin',
			resolvedTarget: 'package/services/metadata/thumbnail.bin',
			featureFamily: 'preservedDocumentProperties',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: '',
			relationshipPartPath: '_rels/.rels',
			id: 'rIdOpaqueCustomProps',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
			rawTarget: 'package/services/custom.bin',
			resolvedTarget: 'package/services/custom.bin',
			featureFamily: 'preservedDocumentProperties',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: '',
			relationshipPartPath: '_rels/.rels',
			id: 'rIdOpaqueSignatureOrigin',
			type: 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin',
			rawTarget: 'package/signatures/origin.sigs',
			resolvedTarget: 'package/signatures/origin.sigs',
			featureFamily: 'preservedSignature',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'package/signatures/origin.sigs',
			relationshipPartPath: 'package/signatures/_rels/origin.sigs.rels',
			id: 'rIdOpaqueSignature',
			type: 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature',
			rawTarget: 'signature-package.bin',
			resolvedTarget: 'package/signatures/signature-package.bin',
			featureFamily: 'preservedSignature',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: '',
			relationshipPartPath: '_rels/.rels',
			id: 'rIdOpaqueCustomUi',
			type: 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility',
			rawTarget: 'package/ui/custom-ui.bin',
			resolvedTarget: 'package/ui/custom-ui.bin',
			featureFamily: 'preservedCustomUi',
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
			sourcePartPath: 'xl/worksheets/sheet1.xml',
			relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
			id: 'rIdQueryTableOpaque',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
			rawTarget: 'opaque-query.bin',
			resolvedTarget: 'xl/worksheets/opaque-query.bin',
			featureFamily: 'preservedQueryTable',
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
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			id: 'rIdXmlMapsOpaque',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps',
			rawTarget: 'opaque-xmlmaps.bin',
			resolvedTarget: 'xl/opaque-xmlmaps.bin',
			featureFamily: 'preservedCustomXml',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			id: 'rIdCustomPropertyOpaque',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty',
			rawTarget: 'opaque-custom-property.bin',
			resolvedTarget: 'xl/opaque-custom-property.bin',
			featureFamily: 'preservedMetadata',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			id: 'rIdDiagramDataOpaque',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData',
			rawTarget: 'opaque-diagram.bin',
			resolvedTarget: 'xl/opaque-diagram.bin',
			featureFamily: 'preservedDrawing',
		})
		expect(graph.relationships).toContainEqual({
			sourcePartPath: 'xl/workbook.xml',
			relationshipPartPath: 'xl/_rels/workbook.xml.rels',
			id: 'rIdRevisionHeadersOpaque',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders',
			rawTarget: 'opaque-revision.bin',
			resolvedTarget: 'xl/opaque-revision.bin',
			featureFamily: 'preservedRevision',
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
		expect(
			graph.parts.find(
				(part) => part.path === 'package/services/metadata/core-properties/source.psmdcp',
			),
		).toMatchObject({
			ownerScope: 'document-properties',
			sourceRelationshipId: 'rIdOpaqueCoreProps',
			featureFamily: 'preservedDocumentProperties',
			preservationPolicy: 'preserve-exact',
		})
		expect(
			graph.parts.find((part) => part.path === 'package/services/metadata/thumbnail.bin'),
		).toMatchObject({
			ownerScope: 'document-properties',
			sourceRelationshipId: 'rIdOpaqueThumbnail',
			featureFamily: 'preservedDocumentProperties',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'package/services/app.bin')).toMatchObject({
			ownerScope: 'document-properties',
			sourceRelationshipId: 'rIdOpaqueAppProps',
			featureFamily: 'preservedDocumentProperties',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'package/services/custom.bin')).toMatchObject({
			ownerScope: 'document-properties',
			sourceRelationshipId: 'rIdOpaqueCustomProps',
			featureFamily: 'preservedDocumentProperties',
			preservationPolicy: 'preserve-exact',
		})
		expect(
			graph.parts.find((part) => part.path === 'package/signatures/origin.sigs'),
		).toMatchObject({
			ownerScope: 'security',
			sourceRelationshipId: 'rIdOpaqueSignatureOrigin',
			featureFamily: 'preservedSignature',
			preservationPolicy: 'invalidate-on-edit',
			bytePreservationExpected: false,
		})
		expect(
			graph.parts.find((part) => part.path === 'package/signatures/signature-package.bin'),
		).toMatchObject({
			ownerScope: 'security',
			sourceRelationshipId: 'rIdOpaqueSignature',
			featureFamily: 'preservedSignature',
			preservationPolicy: 'invalidate-on-edit',
			bytePreservationExpected: false,
		})
		expect(graph.parts.find((part) => part.path === 'package/ui/custom-ui.bin')).toMatchObject({
			ownerScope: 'active-content',
			sourceRelationshipId: 'rIdOpaqueCustomUi',
			featureFamily: 'preservedCustomUi',
			preservationPolicy: 'preserve-exact',
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
		expect(graph.parts.find((part) => part.path === 'xl/opaque-xmlmaps.bin')).toMatchObject({
			ownerScope: 'custom-xml',
			sourceRelationshipId: 'rIdXmlMapsOpaque',
			featureFamily: 'preservedCustomXml',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/customProperty1.bin')).toMatchObject({
			sourceRelationshipId: 'rIdCustomProperty',
			featureFamily: 'preservedMetadata',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/opaque-custom-property.bin')).toMatchObject(
			{
				ownerScope: 'metadata',
				sourceRelationshipId: 'rIdCustomPropertyOpaque',
				featureFamily: 'preservedMetadata',
				preservationPolicy: 'preserve-exact',
			},
		)
		expect(graph.parts.find((part) => part.path === 'xl/diagrams/data1.xml')).toMatchObject({
			sourceRelationshipId: 'rIdDiagramData',
			featureFamily: 'preservedDrawing',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/opaque-diagram.bin')).toMatchObject({
			ownerScope: 'drawing',
			sourceRelationshipId: 'rIdDiagramDataOpaque',
			featureFamily: 'preservedDrawing',
			preservationPolicy: 'preserve-exact',
		})
		expect(graph.parts.find((part) => part.path === 'xl/model/item.data')).toMatchObject({
			ownerScope: 'analytics',
			sourceRelationshipId: 'rIdDataModel',
			featureFamily: 'preservedDataModel',
			preservationPolicy: 'inspect-only',
			bytePreservationExpected: true,
		})
		expect(graph.parts.find((part) => part.path === 'xl/customData/item1.data')).toMatchObject({
			ownerScope: 'analytics',
			sourceRelationshipId: 'rIdPowerQuery',
			featureFamily: 'preservedPowerQuery',
			preservationPolicy: 'inspect-only',
			bytePreservationExpected: true,
		})
		expect(graph.parts.find((part) => part.path === 'xl/opaque-mashup.bin')).toMatchObject({
			ownerScope: 'analytics',
			sourceRelationshipId: 'rIdPowerQueryOpaque',
			featureFamily: 'preservedPowerQuery',
			preservationPolicy: 'inspect-only',
		})
		expect(graph.parts.find((part) => part.path === 'xl/opaque-model.bin')).toMatchObject({
			ownerScope: 'analytics',
			sourceRelationshipId: 'rIdDataModelOpaque',
			featureFamily: 'preservedDataModel',
			preservationPolicy: 'inspect-only',
		})
		expect(
			graph.parts.find((part) => part.path === 'xl/worksheets/opaque-query.bin'),
		).toMatchObject({
			ownerScope: 'analytics',
			sourceRelationshipId: 'rIdQueryTableOpaque',
			featureFamily: 'preservedQueryTable',
			preservationPolicy: 'preserve-exact',
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
			ownerScope: 'analytics',
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
		expect(graph.parts.find((part) => part.path === 'xl/opaque-revision.bin')).toMatchObject({
			ownerScope: 'metadata',
			sourceRelationshipId: 'rIdRevisionHeadersOpaque',
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
