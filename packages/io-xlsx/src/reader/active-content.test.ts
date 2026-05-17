import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { unzipSync } from 'fflate'
import { makeXlsx } from '../../test/helpers.ts'
import { writeXlsx } from '../writer/index.ts'
import {
	parseActiveXControlInfo,
	parseVmlControlInfos,
	parseWorksheetControlInfos,
} from './active-content.ts'
import { readXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('active content inventory', () => {
	test('parses active content relationship ids with non-r prefixes', () => {
		expect(
			parseActiveXControlInfo(
				`<?xml version="1.0"?><ax:ocx ax:classid="{8BD21D40-EC42-11CE-9E0D-00AA006002F3}" ax:persistence="persistStreamInit" rel:id="rIdBinary" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`,
				[],
			),
		).toMatchObject({
			classId: '{8BD21D40-EC42-11CE-9E0D-00AA006002F3}',
			persistence: 'persistStreamInit',
			relationshipId: 'rIdBinary',
		})

		expect(
			parseWorksheetControlInfos(
				`<x:controls xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:control rel:id="rIdControl" shapeId="7" name="Button 1">
    <x:controlPr rel:id="rIdControlPr"/>
  </x:control>
</x:controls>`,
				'xl/worksheets/sheet1.xml',
				[
					{
						id: 'rIdControlPr',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp',
						target: '../ctrlProps/ctrlProp1.xml',
					},
				],
			),
		).toEqual([
			{
				shapeId: 7,
				name: 'Button 1',
				relationshipId: 'rIdControl',
				controlPrRelationshipId: 'rIdControlPr',
				controlPrRelationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp',
				controlPrTarget: 'xl/ctrlProps/ctrlProp1.xml',
			},
		])

		expect(
			parseVmlControlInfos(
				`<xml xmlns:vm="urn:schemas-microsoft-com:vml" xmlns:img="urn:excel-image-links" xmlns:o="urn:schemas-microsoft-com:office:office">
  <vm:shape id="Button 1" o:spid="_x0000_s1025">
    <vm:imagedata img:relid="rIdImage"/>
    <x:ClientData xmlns:x="urn:schemas-microsoft-com:office:excel" ObjectType="Button"/>
  </vm:shape>
</xml>`,
				'xl/drawings/vmlDrawing1.vml',
				[
					{
						id: 'rIdImage',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
						target: '../media/button1.png',
					},
				],
			),
		).toEqual([
			{
				shapeId: 1025,
				shapeName: 'Button 1',
				shapeSpid: '_x0000_s1025',
				objectType: 'Button',
				imageRelationshipId: 'rIdImage',
				imageRelationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
				imageTarget: 'xl/media/button1.png',
			},
		])
	})

	test('discovers and preserves workbook Custom UI callbacks by UI extensibility relationship', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/customUI/customUI2.xml" ContentType="application/vnd.ms-office.customUI+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rIdCustomUi" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="/customUI/customUI2.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'customUI/customUI2.xml': `<?xml version="1.0"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon.OnLoad" loadImage="Ribbon.LoadImage">
  <ribbon><tabs><tab id="tabAscend" label="Ascend">
    <group id="grpActions" label="Actions">
      <button id="runReport" label="Run" onAction="Module1.RunReport" getEnabled="Ribbon.CanRun"/>
    </group>
  </tab></tabs></ribbon>
</customUI>`,
			'customUI/_rels/customUI2.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../xl/media/image1.png"/>
</Relationships>`,
			'xl/media/image1.png': 'image-bytes',
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.activeContent).toContainEqual({
			kind: 'customUi',
			partPath: 'customUI/customUI2.xml',
			contentType: 'application/vnd.ms-office.customUI+xml',
			anchor: 'workbook',
			relType: 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility',
			sourceRelationshipId: 'rIdCustomUi',
			relationshipCount: 1,
			executionPolicy: 'blocked',
			customUi: {
				namespaceUri: 'http://schemas.microsoft.com/office/2009/07/customui',
				callbackCount: 4,
				callbacks: [
					{ attribute: 'onLoad', macro: 'Ribbon.OnLoad' },
					{ attribute: 'loadImage', macro: 'Ribbon.LoadImage' },
					{ attribute: 'onAction', macro: 'Module1.RunReport' },
					{ attribute: 'getEnabled', macro: 'Ribbon.CanRun' },
				],
			},
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedCustomUi'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['customUI/customUI2.xml'],
		})

		const written = writeXlsx(result.value.workbook, result.value.capsules)
		expectOk(written)
		const entries = unzipSync(written.value)
		expect(new TextDecoder().decode(entries['customUI/customUI2.xml'])).toContain(
			'onAction="Module1.RunReport"',
		)
		expect(new TextDecoder().decode(entries['_rels/.rels'])).toContain(
			'Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"',
		)
	})

	test('discovers macros, ActiveX controls, and form control property parts', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
</Relationships>`,
			'xl/vbaProject.bin': 'macro-bytes',
			'xl/activeX/activeX1.xml': `<?xml version="1.0"?><ax:ocx ax:classid="{8BD21D40-EC42-11CE-9E0D-00AA006002F3}" ax:persistence="persistStreamInit" r:id="rId1" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`,
			'xl/activeX/_rels/activeX1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
</Relationships>`,
			'xl/ctrlProps/ctrlProp1.xml': `<?xml version="1.0"?><formControlPr macro="Module1.Run" fmlaLink="$A$1" fmlaRange="$A$2:$A$4" checked="Checked" dropLines="8"/>`,
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.activeContent).toEqual([
			{
				kind: 'vbaProject',
				partPath: 'xl/vbaProject.bin',
				contentType: 'application/vnd.ms-office.vbaProject',
				anchor: 'workbook',
				relType: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
				sourceRelationshipId: 'rId2',
				relationshipCount: 0,
				byteSize: 11,
				opaque: true,
				executionPolicy: 'blocked',
			},
			{
				kind: 'activeX',
				partPath: 'xl/activeX/activeX1.xml',
				contentType: 'application/vnd.ms-office.activeX+xml',
				anchor: 'sheet',
				sheetName: 'Data',
				relType: 'http://schemas.microsoft.com/office/2006/relationships/activeXControl',
				sourceRelationshipId: 'rId1',
				relationshipCount: 1,
				executionPolicy: 'blocked',
				activeX: {
					classId: '{8BD21D40-EC42-11CE-9E0D-00AA006002F3}',
					persistence: 'persistStreamInit',
					relationshipId: 'rId1',
					binaryRelationshipId: 'rId1',
					binaryTarget: 'activeX1.bin',
				},
			},
			{
				kind: 'formControl',
				partPath: 'xl/ctrlProps/ctrlProp1.xml',
				contentType: 'application/vnd.ms-excel.controlproperties+xml',
				anchor: 'sheet',
				sheetName: 'Data',
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp',
				sourceRelationshipId: 'rId2',
				relationshipCount: 0,
				executionPolicy: 'blocked',
				formControl: {
					macro: 'Module1.Run',
					linkedCell: '$A$1',
					listFillRange: '$A$2:$A$4',
					checked: 'Checked',
					dropLines: 8,
				},
			},
		])
		expect(result.value.report.status).toBe('has-preserved')
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedMacro'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/vbaProject.bin'],
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedActiveX'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/activeX/activeX1.xml'],
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedControl'),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/ctrlProps/ctrlProp1.xml'],
		})
	})

	test('summarizes real VBA project modules without exposing source code', () => {
		const bytes = readFileSync(
			new URL('../../../../fixtures/xlsx/calamine/vba.xlsm', import.meta.url),
		)

		const result = readXlsx(bytes)
		expectOk(result)

		const vbaProject = result.value.workbook.activeContent.find(
			(content) => content.kind === 'vbaProject',
		)
		expect(vbaProject).toMatchObject({
			partPath: 'xl/vbaProject.bin',
			opaque: true,
			executionPolicy: 'blocked',
			vbaProject: {
				moduleCount: 5,
				projectStreamPresent: true,
			},
		})
		expect(vbaProject?.vbaProject?.modules).toEqual([
			{ name: 'ThisWorkbook', kind: 'document' },
			{ name: 'Sheet1', kind: 'document' },
			{ name: 'Sheet2', kind: 'document' },
			{ name: 'Sheet3', kind: 'document' },
			{ name: 'testVBA', kind: 'standard' },
		])
		expect(JSON.stringify(vbaProject?.vbaProject)).not.toContain('Attribute VB_Name')
	})

	test('metadata-only reads inventory active content without hydrating preservation capsules', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>`,
			'xl/_rels/vbaProject.bin.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVbaSignature" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdActiveX" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml"/>
  <Relationship Id="rIdControl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
</Relationships>`,
			'xl/vbaProject.bin': 'macro-bytes',
			'xl/vbaProjectSignature.bin': 'signature-bytes',
			'xl/activeX/activeX1.xml': `<?xml version="1.0"?><ax:ocx ax:classid="{8BD21D40-EC42-11CE-9E0D-00AA006002F3}" ax:persistence="persistStreamInit" r:id="rIdBinary" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`,
			'xl/activeX/_rels/activeX1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdBinary" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
</Relationships>`,
			'xl/ctrlProps/ctrlProp1.xml': `<?xml version="1.0"?><formControlPr macro="Module1.Run" fmlaLink="$A$1"/>`,
		})

		const result = readXlsx(bytes, { mode: 'metadata-only' })
		expectOk(result)

		expect(result.value.loadInfo.isPartial).toBe(true)
		expect(result.value.capsules).toEqual([])
		const metadataVbaProject = result.value.workbook.activeContent.find(
			(content) => content.kind === 'vbaProject',
		)
		expect(result.value.workbook.activeContent).toContainEqual(
			expect.objectContaining({
				kind: 'vbaProject',
				partPath: 'xl/vbaProject.bin',
				sourceRelationshipId: 'rIdVba',
				byteSize: 11,
			}),
		)
		expect(metadataVbaProject?.vbaProject).toBeUndefined()
		expect(result.value.workbook.activeContent).toContainEqual({
			kind: 'vbaSignature',
			partPath: 'xl/vbaProjectSignature.bin',
			contentType: 'application/vnd.ms-office.vbaProjectSignature',
			anchor: 'workbook',
			sourcePartPath: 'xl/vbaProject.bin',
			relType: 'http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature',
			sourceRelationshipId: 'rIdVbaSignature',
			relationshipCount: 0,
			invalidationPolicy: 'invalidatedByPackageEdit',
			resigningPolicy: 'notSupported',
		})
		const metadataActiveX = result.value.workbook.activeContent.find(
			(content) => content.kind === 'activeX',
		)
		expect(result.value.workbook.activeContent).toContainEqual(
			expect.objectContaining({
				kind: 'activeX',
				partPath: 'xl/activeX/activeX1.xml',
				anchor: 'sheet',
				sheetName: 'Data',
				sourceRelationshipId: 'rIdActiveX',
				activeX: expect.objectContaining({
					classId: '{8BD21D40-EC42-11CE-9E0D-00AA006002F3}',
					binaryRelationshipId: 'rIdBinary',
				}),
			}),
		)
		expect(metadataActiveX?.worksheetControl).toBeUndefined()
		expect(result.value.workbook.activeContent).toContainEqual(
			expect.objectContaining({
				kind: 'formControl',
				partPath: 'xl/ctrlProps/ctrlProp1.xml',
				anchor: 'sheet',
				sheetName: 'Data',
				sourceRelationshipId: 'rIdControl',
				formControl: expect.objectContaining({ macro: 'Module1.Run' }),
			}),
		)
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedMacro'),
		).toMatchObject({
			count: 1,
			locations: ['xl/vbaProject.bin'],
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedSignature'),
		).toMatchObject({
			count: 1,
			locations: ['xl/vbaProjectSignature.bin'],
		})
	})

	test('links VBA project signatures to their signed VBA project relationship', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>`,
			'xl/_rels/vbaProject.bin.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVbaSignature" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
			'xl/vbaProject.bin': 'macro-bytes',
			'xl/vbaProjectSignature.bin': 'signature-bytes',
		})

		const result = readXlsx(bytes)
		expectOk(result)

		expect(result.value.workbook.activeContent).toContainEqual(
			expect.objectContaining({
				kind: 'vbaProject',
				partPath: 'xl/vbaProject.bin',
				relType: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
				sourceRelationshipId: 'rIdVba',
				relationshipCount: 1,
			}),
		)
		expect(result.value.workbook.activeContent).toContainEqual({
			kind: 'vbaSignature',
			partPath: 'xl/vbaProjectSignature.bin',
			contentType: 'application/vnd.ms-office.vbaProjectSignature',
			anchor: 'workbook',
			sourcePartPath: 'xl/vbaProject.bin',
			relType: 'http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature',
			sourceRelationshipId: 'rIdVbaSignature',
			relationshipCount: 0,
			invalidationPolicy: 'invalidatedByPackageEdit',
			resigningPolicy: 'notSupported',
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedMacro'),
		).toMatchObject({
			count: 1,
			locations: ['xl/vbaProject.bin'],
		})
		expect(
			result.value.report.features.find((feature) => feature.feature === 'preservedSignature'),
		).toMatchObject({
			count: 1,
			locations: ['xl/vbaProjectSignature.bin'],
		})
	})

	test('summarizes real LibreOffice ActiveX relationship metadata', () => {
		const bytes = readFileSync(
			new URL('../../../../fixtures/xlsx/libreoffice/activex_checkbox.xlsx', import.meta.url),
		)

		const result = readXlsx(bytes)
		expectOk(result)

		const activeX = result.value.workbook.activeContent.find(
			(content) => content.kind === 'activeX',
		)
		expect(activeX).toMatchObject({
			kind: 'activeX',
			partPath: 'xl/activeX/activeX1.xml',
			sheetName: 'Sheet1',
			sourceRelationshipId: 'rId3',
			worksheetControl: {
				shapeId: 1025,
				name: 'CheckBox1343',
				relationshipId: 'rId3',
				controlPrRelationshipId: 'rId4',
				controlPrRelationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
				controlPrTarget: 'xl/media/image1.emf',
				anchor: {
					kind: 'twoCell',
					from: { col: 1, row: 3, colOff: 438150, rowOff: 38100 },
					to: { col: 4, row: 6, colOff: 161925, rowOff: 114300 },
				},
				vmlShapeId: 'CheckBox1343',
				vmlShapeSpid: '_x0000_s1025',
				vmlObjectType: 'Pict',
				vmlMapOcx: true,
				vmlImageRelationshipId: 'rId1',
				vmlImageRelationshipType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
				vmlImageTarget: 'xl/media/image1.emf',
			},
			activeX: {
				classId: '{8BD21D40-EC42-11CE-9E0D-00AA006002F3}',
				persistence: 'persistStreamInit',
				relationshipId: 'rId1',
				binaryRelationshipId: 'rId1',
				binaryTarget: 'activeX1.bin',
			},
		})
	})
})
