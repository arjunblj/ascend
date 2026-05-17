import { describe, expect, it } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'
import type { StyleId } from '@ascend/core'
import { createTableId, Workbook } from '@ascend/core'
import { booleanValue, errorValue, numberValue, stringValue } from '@ascend/schema'
import { unzipSync } from 'fflate'
import { applyOperations } from '../../../engine/src/index.ts'
import { fingerprintXlsx } from '../../test/fidelity-harness.ts'
import { makeXlsx } from '../../test/helpers.ts'
import { inspectXlsxPackageGraph } from '../package-graph.ts'
import {
	auditXlsxPackageGraphBytePreservation,
	auditXlsxPackageGraphReadIntegrity,
	auditXlsxPackageGraphSafeEditIntegrity,
} from '../package-graph-fidelity.ts'
import type { PreservationCapsule } from '../preserve.ts'
import { advancedFilterSparklineWorkbook } from '../reader/advanced-filter-sparkline.test.ts'
import { readXlsx } from '../reader/index.ts'
import { ZipArchive } from '../reader/zip.ts'
import { updateConnectionPartXml } from './connection.ts'
import { writeDenseRowsXlsx, writeDenseRowsXlsxStreaming } from './dense-rows.ts'
import { planWriteXlsx, writeXlsx, writeXlsxStreaming } from './index.ts'
import { updatePivotCacheDefinitionXml } from './pivot-cache.ts'
import { updatePivotTableDefinitionXml } from './pivot-table.ts'
import { buildPreservedStylesXml } from './styles.ts'

const S0 = 0 as StyleId
const DASHBOARD_CORPUS_FIXTURE = new URL(
	'../../../../research/excel-corpus/excel-dashboard-v2.xlsx',
	import.meta.url,
)
const POI_COMMENTS_FIXTURE = new URL(
	'../../../../fixtures/xlsx/poi/SimpleWithComments.xlsx',
	import.meta.url,
)

function expectOk<T, E>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(formatTestError(result.error))
}

function formatTestError(error: unknown): string {
	if (error && typeof error === 'object') {
		const candidate = error as { message?: unknown; errors?: Array<{ message?: unknown }> }
		if (typeof candidate.message === 'string') return candidate.message
		if (Array.isArray(candidate.errors)) {
			const messages = candidate.errors
				.map((entry) => (typeof entry?.message === 'string' ? entry.message : null))
				.filter((entry): entry is string => entry !== null)
			if (messages.length > 0) return messages.join('; ')
		}
	}
	return 'Unknown test error'
}

function roundTrip(wb: Workbook, capsules?: PreservationCapsule[]) {
	const written = writeXlsx(wb, capsules)
	if (!written.ok) throw new Error(`write failed: ${written.error.message}`)
	const read = readXlsx(written.value)
	if (!read.ok) throw new Error(`read failed: ${read.error.message}`)
	return { bytes: written.value, result: read.value }
}

function decodeTestXml(bytes: Uint8Array | undefined): string {
	expect(bytes).toBeDefined()
	return new TextDecoder().decode(bytes)
}

function commentsAndThreadedCommentsWorkbook(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
  <Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/>
  </sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>original</t></is></c></row>
  </sheetData>
  <legacyDrawing r:id="rIdVml"/>
</worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
  <Relationship Id="rIdThreaded" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
</Relationships>`,
		'xl/comments1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Ada</author>
    <author>Grace</author>
  </authors>
  <commentList>
    <comment ref="B1" authorId="0"><text><t>Legacy visible note</t></text></comment>
    <comment ref="C3" authorId="1" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid="{comment-c3}"><text><r><rPr><b/></rPr><t>LegacyHidden</t></r><r><t>Note</t></r><phoneticPr fontId="1"/></text></comment>
  </commentList>
</comments>`,
		'xl/drawings/vmlDrawing1.vml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
    <v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype>
  <v:shape id="_x0000_s2048" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:144pt;height:79.5pt;z-index:1;visibility:visible" fillcolor="#ffffe1" o:insetmode="auto">
    <v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
    <v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>
    <x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1, 15, 0, 2, 3, 20, 4, 8</x:Anchor><x:Visible/><x:Row>0</x:Row><x:Column>1</x:Column></x:ClientData>
  </v:shape>
  <v:shape id="_x0000_s2049" type="#_x0000_t202" style="position:absolute;margin-left:90pt;margin-top:36pt;width:120pt;height:60pt;z-index:2;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
    <v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
    <v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>
    <x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>2, 10, 2, 4, 5, 30, 6, 12</x:Anchor><x:Visible>false</x:Visible><x:Row>2</x:Row><x:Column>2</x:Column></x:ClientData>
  </v:shape>
</xml>`,
		'xl/persons/person.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="{person-ada}" displayName="Ada Thread" userId="ada@example.test" providerId="None"/>
  <person id="{person-grace}" displayName="Grace Thread" userId="grace@example.test" providerId="None"/>
</personList>`,
		'xl/threadedComments/threadedComment1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="D4" personId="{person-ada}" id="{root-thread}" dT="2024-03-01T10:11:12.000Z">
    <text>Thread root</text>
    <mentions><mention mentionpersonId="{person-grace}" startIndex="0" length="6"/></mentions>
  </threadedComment>
  <threadedComment ref="D4" personId="{person-grace}" id="{reply-thread}" parentId="{root-thread}" dT="2024-03-02T10:11:12.000Z" done="1">
    <text>Thread reply</text>
    <extLst><ext uri="{reply-ext}"><futureThreadMetadata preserved="1"/></ext></extLst>
  </threadedComment>
</ThreadedComments>`,
	})
}

function commentsAndMixedVmlWorkbook(): Uint8Array {
	const parts = Object.fromEntries(
		Object.entries(unzipSync(commentsAndThreadedCommentsWorkbook())).map(([path, bytes]) => [
			path,
			new TextDecoder().decode(bytes),
		]),
	)
	parts['xl/drawings/vmlDrawing1.vml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
    <v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype>
  <v:shape id="_x0000_s2048" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:144pt;height:79.5pt;z-index:1;visibility:visible" fillcolor="#ffffe1" o:insetmode="auto">
    <v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
    <v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>
    <x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1, 15, 0, 2, 3, 20, 4, 8</x:Anchor><x:Visible/><x:Row>0</x:Row><x:Column>1</x:Column></x:ClientData>
  </v:shape>
  <v:shape id="_x0000_s2049" type="#_x0000_t202" style="position:absolute;margin-left:90pt;margin-top:36pt;width:120pt;height:60pt;z-index:2;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
    <v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
    <v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>
    <x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>2, 10, 2, 4, 5, 30, 6, 12</x:Anchor><x:Visible>false</x:Visible><x:Row>2</x:Row><x:Column>2</x:Column></x:ClientData>
  </v:shape>
  <v:shape id="_x0000_sButton" style="position:absolute;visibility:visible">
    <x:ClientData ObjectType="Button"><x:Anchor>1, 15, 2, 3, 4, 5, 6, 7</x:Anchor><x:Visible/></x:ClientData>
  </v:shape>
  <v:shape id="_x0000_sUnknown" style="position:absolute;visibility:visible">
    <v:textbox><div>preserve unknown VML shape</div></v:textbox>
  </v:shape>
</xml>`
	return makeXlsx(parts)
}

function commentsAndThreadedCommentsWithoutPersonsWorkbook(): Uint8Array {
	const parts = Object.fromEntries(
		Object.entries(unzipSync(commentsAndThreadedCommentsWorkbook())).map(([path, bytes]) => [
			path,
			new TextDecoder().decode(bytes),
		]),
	)
	delete parts['xl/persons/person.xml']
	parts['[Content_Types].xml'] = parts['[Content_Types].xml']?.replace(
		/\n {2}<Override PartName="\/xl\/persons\/person\.xml" ContentType="application\/vnd\.ms-excel\.person\+xml"\/>/,
		'',
	)
	return makeXlsx(parts)
}

describe('writeXlsx', () => {
	it('round-trips cell values correctly', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, { value: stringValue('Hello'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(42), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: booleanValue(true), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(3.14), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: booleanValue(false), formula: null, styleId: S0 })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s).toBeDefined()
		expect(s?.name).toBe('Test')
		expect(s?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'Hello' })
		expect(s?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 42 })
		expect(s?.cells.get(0, 2)?.value).toEqual({ kind: 'boolean', value: true })
		expect(s?.cells.get(1, 0)?.value).toEqual({ kind: 'number', value: 3.14 })
		expect(s?.cells.get(1, 1)?.value).toEqual({ kind: 'boolean', value: false })
	})

	it('round-trips multiple sheets', () => {
		const wb = new Workbook()
		const s1 = wb.addSheet('First')
		s1.cells.set(0, 0, { value: stringValue('A'), formula: null, styleId: S0 })
		const s2 = wb.addSheet('Second')
		s2.cells.set(0, 0, { value: numberValue(99), formula: null, styleId: S0 })

		const { result } = roundTrip(wb)

		expect(result.workbook.sheets).toHaveLength(2)
		expect(result.workbook.sheets[0]?.name).toBe('First')
		expect(result.workbook.sheets[1]?.name).toBe('Second')
		expect(result.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'A',
		})
		expect(result.workbook.sheets[1]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 99,
		})
	})

	it('skips write fact scan for numeric workbooks without formula metadata', async () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Numeric')
		sheet.cells.setExpectedDensity('dense')
		for (let row = 0; row < 100; row++) {
			for (let col = 0; col < 8; col++) {
				sheet.cells.setPlainNumber(row, col, row * 8 + col)
			}
		}
		const failUnexpectedScan = () => {
			throw new Error('unexpected write fact scan')
		}
		sheet.cells.iterate = function* () {
			failUnexpectedScan()
			yield undefined as never
		}

		const written = await writeXlsxStreaming(wb)
		expectOk(written)
		const zip = unzipSync(written.value)
		expect(zip['xl/sharedStrings.xml']).toBeUndefined()
	})

	it('preserves defined-name metadata attributes when workbook XML is regenerated', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <definedNames>
    <definedName name="MacroName" hidden="1" comment="Menu entry" description="Macro &amp; helper" function="1" vbProcedure="1" xlm="1" functionGroupId="7" shortcutKey="K" publishToServer="1" workbookParameter="1">Data!$A$1</definedName>
  </definedNames>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		})
		const read = readXlsx(bytes)
		expectOk(read)
		expect(read.value.workbook.definedNames.getEntry('MacroName')?.extraAttributes).toContainEqual({
			name: 'description',
			value: 'Macro & helper',
		})

		const written = writeXlsx(read.value.workbook, read.value.capsules, { workbookMetaDirty: true })
		expectOk(written)
		const zip = unzipSync(written.value)
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		expect(workbookXml).toContain('name="MacroName"')
		expect(workbookXml).toContain('hidden="1"')
		expect(workbookXml).toContain('comment="Menu entry"')
		expect(workbookXml).toContain('description="Macro &amp; helper"')
		expect(workbookXml).toContain('function="1"')
		expect(workbookXml).toContain('vbProcedure="1"')
		expect(workbookXml).toContain('xlm="1"')
		expect(workbookXml).toContain('functionGroupId="7"')
		expect(workbookXml).toContain('shortcutKey="K"')
		expect(workbookXml).toContain('publishToServer="1"')
		expect(workbookXml).toContain('workbookParameter="1"')
	})

	it('preserves workbook relationship ids and order when dirty sheets force workbook XML regeneration', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	  <Default Extension="xml" ContentType="application/xml"/>
	  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
	  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
	  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
	  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
	  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
	  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
	</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
	  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
	  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
	</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
	  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
	  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
	</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
	<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
	  <sheets><sheet name="Data" sheetId="1" r:id="rId3"/></sheets>
	</workbook>`,
			'xl/styles.xml': `<?xml version="1.0"?>
	<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
	</styleSheet>`,
			'xl/theme/theme1.xml': `<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements/></a:theme>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
			'docProps/core.xml': `<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>`,
			'docProps/app.xml': `<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>`,
		})
		const read = readXlsx(bytes)
		expectOk(read)
		read.value.workbook.sheets[0]?.cells.set(0, 0, {
			value: numberValue(2),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		const rootRels = new TextDecoder().decode(zip['_rels/.rels'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)

		expect(workbookXml).toContain('r:id="rId3"')
		expect(rootRels).toContain(
			'Type="http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties"',
		)
		expect(workbookRels.indexOf('Id="rId1"')).toBeLessThan(workbookRels.indexOf('Id="rId2"'))
		expect(workbookRels.indexOf('Id="rId2"')).toBeLessThan(workbookRels.indexOf('Id="rId3"'))
		expect(workbookRels).toContain('Id="rId1"')
		expect(workbookRels).toContain('relationships/theme')
		expect(workbookRels).toContain('Id="rId2"')
		expect(workbookRels).toContain('relationships/styles')
		expect(workbookRels).toContain('Id="rId3"')
		expect(workbookRels).toContain('relationships/worksheet')
	})

	it('preserves strict relationship type dialect when dirty writes regenerate rels', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const read = readXlsx(bytes)
		expectOk(read)
		read.value.workbook.sheets[0]?.cells.set(0, 0, {
			value: numberValue(2),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const rootRels = new TextDecoder().decode(zip['_rels/.rels'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)

		expect(rootRels).toContain(
			'Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"',
		)
		expect(workbookRels).toContain(
			'Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"',
		)
		expect(readXlsx(written.value).ok).toBe(true)
	})

	it('preserves strict table queryTable relationship dialect when dirty writes regenerate table rels', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="str"><v>Name</v></c><c r="B1" t="str"><v>Value</v></c></row>
    <row r="2"><c r="A2" t="str"><v>A</v></c><c r="B2"><v>1</v></c></row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rIdTable"/></tableParts>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`,
			'xl/tables/table1.xml': `<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="7" name="QueryTable" displayName="QueryTable" ref="A1:B2" headerRowCount="1" totalsRowCount="0" tableType="queryTable">
  <autoFilter ref="A1:B2"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Name" queryTableFieldId="1"/>
    <tableColumn id="2" name="Value" queryTableFieldId="2"/>
  </tableColumns>
</table>`,
			'xl/tables/_rels/table1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdQuery99" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable" Target="../queryTables/queryTable1.xml"/>
</Relationships>`,
			'xl/queryTables/queryTable1.xml': `<?xml version="1.0"?>
<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="QueryTable" connectionId="1"/>`,
		})
		const read = readXlsx(bytes)
		expectOk(read)
		const table = read.value.workbook.sheets[0]?.tables[0]
		expect(table?.queryTable).toMatchObject({
			relationshipId: 'rIdQuery99',
			relationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
			relationshipRawType: 'http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable',
			partPath: 'xl/queryTables/queryTable1.xml',
			target: '../queryTables/queryTable1.xml',
		})
		read.value.workbook.sheets[0]?.cells.set(1, 1, {
			value: numberValue(2),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const tableRels = new TextDecoder().decode(
			zip['xl/tables/_rels/table1.xml.rels'] ?? new Uint8Array(),
		)
		expect(tableRels).toContain(
			'Type="http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable"',
		)
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.tables[0]?.queryTable).toMatchObject({
			relationshipId: 'rIdQuery99',
			relationshipRawType: 'http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable',
		})
	})

	it('writes edited tabular slicer cache item state back into preserved package XML', () => {
		const slicerCacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_State" sourceName="State">
  <pivotTables><pivotTable tabId="1" name="PivotTable1"/></pivotTables>
  <data><tabular pivotCacheId="5"><items count="2"><i x="1"/><i x="0" s="1"/></items></tabular></data>
</slicerCacheDefinition>`
		const wb = new Workbook()
		wb.addSheet('Sheet1')
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 5,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_State',
			sourceName: 'State',
			pivotCacheId: 5,
			pivotTableNames: ['PivotTable1'],
			items: [{ index: 1 }, { index: 0, selected: true }],
		})
		const applied = applyOperations(wb, [
			{
				op: 'setSlicerCacheItem',
				slicerCache: 'Slicer_State',
				item: 1,
				selected: true,
				noData: true,
			},
			{
				op: 'setSlicerCacheItem',
				slicerCache: 'Slicer_State',
				item: 0,
				selected: null,
			},
		])
		expectOk(applied)
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/slicerCaches/slicerCache1.xml',
				contentType: 'application/vnd.ms-excel.slicerCache+xml',
				relationships: [],
				content: new TextEncoder().encode(slicerCacheXml),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
			},
		]

		const written = writeXlsx(wb, capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(
			zip['xl/slicerCaches/slicerCache1.xml'] ?? new Uint8Array(),
		)

		expect(xml).toContain('<i x="1" s="1" nd="1"/>')
		expect(xml).toContain('<i x="0"/>')
		expect(xml).toContain('<items count="2">')
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.slicerCaches[0]?.items).toEqual([
			{ index: 1, selected: true, noData: true },
			{ index: 0 },
		])
	})

	it('writes edited timeline selection range back into preserved cache XML', () => {
		const timelineCacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="Timeline_Order_Date" sourceName="Order Date">
  <pivotTables><pivotTable tabId="1" name="PivotTable1"/></pivotTables>
  <state singleRangeFilterState="1" filterType="dateRange">
    <selection startDate="2024-01-01T00:00:00" endDate="2024-03-31T00:00:00"/>
    <bounds startDate="2023-01-01T00:00:00" endDate="2024-12-31T00:00:00"/>
  </state>
</timelineCacheDefinition>`
		const wb = new Workbook()
		wb.addSheet('Sheet1')
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 5,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		wb.timelineCaches.push({
			partPath: 'xl/timelineCaches/timelineCache1.xml',
			name: 'Timeline_Order_Date',
			sourceName: 'Order Date',
			pivotCacheId: 5,
			pivotTableNames: ['PivotTable1'],
			state: {
				filterType: 'dateRange',
				singleRangeFilterState: true,
				selection: {
					startDate: '2024-01-01T00:00:00',
					endDate: '2024-03-31T00:00:00',
				},
				bounds: {
					startDate: '2023-01-01T00:00:00',
					endDate: '2024-12-31T00:00:00',
				},
			},
		})
		const applied = applyOperations(wb, [
			{
				op: 'setTimelineRange',
				timelineCache: 'Timeline_Order_Date',
				startDate: '2024-04-01T00:00:00',
				endDate: '2024-06-30T00:00:00',
			},
		])
		expectOk(applied)
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/timelineCaches/timelineCache1.xml',
				contentType: 'application/vnd.ms-excel.timelineCache+xml',
				relationships: [],
				content: new TextEncoder().encode(timelineCacheXml),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.microsoft.com/office/2011/relationships/timelineCache',
			},
		]

		const written = writeXlsx(wb, capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(
			zip['xl/timelineCaches/timelineCache1.xml'] ?? new Uint8Array(),
		)

		expect(xml).toContain(
			'<selection startDate="2024-04-01T00:00:00" endDate="2024-06-30T00:00:00"/>',
		)
		expect(xml).toContain('<bounds startDate="2023-01-01T00:00:00" endDate="2024-12-31T00:00:00"/>')
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.timelineCaches[0]?.state?.selection).toEqual({
			startDate: '2024-04-01T00:00:00',
			endDate: '2024-06-30T00:00:00',
		})
	})

	it('writes edited pivot field item state back into preserved pivot XML', () => {
		const pivotTableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="34">
  <pivotFields count="1"><pivotField axis="axisPage" showAll="0"><items count="3"><item h="1" x="0"/><item x="1" sd="0"></item><item t="default"/></items></pivotField></pivotFields>
  <pageFields count="1"><pageField fld="0"></pageField></pageFields>
</pivotTableDefinition>`
		const pivotCacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="2">
  <cacheFields count="1"><cacheField name="Region"><sharedItems count="2"><s v="West"/><s v="East"/></sharedItems></cacheField></cacheFields>
</pivotCacheDefinition>`
		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [
				{
					index: 0,
					axis: 'axisPage',
					showAll: false,
					items: [
						{ index: 0, cacheIndex: 0, hidden: true },
						{ index: 1, cacheIndex: 1, showDetails: false },
						{ index: 2, itemType: 'default' },
					],
				},
			],
			rowFields: [],
			columnFields: [],
			pageFields: [{ index: 0 }],
			dataFields: [],
		})
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			fields: [
				{
					index: 0,
					name: 'Region',
					sharedItems: [
						{ index: 0, kind: 'string', value: 'West' },
						{ index: 1, kind: 'string', value: 'East' },
					],
				},
			],
		})
		const applied = applyOperations(wb, [
			{
				op: 'setPivotFieldItem',
				pivotTable: 'PivotTable1',
				fieldIndex: 0,
				itemIndex: 0,
				hidden: null,
				manualFilter: true,
				selectedPageItem: 1,
			},
		])
		expectOk(applied)
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/pivotTables/pivotTable1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
				relationships: [],
				content: new TextEncoder().encode(pivotTableXml),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: 'Sheet1' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
			},
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				contentType:
					'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
				relationships: [],
				content: new TextEncoder().encode(pivotCacheXml),
				anchor: { kind: 'workbook' },
				relType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
			},
		]

		const written = writeXlsx(wb, capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const pivotXml = new TextDecoder().decode(
			zip['xl/pivotTables/pivotTable1.xml'] ?? new Uint8Array(),
		)
		const cacheXml = new TextDecoder().decode(
			zip['xl/pivotCache/pivotCacheDefinition1.xml'] ?? new Uint8Array(),
		)

		expect(pivotXml).toContain('<item x="0" s="1"/>')
		expect(pivotXml).toContain('<item x="1" sd="0"></item>')
		expect(pivotXml).toContain('<pageField fld="0" item="1"></pageField>')
		expect(cacheXml).toContain('<pivotCacheDefinition')
		expect(cacheXml).toContain('refreshOnLoad="1"')
		expect(cacheXml).toContain('invalid="1"')
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.pivotTables[0]?.fields[0]?.items).toEqual([
			{ index: 0, cacheIndex: 0, manualFilter: true },
			{ index: 1, cacheIndex: 1, showDetails: false },
			{ index: 2, itemType: 'default' },
		])
		expect(reopened.value.workbook.pivotTables[0]?.pageFields).toEqual([{ index: 0, item: 1 }])
	})

	it('scopes pivot item rewrites to direct main pivot fields', () => {
		const pivotTableXml = `<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="PivotTable1" cacheId="34">
  <pivotFields count="3">
    <pivotField><items count="1"><item x="0"/></items><extLst><ext><x14:pivotField><x14:items><x14:item x="99" h="1"/></x14:items></x14:pivotField></ext></extLst></pivotField>
    <pivotField><items count="1"><item x="1"/></items></pivotField>
    <pivotField><items count="1"><item x="2"/></items></pivotField>
  </pivotFields>
</pivotTableDefinition>`
		const updated = updatePivotTableDefinitionXml(pivotTableXml, {
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [
				{ index: 0 },
				{ index: 1 },
				{ index: 2, items: [{ index: 0, cacheIndex: 2, manualFilter: true }] },
			],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})

		expect(updated).toContain('<x14:item x="99" h="1"/>')
		expect(updated).toContain('<pivotField><items count="1"><item x="1"/></items></pivotField>')
		expect(updated).toContain(
			'<pivotField><items count="1"><item x="2" s="1"/></items></pivotField>',
		)
	})

	it('updates XML-legal single-quoted pivot attributes without duplicating them', () => {
		const pivotTableXml = `<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <pivotFields count='1'><pivotField><items count='1'><item x='0' h='1'/></items></pivotField></pivotFields>
  <pageFields count='1'><pageField fld='0' item='0'/></pageFields>
</pivotTableDefinition>`
		const updatedPivot = updatePivotTableDefinitionXml(pivotTableXml, {
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 1,
			fields: [{ index: 0, items: [{ index: 0, cacheIndex: 2, hidden: false }] }],
			rowFields: [],
			columnFields: [],
			pageFields: [{ index: 0, item: 1 }],
			dataFields: [],
		})

		expect(updatedPivot).toContain('<item x="2" h="0"/>')
		expect(updatedPivot).toContain('<pageField fld=\'0\' item="1"/>')
		expect(updatedPivot).not.toContain("x='0'")
		expect(updatedPivot).not.toContain("item='0' item=")

		const pivotCacheXml = `<pivotCacheDefinition refreshOnLoad='0' invalid='0'><cacheSource type='worksheet'><worksheetSource sheet='Old' ref='A1:B2'/></cacheSource></pivotCacheDefinition>`
		const updatedCache = updatePivotCacheDefinitionXml(pivotCacheXml, {
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 1,
			refreshOnLoad: true,
			invalid: false,
			sourceSheet: 'New',
			sourceRef: 'C1:D2',
		})

		expect(updatedCache).toContain('refreshOnLoad="1"')
		expect(updatedCache).toContain('invalid="0"')
		expect(updatedCache).toContain('sheet="New"')
		expect(updatedCache).toContain('ref="C1:D2"')
		expect(updatedCache).not.toContain("refreshOnLoad='0' refreshOnLoad=")
		expect(updatedCache).not.toContain("sheet='Old' sheet=")
	})

	it('updates prefixed pivot cache definition and worksheet source tags', () => {
		const pivotCacheXml = `<?xml version="1.0"?>
<x:pivotCacheDefinition xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad='0' invalid='1' recordCount="2">
  <x:cacheSource type="worksheet">
    <x:worksheetSource sheet='Old' ref='A1:B2'/>
  </x:cacheSource>
  <x:cacheFields count="1"/>
</x:pivotCacheDefinition>`
		const updated = updatePivotCacheDefinitionXml(pivotCacheXml, {
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 1,
			refreshOnLoad: true,
			enableRefresh: false,
			invalid: false,
			saveData: true,
			sourceSheet: 'New Data',
			sourceRef: 'C1:D20',
		})

		expect(updated).toContain('<x:pivotCacheDefinition')
		expect(updated).toContain('refreshOnLoad="1"')
		expect(updated).toContain('enableRefresh="0"')
		expect(updated).toContain('invalid="0"')
		expect(updated).toContain('saveData="1"')
		expect(updated).toContain('<x:cacheSource type="worksheet">')
		expect(updated).toContain('<x:worksheetSource sheet="New Data" ref="C1:D20"/>')
		expect(updated).toContain('<x:cacheFields count="1"/>')
		expect(updated).not.toContain('<pivotCacheDefinition')
		expect(updated).not.toContain('<worksheetSource')
		expect(updated).not.toContain("sheet='Old' sheet=")
	})

	it('inserts prefixed pivot cache source tags when source nodes are missing', () => {
		const withCacheSource = `<x:pivotCacheDefinition xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:cacheSource type="worksheet"></x:cacheSource></x:pivotCacheDefinition>`
		const updatedWithCacheSource = updatePivotCacheDefinitionXml(withCacheSource, {
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 1,
			sourceSheet: 'Raw',
			sourceRef: 'A1:E200',
		})

		expect(updatedWithCacheSource).toContain(
			'<x:cacheSource type="worksheet"><x:worksheetSource ref="A1:E200" sheet="Raw"/>',
		)
		expect(updatedWithCacheSource).not.toContain('<worksheetSource')

		const withSelfClosingCacheSource = `<x:pivotCacheDefinition xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:cacheSource type="worksheet"/></x:pivotCacheDefinition>`
		const updatedWithSelfClosingCacheSource = updatePivotCacheDefinitionXml(
			withSelfClosingCacheSource,
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				cacheId: 1,
				sourceSheet: 'Raw',
				sourceRef: 'A1:E200',
			},
		)

		expect(updatedWithSelfClosingCacheSource).toContain(
			'<x:cacheSource type="worksheet"><x:worksheetSource ref="A1:E200" sheet="Raw"/></x:cacheSource>',
		)
		expect(updatedWithSelfClosingCacheSource).not.toContain('<x:cacheSource type="worksheet"/>')

		const withoutCacheSource = `<x:pivotCacheDefinition xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0"><x:cacheFields count="0"/></x:pivotCacheDefinition>`
		const updatedWithoutCacheSource = updatePivotCacheDefinitionXml(withoutCacheSource, {
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 1,
			sourceSheet: 'Raw',
			sourceRef: 'A1:E200',
		})

		expect(updatedWithoutCacheSource).toContain(
			'<x:pivotCacheDefinition xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0"><x:cacheSource type="worksheet"><x:worksheetSource ref="A1:E200" sheet="Raw"/></x:cacheSource>',
		)
		expect(updatedWithoutCacheSource).toContain('<x:cacheFields count="0"/>')
		expect(updatedWithoutCacheSource).not.toContain('<cacheSource')
	})

	it('updates XML-legal single-quoted connection attributes without duplicating them', () => {
		const xml = `<connections><connection id='1' name='Sales' refreshOnLoad='0' saveData='1' background='1' keepAlive='1' interval='15' refreshedVersion='5'/></connections>`
		const updated = updateConnectionPartXml(xml, [
			{
				kind: 'connection',
				partPath: 'xl/connections.xml',
				connectionId: 1,
				name: 'Sales',
				refreshOnLoad: true,
				saveData: false,
				backgroundRefresh: false,
				keepAlive: false,
				refreshInterval: 30,
				refreshedVersion: 7,
			},
		])

		expect(updated).toContain('refreshOnLoad="1"')
		expect(updated).toContain('saveData="0"')
		expect(updated).toContain('background="0"')
		expect(updated).toContain('keepAlive="0"')
		expect(updated).toContain('interval="30"')
		expect(updated).toContain('refreshedVersion="7"')
		expect(updated).not.toContain("refreshOnLoad='0' refreshOnLoad=")
		expect(updated).not.toContain("background='1' background=")
		expect(updated).not.toContain("interval='15' interval=")
		expect(updated).not.toContain("refreshedVersion='5' refreshedVersion=")
	})

	it('writes real Calamine pivot page-filter edits and dirties the linked cache', () => {
		const source = readXlsx(
			readFileSync(new URL('../../../../fixtures/xlsx/calamine/pivots.xlsx', import.meta.url)),
		)
		expectOk(source)
		const applied = applyOperations(source.value.workbook, [
			{
				op: 'setPivotFieldItem',
				partPath: 'xl/pivotTables/pivotTable1.xml',
				fieldIndex: 5,
				itemIndex: 2,
				hidden: null,
				manualFilter: true,
				selectedPageItem: 1,
			},
		])
		expectOk(applied)

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const pivotXml = new TextDecoder().decode(
			zip['xl/pivotTables/pivotTable1.xml'] ?? new Uint8Array(),
		)
		const cacheXml = new TextDecoder().decode(
			zip['xl/pivotCache/pivotCacheDefinition1.xml'] ?? new Uint8Array(),
		)

		expect(pivotXml).toContain('<pageField fld="5" hier="-1" item="1"/>')
		expect(pivotXml).toContain('<item x="1" s="1"/>')
		expect(cacheXml).toContain('refreshOnLoad="1"')
		expect(cacheXml).toContain('invalid="1"')
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const pivot = reopened.value.workbook.pivotTables.find(
			(entry) => entry.partPath === 'xl/pivotTables/pivotTable1.xml',
		)
		expect(pivot?.pageFields).toEqual([{ index: 5, item: 1, hierarchy: -1 }])
		expect(pivot?.fields[5]?.items?.[2]).toEqual({ index: 2, cacheIndex: 1, manualFilter: true })
		expect(reopened.value.workbook.pivotCaches.find((cache) => cache.cacheId === 65)).toMatchObject(
			{
				refreshOnLoad: true,
				invalid: true,
			},
		)
	})

	it.skipIf(!existsSync(DASHBOARD_CORPUS_FIXTURE))(
		'writes real dashboard slicer edits and refresh flags into linked pivot cache XML',
		() => {
			const source = readXlsx(readFileSync(DASHBOARD_CORPUS_FIXTURE))
			expectOk(source)
			const applied = applyOperations(source.value.workbook, [
				{
					op: 'setSlicerCacheItem',
					slicerCache: 'Slicer_Product_Category',
					item: 2,
					selected: false,
					noData: true,
				},
			])
			expectOk(applied)
			expect(applied.value.warnings?.[0]?.details).toMatchObject({
				slicerCache: 'Slicer_Product_Category',
				cacheIds: [34],
				cachePartPaths: ['xl/pivotCache/pivotCacheDefinition1.xml'],
			})

			const written = writeXlsx(source.value.workbook, source.value.capsules, {
				dirtySheetNames: applied.value.sheetsModified,
			})
			expectOk(written)
			const zip = unzipSync(written.value)
			const slicerXml = new TextDecoder().decode(
				zip['xl/slicerCaches/slicerCache1.xml'] ?? new Uint8Array(),
			)
			const cacheXml = new TextDecoder().decode(
				zip['xl/pivotCache/pivotCacheDefinition1.xml'] ?? new Uint8Array(),
			)

			expect(slicerXml).toContain('<i x="2" s="0" nd="1"/>')
			expect(cacheXml).toContain('refreshOnLoad="1"')
			expect(cacheXml).toContain('invalid="1"')
			const reopened = readXlsx(written.value)
			expectOk(reopened)
			expect(
				reopened.value.workbook.slicerCaches
					.find((cache) => cache.name === 'Slicer_Product_Category')
					?.items?.find((item) => item.index === 2),
			).toEqual({ index: 2, selected: false, noData: true })
			expect(
				reopened.value.workbook.pivotCaches.find((cache) => cache.cacheId === 34),
			).toMatchObject({
				refreshOnLoad: true,
				invalid: true,
			})
		},
		120_000,
	)

	it('preserves unchanged pivot and slicer capsules byte-for-byte', () => {
		const pivotTableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="34">
  <pivotFields count="1"><pivotField axis="axisPage" showAll="false"><items count="2"><item h="true" x="0"></item><item t="default"/></items></pivotField></pivotFields>
  <pageFields count="1"><pageField fld="0"></pageField></pageFields>
</pivotTableDefinition>`
		const pivotCacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="true" invalid="false" saveData="true">
  <cacheSource type="worksheet"><worksheetSource sheet="Data" ref="A1:B2"/></cacheSource>
</pivotCacheDefinition>`
		const slicerCacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_State" sourceName="State">
  <pivotTables><pivotTable tabId="1" name="PivotTable1"/></pivotTables>
  <data><tabular pivotCacheId="34"><items count="2"><i s="true" x="0"></i><i x="1" nd="false"/></items></tabular></data>
</slicerCacheDefinition>`
		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [
				{
					index: 0,
					axis: 'axisPage',
					showAll: false,
					items: [
						{ index: 0, cacheIndex: 0, hidden: true },
						{ index: 1, itemType: 'default' },
					],
				},
			],
			rowFields: [],
			columnFields: [],
			pageFields: [{ index: 0 }],
			dataFields: [],
		})
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			refreshOnLoad: true,
			invalid: false,
			saveData: true,
			sourceSheet: 'Data',
			sourceRef: 'A1:B2',
			fields: [],
		})
		wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_State',
			sourceName: 'State',
			pivotCacheId: 34,
			pivotTableNames: ['PivotTable1'],
			items: [
				{ index: 0, selected: true },
				{ index: 1, noData: false },
			],
		})
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/pivotTables/pivotTable1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
				relationships: [],
				content: new TextEncoder().encode(pivotTableXml),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: 'Sheet1' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
			},
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				contentType:
					'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
				relationships: [],
				content: new TextEncoder().encode(pivotCacheXml),
				anchor: { kind: 'workbook' },
				relType:
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
			},
			{
				partPath: 'xl/slicerCaches/slicerCache1.xml',
				contentType: 'application/vnd.ms-excel.slicerCache+xml',
				relationships: [],
				content: new TextEncoder().encode(slicerCacheXml),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
			},
		]

		const written = writeXlsx(wb, capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const decodePart = (path: string) => new TextDecoder().decode(zip[path] ?? new Uint8Array())

		expect(decodePart('xl/pivotTables/pivotTable1.xml')).toBe(pivotTableXml)
		expect(decodePart('xl/pivotCache/pivotCacheDefinition1.xml')).toBe(pivotCacheXml)
		expect(decodePart('xl/slicerCaches/slicerCache1.xml')).toBe(slicerCacheXml)
	})

	it('writes edited drawing text back into preserved drawing XML', () => {
		const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="10" name="Callout" descr="Revenue note"/></xdr:nvSpPr>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1"/><a:t>Revenue </a:t></a:r><a:r><a:t>up</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`
		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.drawingRefs = { hasDrawing: true, hasLegacyDrawing: false }
		sheet.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'textBox',
			id: 10,
			name: 'Callout',
			description: 'Revenue note',
			text: 'Revenue up',
			anchor: {
				kind: 'twoCell',
				editAs: 'oneCell',
				from: { row: 2, col: 1 },
				to: { row: 4, col: 3 },
			},
		})
		const applied = applyOperations(wb, [
			{
				op: 'setDrawingText',
				sheet: 'Sheet1',
				drawingPartPath: 'xl/drawings/drawing1.xml',
				id: 10,
				text: 'Revenue flat',
			},
		])
		expectOk(applied)
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/drawings/drawing1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
				relationships: [],
				content: new TextEncoder().encode(drawingXml),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: 'Sheet1' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
			},
		]

		const written = writeXlsx(wb, capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/drawings/drawing1.xml'] ?? new Uint8Array())

		expect(xml).toContain('<xdr:twoCellAnchor editAs="oneCell">')
		expect(xml).toContain('<xdr:cNvPr id="10" name="Callout" descr="Revenue note"/>')
		expect(xml).toContain('<a:rPr b="1"/>')
		expect(xml).toContain('<a:t>Revenue flat</a:t>')
		expect(xml).toContain('<a:t></a:t>')
		expect(xml).not.toContain('Revenue </a:t>')
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.drawingObjectRefs[0]).toMatchObject({
			kind: 'textBox',
			id: 10,
			name: 'Callout',
			text: 'Revenue flat',
		})
	})

	it('preserves formula text on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: 'A1*2', styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(30), formula: 'SUM(A1,B1)', styleId: S0 })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.cells.get(0, 1)?.formula).toBe('A1*2')
		expect(s?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 20 })
		expect(s?.cells.get(1, 0)?.formula).toBe('SUM(A1,B1)')
	})

	it('preserves original stored formula text when unrelated edits dirty the sheet', () => {
		const storedFormula = `_xlfn.LET(
  _xlpm.value, A1,
  SUM(--(_xlfn.VSTACK(_xlpm.value, 2)))
)`
		const source = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><f>${storedFormula}</f><v>3</v></c>
      <c r="C1"><v>4</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})
		const read = readXlsx(source)
		expectOk(read)
		const sheet = read.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 1)?.formula).not.toContain('_xlfn.')
		expect(sheet?.storedFormulaText.get('0:1')).toBe(storedFormula)
		sheet?.cells.set(0, 2, { value: numberValue(5), formula: null, styleId: S0 })

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(xml).toContain(`<f>${storedFormula}</f>`)
		expect(xml).toContain('<c r="C1"><v>5</v></c>')
	})

	it('preserves formula cell metadata when unrelated edits dirty the sheet', () => {
		const source = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" cm="1" t="str"><f t="array" ref="A1">SUM(B1:B2)</f><v>7</v></c>
      <c r="B1"><v>3</v></c>
      <c r="B2"><v>4</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})
		const read = readXlsx(source)
		expectOk(read)
		const sheet = read.value.workbook.sheets[0]
		expect(sheet?.preservedCellMetadata.get('0:0')).toEqual({ cm: 1 })
		sheet?.cells.set(0, 1, { value: numberValue(5), formula: null, styleId: S0 })

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(xml).toContain(
			'<c r="A1" cm="1" t="str"><f t="array" ref="A1">SUM(B1:B2)</f><v>7</v></c>',
		)
		expect(xml).toContain('<c r="B1"><v>5</v></c>')
	})

	it('drops preserved dynamic-array metadata attrs when a formula cell is rewritten ordinary', () => {
		const source = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata" Target="metadata.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Dynamic" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/metadata.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
  <metadataTypes count="1">
    <metadataType name="XLDAPR" minSupportedVersion="120000" cellMeta="1"/>
  </metadataTypes>
  <futureMetadata name="XLDAPR" count="1">
    <bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>
  </futureMetadata>
  <cellMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </cellMetadata>
</metadata>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" cm="1"><f>_xlfn.SEQUENCE(2)</f><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const read = readXlsx(source)
		expectOk(read)
		const sheet = read.value.workbook.sheets[0]
		expect(sheet?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		sheet?.cells.set(0, 0, { value: numberValue(2), formula: '1+1', styleId: S0 })

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Dynamic'],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(sheetXml).toContain('<c r="A1"><f>1+1</f><v>2</v></c>')
		expect(sheetXml).not.toContain('<c r="A1" cm=')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.formula).toBe('1+1')
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.formulaInfo).toBeUndefined()
	})

	it('round-trips shared formula bindings', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Shared')
		sheet.cells.set(0, 1, { value: numberValue(42), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(84), formula: null, styleId: S0 })
		sheet.cells.set(0, 0, {
			value: numberValue(84),
			formula: 'B1*2',
			styleId: S0,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: true,
				masterRef: 'A1',
				ref: 'A1:A2',
			},
		})
		sheet.cells.set(1, 0, {
			value: numberValue(168),
			formula: null,
			styleId: S0,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'A1' },
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(xml).toContain('t="shared" si="0"')
		expect(xml).toContain('ref="A1:A2"')
		expect(xml).toContain('>B1*2</f>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const s = reopened.value.workbook.sheets[0]
		expect(s?.cells.get(0, 0)?.formula).toBe('B1*2')
		expect(s?.cells.get(0, 0)?.formulaInfo).toMatchObject({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			masterRef: 'A1',
		})
		expect(s?.cells.get(1, 0)?.formula).toBeNull()
		expect(s?.cells.get(1, 0)?.formulaInfo).toMatchObject({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'A1',
		})
	})

	it('splits oversized shared formulas into explicit formulas per cell', () => {
		const wb = new Workbook()
		const other = wb.addSheet('Other')
		other.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: S0 })
		other.cells.set(1, 1, { value: numberValue(3), formula: null, styleId: S0 })
		const sheet = wb.addSheet('SharedLong')
		const longFormula = Array.from({ length: 1000 }, () => 'Other!B1').join('+')
		sheet.cells.set(0, 0, {
			value: numberValue(1800),
			formula: longFormula,
			styleId: S0,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: true,
				masterRef: 'A1',
				ref: 'A1:A2',
			},
		})
		sheet.cells.set(1, 0, {
			value: numberValue(2700),
			formula: null,
			styleId: S0,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: false,
				masterRef: 'A1',
			},
		})

		const written = writeXlsx(wb)
		expectOk(written)
		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
		expect(xml).not.toContain('t="shared"')
		expect(xml).toContain('<f>Other!B1+Other!B1')
		expect(xml).toContain('<f>Other!B2+Other!B2')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[1]?.cells.get(0, 0)?.formula).toBe(longFormula)
		expect(
			reopened.value.workbook.sheets[1]?.cells.get(1, 0)?.formula?.startsWith('Other!B2'),
		).toBe(true)
	})

	it('round-trips regular formulas that could be shared, preserving formula text', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 1, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(20), formula: null, styleId: S0 })
		sheet.cells.set(2, 1, { value: numberValue(30), formula: null, styleId: S0 })
		sheet.cells.set(0, 0, { value: numberValue(20), formula: 'B1*2', styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(40), formula: 'B2*2', styleId: S0 })
		sheet.cells.set(2, 0, { value: numberValue(60), formula: 'B3*2', styleId: S0 })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.cells.get(0, 0)?.formula).toBe('B1*2')
		expect(s?.cells.get(1, 0)?.formula).toBe('B2*2')
		expect(s?.cells.get(2, 0)?.formula).toBe('B3*2')
		expect(s?.cells.get(0, 0)?.value).toMatchObject({ kind: 'number', value: 20 })
		expect(s?.cells.get(1, 0)?.value).toMatchObject({ kind: 'number', value: 40 })
		expect(s?.cells.get(2, 0)?.value).toMatchObject({ kind: 'number', value: 60 })
	})

	it('round-trips array formula bindings', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 0, {
			value: numberValue(3),
			formula: 'SUM(B1:B2)',
			styleId: S0,
			formulaInfo: { kind: 'array', ref: 'A1:A2' },
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(xml).toContain('<f t="array" ref="A1:A2">SUM(B1:B2)</f>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'array',
			ref: 'A1:A2',
		})
	})

	it('round-trips data-table formula bindings', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('DataTable')
		sheet.cells.set(0, 0, { value: numberValue(0.5), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, {
			value: numberValue(10),
			formula: null,
			styleId: S0,
			formulaInfo: {
				kind: 'dataTable',
				ref: 'B2:B4',
				dt2D: false,
				dtr: true,
				r1: 'A1',
			},
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const xml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(xml).toContain('<f t="dataTable" ref="B2:B4" dt2D="0" dtr="1" r1="A1"/>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const cell = reopened.value.workbook.sheets[0]?.cells.get(1, 1)
		expect(cell?.formula).toBeNull()
		expect(cell?.formulaInfo).toEqual({
			kind: 'dataTable',
			ref: 'B2:B4',
			dt2D: false,
			dtr: true,
			r1: 'A1',
		})
		expect(cell?.value).toEqual({ kind: 'number', value: 10 })
	})

	it('writes dynamic-array metadata and storage formula syntax', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Dynamic')
		sheet.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'SEQUENCE(3)',
			styleId: S0,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
		})
		sheet.cells.set(0, 1, {
			value: numberValue(6),
			formula: 'SUM(A1#)',
			styleId: S0,
		})
		sheet.cells.set(0, 2, {
			value: numberValue(1),
			formula: '@A1',
			styleId: S0,
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)
		const metadataXml = new TextDecoder().decode(zip['xl/metadata.xml'] ?? new Uint8Array())
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(contentTypes).toContain(
			'/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"',
		)
		expect(workbookRels).toContain(
			'Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata"',
		)
		expect(metadataXml).toContain('name="XLDAPR"')
		expect(sheetXml).toContain('cm="1"')
		expect(sheetXml).toContain('_xlfn.SEQUENCE(3)')
		expect(sheetXml).toContain('SUM(_xlfn.ANCHORARRAY(A1))')
		expect(sheetXml).toContain('_xlfn.SINGLE(A1)')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.formula).toBe('SEQUENCE(3)')
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 1)?.formula).toBe('SUM(A1#)')
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 2)?.formula).toBe('@A1')
	})

	it('preserves formula spacing when storage rewriting is a no-op', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 0, {
			value: errorValue('#REF!'),
			formula: 'CHOOSE(#REF!, 1)',
			styleId: S0,
		})
		sheet.cells.set(0, 1, {
			value: numberValue(1),
			formula: 'COUNT(2, "A",  "",#REF!, #DIV/0!)',
			styleId: S0,
		})

		const written = writeXlsx(wb)
		expectOk(written)
		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<f>CHOOSE(#REF!, 1)</f>')
		expect(sheetXml).toContain('<f>COUNT(2, &quot;A&quot;,  &quot;&quot;,#REF!, #DIV/0!)</f>')
	})

	it('writes all dynamic-array metadata records referenced by cells', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Dynamic')
		sheet.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'SEQUENCE(2)',
			styleId: S0,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
		})
		sheet.cells.set(0, 2, {
			value: numberValue(10),
			formula: 'SEQUENCE(2,1,10)',
			styleId: S0,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 2, collapsed: true },
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const metadataXml = new TextDecoder().decode(zip['xl/metadata.xml'] ?? new Uint8Array())
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(metadataXml).toContain('<futureMetadata name="XLDAPR" count="2">')
		expect(metadataXml).toContain('<cellMetadata count="2">')
		expect(metadataXml).toContain('v="0"')
		expect(metadataXml).toContain('v="1"')
		expect(metadataXml).toContain('fCollapsed="1"')
		expect(sheetXml).toContain('<c r="A1" cm="1"')
		expect(sheetXml).toContain('<c r="C1" cm="2"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 1,
			collapsed: false,
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 2)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 2,
			collapsed: true,
		})
	})

	it('regenerates preserved dynamic-array metadata when edits add records', () => {
		const source = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata" Target="metadata.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Dynamic" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/metadata.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
  <metadataTypes count="1">
    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>
  </metadataTypes>
  <futureMetadata name="XLDAPR" count="1">
    <bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>
  </futureMetadata>
  <cellMetadata count="1">
    <bk><rc t="1" v="0"/></bk>
  </cellMetadata>
</metadata>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" cm="1"><f>_xlfn.SEQUENCE(2)</f><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const read = readXlsx(source)
		expectOk(read)
		const sheet = read.value.workbook.sheets[0]
		sheet?.cells.set(0, 2, {
			value: numberValue(10),
			formula: 'SEQUENCE(2,1,10)',
			styleId: S0,
			formulaInfo: { kind: 'dynamicArray', metadataIndex: 2, collapsed: true },
		})

		const written = writeXlsx(read.value.workbook, undefined, { dirtySheetNames: ['Dynamic'] })
		expectOk(written)
		const zip = unzipSync(written.value)
		const metadataXml = new TextDecoder().decode(zip['xl/metadata.xml'] ?? new Uint8Array())
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(metadataXml).toContain('<futureMetadata name="XLDAPR" count="2">')
		expect(metadataXml).toContain('<cellMetadata count="2">')
		expect(metadataXml).toContain('fCollapsed="1"')
		expect(sheetXml).toContain('<c r="C1" cm="2"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 2)?.formulaInfo).toEqual({
			kind: 'dynamicArray',
			metadataIndex: 2,
			collapsed: true,
		})
	})

	it('round-trips _xlfn. prefix for future functions', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Xlfn')
		const formulas = [
			'IFS(A1>0,"pos",A1<0,"neg",TRUE,"zero")',
			'TEXTJOIN(",",TRUE,A1:A3)',
			'SWITCH(A1,1,"one",2,"two","other")',
			'MAXIFS(A1:A3,B1:B3,">0")',
			'CONCAT(A1,B1)',
			'IFNA(A1,0)',
			'XOR(TRUE,FALSE)',
		]
		for (const [i, formula] of formulas.entries()) {
			sheet.cells.set(i, 0, {
				value: stringValue(''),
				formula,
				styleId: S0,
			})
		}

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(sheetXml).toContain('_xlfn.IFS(')
		expect(sheetXml).toContain('_xlfn.TEXTJOIN(')
		expect(sheetXml).toContain('_xlfn.SWITCH(')
		expect(sheetXml).toContain('_xlfn.MAXIFS(')
		expect(sheetXml).toContain('_xlfn.CONCAT(')
		expect(sheetXml).toContain('_xlfn.IFNA(')
		expect(sheetXml).toContain('_xlfn.XOR(')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const cells = reopened.value.workbook.sheets[0]?.cells
		for (let i = 0; i < formulas.length; i++) {
			expect(cells?.get(i, 0)?.formula).toBe(formulas[i])
		}
	})

	it('preserves sharedStrings.xml when string indices are unchanged', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="2">
  <si><t>World</t></si>
  <si><t>Hello</t></si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>1</v></c>
      <c r="B1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)

		const sheet = source.value.workbook.sheets[0]
		expect(sheet).toBeDefined()
		sheet?.cells.set(0, 1, { value: numberValue(99), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
			sharedStringsDirty: false,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sharedStrings = new TextDecoder().decode(zip['xl/sharedStrings.xml'] ?? new Uint8Array())
		expect(sharedStrings).toContain('<t>World</t>')
		expect(sharedStrings).toContain('<t>Hello</t>')
		expect(sharedStrings).not.toBe('')
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'Hello',
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(1, 0)?.value).toEqual({
			kind: 'string',
			value: 'World',
		})
	})

	it('passes through preserved source ZIP parts on numeric dirty-sheet edits', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Shared payload that should remain compressed exactly once.</t></si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		source.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: numberValue(99),
			formula: null,
			styleId: S0,
		})
		const sourceArchive = new ZipArchive(sourceBytes)
		let sharedStringsTextReads = 0
		const readText = sourceArchive.readText.bind(sourceArchive)
		sourceArchive.readText = (path: string) => {
			if (path === 'xl/sharedStrings.xml') sharedStringsTextReads++
			return readText(path)
		}

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
			dirtyCellPatches: [{ sheetName: 'Data', refs: ['B1'] }],
			sharedStringsDirty: false,
			sourceArchive,
		})
		expectOk(written)
		expect(sharedStringsTextReads).toBe(0)

		const writtenZip = new ZipArchive(written.value)
		expect(writtenZip.readCompressedBytes('xl/sharedStrings.xml')).toEqual(
			sourceArchive.readCompressedBytes('xl/sharedStrings.xml'),
		)
		expect(writtenZip.get('xl/worksheets/sheet1.xml')?.crc).not.toBe(
			sourceArchive.get('xl/worksheets/sheet1.xml')?.crc,
		)
	})

	it('passes through preserved source ZIP parts on streaming dirty-sheet edits', async () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Shared payload that should remain compressed exactly once.</t></si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		source.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: numberValue(99),
			formula: null,
			styleId: S0,
		})
		const sourceArchive = new ZipArchive(sourceBytes)
		let sharedStringsTextReads = 0
		const readText = sourceArchive.readText.bind(sourceArchive)
		sourceArchive.readText = (path: string) => {
			if (path === 'xl/sharedStrings.xml') sharedStringsTextReads++
			return readText(path)
		}

		const written = await writeXlsxStreaming(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
			dirtyCellPatches: [{ sheetName: 'Data', refs: ['B1'] }],
			sharedStringsDirty: false,
			sourceArchive,
		})
		expectOk(written)
		expect(sharedStringsTextReads).toBe(0)

		const writtenZip = new ZipArchive(written.value)
		expect(writtenZip.readCompressedBytes('xl/sharedStrings.xml')).toEqual(
			sourceArchive.readCompressedBytes('xl/sharedStrings.xml'),
		)
		expect(writtenZip.get('xl/worksheets/sheet1.xml')?.crc).not.toBe(
			sourceArchive.get('xl/worksheets/sheet1.xml')?.crc,
		)
	})

	it('keeps shared string positions when preserved entries contain duplicates', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Alpha</t></si>
  <si><t>Beta</t></si>
  <si><t>Beta</t></si>
  <si><t>Gamma</t></si>
</sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>2</v></c>
      <c r="B1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)
		source.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: numberValue(99),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
			sharedStringsDirty: false,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sharedStrings = new TextDecoder().decode(zip['xl/sharedStrings.xml'] ?? new Uint8Array())
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sharedStrings.match(/<si>/g)).toHaveLength(4)
		expect(sheetXml).toContain('<row r="2"><c t="s"><v>3</v></c></row>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'Beta',
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(1, 0)?.value).toEqual({
			kind: 'string',
			value: 'Gamma',
		})
	})

	it('reuses rich sharedStrings.xml on numeric dirty-sheet edits', () => {
		const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><r><t>課きく</t></r><rPh sb="0" eb="1"><r><t>カ</t></r></rPh><phoneticPr fontId="1"/></si>
</sst>`
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
			'xl/sharedStrings.xml': sharedStringsXml,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		source.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: numberValue(2),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sharedStrings = new TextDecoder().decode(zip['xl/sharedStrings.xml'] ?? new Uint8Array())
		expect(sharedStrings).toBe(sharedStringsXml)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<row r="1"><c t="s"><v>0</v></c><c><v>2</v></c></row>')
	})

	it('appends to preserved rich sharedStrings.xml without rewriting existing entries', () => {
		const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><r><t>課きく</t></r><rPh sb="0" eb="1"><r><t>カ</t></r></rPh><phoneticPr fontId="1"/></si>
</sst>`
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
			'xl/sharedStrings.xml': sharedStringsXml,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		source.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: stringValue('New'),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sharedStrings = new TextDecoder().decode(zip['xl/sharedStrings.xml'] ?? new Uint8Array())
		expect(sharedStrings).toContain(
			'<si><r><t>課きく</t></r><rPh sb="0" eb="1"><r><t>カ</t></r></rPh><phoneticPr fontId="1"/></si>',
		)
		expect(sharedStrings).toContain(
			'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">',
		)
		expect(sharedStrings).not.toContain('standalone="yes" count=')
		expect(sharedStrings).not.toContain('</sst><si>')
		expect(sharedStrings).toContain('<si><t>New</t></si>')
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<row r="1"><c t="s"><v>0</v></c><c t="s"><v>1</v></c></row>')
	})

	it('keeps dirty inline-string workbooks inline when the source has no sharedStrings part', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Inline</t></is></c><c r="B1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		source.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: numberValue(2),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		expect(zip['xl/sharedStrings.xml']).toBeUndefined()
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<c t="inlineStr"><is><t>Inline</t></is></c>')
		expect(sheetXml).toContain('<c><v>2</v></c>')
	})

	it('does not add styles or docProps to dirty minimal packages that do not need them', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
			'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Old</t></si></sst>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		sheet.cells.set(0, 0, { value: stringValue('New'), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		const rootRels = new TextDecoder().decode(zip['_rels/.rels'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)
		expect(zip['xl/styles.xml']).toBeUndefined()
		expect(zip['docProps/core.xml']).toBeUndefined()
		expect(zip['docProps/app.xml']).toBeUndefined()
		expect(contentTypes).not.toContain('styles+xml')
		expect(contentTypes).not.toContain('/docProps/')
		expect(rootRels).not.toContain('core-properties')
		expect(rootRels).not.toContain('extended-properties')
		expect(workbookRels).not.toContain('relationships/styles')
	})

	it('preserves package docProps parts and custom property relationships', () => {
		const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Original Author</dc:creator>
  <cp:lastModifiedBy>Reviewer</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created>
</cp:coreProperties>`
		const appXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Excel</Application>
  <Company>Acme Analytics</Company>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>1</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>Data</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
</Properties>`
		const customXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Desk">
    <vt:lpwstr>Research</vt:lpwstr>
  </property>
</Properties>`
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
</Relationships>`,
			'docProps/core.xml': coreXml,
			'docProps/app.xml': appXml,
			'docProps/custom.xml': customXml,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)
		expect(source.value.workbook.documentProperties).toEqual({
			core: {
				creator: 'Original Author',
				lastModifiedBy: 'Reviewer',
				created: '2024-01-02T03:04:05Z',
			},
			app: {
				Application: 'Excel',
				Company: 'Acme Analytics',
				HeadingPairs: ['Worksheets', 1],
				TitlesOfParts: ['Data'],
			},
			custom: [
				{
					name: 'Desk',
					value: 'Research',
					type: 'lpwstr',
					pid: 2,
					fmtid: '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
				},
			],
		})
		expect(source.value.capsules.map((capsule) => capsule.partPath).sort()).toContain(
			'docProps/custom.xml',
		)

		const written = writeXlsx(source.value.workbook, source.value.capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const decode = (path: string) => new TextDecoder().decode(zip[path] ?? new Uint8Array())

		expect(decode('docProps/core.xml')).toBe(coreXml)
		expect(decode('docProps/app.xml')).toBe(appXml)
		expect(decode('docProps/custom.xml')).toBe(customXml)
		expect(decode('_rels/.rels')).toContain(
			'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"',
		)
		expect(decode('_rels/.rels')).toContain('Target="docProps/custom.xml"')
		expect(decode('xl/_rels/workbook.xml.rels')).not.toContain('docProps/custom.xml')
	})

	it('writes edited package document properties without preserving stale docProps XML', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
</Relationships>`,
			'docProps/core.xml':
				'<cp:coreProperties xmlns:cp="old"><dc:title>Old</dc:title></cp:coreProperties>',
			'docProps/app.xml': '<Properties><Application>Old</Application></Properties>',
			'docProps/custom.xml':
				'<Properties><property name="Old"><vt:lpwstr>stale</vt:lpwstr></property></Properties>',
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		source.value.workbook.documentProperties = {
			core: { title: 'Board Pack', creator: 'Finance' },
			app: { Application: 'Ascend', Company: 'Ascend Fixtures' },
			custom: [{ name: 'Reviewed', value: true, type: 'bool' }],
		}

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			workbookMetaDirty: true,
			documentPropertiesDirty: true,
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const decode = (path: string) => new TextDecoder().decode(zip[path] ?? new Uint8Array())

		expect(decode('docProps/core.xml')).toContain('<dc:title>Board Pack</dc:title>')
		expect(decode('docProps/core.xml')).toContain('<dc:creator>Finance</dc:creator>')
		expect(decode('docProps/app.xml')).toContain('<Application>Ascend</Application>')
		expect(decode('docProps/custom.xml')).toContain('name="Reviewed"')
		expect(decode('docProps/custom.xml')).toContain('<vt:bool>true</vt:bool>')
		expect(decode('docProps/custom.xml')).not.toContain('stale')
		expect(decode('_rels/.rels')).toContain(
			'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"',
		)

		source.value.workbook.documentProperties = {
			core: { title: 'Board Pack' },
			app: {
				Application: 'Ascend',
				HeadingPairs: ['Worksheets', 1],
				TitlesOfParts: ['Data'],
			},
		}
		const withVectors = writeXlsx(source.value.workbook, source.value.capsules, {
			workbookMetaDirty: true,
			documentPropertiesDirty: true,
		})
		expectOk(withVectors)
		const zipWithVectors = unzipSync(withVectors.value)
		const appWithVectors = new TextDecoder().decode(
			zipWithVectors['docProps/app.xml'] ?? new Uint8Array(),
		)
		expect(appWithVectors).toContain('<HeadingPairs><vt:vector size="2" baseType="variant">')
		expect(appWithVectors).toContain('<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>')
		expect(appWithVectors).toContain('<vt:variant><vt:i4>1</vt:i4></vt:variant>')
		expect(appWithVectors).toContain('<TitlesOfParts><vt:vector size="1" baseType="lpstr">')
		expect(appWithVectors).toContain('<vt:lpstr>Data</vt:lpstr>')

		source.value.workbook.documentProperties = {
			core: { title: 'Board Pack' },
			app: { Application: 'Ascend' },
		}
		const withoutCustom = writeXlsx(source.value.workbook, source.value.capsules, {
			workbookMetaDirty: true,
			documentPropertiesDirty: true,
		})
		expectOk(withoutCustom)
		const zipWithoutCustom = unzipSync(withoutCustom.value)
		const decodeWithoutCustom = (path: string) =>
			new TextDecoder().decode(zipWithoutCustom[path] ?? new Uint8Array())
		expect(zipWithoutCustom['docProps/custom.xml']).toBeUndefined()
		expect(decodeWithoutCustom('_rels/.rels')).not.toContain('custom-properties')
		expect(decodeWithoutCustom('[Content_Types].xml')).not.toContain('/docProps/custom.xml')
	})

	it('preserves nonstandard root core-properties topology', () => {
		const corePath = 'package/services/metadata/core-properties/source.psmdcp'
		const coreXml = `<?xml version="1.0" encoding="utf-8"?>
<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
  <dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">ClosedXML</dc:creator>
</coreProperties>`
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Default Extension="psmdcp" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="/docProps/app.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="/package/services/metadata/core-properties/source.psmdcp"/>
</Relationships>`,
			'docProps/app.xml': '<Properties/>',
			[corePath]: coreXml,
			'xl/workbook.xml': `<?xml version="1.0" encoding="utf-8"?>
<x:workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheets><x:sheet name="Data" sheetId="1" r:id="rId1"/></x:sheets>
</x:workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="utf-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		sheet?.cells.set(0, 0, { value: numberValue(424242), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules)
		expectOk(written)
		const zip = unzipSync(written.value)
		const decode = (path: string) => new TextDecoder().decode(zip[path] ?? new Uint8Array())

		expect(zip['docProps/core.xml']).toBeUndefined()
		expect(decode(corePath)).toBe(coreXml)
		expect(decode('[Content_Types].xml')).toContain(
			'<Default Extension="psmdcp" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
		)
		expect(decode('[Content_Types].xml')).not.toContain('PartName="/xl/workbook.xml"')
		expect(decode('[Content_Types].xml')).not.toContain('PartName="/docProps/core.xml"')
		expect(decode('_rels/.rels')).toContain(
			'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="/package/services/metadata/core-properties/source.psmdcp"',
		)
		expect(decode('xl/_rels/workbook.xml.rels')).not.toContain('core-properties')
	})

	it('classifies generated and preserved parts in the write plan', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`,
			'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
			'xl/theme/theme1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test Theme"/>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)

		const plan = planWriteXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Sheet1'],
		})
		expectOk(plan)

		const workbookPart = plan.value.descriptors.find((entry) => entry.path === 'xl/workbook.xml')
		const stylesPart = plan.value.descriptors.find((entry) => entry.path === 'xl/styles.xml')
		const themePart = plan.value.descriptors.find((entry) => entry.path === 'xl/theme/theme1.xml')
		const sheetPart = plan.value.descriptors.find(
			(entry) => entry.path === 'xl/worksheets/sheet1.xml',
		)
		expect(workbookPart?.origin).toBe('generated')
		expect(stylesPart?.origin).toBe('generated')
		expect(themePart?.origin).toBe('preserved-source')
		expect(sheetPart?.origin).toBe('generated')

		const summaryOnlyPlan = planWriteXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Sheet1'],
			summaryOnly: true,
		})
		expectOk(summaryOnlyPlan)
		const summarySheetPart = summaryOnlyPlan.value.descriptors.find(
			(entry) => entry.path === 'xl/worksheets/sheet1.xml',
		)
		expect(summaryOnlyPlan.value.parts.size).toBe(0)
		expect(summarySheetPart?.origin).toBe('generated')
	})

	it('dirty-part patching: only regenerates modified sheet when dirtySheetNames provided', () => {
		const wb = new Workbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		const s3 = wb.addSheet('Sheet3')
		s1.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		s2.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })
		s3.cells.set(0, 0, { value: numberValue(3), formula: null, styleId: S0 })

		const initialBytes = writeXlsx(wb)
		expectOk(initialBytes)
		const read = readXlsx(initialBytes.value)
		expectOk(read)

		read.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: numberValue(99),
			formula: null,
			styleId: S0,
		})

		const dirtyWrite = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Sheet1'],
		})
		expectOk(dirtyWrite)

		const reopened = readXlsx(dirtyWrite.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets).toHaveLength(3)
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 1,
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(0, 1)?.value).toEqual({
			kind: 'number',
			value: 99,
		})
		expect(reopened.value.workbook.sheets[1]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 2,
		})
		expect(reopened.value.workbook.sheets[2]?.cells.get(0, 0)?.value).toEqual({
			kind: 'number',
			value: 3,
		})

		const fullWrite = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Sheet1', 'Sheet2', 'Sheet3'],
		})
		expectOk(fullWrite)
		const fullReopened = readXlsx(fullWrite.value)
		expectOk(fullReopened)
		expect(fullReopened.value.workbook.sheets[0]?.cells.get(0, 1)?.value).toEqual({
			kind: 'number',
			value: 99,
		})
	})

	it('dirty-part patching preserves worksheet relationship sidecar casing', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/Sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/Sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Page1_1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/Sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <drawing r:id="rIdDrawing"/>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
			'xl/drawings/drawing1.xml': `<?xml version="1.0"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>`,
		})
		const source = readXlsx(bytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		expect(sheet?.preservedXml?.relsPath).toBe('xl/worksheets/_rels/sheet1.xml.rels')
		expect(sheet?.drawingRefs).toEqual({ hasDrawing: true, hasLegacyDrawing: false })

		sheet?.cells.set(44, 0, { value: numberValue(424242), formula: null, styleId: S0 })
		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Page1_1'],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		expect(zip['xl/worksheets/_rels/sheet1.xml.rels']).toBeDefined()
		expect(zip['xl/worksheets/_rels/Sheet1.xml.rels']).toBeUndefined()
		expect(zip['xl/drawings/drawing1.xml']).toBeDefined()
		const contentTypesXml = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		expect(contentTypesXml).toContain('PartName="/xl/worksheets/Sheet1.xml"')
		expect(contentTypesXml).not.toContain('PartName="/xl/worksheets/sheet1.xml"')
		const sheetRelsXml = new TextDecoder().decode(
			zip['xl/worksheets/_rels/sheet1.xml.rels'] ?? new Uint8Array(),
		)
		expect(sheetRelsXml).toContain('Target="../drawings/drawing1.xml"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.drawingRefs).toEqual({
			hasDrawing: true,
			hasLegacyDrawing: false,
		})
		expect(reopened.value.workbook.sheets[0]?.cells.get(44, 0)?.value).toEqual({
			kind: 'number',
			value: 424242,
		})
	})

	it('dirty-part patching preserves shared string indexes for untouched sheets', () => {
		const wb = new Workbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		s1.cells.set(0, 0, { value: stringValue('alpha'), formula: null, styleId: S0 })
		s2.cells.set(0, 0, { value: stringValue('beta'), formula: null, styleId: S0 })
		const initial = writeXlsx(wb)
		expectOk(initial)
		const reopened = readXlsx(initial.value)
		expectOk(reopened)
		reopened.value.workbook.sheets[0]?.cells.set(0, 1, {
			value: stringValue('gamma'),
			formula: null,
			styleId: S0,
		})
		const patched = writeXlsx(reopened.value.workbook, reopened.value.capsules, {
			dirtySheetNames: ['Sheet1'],
		})
		expectOk(patched)
		const roundTripped = readXlsx(patched.value)
		expectOk(roundTripped)
		expect(roundTripped.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual(
			stringValue('alpha'),
		)
		expect(roundTripped.value.workbook.sheets[0]?.cells.get(0, 1)?.value).toEqual(
			stringValue('gamma'),
		)
		expect(roundTripped.value.workbook.sheets[1]?.cells.get(0, 0)?.value).toEqual(
			stringValue('beta'),
		)
	})

	it('preserves bold style on round-trip', () => {
		const wb = new Workbook()
		const boldId = wb.styles.register({ font: { bold: true } })
		const sheet = wb.addSheet('Styled')
		sheet.cells.set(0, 0, { value: stringValue('Bold'), formula: null, styleId: boldId })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		const cell = s?.cells.get(0, 0)
		expect(cell).toBeDefined()
		const style = result.workbook.styles.get(cell?.styleId ?? (0 as StyleId))
		expect(style?.font?.bold).toBe(true)
	})

	it('preserves number format on round-trip', () => {
		const wb = new Workbook()
		const pctId = wb.styles.register({ numberFormat: '0.00%' })
		const sheet = wb.addSheet('Fmt')
		sheet.cells.set(0, 0, { value: numberValue(0.75), formula: null, styleId: pctId })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		const cell = s?.cells.get(0, 0)
		expect(cell).toBeDefined()
		const style = result.workbook.styles.get(cell?.styleId ?? (0 as StyleId))
		expect(style?.numberFormat).toBe('0.00%')
	})

	it('1904 date system round-trips correctly', () => {
		const wb = new Workbook()
		wb.calcSettings = { ...wb.calcSettings, dateSystem: '1904' }
		const dateFmtId = wb.styles.register({ numberFormat: 'yyyy-mm-dd' })
		const sheet = wb.addSheet('Dates')
		sheet.cells.set(0, 0, {
			value: { kind: 'date', serial: 0 },
			formula: null,
			styleId: dateFmtId,
		})
		sheet.cells.set(1, 0, {
			value: { kind: 'date', serial: 1 },
			formula: null,
			styleId: dateFmtId,
		})
		sheet.cells.set(2, 0, {
			value: { kind: 'date', serial: 100 },
			formula: null,
			styleId: dateFmtId,
		})

		const { bytes, result } = roundTrip(wb)
		expect(result.workbook.calcSettings?.dateSystem).toBe('1904')
		const s = result.workbook.sheets[0]
		expect(s?.cells.get(0, 0)?.value).toEqual({ kind: 'date', serial: 0 })
		expect(s?.cells.get(1, 0)?.value).toEqual({ kind: 'date', serial: 1 })
		expect(s?.cells.get(2, 0)?.value).toEqual({ kind: 'date', serial: 100 })

		const zip = unzipSync(bytes)
		const wbXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		expect(wbXml).toContain('date1904="1"')
	})

	describe('number format edge cases', () => {
		it('round-trips custom number format #,##0.00', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '#,##0.00' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: numberValue(1234.5), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('#,##0.00')
		})

		it('round-trips percentage format 0.00%', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '0.00%' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: numberValue(0.125), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('0.00%')
		})

		it('round-trips currency format $#,##0', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '$#,##0' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: numberValue(999), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('$#,##0')
		})

		it('round-trips accounting format _($* #,##0.00_)', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '_($* #,##0.00_)' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: numberValue(-1234.56), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('_($* #,##0.00_)')
		})

		it('round-trips date format yyyy-mm-dd', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: 'yyyy-mm-dd' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, {
				value: { kind: 'date', serial: 45292 },
				formula: null,
				styleId: fmtId,
			})

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('yyyy-mm-dd')
		})

		it('round-trips date format dd/mm/yyyy', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: 'dd/mm/yyyy' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, {
				value: { kind: 'date', serial: 44927 },
				formula: null,
				styleId: fmtId,
			})

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('dd/mm/yyyy')
		})

		it('round-trips negative format with color [Red]', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '#,##0.00;[Red]-#,##0.00' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: numberValue(-100), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('#,##0.00;[Red]-#,##0.00')
		})

		it('round-trips text format @', () => {
			const wb = new Workbook()
			const fmtId = wb.styles.register({ numberFormat: '@' })
			const sheet = wb.addSheet('Fmt')
			sheet.cells.set(0, 0, { value: stringValue('ID-001'), formula: null, styleId: fmtId })

			const { result } = roundTrip(wb)
			const style = result.workbook.styles.get(
				result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
			)
			expect(style?.numberFormat).toBe('@')
		})
	})

	it('round-trips gradient fills', () => {
		const wb = new Workbook()
		const gradientId = wb.styles.register({
			fill: {
				gradient: {
					type: 'linear',
					degree: 45,
					stops: [
						{ position: 0, color: { kind: 'rgb', rgb: 'FFFF0000' } },
						{ position: 1, color: { kind: 'theme', theme: 1, tint: 0.25 } },
					],
				},
			},
		})
		const sheet = wb.addSheet('Styled')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: gradientId })

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const stylesXml = new TextDecoder().decode(zip['xl/styles.xml'] ?? new Uint8Array())
		expect(stylesXml).toContain('<gradientFill type="linear" degree="45">')
		expect(stylesXml).toContain('<stop position="0">')
		expect(stylesXml).toContain('<color rgb="FFFF0000"/>')
		expect(stylesXml).toContain('<color theme="1" tint="0.25"/>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const style = reopened.value.workbook.styles.get(
			reopened.value.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
		)
		expect(style?.fill).toEqual({
			gradient: {
				type: 'linear',
				degree: 45,
				stops: [
					{ position: 0, color: { kind: 'rgb', rgb: 'FFFF0000' } },
					{ position: 1, color: { kind: 'theme', theme: 1, tint: 0.25 } },
				],
			},
		})
	})

	it('appends number-format styles onto preserved styles.xml without rebuilding it', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
			'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count='1'><numFmt numFmtId='165' formatCode='0.0%'/></numFmts>
  <fonts count="1"><font/></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId='165' fontId='0' fillId='0' borderId='0' applyNumberFormat='1'/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">
    <tableStyle name="TableStyleMedium2"/>
  </tableStyles>
</styleSheet>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" s="1"><v>0.25</v></c><c r="B1"><v>0.5</v></c></row></sheetData>
</worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)

		const applied = applyOperations(source.value.workbook, [
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'B1:B1', format: '0.0%' },
		])
		expectOk(applied)

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Sheet1'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const stylesXml = new TextDecoder().decode(zip['xl/styles.xml'] ?? new Uint8Array())
		expect(stylesXml).toContain("formatCode='0.0%'")
		expect(stylesXml).toContain("numFmtId='165'")
		expect(stylesXml).not.toContain('numFmtId="164" formatCode="0.0%"')
		expect(stylesXml).toContain("applyNumberFormat='1'")
		expect(stylesXml).toContain('<cellStyleXfs')
		expect(stylesXml).toContain('<cellStyles')
		expect(stylesXml).toContain('defaultTableStyle="TableStyleMedium2"')
		expect(stylesXml).toContain('defaultPivotStyle="PivotStyleLight16"')
		expect(stylesXml).toContain('<tableStyle name="TableStyleMedium2"/>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const cell = reopened.value.workbook.sheets[0]?.cells.get(0, 0)
		const style = reopened.value.workbook.styles.get(cell?.styleId ?? (0 as StyleId))
		expect(style?.numberFormat).toBe('0.0%')
	})

	it('reuses XML-legal single-quoted preserved number formats when patching styles', () => {
		const wb = new Workbook()
		const baseStyleId = wb.styles.register({})
		const formattedStyleId = wb.styles.register({ numberFormat: '0.0%' })
		const sourceStylesXml = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count='1'><numFmt numFmtId='165' formatCode='0.0%'/></numFmts>
  <fonts count="1"><font/></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count='1'><xf numFmtId='0' fontId='0' fillId='0' borderId='0'/></cellXfs>
</styleSheet>`

		const result = buildPreservedStylesXml(
			sourceStylesXml,
			{
				path: 'xl/styles.xml',
				xfByStyleId: { [baseStyleId]: 0 },
				baseStyleIdByStyleId: { [formattedStyleId]: baseStyleId },
			},
			wb.styles,
		)

		expect(result?.xml).toContain("formatCode='0.0%'")
		expect(result?.xml).toContain('numFmtId="165"')
		expect(result?.xml).not.toContain('numFmtId="164" formatCode="0.0%"')
		expect(result?.xml).toContain('count="2"')
	})

	it('preserves merges on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Merges')
		sheet.cells.set(0, 0, { value: stringValue('Merged'), formula: null, styleId: S0 })
		sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 1, col: 2 } })
		sheet.merges.push({ start: { row: 3, col: 0 }, end: { row: 3, col: 1 } })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.merges).toHaveLength(2)
		expect(s?.merges[0]).toEqual({ start: { row: 0, col: 0 }, end: { row: 1, col: 2 } })
		expect(s?.merges[1]).toEqual({ start: { row: 3, col: 0 }, end: { row: 3, col: 1 } })
	})

	it('preserves worksheet layout metadata on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Layout')
		sheet.cells.set(0, 0, { value: stringValue('Header'), formula: null, styleId: S0 })
		sheet.frozenRows = 1
		sheet.frozenCols = 2
		sheet.colWidths.set(0, 18.5)
		sheet.colWidths.set(1, 18.5)
		sheet.rowHeights.set(0, 24)
		sheet.rowDefs.set(0, {
			spans: '1:2',
			style: 2,
			customFormat: true,
			customHeight: false,
			thickTop: true,
			thickBot: true,
			dyDescent: 0.3,
		})
		sheet.autoFilter = {
			ref: 'A1:B10',
			columns: [],
		}
		sheet.pageMargins = {
			left: 0.7,
			right: 0.7,
			top: 0.75,
			bottom: 0.75,
			header: 0.3,
			footer: 0.3,
		}
		sheet.pageSetup = {
			orientation: 'landscape',
			fitToWidth: 1,
			fitToHeight: 2,
			firstPageNumber: 2,
			copies: 3,
			horizontalDpi: 600,
			verticalDpi: 600,
			pageOrder: 'overThenDown',
			cellComments: 'asDisplayed',
			errors: 'dash',
			blackAndWhite: true,
			draft: true,
			useFirstPageNumber: true,
			usePrinterDefaults: false,
		}
		sheet.printOptions = {
			gridLines: true,
			gridLinesSet: true,
			headings: true,
		}
		sheet.headerFooter = {
			differentOddEven: true,
			differentFirst: true,
			scaleWithDoc: false,
			alignWithMargins: true,
			oddHeader: '&LTest',
			oddFooter: '&R1',
		}
		sheet.phoneticPr = {
			fontId: 1,
			type: 'noConversion',
			alignment: 'center',
		}
		sheet.rowBreaks = [{ id: 5, min: 0, max: 16383, man: true }]
		sheet.colBreaks = [{ id: 2, min: 0, max: 1048575, man: true }]

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.frozenRows).toBe(1)
		expect(s?.frozenCols).toBe(2)
		expect(s?.colWidths.get(0)).toBe(18.5)
		expect(s?.colWidths.get(1)).toBe(18.5)
		expect(s?.rowHeights.get(0)).toBe(24)
		expect(s?.rowDefs.get(0)).toEqual({
			spans: '1:2',
			style: 2,
			customFormat: true,
			customHeight: false,
			thickTop: true,
			thickBot: true,
			dyDescent: 0.3,
		})
		expect(s?.autoFilter).toEqual({
			ref: 'A1:B10',
			columns: [],
		})
		expect(s?.pageMargins).toEqual({
			left: 0.7,
			right: 0.7,
			top: 0.75,
			bottom: 0.75,
			header: 0.3,
			footer: 0.3,
		})
		expect(s?.pageSetup).toEqual({
			orientation: 'landscape',
			fitToWidth: 1,
			fitToHeight: 2,
			firstPageNumber: 2,
			copies: 3,
			horizontalDpi: 600,
			verticalDpi: 600,
			pageOrder: 'overThenDown',
			cellComments: 'asDisplayed',
			errors: 'dash',
			blackAndWhite: true,
			draft: true,
			useFirstPageNumber: true,
			usePrinterDefaults: false,
		})
		expect(s?.printOptions).toEqual({
			gridLines: true,
			gridLinesSet: true,
			headings: true,
		})
		expect(s?.headerFooter).toEqual({
			differentOddEven: true,
			differentFirst: true,
			scaleWithDoc: false,
			alignWithMargins: true,
			oddHeader: '&LTest',
			oddFooter: '&R1',
		})
		expect(s?.phoneticPr).toEqual({
			fontId: 1,
			type: 'noConversion',
			alignment: 'center',
		})
		expect(s?.rowBreaks).toEqual([{ id: 5, min: 0, max: 16383, man: true }])
		expect(s?.colBreaks).toEqual([{ id: 2, min: 0, max: 1048575, man: true }])
	})

	it('preserves split pane metadata without converting it to frozen panes', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Split')
		sheet.cells.set(0, 0, { value: stringValue('Header'), formula: null, styleId: S0 })
		sheet.preservedPaneAttributes = {
			xSplit: '12540',
			ySplit: '3840',
			topLeftCell: 'K23',
			activePane: 'bottomRight',
		}

		const { result, bytes } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.frozenRows).toBe(0)
		expect(s?.frozenCols).toBe(0)
		expect(s?.preservedPaneAttributes).toEqual({
			xSplit: '12540',
			ySplit: '3840',
			topLeftCell: 'K23',
			activePane: 'bottomRight',
		})
		const sheetXml = new TextDecoder().decode(unzipSync(bytes)['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain(
			'<pane xSplit="12540" ySplit="3840" topLeftCell="K23" activePane="bottomRight"/>',
		)
		expect(sheetXml).not.toContain('state="frozen"')
	})

	it('preserves macro-enabled workbook content type when workbook capsules are present', () => {
		const wb = new Workbook()
		wb.preservedXml = {
			workbookXml:
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
			workbookRelsXml:
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
			contentType: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
		}
		wb.addSheet('Sheet1')
		const macroCapsule: PreservationCapsule = {
			partPath: 'xl/vbaProject.bin',
			contentType: 'application/vnd.ms-office.vbaProject',
			relationships: [],
			content: new Uint8Array([0xde, 0xad, 0xbe, 0xef]),
			anchor: { kind: 'workbook' },
			relType: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
		}

		const written = writeXlsx(wb, [macroCapsule])
		if (!written.ok) throw new Error(`write failed: ${written.error.message}`)

		const zip = unzipSync(written.value)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)

		expect(contentTypes).toContain('application/vnd.ms-excel.sheet.macroEnabled.main+xml')
		expect(contentTypes).toContain('application/vnd.ms-office.vbaProject')
		expect(workbookRels).toContain('relationships/vbaProject')
		expect(zip['xl/vbaProject.bin']).toBeDefined()
	})

	it('preserves default-covered capsule content types and workbook-relative targets', () => {
		const wb = new Workbook()
		wb.preservedXml = {
			contentTypeDefaults: [{ extension: 'png', contentType: 'image/png' }],
		}
		wb.addSheet('Sheet1')
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'customXml/item1.xml',
				contentType: 'application/xml',
				contentTypeSource: 'default',
				relationships: [],
				content: new TextEncoder().encode('<root/>'),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
			},
			{
				partPath: 'xl/media/image1.png',
				contentType: 'image/png',
				contentTypeSource: 'default',
				relationships: [],
				content: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
				anchor: { kind: 'workbook' },
			},
		]

		const written = writeXlsx(wb, capsules)
		if (!written.ok) throw new Error(`write failed: ${written.error.message}`)

		const zip = unzipSync(written.value)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)

		expect(contentTypes).toContain('<Default Extension="png" ContentType="image/png"/>')
		expect(contentTypes).not.toContain('PartName="/customXml/item1.xml"')
		expect(contentTypes).not.toContain('PartName="/xl/media/image1.png"')
		expect(workbookRels).toContain('Target="../customXml/item1.xml"')
	})

	it('preserves explicit overrides for final parts even when defaults cover them', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)

		const contentTypes = new TextDecoder().decode(
			unzipSync(written.value)['[Content_Types].xml'] ?? new Uint8Array(),
		)
		expect(contentTypes).toContain('PartName="/_rels/.rels"')
		expect(contentTypes).toContain('PartName="/xl/_rels/workbook.xml.rels"')
	})

	it('preserves content type root attributes when package content types are regenerated', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:ct2="urn:ascend:test-content-types"
  mc:Ignorable="ct2"
  ct2:packageFlavor="review">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)

		const contentTypes = new TextDecoder().decode(
			unzipSync(written.value)['[Content_Types].xml'] ?? new Uint8Array(),
		)
		expect(contentTypes).toContain(
			'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
		)
		expect(contentTypes).toContain('xmlns:ct2="urn:ascend:test-content-types"')
		expect(contentTypes).toContain('mc:Ignorable="ct2"')
		expect(contentTypes).toContain('ct2:packageFlavor="review"')
		expect(contentTypes).toContain('PartName="/xl/workbook.xml"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.name).toBe('Data')
	})

	it('preserves relationship root attributes when relationship parts are regenerated', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:pkgrel="urn:ascend:package-relationships"
  mc:Ignorable="pkgrel"
  pkgrel:origin="root">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:wbrel="urn:ascend:workbook-relationships"
  mc:Ignorable="wbrel"
  wbrel:origin="workbook">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			workbookMetaDirty: true,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const rootRels = new TextDecoder().decode(zip['_rels/.rels'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)
		expect(rootRels).toContain('xmlns:pkgrel="urn:ascend:package-relationships"')
		expect(rootRels).toContain('mc:Ignorable="pkgrel"')
		expect(rootRels).toContain('pkgrel:origin="root"')
		expect(workbookRels).toContain('xmlns:wbrel="urn:ascend:workbook-relationships"')
		expect(workbookRels).toContain('mc:Ignorable="wbrel"')
		expect(workbookRels).toContain('wbrel:origin="workbook"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.name).toBe('Data')
	})

	it('preserves worksheet autoFilter criteria and sort state on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Filter')
		sheet.autoFilter = {
			ref: 'A1:B10',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					blank: true,
					values: ['Open', 'Closed'],
				},
				{
					colId: 1,
					kind: 'dynamicFilter',
					dynamicFilterType: 'thisMonth',
					dynamicFilterValIso: '2026-03-01T00:00:00',
					dynamicFilterMaxValIso: '2026-04-01T00:00:00',
				},
			],
			sortState: {
				ref: 'A2:B10',
				caseSensitive: true,
				conditions: [{ ref: 'B2:B10', descending: true }],
			},
		}

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.autoFilter).toEqual({
			ref: 'A1:B10',
			columns: [
				{
					colId: 0,
					kind: 'filters',
					blank: true,
					values: ['Open', 'Closed'],
				},
				{
					colId: 1,
					kind: 'dynamicFilter',
					dynamicFilterType: 'thisMonth',
					dynamicFilterValIso: '2026-03-01T00:00:00',
					dynamicFilterMaxValIso: '2026-04-01T00:00:00',
				},
			],
			sortState: {
				ref: 'A2:B10',
				caseSensitive: true,
				conditions: [{ ref: 'B2:B10', descending: true }],
			},
		})
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			autoFilter: 1,
			filterColumn: 2,
			filters: 1,
			filter: 2,
			dynamicFilter: 1,
			sortState: 1,
			sortCondition: 1,
		})
	})

	it('preserves top-level worksheet sort state on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Sorted')
		sheet.cells.set(0, 3, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.sortState = {
			ref: 'D1:I2707',
			conditions: [{ ref: 'D1' }],
		}

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.sortState).toEqual({
			ref: 'D1:I2707',
			conditions: [{ ref: 'D1' }],
		})
		const sheetXml = new TextDecoder().decode(unzipSync(bytes)['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain('<sortState ref="D1:I2707"><sortCondition ref="D1"/></sortState>')
	})

	it('preserves raw sortState attributes on regenerated sheets', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sorted" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <sortState ref="A1:A10" xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"><sortCondition ref="A1:A10"/></sortState>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Sorted'],
		})
		expectOk(written)

		const sheetXml = new TextDecoder().decode(
			unzipSync(written.value)['xl/worksheets/sheet1.xml'] ?? new Uint8Array(),
		)
		expect(sheetXml).toContain(
			'<sortState ref="A1:A10" xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2">',
		)
	})

	it('preserves raw autoFilter sortState attributes on regenerated sheets', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Filter" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <autoFilter ref="A1:A10"><sortState ref="A1:A10" xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"><sortCondition ref="A1:A10"/></sortState></autoFilter>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Filter'],
		})
		expectOk(written)

		const sheetXml = new TextDecoder().decode(
			unzipSync(written.value)['xl/worksheets/sheet1.xml'] ?? new Uint8Array(),
		)
		expect(sheetXml).toContain(
			'<sortState ref="A1:A10" xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2">',
		)
	})

	it('preserves workbook and sheet protection on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Protected')
		sheet.cells.set(0, 0, { value: stringValue('Locked'), formula: null, styleId: S0 })
		wb.workbookProtection = {
			lockStructure: true,
			workbookPassword: 'ABCD',
			workbookAlgorithmName: 'SHA-512',
			workbookSpinCount: 100000,
			extraAttributes: [{ name: 'futureProtectionMode', value: 'strict' }],
		}
		sheet.protection = {
			sheet: true,
			objects: true,
			scenarios: true,
			password: '1234',
			sort: false,
			autoFilter: false,
			selectUnlockedCells: true,
		}
		sheet.protectedRanges = [{ name: 'Editable', sqref: 'C:C', password: '1234' }]

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.workbookProtection).toEqual({
			lockStructure: true,
			workbookPassword: 'ABCD',
			workbookAlgorithmName: 'SHA-512',
			workbookSpinCount: 100000,
			extraAttributes: [{ name: 'futureProtectionMode', value: 'strict' }],
		})
		expect(result.workbook.sheets[0]?.protection).toEqual({
			sheet: true,
			objects: true,
			scenarios: true,
			password: '1234',
			sort: false,
			autoFilter: false,
			selectUnlockedCells: true,
		})
		expect(result.workbook.sheets[0]?.protectedRanges).toEqual([
			{ name: 'Editable', sqref: 'C:C', password: '1234' },
		])
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.workbook?.tagCounts).toMatchObject({
			workbookProtection: 1,
		})
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			sheetProtection: 1,
			protectedRanges: 1,
			protectedRange: 1,
		})
		const workbookXml = new TextDecoder().decode(
			unzipSync(bytes)['xl/workbook.xml'] ?? new Uint8Array(),
		)
		expect(workbookXml).toContain('futureProtectionMode="strict"')
	})

	it('preserves hyperlinks on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Links')
		sheet.cells.set(0, 0, { value: stringValue('Docs'), formula: null, styleId: S0 })
		sheet.hyperlinks.set('A1', {
			target: 'https://example.com/docs',
			display: 'Docs',
			tooltip: 'Open docs',
		})

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.hyperlinks.get('A1')).toEqual({
			target: 'https://example.com/docs',
			display: 'Docs',
			tooltip: 'Open docs',
		})
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			hyperlinks: 1,
			hyperlink: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.normalized).toContain(
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
		)
		expect(fingerprint.sheetRels[0]?.xml.normalized).toContain('TargetMode="External"')
	})

	it('preserves drawing and legacyDrawing references when sheet capsules exist', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Visuals')
		sheet.cells.set(0, 0, { value: stringValue('Chart host'), formula: null, styleId: S0 })
		sheet.drawingRefs = { hasDrawing: true, hasLegacyDrawing: true }

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/drawings/drawing1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
				relationships: [],
				content: new TextEncoder().encode(
					'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
				),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
			},
			{
				partPath: 'xl/drawings/vmlDrawing1.vml',
				contentType: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
				relationships: [],
				content: new TextEncoder().encode('<xml xmlns:v="urn:schemas-microsoft-com:vml"/>'),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
			},
		]

		const { result, bytes } = roundTrip(wb, capsules)
		expect(result.workbook.sheets[0]?.drawingRefs).toEqual({
			hasDrawing: true,
			hasLegacyDrawing: true,
		})
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			drawing: 1,
			legacyDrawing: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 2,
		})
	})

	it('generates drawing XML and media parts for programmatic image refs', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Images')
		sheet.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			content: new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
			anchor: {
				kind: 'oneCell',
				from: { row: 1, col: 1 },
				cx: 320000,
				cy: 240000,
			},
			name: 'Logo',
			description: 'Brand logo',
		})

		const written = writeXlsx(wb)
		expectOk(written)
		const entries = unzipSync(written.value)
		expect(entries['xl/drawings/drawing1.xml']).toBeDefined()
		expect(entries['xl/drawings/_rels/drawing1.xml.rels']).toBeDefined()
		expect(entries['xl/media/image1.png']).toBeDefined()

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.drawingRefs.hasDrawing).toBe(true)
		expect(reopened.value.workbook.sheets[0]?.imageRefs).toHaveLength(1)
		expect(reopened.value.workbook.sheets[0]?.imageRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			targetPath: 'xl/media/image1.png',
			name: 'Logo',
			description: 'Brand logo',
			anchor: {
				kind: 'oneCell',
			},
		})
	})

	it('overrides replaced image media while preserving existing drawing capsules', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Images')
		sheet.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImage1',
			targetPath: 'xl/media/image1.png',
			contentType: 'image/png',
			content: new Uint8Array([1, 2, 3]),
			anchor: {
				kind: 'oneCell',
				from: { row: 1, col: 1 },
				cx: 320000,
				cy: 240000,
			},
			name: 'Logo',
		})
		const source = writeXlsx(wb)
		expectOk(source)
		const read = readXlsx(source.value)
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		const image = readSheet?.imageRefs[0]
		expect(image).toBeDefined()
		if (!readSheet || !image) return
		readSheet.imageRefs[0] = {
			...image,
			contentType: 'image/png',
			content: new Uint8Array([4, 5, 6]),
		}

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Images'],
		})
		expectOk(written)
		const entries = unzipSync(written.value)
		expect(Array.from(entries['xl/media/image1.png'] ?? [])).toEqual([4, 5, 6])

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.imageRefs[0]).toMatchObject({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			targetPath: 'xl/media/image1.png',
			name: 'Logo',
			anchor: { kind: 'oneCell' },
		})
	})

	it('keeps sheet-owned drawing capsules attached after renaming a sheet', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Visuals')
		sheet.cells.set(0, 0, { value: stringValue('Chart host'), formula: null, styleId: S0 })
		sheet.drawingRefs = { hasDrawing: true, hasLegacyDrawing: true }

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/drawings/drawing1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
				relationships: [],
				content: new TextEncoder().encode(
					'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
				),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
			},
			{
				partPath: 'xl/drawings/vmlDrawing1.vml',
				contentType: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
				relationships: [],
				content: new TextEncoder().encode('<xml xmlns:v="urn:schemas-microsoft-com:vml"/>'),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
			},
		]

		sheet.name = 'Dashboard'
		const { result, bytes } = roundTrip(wb, capsules)
		expect(result.workbook.sheets[0]?.name).toBe('Dashboard')
		expect(result.workbook.sheets[0]?.drawingRefs).toEqual({
			hasDrawing: true,
			hasLegacyDrawing: true,
		})
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			drawing: 1,
			legacyDrawing: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 2,
		})
	})

	it('writes classic comments with generated comments and VML parts', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Notes')
		sheet.cells.set(1, 1, { value: stringValue('Cell'), formula: null, styleId: S0 })
		sheet.comments.set('B2', { text: 'Hello', author: 'Ada' })

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.comments.get('B2')).toMatchObject({
			text: 'Hello',
			author: 'Ada',
		})
		expect(result.workbook.sheets[0]?.drawingRefs).toEqual({
			hasDrawing: false,
			hasLegacyDrawing: true,
		})

		const entries = unzipSync(bytes)
		expect(entries['xl/comments1.xml']).toBeDefined()
		expect(entries['xl/drawings/vmlDrawing1.vml']).toBeDefined()
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			legacyDrawing: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 2,
		})
	})

	it.skipIf(!existsSync(POI_COMMENTS_FIXTURE))(
		'preserves source comment VML layout for text-only comment edits',
		() => {
			const source = readFileSync(POI_COMMENTS_FIXTURE)
			const opened = readXlsx(source)
			expectOk(opened)
			const sheet = opened.value.workbook.sheets[0]
			expect(sheet).toBeDefined()
			if (!sheet) return
			const original = sheet.comments.get('B1')
			expect(original).toBeDefined()
			if (!original) return
			sheet.comments.set('B1', { ...original, text: 'Updated comment text' })

			const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
				dirtySheetNames: [sheet.name],
			})
			expectOk(written)
			const sourceEntries = unzipSync(source)
			const writtenEntries = unzipSync(written.value)
			expect(new TextDecoder().decode(writtenEntries['xl/comments1.xml'])).toContain(
				'Updated comment text',
			)
			expect(writtenEntries['xl/drawings/vmlDrawing1.vml']).toEqual(
				sourceEntries['xl/drawings/vmlDrawing1.vml'],
			)
		},
	)

	it('patches legacy comment text without rewriting untouched comment payloads', () => {
		const source = commentsAndThreadedCommentsWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)
		const sheet = opened.value.workbook.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		const original = sheet.comments.get('B1')
		expect(original).toBeDefined()
		if (!original) return
		sheet.comments.set('B1', { ...original, text: 'Updated legacy text' })

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: [sheet.name],
		})
		expectOk(written)
		const sourceEntries = unzipSync(source)
		const writtenEntries = unzipSync(written.value)
		const commentsXml = new TextDecoder().decode(writtenEntries['xl/comments1.xml'])

		expect(commentsXml).toContain('Updated legacy text')
		expect(commentsXml).toContain('xr:uid="{comment-c3}"')
		expect(commentsXml).toContain('<rPr><b/></rPr><t>LegacyHidden</t>')
		expect(commentsXml).toContain('<phoneticPr fontId="1"/>')
		expect(writtenEntries['xl/drawings/vmlDrawing1.vml']).toEqual(
			sourceEntries['xl/drawings/vmlDrawing1.vml'],
		)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.getSheet('Sheet1')?.comments.get('C3')).toMatchObject({
			text: 'LegacyHiddenNote',
			author: 'Grace',
		})
	})

	it('preserves comments, VML, threaded comments, persons, and package graph through dirty cell edits', () => {
		const source = commentsAndThreadedCommentsWorkbook()
		const sourceEntries = unzipSync(source)
		const sourceGraph = inspectXlsxPackageGraph(source)
		expect(auditXlsxPackageGraphReadIntegrity(sourceGraph)).toEqual([])

		const opened = readXlsx(source)
		expectOk(opened)
		const sheet = opened.value.workbook.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		expect(sheet.comments.get('B1')).toMatchObject({
			text: 'Legacy visible note',
			author: 'Ada',
			legacyDrawing: {
				shapeId: '_x0000_s2048',
				anchor: [1, 15, 0, 2, 3, 20, 4, 8],
				visible: true,
			},
		})
		expect(sheet.comments.get('C3')).toMatchObject({
			text: 'LegacyHiddenNote',
			author: 'Grace',
			legacyDrawing: {
				shapeId: '_x0000_s2049',
				anchor: [2, 10, 2, 4, 5, 30, 6, 12],
				visible: false,
			},
		})
		expect(sheet.threadedComments).toEqual([
			{
				ref: 'D4',
				text: 'Thread root',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				id: '{root-thread}',
				personId: '{person-ada}',
				author: 'Ada Thread',
				dateTime: '2024-03-01T10:11:12.000Z',
			},
			{
				ref: 'D4',
				text: 'Thread reply',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				id: '{reply-thread}',
				parentId: '{root-thread}',
				personId: '{person-grace}',
				author: 'Grace Thread',
				dateTime: '2024-03-02T10:11:12.000Z',
				done: true,
			},
		])

		const applied = applyOperations(opened.value.workbook, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'edited' }] },
		])
		expectOk(applied)
		expect(applied.value.sheetsModified).toEqual(['Sheet1'])

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)

		const writtenEntries = unzipSync(written.value)
		expect(writtenEntries['xl/comments1.xml']).toEqual(sourceEntries['xl/comments1.xml'])
		expect(writtenEntries['xl/drawings/vmlDrawing1.vml']).toEqual(
			sourceEntries['xl/drawings/vmlDrawing1.vml'],
		)
		expect(writtenEntries['xl/threadedComments/threadedComment1.xml']).toEqual(
			sourceEntries['xl/threadedComments/threadedComment1.xml'],
		)
		expect(writtenEntries['xl/persons/person.xml']).toEqual(sourceEntries['xl/persons/person.xml'])

		const worksheetXml = decodeTestXml(writtenEntries['xl/worksheets/sheet1.xml'])
		expect(worksheetXml).toContain('r:id="rIdVml"')
		expect(worksheetXml).toContain('<t>edited</t>')
		const worksheetRelsXml = writtenEntries['xl/worksheets/_rels/sheet1.xml.rels']
			? decodeTestXml(writtenEntries['xl/worksheets/_rels/sheet1.xml.rels'])
			: ''
		expect(worksheetRelsXml).toContain('Id="rIdComments"')
		expect(worksheetRelsXml).toContain(
			'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"',
		)
		expect(worksheetRelsXml).toContain('Target="../comments1.xml"')
		expect(worksheetRelsXml).toContain('Id="rIdVml"')
		expect(worksheetRelsXml).toContain(
			'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"',
		)
		expect(worksheetRelsXml).toContain('Target="../drawings/vmlDrawing1.vml"')
		expect(worksheetRelsXml).toContain('Id="rIdThreaded"')
		expect(worksheetRelsXml).toContain(
			'Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"',
		)
		expect(worksheetRelsXml).toContain('Target="../threadedComments/threadedComment1.xml"')
		const contentTypesXml = decodeTestXml(writtenEntries['[Content_Types].xml'])
		expect(contentTypesXml).toContain(
			'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"',
		)
		expect(contentTypesXml).toContain('ContentType="application/vnd.ms-excel.threadedcomments+xml"')
		expect(contentTypesXml).toContain('ContentType="application/vnd.ms-excel.person+xml"')

		const writtenGraph = inspectXlsxPackageGraph(written.value)
		expect(auditXlsxPackageGraphReadIntegrity(writtenGraph)).toEqual([])
		expect(auditXlsxPackageGraphSafeEditIntegrity(sourceGraph, writtenGraph)).toEqual([])
		expect(auditXlsxPackageGraphBytePreservation(sourceGraph, source, written.value)).toEqual([])

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedSheet = reopened.value.workbook.getSheet('Sheet1')
		expect(reopenedSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'edited' })
		expect(reopenedSheet?.comments.get('B1')).toMatchObject({
			text: 'Legacy visible note',
			author: 'Ada',
			legacyDrawing: {
				shapeId: '_x0000_s2048',
				anchor: [1, 15, 0, 2, 3, 20, 4, 8],
				visible: true,
			},
		})
		expect(reopenedSheet?.comments.get('C3')).toMatchObject({
			text: 'LegacyHiddenNote',
			author: 'Grace',
			legacyDrawing: {
				shapeId: '_x0000_s2049',
				anchor: [2, 10, 2, 4, 5, 30, 6, 12],
				visible: false,
			},
		})
		expect(reopenedSheet?.threadedComments).toEqual(sheet.threadedComments)
	})

	it('drops stale comment sidecars after deleting all legacy and threaded comments', () => {
		const source = commentsAndThreadedCommentsWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)
		const applied = applyOperations(opened.value.workbook, [
			{ op: 'deleteComment', sheet: 'Sheet1', ref: 'B1' },
			{ op: 'deleteComment', sheet: 'Sheet1', ref: 'C3' },
			{ op: 'deleteComment', sheet: 'Sheet1', ref: 'D4' },
		])
		expectOk(applied)
		expect(applied.value.sheetsModified).toEqual(['Sheet1'])

		const sheet = opened.value.workbook.getSheet('Sheet1')
		expect(sheet?.comments.size).toBe(0)
		expect(sheet?.threadedComments).toEqual([])

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)
		const writtenEntries = unzipSync(written.value)
		expect(writtenEntries['xl/comments1.xml']).toBeUndefined()
		expect(writtenEntries['xl/drawings/vmlDrawing1.vml']).toBeUndefined()
		expect(writtenEntries['xl/threadedComments/threadedComment1.xml']).toBeUndefined()
		expect(writtenEntries['xl/persons/person.xml']).toBeUndefined()

		const worksheetXml = decodeTestXml(writtenEntries['xl/worksheets/sheet1.xml'])
		expect(worksheetXml).not.toContain('<legacyDrawing')
		const worksheetRelsXml = writtenEntries['xl/worksheets/_rels/sheet1.xml.rels']
			? decodeTestXml(writtenEntries['xl/worksheets/_rels/sheet1.xml.rels'])
			: ''
		expect(worksheetRelsXml).not.toContain('/relationships/comments')
		expect(worksheetRelsXml).not.toContain('/relationships/vmlDrawing')
		expect(worksheetRelsXml).not.toContain('/relationships/threadedComment')
		const contentTypesXml = decodeTestXml(writtenEntries['[Content_Types].xml'])
		expect(contentTypesXml).not.toContain('spreadsheetml.comments+xml')
		expect(contentTypesXml).not.toContain('vnd.ms-excel.threadedcomments+xml')
		expect(contentTypesXml).not.toContain('vnd.ms-excel.person+xml')

		const writtenGraph = inspectXlsxPackageGraph(written.value)
		expect(auditXlsxPackageGraphReadIntegrity(writtenGraph)).toEqual([])
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedSheet = reopened.value.workbook.getSheet('Sheet1')
		expect(reopenedSheet?.comments.size).toBe(0)
		expect(reopenedSheet?.threadedComments).toEqual([])
	})

	it('strips deleted legacy comment shapes from mixed VML drawings without dropping controls', () => {
		const source = commentsAndMixedVmlWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)
		const applied = applyOperations(opened.value.workbook, [
			{ op: 'deleteComment', sheet: 'Sheet1', ref: 'B1' },
			{ op: 'deleteComment', sheet: 'Sheet1', ref: 'C3' },
		])
		expectOk(applied)

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)
		const writtenEntries = unzipSync(written.value)
		expect(writtenEntries['xl/comments1.xml']).toBeUndefined()
		expect(writtenEntries['xl/drawings/vmlDrawing1.vml']).toBeDefined()

		const vmlXml = decodeTestXml(writtenEntries['xl/drawings/vmlDrawing1.vml'])
		expect(vmlXml).not.toContain('_x0000_s2048')
		expect(vmlXml).not.toContain('_x0000_s2049')
		expect(vmlXml).toContain('_x0000_sButton')
		expect(vmlXml).toContain('ObjectType="Button"')
		expect(vmlXml).toContain('_x0000_sUnknown')
		expect(vmlXml).toContain('preserve unknown VML shape')

		const worksheetXml = decodeTestXml(writtenEntries['xl/worksheets/sheet1.xml'])
		expect(worksheetXml).toContain('<legacyDrawing')
		const worksheetRelsXml = decodeTestXml(writtenEntries['xl/worksheets/_rels/sheet1.xml.rels'])
		expect(worksheetRelsXml).not.toContain('/relationships/comments')
		expect(worksheetRelsXml).toContain('/relationships/vmlDrawing')
		expect(worksheetRelsXml).toContain('/relationships/threadedComment')
		const contentTypesXml = decodeTestXml(writtenEntries['[Content_Types].xml'])
		expect(contentTypesXml).not.toContain('spreadsheetml.comments+xml')
		expect(contentTypesXml).toContain('vnd.ms-excel.threadedcomments+xml')
		expect(contentTypesXml).toContain('vnd.ms-excel.person+xml')

		const writtenGraph = inspectXlsxPackageGraph(written.value)
		expect(auditXlsxPackageGraphReadIntegrity(writtenGraph)).toEqual([])
		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedSheet = reopened.value.workbook.getSheet('Sheet1')
		expect(reopenedSheet?.comments.size).toBe(0)
		expect(reopenedSheet?.threadedComments).toEqual(
			opened.value.workbook.getSheet('Sheet1')?.threadedComments,
		)
		expect(reopenedSheet?.drawingObjectRefs).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ source: 'vml', vmlObjectType: 'Button' }),
				expect.objectContaining({
					source: 'vml',
					vmlShapeId: '_x0000_sUnknown',
					text: 'preserve unknown VML shape',
				}),
			]),
		)
	})

	it('regenerates legacy comment VML from shifted layout metadata', () => {
		const source = commentsAndThreadedCommentsWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)

		const applied = applyOperations(opened.value.workbook, [
			{ op: 'insertRows', sheet: 'Sheet1', at: 0, count: 1 },
			{ op: 'insertCols', sheet: 'Sheet1', at: 0, count: 1 },
		])
		expectOk(applied)
		expect(applied.value.sheetsModified).toEqual(['Sheet1'])

		const sheet = opened.value.workbook.getSheet('Sheet1')
		expect(sheet?.comments.get('C2')?.legacyDrawing).toMatchObject({
			shapeId: '_x0000_s2048',
			anchor: [2, 15, 1, 2, 4, 20, 5, 8],
			visible: true,
		})
		expect(sheet?.comments.get('D4')?.legacyDrawing).toMatchObject({
			shapeId: '_x0000_s2049',
			anchor: [3, 10, 3, 4, 6, 30, 7, 12],
			visible: false,
		})

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)
		const writtenEntries = unzipSync(written.value)
		const vmlXml = decodeTestXml(writtenEntries['xl/drawings/vmlDrawing1.vml'])
		expect(vmlXml).toContain('id="_x0000_s2048"')
		expect(vmlXml).toContain('<x:Anchor>2, 15, 1, 2, 4, 20, 5, 8</x:Anchor>')
		expect(vmlXml).toContain('<x:Row>1</x:Row>')
		expect(vmlXml).toContain('<x:Column>2</x:Column>')
		expect(vmlXml).toContain('<x:Visible/>')
		expect(vmlXml).toContain('id="_x0000_s2049"')
		expect(vmlXml).toContain('<x:Anchor>3, 10, 3, 4, 6, 30, 7, 12</x:Anchor>')
		expect(vmlXml).toContain('<x:Row>3</x:Row>')
		expect(vmlXml).toContain('<x:Column>3</x:Column>')
		expect(vmlXml).toContain('<x:Visible>false</x:Visible>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.getSheet('Sheet1')?.comments.get('C2')).toMatchObject({
			text: 'Legacy visible note',
			legacyDrawing: {
				shapeId: '_x0000_s2048',
				anchor: [2, 15, 1, 2, 4, 20, 5, 8],
				visible: true,
			},
		})
		expect(reopened.value.workbook.getSheet('Sheet1')?.comments.get('D4')).toMatchObject({
			text: 'LegacyHiddenNote',
			legacyDrawing: {
				shapeId: '_x0000_s2049',
				anchor: [3, 10, 3, 4, 6, 30, 7, 12],
				visible: false,
			},
		})
	})

	it('deduplicates legacy comment VML shape ids after copying notes', () => {
		const source = commentsAndThreadedCommentsWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)
		const applied = applyOperations(opened.value.workbook, [
			{ op: 'copyRange', sheet: 'Sheet1', source: 'B1', target: 'E5', mode: 'comments' },
		])
		expectOk(applied)
		expect(applied.value.sheetsModified).toEqual(['Sheet1'])

		const sheet = opened.value.workbook.getSheet('Sheet1')
		expect(sheet?.comments.get('E5')?.legacyDrawing).toMatchObject({
			shapeId: '_x0000_s2048',
			anchor: [4, 15, 4, 2, 6, 20, 8, 8],
			row: 4,
			column: 4,
		})

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)
		const entries = unzipSync(written.value)
		const commentsXml = decodeTestXml(entries['xl/comments1.xml'])
		expect(commentsXml).toContain('xr:uid="{comment-c3}"')
		expect(commentsXml).toContain('<rPr><b/></rPr><t>LegacyHidden</t>')
		expect(commentsXml).toContain('<phoneticPr fontId="1"/>')
		expect(commentsXml).toContain('<comment ref="E5" authorId="0">')
		const vmlXml = decodeTestXml(entries['xl/drawings/vmlDrawing1.vml'])
		const noteShapeIds = [
			...vmlXml.matchAll(/<v:shape\b[^>]*\bid="([^"]+)"[\s\S]*?<x:ClientData ObjectType="Note">/g),
		].map((match) => match[1])
		expect(noteShapeIds).toHaveLength(3)
		expect(new Set(noteShapeIds).size).toBe(3)
		expect(noteShapeIds).toEqual(['_x0000_s2048', '_x0000_s2049', '_x0000_s2050'])
		expect(vmlXml).toContain('<x:Anchor>4, 15, 4, 2, 6, 20, 8, 8</x:Anchor>')
		expect(vmlXml).toContain('<x:Row>4</x:Row>')
		expect(vmlXml).toContain('<x:Column>4</x:Column>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedSheet = reopened.value.workbook.getSheet('Sheet1')
		expect(reopenedSheet?.comments.get('B1')).toBeDefined()
		expect(reopenedSheet?.comments.get('C3')).toBeDefined()
		expect(reopenedSheet?.comments.get('E5')).toMatchObject({
			text: 'Legacy visible note',
			author: 'Ada',
			legacyDrawing: {
				shapeId: '_x0000_s2050',
				anchor: [4, 15, 4, 2, 6, 20, 8, 8],
				row: 4,
				column: 4,
			},
		})
	})

	it('preserves untouched rich legacy comment XML when deleting another note', () => {
		const source = commentsAndThreadedCommentsWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)
		const applied = applyOperations(opened.value.workbook, [
			{ op: 'deleteComment', sheet: 'Sheet1', ref: 'B1' },
		])
		expectOk(applied)
		expect(applied.value.sheetsModified).toEqual(['Sheet1'])

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)
		const entries = unzipSync(written.value)
		const commentsXml = decodeTestXml(entries['xl/comments1.xml'])
		expect(commentsXml).not.toContain('ref="B1"')
		expect(commentsXml).toContain('ref="C3"')
		expect(commentsXml).toContain('xr:uid="{comment-c3}"')
		expect(commentsXml).toContain('<rPr><b/></rPr><t>LegacyHidden</t>')
		expect(commentsXml).toContain('<phoneticPr fontId="1"/>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const reopenedSheet = reopened.value.workbook.getSheet('Sheet1')
		expect(reopenedSheet?.comments.has('B1')).toBe(false)
		expect(reopenedSheet?.comments.get('C3')).toMatchObject({
			text: 'LegacyHiddenNote',
			author: 'Grace',
		})
	})

	it('preserves conditional formatting and data validations on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Rules')
		sheet.cells.set(0, 0, { value: numberValue(4), formula: null, styleId: S0 })
		wb.differentialStyles.push({
			font: { bold: true },
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
		})
		sheet.conditionalFormats.push({
			sqref: 'A1:A10',
			pivot: true,
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					dxfId: 0,
					priority: 1,
					stopIfTrue: true,
					formulas: ['3'],
					style: wb.differentialStyles[0],
				},
				{
					type: 'aboveAverage',
					priority: 2,
					formulas: [],
					aboveAverage: false,
					stdDev: 2,
				},
				{
					type: 'containsText',
					operator: 'containsText',
					priority: 3,
					formulas: [],
					text: 'Grain',
				},
			],
		})
		sheet.dataValidations.push({
			sqref: 'B2:B4',
			uid: '{CAFD7DE3-F94F-4BD6-B5E6-7E794FD6EC31}',
			type: 'list',
			allowBlank: true,
			showInputMessage: true,
			formula1: '"Q1,Q2,Q3"',
		})
		sheet.dataValidationSettings = {
			disablePrompts: true,
			xWindow: 220,
			yWindow: 120,
		}

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.conditionalFormats).toEqual([
			{
				sqref: 'A1:A10',
				pivot: true,
				rules: [
					{
						type: 'cellIs',
						operator: 'greaterThan',
						dxfId: 0,
						priority: 1,
						stopIfTrue: true,
						formulas: ['3'],
						style: {
							font: { bold: true },
							fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
						},
					},
					{
						type: 'aboveAverage',
						priority: 2,
						formulas: [],
						aboveAverage: false,
						stdDev: 2,
					},
					{
						type: 'containsText',
						operator: 'containsText',
						priority: 3,
						formulas: [],
						text: 'Grain',
					},
				],
			},
		])
		expect(result.workbook.sheets[0]?.dataValidations).toEqual([
			{
				sqref: 'B2:B4',
				uid: '{CAFD7DE3-F94F-4BD6-B5E6-7E794FD6EC31}',
				type: 'list',
				allowBlank: true,
				showInputMessage: true,
				formula1: '"Q1,Q2,Q3"',
			},
		])
		expect(result.workbook.sheets[0]?.dataValidationSettings).toEqual({
			disablePrompts: true,
			xWindow: 220,
			yWindow: 120,
		})
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.styles?.tagCounts).toMatchObject({
			dxfs: 1,
			dxf: 1,
		})
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			conditionalFormatting: 1,
			cfRule: 3,
			dataValidations: 1,
			dataValidation: 1,
			formula: 1,
			formula1: 1,
		})
		const sheetXml = new TextDecoder().decode(
			unzipSync(bytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array(),
		)
		expect(sheetXml).toContain('stdDev="2"')
		expect(sheetXml).toContain('text="Grain"')
		expect(sheetXml).toContain('pivot="1"')
		expect(sheetXml).toContain('disablePrompts="1"')
		expect(sheetXml).toContain('xWindow="220"')
		expect(sheetXml).toContain('yWindow="120"')
		expect(sheetXml).toContain(
			'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"',
		)
		expect(sheetXml).toContain('xr:uid="{CAFD7DE3-F94F-4BD6-B5E6-7E794FD6EC31}"')
	})

	it('preserves rich text inline strings when useSharedStrings is false (xlsx-2)', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, {
			value: {
				kind: 'richText',
				runs: [{ text: 'bold text', bold: true }, { text: ' normal' }],
			},
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(wb, undefined, { useSharedStrings: false })
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<r><rPr><b/></rPr><t>bold text</t></r>')
		expect(sheetXml).toContain('<r><t> normal</t></r>')
		expect(sheetXml).toContain('t="inlineStr"')
		expect(sheetXml).toContain('<is>')

		const read = readXlsx(written.value)
		expectOk(read)
		expect(read.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'richText',
			runs: [{ text: 'bold text', bold: true }, { text: ' normal' }],
		})
	})

	it('writes inline strings with useInlineStrings and round-trips correctly', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		for (let r = 0; r < 50; r++) {
			for (let c = 0; c < 10; c++) {
				sheet.cells.set(r, c, {
					value: stringValue(`unique-${r}-${c}`),
					formula: null,
					styleId: S0,
				})
			}
		}

		const written = writeXlsx(wb, undefined, { useInlineStrings: true })
		expectOk(written)

		const zip = unzipSync(written.value)
		expect(zip['xl/sharedStrings.xml']).toBeUndefined()

		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('t="inlineStr"')
		expect(sheetXml).toContain('<is><t>unique-0-0</t></is>')
		expect(sheetXml).toContain('<is><t>unique-25-5</t></is>')

		const read = readXlsx(written.value)
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'unique-0-0',
		})
		expect(readSheet?.cells.get(25, 5)?.value).toEqual({
			kind: 'string',
			value: 'unique-25-5',
		})
	})

	it('writes plain string cells with usePlainStrings and round-trips correctly', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, {
			value: stringValue('plain & fast'),
			formula: null,
			styleId: S0,
		})

		const written = writeXlsx(wb, undefined, { usePlainStrings: true })
		expectOk(written)

		const zip = unzipSync(written.value)
		expect(zip['xl/sharedStrings.xml']).toBeUndefined()

		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('t="str"')
		expect(sheetXml).toContain('<v>plain &amp; fast</v>')

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		expect(read.value.workbook.sheets[0]?.cells.get(0, 0)?.value).toEqual({
			kind: 'string',
			value: 'plain & fast',
		})
	})

	it('omits dense cell refs only when positions can be inferred', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('b'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: stringValue('sparse'), formula: null, styleId: S0 })

		const written = writeXlsx(wb, undefined, {
			useSharedStrings: false,
			usePlainStrings: true,
			omitDenseCellRefs: true,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<row r="1"><c t="str"><v>a</v></c><c t="str"><v>b</v></c></row>')
		expect(sheetXml).toContain('<row r="2"><c r="B2" t="str"><v>sparse</v></c></row>')

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'a' })
		expect(readSheet?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: 'b' })
		expect(readSheet?.cells.get(1, 1)?.value).toEqual({ kind: 'string', value: 'sparse' })
	})

	it('omits dense shared-string refs when positions can be inferred', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('b'), formula: null, styleId: S0 })

		const written = writeXlsx(wb, undefined, {
			useSharedStrings: true,
			omitDenseCellRefs: true,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = decodeTestXml(zip['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain('<row r="1"><c t="s"><v>0</v></c><c t="s"><v>1</v></c></row>')

		const sharedStringsXml = decodeTestXml(zip['xl/sharedStrings.xml'])
		expect(sharedStringsXml).toContain('count="2" uniqueCount="2"')
		expect(sharedStringsXml).toContain('<si><t>a</t></si>')
		expect(sharedStringsXml).toContain('<si><t>b</t></si>')

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'a' })
		expect(readSheet?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: 'b' })
	})

	it('omits dense shared-string refs when a numeric cell starts the row', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('b'), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: numberValue(3), formula: null, styleId: S0 })

		const written = writeXlsx(wb, undefined, {
			useSharedStrings: true,
			omitDenseCellRefs: true,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = decodeTestXml(zip['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain(
			'<row r="1"><c><v>1</v></c><c t="s"><v>0</v></c><c><v>3</v></c></row>',
		)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 1 })
		expect(readSheet?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: 'b' })
		expect(readSheet?.cells.get(0, 2)?.value).toEqual({ kind: 'number', value: 3 })
	})

	it('keeps dense cell refs when formula state follows scalar cells', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: 'A1+1', styleId: S0 })

		const written = writeXlsx(wb, undefined, { omitDenseCellRefs: true })
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain(
			'<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1+1</f><v>2</v></c></row>',
		)
	})

	it('does not count shared strings during dense-ref fallback preflight', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Test')
		sheet.cells.set(0, 0, { value: stringValue('a'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: 'LEN(A1)', styleId: S0 })

		const written = writeXlsx(wb, undefined, {
			useSharedStrings: true,
			omitDenseCellRefs: true,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = decodeTestXml(zip['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain(
			'<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>LEN(A1)</f><v>2</v></c></row>',
		)

		const sharedStringsXml = decodeTestXml(zip['xl/sharedStrings.xml'])
		expect(sharedStringsXml).toContain('count="1" uniqueCount="1"')
		expect(sharedStringsXml.match(/<si>/g)).toHaveLength(1)
	})

	it('writes dense rows without materializing a workbook', () => {
		const written = writeDenseRowsXlsx({
			rows: 2,
			cols: 3,
			omitCellRefs: true,
			allCellsPresent: true,
			valueAt: (row, col) => (row === 1 && col === 1 ? null : `r${row}c${col}`),
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<row r="1"><c t="str"><v>r0c0</v></c>')
		expect(sheetXml).toContain(
			'<row r="2"><c r="A2" t="str"><v>r1c0</v></c><c r="C2" t="str"><v>r1c2</v></c></row>',
		)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'r0c0' })
		expect(readSheet?.cells.get(1, 1)).toBeUndefined()
		expect(readSheet?.cells.get(1, 2)?.value).toEqual({ kind: 'string', value: 'r1c2' })
	})

	it('can omit dense row refs for sequential generated exports', () => {
		const written = writeDenseRowsXlsx({
			rows: 3,
			cols: 2,
			omitCellRefs: true,
			omitRowRefs: true,
			allCellsPresent: true,
			valueType: 'number',
			valueAt: (row, col) => (row === 1 ? null : row * 10 + col),
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<row><c><v>0</v></c><c><v>1</v></c></row>')
		expect(sheetXml).toContain('<row></row>')
		expect(sheetXml).not.toContain('<row r=')

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 0 })
		expect(readSheet?.cells.get(1, 0)).toBeUndefined()
		expect(readSheet?.cells.get(2, 1)?.value).toEqual({ kind: 'number', value: 21 })
	})

	it('applies omitted dense row refs to boolean and repeated row fast paths', () => {
		const booleanRows = writeDenseRowsXlsx({
			rows: 1,
			cols: 2,
			omitCellRefs: true,
			omitRowRefs: true,
			allCellsPresent: true,
			valueType: 'boolean',
			valueAt: (_row, col) => col === 0,
		})
		expectOk(booleanRows)
		const repeatedRows = writeDenseRowsXlsx({
			rows: 2,
			cols: 2,
			omitCellRefs: true,
			omitRowRefs: true,
			cacheRepeatedRows: true,
			valueAt: (_row, col) => col,
		})
		expectOk(repeatedRows)

		for (const bytes of [booleanRows.value, repeatedRows.value]) {
			const zip = unzipSync(bytes)
			const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
			expect(sheetXml).toContain('<row>')
			expect(sheetXml).not.toContain('<row r=')
		}
	})

	it('reuses constant dense row bodies when explicitly requested', () => {
		let calls = 0
		const written = writeDenseRowsXlsx({
			rows: 4,
			cols: 3,
			omitCellRefs: true,
			constantRows: true,
			valueAt: (_row, col) => {
				calls++
				return col < 2 ? 'repeat' : col
			},
		})
		expectOk(written)
		expect(calls).toBe(3)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'repeat' })
		expect(readSheet?.cells.get(3, 0)?.value).toEqual({ kind: 'string', value: 'repeat' })
		expect(readSheet?.cells.get(3, 2)?.value).toEqual({ kind: 'number', value: 2 })
	})

	it('can skip string escaping for generated XML-safe dense values', () => {
		const written = writeDenseRowsXlsx({
			rows: 1,
			cols: 2,
			omitCellRefs: true,
			stringsAreXmlSafe: true,
			valueAt: (_row, col) => `safe-${col}`,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<c t="str"><v>safe-0</v></c>')

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		expect(read.value.workbook.sheets[0]?.cells.get(0, 1)?.value).toEqual({
			kind: 'string',
			value: 'safe-1',
		})
	})

	it('uses dense value type hints while preserving fallback semantics', () => {
		const written = writeDenseRowsXlsx({
			rows: 1,
			cols: 3,
			omitCellRefs: true,
			valueType: 'number',
			valueAt: (_row, col) => (col === 2 ? 'fallback' : col + 1),
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain(
			'<row r="1"><c><v>1</v></c><c><v>2</v></c><c t="str"><v>fallback</v></c></row>',
		)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 1 })
		expect(readSheet?.cells.get(0, 2)?.value).toEqual({ kind: 'string', value: 'fallback' })
	})

	it('uses dense column value type hints while preserving fallback semantics', () => {
		const written = writeDenseRowsXlsx({
			rows: 2,
			cols: 4,
			omitCellRefs: true,
			allCellsPresent: true,
			stringsAreXmlSafe: true,
			valueTypes: ['string', 'number', 'string', 'number'],
			valueAt: (row, col) => {
				if (row === 1 && col === 3) return 'fallback'
				return col % 2 === 0 ? `text-${row}-${col}` : row * 10 + col
			},
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain(
			'<row r="1"><c t="str"><v>text-0-0</v></c><c><v>1</v></c><c t="str"><v>text-0-2</v></c><c><v>3</v></c></row>',
		)
		expect(sheetXml).toContain(
			'<row r="2"><c t="str"><v>text-1-0</v></c><c><v>11</v></c><c t="str"><v>text-1-2</v></c><c t="str"><v>fallback</v></c></row>',
		)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'text-0-0' })
		expect(readSheet?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 1 })
		expect(readSheet?.cells.get(1, 3)?.value).toEqual({ kind: 'string', value: 'fallback' })
	})

	it('streams dense rows without materializing a workbook', async () => {
		const written = await writeDenseRowsXlsxStreaming({
			rows: 2,
			cols: 2,
			omitCellRefs: true,
			valueAt: (row, col) => row * 2 + col,
		})
		expectOk(written)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 0 })
		expect(readSheet?.cells.get(1, 1)?.value).toEqual({ kind: 'number', value: 3 })
	})

	it('keeps small dense streaming output equivalent to the buffered dense writer', async () => {
		const options = {
			rows: 64,
			cols: 8,
			omitCellRefs: true,
			omitRowRefs: true,
			allCellsPresent: true,
			stringsAreXmlSafe: true,
			valueType: 'string' as const,
			valueAt: (row: number, col: number) => `text-${row}-${col}`,
		}
		const buffered = writeDenseRowsXlsx(options)
		expectOk(buffered)
		const streaming = await writeDenseRowsXlsxStreaming(options)
		expectOk(streaming)
		expect(streaming.value).toEqual(buffered.value)

		const read = readXlsx(streaming.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'text-0-0' })
		expect(readSheet?.cells.get(63, 7)?.value).toEqual({ kind: 'string', value: 'text-63-7' })
	})

	it('supports compact dense writer compression without changing values', async () => {
		const fast = await writeDenseRowsXlsxStreaming({
			rows: 200,
			cols: 20,
			omitCellRefs: true,
			allCellsPresent: true,
			stringsAreXmlSafe: true,
			valueType: 'string',
			valueAt: (row, col) => `text-${row}-${col}`,
		})
		expectOk(fast)
		const written = await writeDenseRowsXlsxStreaming({
			rows: 200,
			cols: 20,
			omitCellRefs: true,
			allCellsPresent: true,
			stringsAreXmlSafe: true,
			valueType: 'string',
			compressionProfile: 'compact',
			valueAt: (row, col) => `text-${row}-${col}`,
		})
		expectOk(written)
		expect(written.value.byteLength).toBeLessThan(fast.value.byteLength)

		const read = readXlsx(written.value, { mode: 'values' })
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'string', value: 'text-0-0' })
		expect(readSheet?.cells.get(199, 19)?.value).toEqual({
			kind: 'string',
			value: 'text-199-19',
		})
	})

	it('creates dxfId for CF rules with style but no dxfId (xlsx-4)', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Rules')
		sheet.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: S0 })
		sheet.conditionalFormats.push({
			sqref: 'A1:A5',
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					priority: 1,
					formulas: ['3'],
					style: {
						font: { bold: true },
						fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
					},
				},
			],
		})

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		const stylesXml = new TextDecoder().decode(zip['xl/styles.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('dxfId="0"')
		expect(stylesXml).toContain('<dxfs count="1">')
		expect(stylesXml).toContain('<dxf>')
		expect(stylesXml).toContain('<b/>')
		expect(stylesXml).toContain('FFC6EFCE')

		const { result } = roundTrip(wb)
		const rule = result.workbook.sheets[0]?.conditionalFormats[0]?.rules[0]
		expect(rule?.dxfId).toBeDefined()
		expect(rule?.dxfId).toBeGreaterThanOrEqual(0)
		expect(rule?.style).toEqual({
			font: { bold: true },
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
		})
	})

	it('preserves regular conditional formatting extension attributes and child XML', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Rules')
		sheet.conditionalFormats.push({
			sqref: 'A1:A5',
			rules: [
				{
					type: 'dataBar',
					priority: 1,
					formulas: [],
					preservedRuleAttributes: {
						activePresent: '1',
						'xr:uid': '{REGULAR-CF-UID}',
					},
					preservedRuleChildXml: [
						'<extLst><ext uri="{regular-cf-extension}"><x14ac:metadata flag="1"/></ext></extLst>',
					],
					dataBar: {
						cfvo: [{ type: 'min' }, { type: 'max' }],
						color: { rgb: 'FF638EC6' },
					},
				},
			],
		})

		const { bytes, result } = roundTrip(wb)
		const zip = unzipSync(bytes)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

		expect(sheetXml).toContain(
			'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"',
		)
		expect(sheetXml).toContain(
			'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"',
		)
		expect(sheetXml).toContain('activePresent="1"')
		expect(sheetXml).toContain('xr:uid="{REGULAR-CF-UID}"')
		expect(sheetXml).toContain(
			'<extLst><ext uri="{regular-cf-extension}"><x14ac:metadata flag="1"/></ext></extLst>',
		)
		expect(result.workbook.sheets[0]?.conditionalFormats[0]?.rules[0]).toMatchObject({
			type: 'dataBar',
			priority: 1,
			preservedRuleAttributes: {
				activePresent: '1',
				'xr:uid': '{REGULAR-CF-UID}',
			},
			preservedRuleChildXml: [
				'<extLst><ext uri="{regular-cf-extension}"><x14ac:metadata flag="1"/></ext></extLst>',
			],
		})
	})

	it('preserves defined names on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.definedNames.set('Total', 'Data!$A$1')

		const { result } = roundTrip(wb)
		expect(result.workbook.definedNames.get('Total')).toBe('Data!$A$1')
	})

	it('preserves sheet-scoped defined names on round-trip', () => {
		const wb = new Workbook()
		const data = wb.addSheet('Data')
		const summary = wb.addSheet('Summary')
		data.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		summary.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })
		wb.definedNames.set('Budget', 'Summary!$A$1', { kind: 'sheet', sheetId: summary.id })

		const { result } = roundTrip(wb)
		const resolved = result.workbook.definedNames.resolve(
			'Budget',
			result.workbook.getSheet('Summary')?.id,
		)
		expect(resolved?.scope.kind).toBe('sheet')
		expect(resolved?.formula).toBe('Summary!$A$1')
	})

	it('preserves hidden defined names on round-trip', () => {
		const wb = new Workbook()
		const data = wb.addSheet('Data')
		data.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.definedNames.set(
			'_xlnm._FilterDatabase',
			'Data!$A$1:$B$10',
			{ kind: 'sheet', sheetId: data.id },
			{ hidden: true },
		)

		const { bytes } = roundTrip(wb)
		const workbookXml = new TextDecoder().decode(unzipSync(bytes)['xl/workbook.xml'])
		expect(workbookXml).toContain(
			'<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$B$10</definedName>',
		)

		const result = readXlsx(bytes)
		expectOk(result)
		const resolved = result.value.workbook.definedNames.resolve(
			'_xlnm._FilterDatabase',
			result.value.workbook.getSheet('Data')?.id,
		)
		expect(resolved?.hidden).toBe(true)
		expect(resolved?.formula).toBe('Data!$A$1:$B$10')
	})

	it('preserves workbook views and external reference wiring on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.workbookFileVersion = {
			appName: 'xl',
			lastEdited: '7',
			lowestEdited: '7',
			rupBuild: '23420',
			extraAttributes: [{ name: 'productRelease', value: '2021' }],
		}
		wb.workbookFileSharing = {
			readOnlyRecommended: true,
			userName: 'Analyst',
			reservationPassword: 'ABCD',
			algorithmName: 'SHA-512',
			hashValue: 'HASH',
			saltValue: 'SALT',
			spinCount: 100000,
			extraAttributes: [{ name: 'sharingMode', value: 'review' }],
		}
		wb.workbookViews.push({
			activeTab: 1,
			firstSheet: 0,
			visibility: 'visible',
			tabRatio: 600,
			extraAttributes: [
				{ name: 'minimized', value: '1' },
				{ name: 'showSheetTabs', value: '0' },
				{ name: 'windowWidth', value: '16800' },
			],
		})
		wb.workbookProperties = {
			codeName: 'Model',
			filterPrivacy: true,
			extraAttributes: [
				{ name: 'checkCompatibility', value: '1' },
				{ name: 'autoCompressPictures', value: '0' },
			],
		}
		wb.calcSettings = {
			...wb.calcSettings,
			extraAttributes: [
				{ name: 'refMode', value: 'R1C1' },
				{ name: 'fullPrecision', value: '0' },
			],
		}
		wb.externalReferences.push('xl/externalLinks/externalLink1.xml')

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
				relationships: [],
				content: new TextEncoder().encode(
					'<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
				),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
			},
		]

		const { result, bytes } = roundTrip(wb, capsules)
		expect(result.workbook.workbookFileVersion).toEqual({
			appName: 'xl',
			lastEdited: '7',
			lowestEdited: '7',
			rupBuild: '23420',
			extraAttributes: [{ name: 'productRelease', value: '2021' }],
		})
		expect(result.workbook.workbookFileSharing).toEqual({
			readOnlyRecommended: true,
			userName: 'Analyst',
			reservationPassword: 'ABCD',
			algorithmName: 'SHA-512',
			hashValue: 'HASH',
			saltValue: 'SALT',
			spinCount: 100000,
			extraAttributes: [{ name: 'sharingMode', value: 'review' }],
		})
		expect(result.workbook.workbookProperties).toEqual({
			codeName: 'Model',
			filterPrivacy: true,
			extraAttributes: [
				{ name: 'checkCompatibility', value: '1' },
				{ name: 'autoCompressPictures', value: '0' },
			],
		})
		expect(result.workbook.workbookViews).toEqual([
			{
				activeTab: 1,
				firstSheet: 0,
				visibility: 'visible',
				tabRatio: 600,
				extraAttributes: [
					{ name: 'minimized', value: '1' },
					{ name: 'showSheetTabs', value: '0' },
					{ name: 'windowWidth', value: '16800' },
				],
			},
		])
		expect(result.workbook.calcSettings.extraAttributes).toEqual([
			{ name: 'refMode', value: 'R1C1' },
			{ name: 'fullPrecision', value: '0' },
		])
		expect(result.workbook.externalReferences).toEqual(['xl/externalLinks/externalLink1.xml'])

		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.workbook?.tagCounts).toMatchObject({
			fileVersion: 1,
			fileSharing: 1,
			bookViews: 1,
			workbookView: 1,
			externalReferences: 1,
			externalReference: 1,
			calcPr: 1,
		})
		const workbookXml = new TextDecoder().decode(unzipSync(bytes)['xl/workbook.xml'])
		expect(workbookXml).toContain('appName="xl"')
		expect(workbookXml).toContain('rupBuild="23420"')
		expect(workbookXml).toContain('productRelease="2021"')
		expect(workbookXml).toContain('readOnlyRecommended="1"')
		expect(workbookXml).toContain('userName="Analyst"')
		expect(workbookXml).toContain('reservationPassword="ABCD"')
		expect(workbookXml).toContain('sharingMode="review"')
		expect(workbookXml).toContain('checkCompatibility="1"')
		expect(workbookXml).toContain('autoCompressPictures="0"')
		expect(workbookXml).toContain('minimized="1"')
		expect(workbookXml).toContain('showSheetTabs="0"')
		expect(workbookXml).toContain('windowWidth="16800"')
		expect(workbookXml).toContain('refMode="R1C1"')
		expect(workbookXml).toContain('fullPrecision="0"')
	})

	it('preserves unsupported workbook child XML when workbook metadata is regenerated', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.preservedXml = {
			workbookXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:mx="urn:ascend:test-workbook-child"
  xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
  mc:Ignorable="xr mx"
  xr:uid="{WORKBOOK-UID}">
  <workbookPr codeName="Book1"/>
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <functionGroups><functionGroup name="Engineering"/></functionGroups>
  <customWorkbookViews><customWorkbookView name="Review" guid="{11111111-1111-1111-1111-111111111111}" activeSheetId="1"/></customWorkbookViews>
  <webPublishing css="1" thicket="0"/>
  <fileRecoveryPr autoRecover="1" crashSave="1"/>
  <mx:workbookProbe mx:flag="1"/>
  <calcPr calcMode="manual"/>
</workbook>`,
		}

		const written = writeXlsx(wb, undefined, { workbookMetaDirty: true })
		expectOk(written)

		const zip = unzipSync(written.value)
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		expect(workbookXml).toContain(
			'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
		)
		expect(workbookXml).toContain('xmlns:mx="urn:ascend:test-workbook-child"')
		expect(workbookXml).toContain(
			'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"',
		)
		expect(workbookXml).toContain('mc:Ignorable="xr mx"')
		expect(workbookXml).toContain('xr:uid="{WORKBOOK-UID}"')
		expect(workbookXml).toContain(
			'<functionGroups><functionGroup name="Engineering"/></functionGroups>',
		)
		expect(workbookXml).toContain(
			'<customWorkbookViews><customWorkbookView name="Review" guid="{11111111-1111-1111-1111-111111111111}" activeSheetId="1"/></customWorkbookViews>',
		)
		expect(workbookXml).toContain('<webPublishing css="1" thicket="0"/>')
		expect(workbookXml).toContain('<fileRecoveryPr autoRecover="1" crashSave="1"/>')
		expect(workbookXml).toContain('<mx:workbookProbe mx:flag="1"/>')
		expect(workbookXml).not.toContain('<workbookPr codeName="Book1"/>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets.map((entry) => entry.name)).toEqual(['Data'])
	})

	it('rewrites external link relationship targets while preserving link parts', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.externalReferences.push('xl/externalLinks/externalLink1.xml')
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rId2',
			linkRelId: 'rIdExt',
			target: '../sources/source.xlsx',
			targetMode: 'External',
		})

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
				relationships: [
					{
						id: 'rIdMetadata',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
						target: '../customXml/item1.xml',
					},
					{
						id: 'rIdExt',
						type: 'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup',
						target: '../sources/source.xlsx',
						targetMode: 'External',
					},
				],
				content: new TextEncoder().encode(
					'<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
				),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
			},
		]

		const applied = applyOperations(wb, [
			{
				op: 'rewriteExternalLink',
				partPath: 'xl/externalLinks/externalLink1.xml',
				newTarget: '../sources/reforecast & final.xlsx',
			},
		])
		expectOk(applied)
		const written = writeXlsx(wb, capsules, { workbookMetaDirty: true })
		expectOk(written)

		const parts = unzipSync(written.value)
		const rels = new TextDecoder().decode(
			parts['xl/externalLinks/_rels/externalLink1.xml.rels'] ?? new Uint8Array(),
		)
		expect(rels).toContain('Target="../sources/reforecast &amp; final.xlsx"')
		expect(rels).toContain('TargetMode="External"')
		expect(rels).toContain(
			'Type="http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup"',
		)
		expect(rels).toContain('Id="rIdMetadata"')
		expect(rels).toContain('Target="../customXml/item1.xml"')

		const read = readXlsx(written.value)
		expectOk(read)
		expect(read.value.workbook.externalReferenceDetails[0]).toMatchObject({
			partPath: 'xl/externalLinks/externalLink1.xml',
			linkRelId: 'rIdExt',
			linkRelationshipType:
				'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup',
			target: '../sources/reforecast & final.xlsx',
			targetMode: 'External',
		})
	})

	it('repairs missing externalBook relationship ids when rewriting external link targets', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		wb.externalReferences.push('xl/externalLinks/externalLink1.xml')
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rId2',
			externalBookRelId: 'rIdExt',
			linkBindingStatus: 'missingPathRelationship',
		})

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/externalLinks/externalLink1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
				relationships: [
					{
						id: 'rIdMetadata',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
						target: '../customXml/item1.xml',
					},
				],
				content: new TextEncoder().encode(
					'<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rIdExt"/></externalLink>',
				),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
			},
		]

		const applied = applyOperations(wb, [
			{
				op: 'rewriteExternalLink',
				partPath: 'xl/externalLinks/externalLink1.xml',
				newTarget: '../sources/repaired.xlsx',
			},
		])
		expectOk(applied)
		const written = writeXlsx(wb, capsules, { workbookMetaDirty: true })
		expectOk(written)

		const parts = unzipSync(written.value)
		const rels = new TextDecoder().decode(
			parts['xl/externalLinks/_rels/externalLink1.xml.rels'] ?? new Uint8Array(),
		)
		expect(rels).toContain('Id="rIdMetadata"')
		expect(rels).toContain('Target="../customXml/item1.xml"')
		expect(rels).toContain('Id="rIdExt"')
		expect(rels).toContain(
			'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"',
		)
		expect(rels).toContain('Target="../sources/repaired.xlsx"')

		const read = readXlsx(written.value)
		expectOk(read)
		expect(read.value.workbook.externalReferenceDetails[0]).toMatchObject({
			partPath: 'xl/externalLinks/externalLink1.xml',
			externalBookRelId: 'rIdExt',
			linkRelId: 'rIdExt',
			linkRelationshipKind: 'externalLinkPath',
			linkBindingStatus: 'externalBookRelId',
			target: '../sources/repaired.xlsx',
			targetMode: 'External',
		})
	})

	it('preserves workbook theme parts on round-trip', () => {
		const wb = new Workbook()
		const themedStyle = wb.styles.register({
			font: { color: { kind: 'theme', theme: 4, tint: -0.25 } },
			fill: { pattern: 'solid', fgColor: { kind: 'theme', theme: 5 } },
		})
		const sheet = wb.addSheet('Theme')
		sheet.cells.set(0, 0, { value: stringValue('Brand'), formula: null, styleId: themedStyle })
		wb.preservedTheme = {
			path: 'xl/theme/theme1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
			xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme">
  <a:themeElements>
    <a:clrScheme name="Brand">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="123456"/></a:dk2>
      <a:lt2><a:srgbClr val="F2F2F2"/></a:lt2>
      <a:accent1><a:srgbClr val="004488"/></a:accent1>
      <a:accent2><a:srgbClr val="D64545"/></a:accent2>
      <a:accent3><a:srgbClr val="4CAF50"/></a:accent3>
      <a:accent4><a:srgbClr val="7E57C2"/></a:accent4>
      <a:accent5><a:srgbClr val="00ACC1"/></a:accent5>
      <a:accent6><a:srgbClr val="FB8C00"/></a:accent6>
      <a:hlink><a:srgbClr val="1A73E8"/></a:hlink>
      <a:folHlink><a:srgbClr val="7B1FA2"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Brand Fonts">
      <a:majorFont><a:latin typeface="Inter Display"/></a:majorFont>
      <a:minorFont><a:latin typeface="Inter"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Brand Formats"/>
  </a:themeElements>
</a:theme>`,
		}
		wb.themeMetadata = {
			name: 'Custom Theme',
			colorSchemeName: 'Brand',
			colorCount: 12,
			majorFontLatin: 'Inter Display',
			minorFontLatin: 'Inter',
		}

		const { result, bytes } = roundTrip(wb)
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.partPaths).toContain('xl/theme/theme1.xml')
		expect(result.workbook.preservedTheme).toEqual({
			path: 'xl/theme/theme1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
		})
		expect(result.workbook.themeMetadata).toEqual({
			name: 'Custom Theme',
			colorSchemeName: 'Brand',
			colorCount: 12,
			majorFontLatin: 'Inter Display',
			minorFontLatin: 'Inter',
		})
		const style = result.workbook.styles.get(
			result.workbook.sheets[0]?.cells.get(0, 0)?.styleId ?? S0,
		)
		expect(style?.font?.color).toEqual({ kind: 'theme', theme: 4, tint: -0.25 })
		expect(style?.fill?.fgColor).toEqual({ kind: 'theme', theme: 5 })
		expect(fingerprint.workbookRels?.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 4,
		})
		expect(fingerprint.contentTypes?.normalized).toContain('/xl/theme/theme1.xml')
	})

	it('generates workbook theme part from theme metadata when no preserved theme exists', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('ThemeMeta')
		sheet.cells.set(0, 0, { value: stringValue('Brand'), formula: null, styleId: S0 })
		wb.themeMetadata = {
			name: 'Generated Theme',
			colorSchemeName: 'Generated Colors',
			colorCount: 12,
			majorFontLatin: 'Aptos Display',
			minorFontLatin: 'Aptos',
		}

		const { result, bytes } = roundTrip(wb)
		const zip = unzipSync(bytes)
		const themeXml = new TextDecoder().decode(zip['xl/theme/theme1.xml'] ?? new Uint8Array())
		expect(themeXml).toContain('name="Generated Theme"')
		expect(themeXml).toContain('name="Generated Colors"')
		expect(themeXml).toContain('typeface="Aptos Display"')
		expect(themeXml).toContain('typeface="Aptos"')
		expect(result.workbook.preservedTheme).toEqual({
			path: 'xl/theme/theme1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
		})
		expect(result.workbook.themeMetadata).toEqual({
			name: 'Generated Theme',
			colorSchemeName: 'Generated Colors',
			colorCount: 12,
			majorFontLatin: 'Aptos Display',
			minorFontLatin: 'Aptos',
		})
	})

	it('updates preserved theme metadata and colors without dropping format scheme XML', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Theme')
		sheet.cells.set(0, 0, { value: stringValue('Brand'), formula: null, styleId: S0 })
		wb.preservedTheme = {
			path: 'xl/theme/theme1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
			xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<theme xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Twist">
  <themeElements>
    <clrScheme name="Office">
      <dk1><sysClr val="windowText" lastClr="000000"/></dk1>
      <lt1><sysClr val="window" lastClr="FFFFFF"/></lt1>
      <accent1><srgbClr val="4F81BD"/></accent1>
      <hlink><srgbClr val="0000FF"/></hlink>
    </clrScheme>
    <fontScheme name="Office">
      <majorFont><latin typeface="Aptos Display"/></majorFont>
      <minorFont><latin typeface="Aptos"/></minorFont>
    </fontScheme>
    <fmtScheme name="Keep Me"><fillStyleLst/></fmtScheme>
  </themeElements>
</theme>`,
		}
		wb.themeMetadata = {
			name: 'Office Twist',
			colorSchemeName: 'Office',
			colorCount: 4,
			majorFontLatin: 'Aptos Display',
			minorFontLatin: 'Aptos',
		}
		wb.themeColors.push(
			{ slot: 'dk1', systemColor: 'windowText', lastColor: '000000' },
			{ slot: 'lt1', systemColor: 'window', lastColor: 'FFFFFF' },
			{ slot: 'accent1', rgb: '4F81BD' },
			{ slot: 'hlink', rgb: '0000FF' },
		)

		const applied = applyOperations(wb, [
			{
				op: 'setTheme',
				themeName: 'Brand Theme',
				colorSchemeName: 'Brand Colors',
				majorFontLatin: 'Inter Display',
				minorFontLatin: 'Inter',
				themeColors: [{ slot: 'accent1', rgb: '0F6CBD' }],
			},
		])
		expectOk(applied)

		const written = writeXlsx(wb)
		expectOk(written)
		const zip = unzipSync(written.value)
		const themeXml = new TextDecoder().decode(zip['xl/theme/theme1.xml'] ?? new Uint8Array())
		expect(themeXml).toContain('name="Brand Theme"')
		expect(themeXml).toContain('<clrScheme name="Brand Colors">')
		expect(themeXml).toContain('<majorFont><latin typeface="Inter Display"/></majorFont>')
		expect(themeXml).toContain('<minorFont><latin typeface="Inter"/></minorFont>')
		expect(themeXml).toContain('<accent1><srgbClr val="0F6CBD"/></accent1>')
		expect(themeXml).toContain('<fmtScheme name="Keep Me"><fillStyleLst/></fmtScheme>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.themeMetadata).toEqual({
			name: 'Brand Theme',
			colorSchemeName: 'Brand Colors',
			colorCount: 4,
			majorFontLatin: 'Inter Display',
			minorFontLatin: 'Inter',
		})
		expect(reopened.value.workbook.themeColors.find((color) => color.slot === 'accent1')).toEqual({
			slot: 'accent1',
			rgb: '0F6CBD',
		})
	})

	it('preserves table-part sheet wiring when table capsules are present', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Balance')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'BalanceTable',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [
				{ id: 1, name: 'Name' },
				{ id: 2, name: 'Value' },
			],
			hasHeaders: true,
			hasTotals: false,
		})

		const capsuleContent = new TextEncoder().encode(`<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="BalanceTable" ref="A1:B2" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="2">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Value"/>
  </tableColumns>
</table>`)
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/tables/table1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
				relationships: [],
				content: capsuleContent,
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
			},
		]

		const { result, bytes } = roundTrip(wb, capsules)
		expect(result.workbook.sheets[0]?.tables).toHaveLength(1)
		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			tableParts: 1,
			tablePart: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 1,
		})
		expect(unzipSync(bytes)['xl/tables/table1.xml']).toEqual(capsuleContent)
	})

	it('matches table capsules by model when sheet relationships are out of part-name order', () => {
		const table1Xml = `<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="FirstTable" ref="A1:A2" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="1"><tableColumn id="1" name="First"/></tableColumns>
</table>`
		const table2Xml = `<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="Table2" displayName="SecondTable" ref="B1:B2" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="1"><tableColumn id="1" name="Second"/></tableColumns>
</table>`
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/tables/table1.xml': table1Xml,
			'xl/tables/table2.xml': table2Xml,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>First</t></is></c><c r="B1" t="inlineStr"><is><t>Second</t></is></c></row>
    <row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row>
  </sheetData>
  <tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		expect(sheet.tables.map((table) => table.name)).toEqual(['SecondTable', 'FirstTable'])
		sheet.cells.set(1, 0, { value: numberValue(3), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Data'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const table1 = new TextDecoder().decode(zip['xl/tables/table1.xml'] ?? new Uint8Array())
		const table2 = new TextDecoder().decode(zip['xl/tables/table2.xml'] ?? new Uint8Array())
		const sheetRels = new TextDecoder().decode(
			zip['xl/worksheets/_rels/sheet1.xml.rels'] ?? new Uint8Array(),
		)
		expect(table1).toContain('displayName="FirstTable"')
		expect(table2).toContain('displayName="SecondTable"')
		expect(sheetRels).toContain('Id="rId1"')
		expect(sheetRels).toContain('Target="../tables/table2.xml"')
		expect(sheetRels).toContain('Id="rId2"')
		expect(sheetRels).toContain('Target="../tables/table1.xml"')
	})

	it('regenerates table parts when capsule-backed table metadata changes', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Balance')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Value'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.cells.set(2, 0, { value: stringValue('Debt'), formula: null, styleId: S0 })
		sheet.cells.set(2, 1, { value: numberValue(20), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'BalanceTable',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [
				{ id: 1, name: 'Name' },
				{ id: 2, name: 'Value' },
			],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:B2',
				columns: [{ colId: 1, kind: 'filters', values: ['10'] }],
				sortState: { ref: 'A1:B2', conditions: [{ ref: 'B2:B2', descending: true }] },
			},
		})
		const resized = applyOperations(wb, [
			{ op: 'resizeTable', table: 'BalanceTable', ref: 'A1:B3' },
		])
		expectOk(resized)

		const staleCapsuleContent = new TextEncoder().encode(`<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="BalanceTable" ref="A1:B2" headerRowCount="1" totalsRowCount="0">
  <tableColumns count="2">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Value"/>
  </tableColumns>
</table>`)
		const { bytes } = roundTrip(wb, [
			{
				partPath: 'xl/tables/table1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
				relationships: [],
				content: staleCapsuleContent,
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
			},
		])

		const tableEntry = unzipSync(bytes)['xl/tables/table1.xml']
		expect(tableEntry).toBeDefined()
		if (!tableEntry) return
		expect(tableEntry).not.toEqual(staleCapsuleContent)
		const tableXml = new TextDecoder().decode(tableEntry)
		expect(tableXml).toContain('ref="A1:B3"')
		expect(tableXml).toContain('<autoFilter ref="A1:B3">')
		expect(tableXml).toContain('<sortState ref="A1:B3">')
		expect(tableXml).toContain('<sortCondition ref="B2:B3" descending="1"/>')
	})

	it('persists table resize filter remaps through generated table XML', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: stringValue('Region'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: stringValue('Status'), formula: null, styleId: S0 })
		sheet.cells.set(0, 3, { value: stringValue('Forecast'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('West'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.cells.set(1, 2, { value: stringValue('Open'), formula: null, styleId: S0 })
		sheet.cells.set(1, 3, { value: numberValue(20), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [
				{ id: 1, name: 'Region' },
				{ id: 2, name: 'Amount' },
				{ id: 3, name: 'Status' },
			],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:C4',
				columns: [
					{ colId: 0, kind: 'filters', values: ['West'] },
					{ colId: 2, kind: 'filters', values: ['Open'] },
				],
				sortState: {
					ref: 'A1:C4',
					conditions: [{ ref: 'A2:A4' }, { ref: 'C2:C4', descending: true }],
				},
			},
			sortState: {
				ref: 'A1:C4',
				conditions: [{ ref: 'A2:A4' }, { ref: 'C2:C4', descending: true }],
			},
		})

		const resized = applyOperations(wb, [{ op: 'resizeTable', table: 'Sales', ref: 'B1:D4' }])
		expectOk(resized)

		const { result, bytes } = roundTrip(wb)
		const table = result.workbook.sheets[0]?.tables[0]
		expect(table?.ref).toEqual({ start: { row: 0, col: 1 }, end: { row: 3, col: 3 } })
		expect(table?.columns.map((column) => column.name)).toEqual(['Amount', 'Status', 'Forecast'])
		expect(table?.autoFilter).toEqual({
			ref: 'B1:D4',
			columns: [{ colId: 1, kind: 'filters', values: ['Open'] }],
			sortState: {
				ref: 'B1:D4',
				conditions: [{ ref: 'C2:C4', descending: true }],
			},
		})
		expect(table?.sortState).toEqual({
			ref: 'B1:D4',
			conditions: [{ ref: 'C2:C4', descending: true }],
		})

		const tableEntry = unzipSync(bytes)['xl/tables/table1.xml']
		expect(tableEntry).toBeDefined()
		if (!tableEntry) return
		const tableXml = new TextDecoder().decode(tableEntry)
		expect(tableXml).toContain('<autoFilter ref="B1:D4">')
		expect(tableXml).toContain('<filterColumn colId="1">')
		expect(tableXml).toContain('<sortState ref="B1:D4">')
		expect(tableXml).toContain('<sortCondition ref="C2:C4" descending="1"/>')
		expect(tableXml).not.toContain('colId="0"')
		expect(tableXml).not.toContain('ref="A2:A4"')
	})

	it('preserves richer table metadata on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Inventory')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Qty'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Bolts'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(5), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			ooxmlId: 42,
			name: 'InventoryTable',
			nameAttribute: null,
			sheetId: sheet.id,
			uid: '{TABLE-UID}',
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [
				{ id: 1, uid: '{COLUMN-1-UID}', name: 'Name', totalsRowLabel: 'Total', dataDxfId: 7 },
				{
					id: 2,
					uid: '{COLUMN-2-UID}',
					name: 'Qty',
					formula: '[@Name]&"-"&[@Qty]',
					formulaIsArray: true,
					xmlColumnPr: {
						mapId: 2,
						xpath: '/ns1:Inventory/ns1:Qty',
						xmlDataType: 'unsignedInt',
					},
					totalsRowFunction: 'sum',
					totalsRowDxfId: 8,
				},
			],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: {
				ref: 'A1:B2',
				uid: '{FILTER-UID}',
				columns: [],
				sortState: {
					ref: 'A2:B2',
					preservedAttributes: {
						'xmlns:xlrd2': 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2',
					},
					conditions: [{ ref: 'B2:B2' }],
				},
			},
			altText: 'Inventory',
			altTextSummary: 'Inventory summary table',
			dataCellStyle: 'InventoryData',
			headerRowDxfId: 5,
			headerRowCellStyle: 'InventoryHeader',
			dataDxfId: 6,
			tableBorderDxfId: 9,
			tableStyleInfo: {
				name: 'TableStyleMedium2',
				showFirstColumn: false,
				showLastColumn: false,
				showRowStripes: true,
				showColumnStripes: false,
			},
		})

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.tables[0]).toEqual(
			expect.objectContaining({
				name: 'InventoryTable',
				nameAttribute: null,
				ooxmlId: 42,
				uid: '{TABLE-UID}',
				altText: 'Inventory',
				altTextSummary: 'Inventory summary table',
				dataCellStyle: 'InventoryData',
				headerRowDxfId: 5,
				headerRowCellStyle: 'InventoryHeader',
				dataDxfId: 6,
				tableBorderDxfId: 9,
				tableStyleInfo: {
					name: 'TableStyleMedium2',
					showFirstColumn: false,
					showLastColumn: false,
					showRowStripes: true,
					showColumnStripes: false,
				},
				columns: [
					{
						id: 1,
						uid: '{COLUMN-1-UID}',
						name: 'Name',
						totalsRowLabel: 'Total',
						dataDxfId: 7,
					},
					{
						id: 2,
						uid: '{COLUMN-2-UID}',
						name: 'Qty',
						formula: '[@Name]&"-"&[@Qty]',
						formulaIsArray: true,
						xmlColumnPr: {
							mapId: 2,
							xpath: '/ns1:Inventory/ns1:Qty',
							xmlDataType: 'unsignedInt',
						},
						totalsRowFunction: 'sum',
						totalsRowDxfId: 8,
					},
				],
			}),
		)
		const entries = unzipSync(bytes)
		const tableEntry = entries['xl/tables/table1.xml']
		expect(tableEntry).toBeDefined()
		if (!tableEntry) return
		const tableXml = new TextDecoder().decode(tableEntry)
		expect(tableXml).toContain('id="42"')
		expect(tableXml).toContain('displayName="InventoryTable"')
		expect(tableXml).not.toContain(' name="InventoryTable"')
		expect(tableXml).toContain(
			'<extLst><ext uri="{504A1905-F514-4f6f-8877-14C23A59335A}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:table altText="Inventory" altTextSummary="Inventory summary table"/></ext></extLst>',
		)
		expect(tableXml).toContain(
			'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"',
		)
		expect(tableXml).toContain(
			'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"',
		)
		expect(tableXml).toContain('mc:Ignorable="xr xr3"')
		expect(tableXml).toContain('xr:uid="{TABLE-UID}"')
		expect(tableXml).toContain('<autoFilter ref="A1:B2" xr:uid="{FILTER-UID}">')
		expect(tableXml).toContain(
			'<sortState ref="A2:B2" xmlns:xlrd2="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2">',
		)
		expect(tableXml).toContain('xr3:uid="{COLUMN-1-UID}"')
		expect(tableXml).toContain('xr3:uid="{COLUMN-2-UID}"')
		expect(tableXml).toContain(
			'<calculatedColumnFormula array="1">[@Name]&amp;&quot;-&quot;&amp;[@Qty]</calculatedColumnFormula>',
		)
		expect(tableXml).toContain(
			'<xmlColumnPr mapId="2" xpath="/ns1:Inventory/ns1:Qty" xmlDataType="unsignedInt"/>',
		)
		expect(tableXml).toContain('dataCellStyle="InventoryData"')
		expect(tableXml).toContain('headerRowDxfId="5"')
		expect(tableXml).toContain('headerRowCellStyle="InventoryHeader"')
		expect(tableXml).toContain('dataDxfId="6"')
		expect(tableXml).toContain('tableBorderDxfId="9"')
		expect(tableXml).toContain('totalsRowFunction="sum"')
		expect(tableXml).toContain('tableStyleInfo')
	})

	it('writes table style edits through generated table XML', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Inventory')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Qty'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Bolts'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(5), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'InventoryTable',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [
				{ id: 1, name: 'Name' },
				{ id: 2, name: 'Qty' },
			],
			hasHeaders: true,
			hasTotals: false,
			tableStyleInfo: {
				name: 'TableStyleMedium4',
				showRowStripes: false,
			},
		})

		const applied = applyOperations(wb, [
			{
				op: 'setTableStyle',
				table: 'InventoryTable',
				styleName: 'TableStyleMedium2',
				showFirstColumn: false,
				showLastColumn: true,
				showRowStripes: true,
				showColumnStripes: false,
			},
		])
		expectOk(applied)

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.tables[0]?.tableStyleInfo).toEqual({
			name: 'TableStyleMedium2',
			showFirstColumn: false,
			showLastColumn: true,
			showRowStripes: true,
			showColumnStripes: false,
		})

		const tableEntry = unzipSync(bytes)['xl/tables/table1.xml']
		expect(tableEntry).toBeDefined()
		if (!tableEntry) return
		const tableXml = new TextDecoder().decode(tableEntry)
		expect(tableXml).toContain(
			'<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="1" showRowStripes="1" showColumnStripes="0"/>',
		)
	})

	it('writes table column renames through generated table XML', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Inventory')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Price'), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: stringValue('Total'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(5), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(4), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'InventoryTable',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 2 } },
			columns: [
				{ id: 1, name: 'Qty' },
				{ id: 2, name: 'Price' },
				{ id: 3, name: 'Total', formula: '[@Qty]*[@Price]' },
			],
			hasHeaders: true,
			hasTotals: false,
		})

		const applied = applyOperations(wb, [
			{
				op: 'setTableColumn',
				table: 'InventoryTable',
				column: 'Qty',
				newName: 'Units',
			},
		])
		expectOk(applied)

		const { result, bytes } = roundTrip(wb)
		const roundTrippedSheet = result.workbook.sheets[0]
		expect(roundTrippedSheet?.tables[0]?.columns.map((column) => column.name)).toEqual([
			'Units',
			'Price',
			'Total',
		])
		expect(roundTrippedSheet?.cells.get(0, 0)?.value).toEqual(stringValue('Units'))
		expect(roundTrippedSheet?.tables[0]?.columns[2]?.formula).toBe('[@Units]*[@Price]')

		const tableEntry = unzipSync(bytes)['xl/tables/table1.xml']
		expect(tableEntry).toBeDefined()
		if (!tableEntry) return
		const tableXml = new TextDecoder().decode(tableEntry)
		expect(tableXml).toContain('<tableColumn id="1" name="Units">')
		expect(tableXml).toContain(
			'<calculatedColumnFormula>[@Units]*[@Price]</calculatedColumnFormula>',
		)
	})

	it('persists table structured-reference rewrites in worksheet metadata', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: stringValue('Qty'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Price'), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: stringValue('Total'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(5), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(4), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 2 } },
			columns: [
				{ id: 1, name: 'Qty' },
				{ id: 2, name: 'Price' },
				{ id: 3, name: 'Total', formula: '[@Qty]*[@Price]' },
			],
			hasHeaders: true,
			hasTotals: false,
		})
		sheet.dataValidations.push({
			sqref: 'D2:D3',
			type: 'list',
			formula1: 'SUM(Sales[Qty])',
		})
		sheet.conditionalFormats.push({
			sqref: 'E2:E3',
			rules: [
				{
					type: 'expression',
					formulas: ['SUM(Sales[Qty])>0'],
					colorScale: {
						cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }],
						colors: [{ rgb: 'FFFF0000' }],
					},
					dataBar: {
						cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }],
						color: { rgb: 'FF00AA00' },
					},
					iconSet: {
						cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }],
					},
				},
			],
		})
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'F2:F3',
			type: 'list',
			formula1: 'SUM(Sales[Qty])',
		})
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'G2:G3',
			formulas: ['SUM(Sales[Qty])>0'],
			colorScale: {
				cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }],
				colors: [{ rgb: 'FF63BE7B' }],
			},
			dataBar: { cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }] },
		})

		const applied = applyOperations(wb, [
			{ op: 'renameTable', table: 'Sales', newName: 'Inventory' },
			{ op: 'setTableColumn', table: 'Inventory', column: 'Qty', newName: 'Units' },
		])
		expectOk(applied)

		const { result, bytes } = roundTrip(wb)
		const roundTrippedSheet = result.workbook.sheets[0]
		expect(roundTrippedSheet?.tables[0]?.name).toBe('Inventory')
		expect(roundTrippedSheet?.tables[0]?.columns.map((column) => column.name)).toEqual([
			'Units',
			'Price',
			'Total',
		])
		expect(roundTrippedSheet?.tables[0]?.columns[2]?.formula).toBe('[@Units]*[@Price]')
		expect(roundTrippedSheet?.dataValidations[0]?.formula1).toBe('SUM(Inventory[Units])')
		const rule = roundTrippedSheet?.conditionalFormats[0]?.rules[0]
		expect(rule?.formulas[0]).toBe('SUM(Inventory[Units])>0')
		expect(rule?.colorScale?.cfvo[0]?.value).toBe('SUM(Inventory[Units])')
		expect(rule?.dataBar?.cfvo[0]?.value).toBe('SUM(Inventory[Units])')
		expect(rule?.iconSet?.cfvo[0]?.value).toBe('SUM(Inventory[Units])')
		expect(roundTrippedSheet?.x14DataValidations[0]?.formula1).toBe('SUM(Inventory[Units])')
		expect(roundTrippedSheet?.x14ConditionalFormats[0]?.formulas[0]).toBe('SUM(Inventory[Units])>0')
		expect(roundTrippedSheet?.x14ConditionalFormats[0]?.colorScale?.cfvo[0]?.value).toBe(
			'SUM(Inventory[Units])',
		)
		expect(roundTrippedSheet?.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe(
			'SUM(Inventory[Units])',
		)
		expect(roundTrippedSheet?.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe(
			'SUM(Inventory[Units])',
		)

		const entries = unzipSync(bytes)
		const sheetXml = new TextDecoder().decode(
			entries['xl/worksheets/sheet1.xml'] ?? new Uint8Array(),
		)
		expect(sheetXml).toContain('SUM(Inventory[Units])')
		expect(sheetXml).toContain('SUM(Inventory[Units])&gt;0')
		expect(sheetXml).not.toContain('Sales[Qty]')
		expect(sheetXml).not.toContain('Inventory[Qty]')
	})

	it('writes materialized table totals rows from metadata edits', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Cash'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [
				{ id: 1, name: 'Name' },
				{ id: 2, name: 'Amount' },
			],
			hasHeaders: true,
			hasTotals: true,
		})

		const applied = applyOperations(wb, [
			{ op: 'setTableColumn', table: 'Sales', column: 'Name', totalsRowLabel: 'Total' },
			{ op: 'setTableColumn', table: 'Sales', column: 'Amount', totalsRowFunction: 'sum' },
		])
		expectOk(applied)

		const { result, bytes } = roundTrip(wb)
		const roundTrippedSheet = result.workbook.sheets[0]
		expect(roundTrippedSheet?.tables[0]?.hasTotals).toBe(true)
		expect(roundTrippedSheet?.cells.get(2, 0)?.value).toEqual(stringValue('Total'))
		expect(roundTrippedSheet?.cells.get(2, 1)?.formula).toBe('SUBTOTAL(109,Sales[Amount])')

		const entries = unzipSync(bytes)
		const sheetXml = new TextDecoder().decode(
			entries['xl/worksheets/sheet1.xml'] ?? new Uint8Array(),
		)
		expect(sheetXml).toContain('<dimension ref="A1:B3"/>')
		expect(sheetXml).toContain('<f>SUBTOTAL(109,Sales[Amount])</f>')
		const tableXml = new TextDecoder().decode(entries['xl/tables/table1.xml'] ?? new Uint8Array())
		expect(tableXml).toContain('totalsRowCount="1"')
		expect(tableXml).toContain('<autoFilter ref="A1:B2"/>')
		expect(tableXml).toContain('totalsRowFunction="sum"')
	})

	it('round-trips table created via createTable op through XLSX', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: stringValue('Product'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Amount'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Widget'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(100), formula: null, styleId: S0 })
		sheet.cells.set(2, 0, { value: stringValue('Gadget'), formula: null, styleId: S0 })
		sheet.cells.set(2, 1, { value: numberValue(200), formula: null, styleId: S0 })

		const applyResult = applyOperations(wb, [
			{
				op: 'createTable',
				sheet: 'Data',
				ref: 'A1:B3',
				name: 'SalesTable',
				hasHeaders: true,
			},
		])
		expectOk(applyResult)

		const { result } = roundTrip(wb)
		const roundTrippedSheet = result.workbook.sheets[0]
		expect(roundTrippedSheet?.tables).toHaveLength(1)
		const table = roundTrippedSheet?.tables[0]
		expect(table?.name).toBe('SalesTable')
		expect(table?.hasHeaders).toBe(true)
		expect(table?.ref).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 1 } })
		expect(table?.columns).toHaveLength(2)
		expect(table?.columns[0]?.name).toBe('Product')
		expect(table?.columns[1]?.name).toBe('Amount')
	})

	it('emits table parts for semantic tables without capsules', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Inventory')
		sheet.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: stringValue('Qty'), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: stringValue('Bolts'), formula: null, styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(5), formula: null, styleId: S0 })
		sheet.tables.push({
			id: createTableId(),
			name: 'InventoryTable',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Qty' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets[0]?.tables).toHaveLength(1)
		expect(result.workbook.sheets[0]?.tables[0]?.name).toBe('InventoryTable')

		const fingerprint = fingerprintXlsx(bytes)
		expect(fingerprint.partPaths).toContain('xl/tables/table1.xml')
		expect(fingerprint.sheets[0]?.xml.tagCounts).toMatchObject({
			tableParts: 1,
			tablePart: 1,
		})
		expect(fingerprint.sheetRels[0]?.xml.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 1,
		})

		const entries = unzipSync(bytes)
		const tableEntry = entries['xl/tables/table1.xml']
		expect(tableEntry).toBeDefined()
		if (!tableEntry) return
		const tableXml = new TextDecoder().decode(tableEntry)
		expect(tableXml).toContain('InventoryTable')
		expect(tableXml).toContain('<tableColumns count="2">')
	})

	it('preserves capsule parts through write-read cycle', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })

		const capsuleContent = new TextEncoder().encode('<chart>test chart data</chart>')
		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/charts/chart1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
				relationships: [],
				content: capsuleContent,
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
			},
		]

		const { bytes, result } = roundTrip(wb, capsules)

		const entries = unzipSync(bytes)
		expect(entries['xl/charts/chart1.xml']).toBeDefined()

		const decoded = new TextDecoder().decode(entries['xl/charts/chart1.xml'])
		expect(decoded).toBe('<chart>test chart data</chart>')

		const readCapsules = result.capsules
		const chart = readCapsules.find((c) => c.partPath === 'xl/charts/chart1.xml')
		expect(chart).toBeDefined()
		expect(chart?.contentType).toContain('chart')
		expect(chart?.content).toBeUndefined()
	})

	it('updates chart series refs inside preserved chart capsules', () => {
		const wb = new Workbook()
		wb.addSheet('Sheet1')
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			chartType: 'barChart',
			series: [
				{
					nameRef: 'Data!$C$1',
					categoryRef: 'Data!$A$2:$A$10',
					valueRef: 'Data!$C$2:$C$10',
				},
			],
		})
		const chartXml = `<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:barChart><c:ser>
    <c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>
    <c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:pt idx="0"><c:v>Jan</c:v></c:pt></c:strCache></c:strRef></c:cat>
    <c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>
  </c:ser></c:barChart></c:plotArea></c:chart>
</c:chartSpace>`
		const written = writeXlsx(wb, [
			{
				partPath: 'xl/charts/chart1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
				relationships: [],
				content: new TextEncoder().encode(chartXml),
				anchor: { kind: 'workbook' },
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
			},
		])
		expectOk(written)

		const entries = unzipSync(written.value)
		const updated = entries['xl/charts/chart1.xml']
		expect(updated).toBeDefined()
		if (!updated) return
		const xml = new TextDecoder().decode(updated)
		expect(xml).toContain('<c:f>Data!$C$1</c:f>')
		expect(xml).toContain('<c:f>Data!$A$2:$A$10</c:f>')
		expect(xml).toContain('<c:f>Data!$C$2:$C$10</c:f>')
		expect(xml).not.toContain('<c:f>Data!$B$2:$B$4</c:f>')
		expect(xml).not.toContain('<c:strCache>')
		expect(xml).not.toContain('<c:numCache>')
	})

	it('invalidates stale caches in real ClosedXML chart edits', () => {
		const source = readXlsx(
			readFileSync(
				new URL(
					'../../../../fixtures/xlsx/closedxml/Other_Charts_PreserveCharts_inputfile.xlsx',
					import.meta.url,
				),
			),
		)
		expectOk(source)
		const chart = source.value.workbook.chartParts[0]
		expect(chart?.partPath).toBe('xl/charts/chart1.xml')

		const applied = applyOperations(source.value.workbook, [
			{
				op: 'setChartSeriesSource',
				partPath: 'xl/charts/chart1.xml',
				seriesIndex: 0,
				valueRef: 'Sheet1!$C$2:$C$8',
			},
		])
		expectOk(applied)
		const written = writeXlsx(source.value.workbook, source.value.capsules)
		expectOk(written)

		const entries = unzipSync(written.value)
		const updated = entries['xl/charts/chart1.xml']
		expect(updated).toBeDefined()
		if (!updated) return
		const xml = new TextDecoder().decode(updated)
		const categoryXml = xml.match(/<c:cat>[\s\S]*?<\/c:cat>/)?.[0]
		const valueXml = xml.match(/<c:val>[\s\S]*?<\/c:val>/)?.[0]

		expect(categoryXml).toContain('<c:f>Sheet1!$A$2:$A$8</c:f>')
		expect(categoryXml).toContain('<c:numCache>')
		expect(valueXml).toContain('<c:f>Sheet1!$C$2:$C$8</c:f>')
		expect(valueXml).not.toContain('<c:numCache>')
		expect(valueXml).not.toContain('<c:v>3</c:v>')
	})

	it('rewrites x14 data-validation ranges in real ClosedXML structural edits', () => {
		const source = readXlsx(
			readFileSync(
				new URL('../../../../fixtures/xlsx/closedxml/Misc_DataValidation.xlsx', import.meta.url),
			),
		)
		expectOk(source)
		const sheet = source.value.workbook.sheets.find(
			(entry) => entry.name === 'Data Validation - Copy',
		)
		expect(sheet?.x14DataValidations).toEqual([
			{
				index: 0,
				sqref: 'A5:A5',
				type: 'list',
				operator: 'between',
				allowBlank: true,
				showInputMessage: true,
				showErrorMessage: true,
				showDropDown: false,
				errorStyle: 'stop',
				formula1: "'Data Validation'!$C$1:$C$2",
			},
		])

		const applied = applyOperations(source.value.workbook, [
			{
				op: 'insertRows',
				sheet: 'Data Validation - Copy',
				at: 4,
				count: 1,
			},
		])
		expectOk(applied)
		expect(sheet?.x14DataValidations[0]?.sqref).toBe('A6')
		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)

		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet3.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)
		expect(xml).toContain('<xm:sqref>A6</xm:sqref>')
		expect(xml).not.toContain('<xm:sqref>A5:A5</xm:sqref>')
		expect(xml).toContain("<xm:f>'Data Validation'!$C$1:$C$2</xm:f>")
	})

	it('rewrites same-sheet x14 extension formulas after structural edits', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B2:B3',
			formulas: ['A2>0'],
		})
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'C2:C3',
			formula1: 'A2:A3',
		})
		sheet.preservedExtLst = `<extLst xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression"><xm:f>A2>0</xm:f></x14:cfRule><xm:sqref>B2:B3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings><x14:dataValidations count="1"><x14:dataValidation type="list"><x14:formula1><xm:f>A2:A3</xm:f></x14:formula1><xm:sqref>C2:C3</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`

		const applied = applyOperations(wb, [{ op: 'insertRows', sheet: 'Data', at: 1, count: 1 }])
		expectOk(applied)
		const written = writeXlsx(wb, [], { dirtySheetNames: applied.value.sheetsModified })
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)

		expect(xml).toContain('<xm:sqref>B3:B4</xm:sqref>')
		expect(xml).toContain('<xm:f>A3&gt;0</xm:f>')
		expect(xml).toContain('<xm:sqref>C3:C4</xm:sqref>')
		expect(xml).toContain('<xm:f>A3:A4</xm:f>')
		expect(xml).not.toContain('<xm:f>A2&gt;0</xm:f>')
		expect(xml).not.toContain('<xm:f>A2:A3</xm:f>')
	})

	it('writes generated x14 extension validation and conditional-format metadata', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'A2:B2',
			type: 'whole',
			formula1: 'Data!A2',
			formula2: 'Data!B2',
			preservedAttributes: {
				'xr:uid': '{DV-UID}',
				customFlag: '1',
			},
			preservedChildXml: ['<x14ac:metadata flag="1"><x14ac:item val="keep"/></x14ac:metadata>'],
		})
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'C2:D2',
			formulas: ['Data!A2>0'],
			type: 'expression',
			preservedRuleAttributes: {
				activePresent: '1',
				'xr:uid': '{CF-UID}',
			},
			preservedRuleChildXml: [
				'<x14:extLst><x14:ext uri="{cf-extension}"><x14ac:metadata flag="1"/></x14:ext></x14:extLst>',
			],
			colorScale: {
				cfvo: [{ type: 'formula', value: 'Data!A2' }],
				colors: [{ rgb: 'FF63BE7B' }, { theme: 4, tint: 0.25 }],
				preservedAttributes: {
					customScale: '1',
					'x14ac:scaleId': '{SCALE-UID}',
				},
				preservedChildXml: ['<x14ac:metadata flag="scale"/>'],
			},
			dataBar: { cfvo: [{ type: 'formula', value: 'Data!A2' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'Data!B2' }] },
		})

		const written = writeXlsx(wb)
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)

		expect(xml).toContain('<extLst')
		expect(xml).toContain('<x14:conditionalFormattings>')
		expect(xml).toContain('<x14:dataValidations count="1">')
		expect(xml).toContain('xr:uid="{DV-UID}"')
		expect(xml).toContain(
			`xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"`,
		)
		expect(xml).toContain(
			`xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"`,
		)
		expect(xml).toContain('customFlag="1"')
		expect(xml).toContain('activePresent="1"')
		expect(xml).toContain('xr:uid="{CF-UID}"')
		expect(xml).toContain('<x14ac:metadata flag="1"><x14ac:item val="keep"/></x14ac:metadata>')
		expect(xml).toContain(
			'<x14:extLst><x14:ext uri="{cf-extension}"><x14ac:metadata flag="1"/></x14:ext></x14:extLst>',
		)
		expect(xml).toContain('<xm:sqref>A2:B2</xm:sqref>')
		expect(xml).toContain('<xm:sqref>C2:D2</xm:sqref>')
		expect(xml).toContain('<xm:f>Data!A2&gt;0</xm:f>')
		expect(xml).toContain('customScale="1"')
		expect(xml).toContain('x14ac:scaleId="{SCALE-UID}"')
		expect(xml).toContain('<x14:colorScale')
		expect(xml).toContain('<x14:cfvo type="formula"><xm:f>Data!A2</xm:f></x14:cfvo>')
		expect(xml).toContain('<x14:color rgb="FF63BE7B"/>')
		expect(xml).toContain('<x14:color theme="4" tint="0.25"/>')
		expect(xml).toContain('<x14ac:metadata flag="scale"/>')
		expect(xml).toContain('<x14:cfvo type="formula"><xm:f>Data!A2</xm:f></x14:cfvo>')
		expect(xml).toContain('<x14:cfvo type="formula"><xm:f>Data!B2</xm:f></x14:cfvo>')

		const read = readXlsx(written.value)
		expectOk(read)
		const roundTripped = read.value.workbook.sheets[0]
		expect(roundTripped?.x14DataValidations[0]).toMatchObject({
			sqref: 'A2:B2',
			formula1: 'Data!A2',
			formula2: 'Data!B2',
			preservedAttributes: {
				'xr:uid': '{DV-UID}',
				customFlag: '1',
			},
			preservedChildXml: ['<x14ac:metadata flag="1"><x14ac:item val="keep"/></x14ac:metadata>'],
		})
		expect(roundTripped?.x14ConditionalFormats[0]?.colorScale).toMatchObject({
			cfvo: [{ type: 'formula', value: 'Data!A2' }],
			colors: [{ rgb: 'FF63BE7B' }, { theme: 4, tint: 0.25 }],
			preservedAttributes: {
				customScale: '1',
				'x14ac:scaleId': '{SCALE-UID}',
			},
			preservedChildXml: ['<x14ac:metadata flag="scale"/>'],
		})
		expect(roundTripped?.x14ConditionalFormats[0]?.dataBar?.cfvo[0]?.value).toBe('Data!A2')
		expect(roundTripped?.x14ConditionalFormats[0]?.iconSet?.cfvo[0]?.value).toBe('Data!B2')
		expect(roundTripped?.x14ConditionalFormats[0]?.preservedRuleAttributes).toEqual({
			activePresent: '1',
			'xr:uid': '{CF-UID}',
		})
		expect(roundTripped?.x14ConditionalFormats[0]?.preservedRuleChildXml).toEqual([
			'<x14:extLst><x14:ext uri="{cf-extension}"><x14ac:metadata flag="1"/></x14:ext></x14:extLst>',
		])
	})

	it('removes x14 data-validation extension entries deleted by structural edits', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'C2:C3',
			formula1: 'A2:A3',
		})
		sheet.preservedExtLst = `<extLst xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1"><x14:dataValidation type="list"><x14:formula1><xm:f>A2:A3</xm:f></x14:formula1><xm:sqref>C2:C3</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`

		const applied = applyOperations(wb, [{ op: 'deleteRows', sheet: 'Data', at: 1, count: 2 }])
		expectOk(applied)
		expect(sheet.x14DataValidations[0]?.deleted).toBe(true)
		const written = writeXlsx(wb, [], { dirtySheetNames: applied.value.sheetsModified })
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)

		expect(xml).not.toContain('<x14:dataValidations')
		expect(xml).not.toContain('<x14:dataValidation ')
		expect(xml).not.toContain('<xm:sqref>C2:C3</xm:sqref>')
		expect(xml).not.toContain('<xm:f>A2:A3</xm:f>')
	})

	it('removes x14 conditional-format extension containers deleted by structural edits', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'C2:C3',
			formulas: ['A2>0'],
			type: 'expression',
			preservedRuleChildXml: [
				'<x14:extLst><x14:ext uri="{cf-extension}"><x14ac:metadata flag="1"/></x14:ext></x14:extLst>',
			],
		})
		sheet.preservedExtLst = `<extLst xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression"><xm:f>A2&gt;0</xm:f><x14:extLst><x14:ext uri="{cf-extension}"><x14ac:metadata flag="1"/></x14:ext></x14:extLst></x14:cfRule><xm:sqref>C2:C3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings><x14ac:metadata flag="keep"/></ext></extLst>`

		const applied = applyOperations(wb, [{ op: 'deleteRows', sheet: 'Data', at: 1, count: 2 }])
		expectOk(applied)
		expect(sheet.x14ConditionalFormats[0]?.deleted).toBe(true)
		const written = writeXlsx(wb, [], { dirtySheetNames: applied.value.sheetsModified })
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)

		expect(xml).not.toContain('<x14:conditionalFormattings')
		expect(xml).not.toContain('<x14:conditionalFormatting')
		expect(xml).not.toContain('<x14:cfRule')
		expect(xml).not.toContain('<xm:sqref>C2:C3</xm:sqref>')
		expect(xml).not.toContain('<xm:f>A2&gt;0</xm:f>')
		expect(xml).toContain('<x14ac:metadata flag="keep"/>')
	})

	it('updates mixed self-closing x14 data-validation entries by document-order index', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.x14DataValidations.push(
			{ index: 0, sqref: 'A2' },
			{ index: 1, sqref: 'B1', formula1: 'Lookup!$A$1:$A$2', deleted: true },
			{ index: 2, sqref: 'C1', deleted: true },
		)
		sheet.preservedExtLst = `<extLst xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="3"><x14:dataValidation sqref="A1"/><x14:dataValidation type="list"><x14:formula1><xm:f>Lookup!$A$1:$A$2</xm:f></x14:formula1><xm:sqref>B1</xm:sqref></x14:dataValidation><x14:dataValidation sqref="C1"/></x14:dataValidations></ext></extLst>`

		const written = writeXlsx(wb)
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)

		expect(xml).toContain('<x14:dataValidations count="1">')
		expect(xml).toContain('<x14:dataValidation><xm:sqref>A2</xm:sqref></x14:dataValidation>')
		expect(xml).not.toContain('sqref="A1"')
		expect(xml).not.toContain('<xm:sqref>B1</xm:sqref>')
		expect(xml).not.toContain('sqref="C1"')
		expect(xml).not.toContain('Lookup!$A$1:$A$2')
	})

	it('preserves x14 validation and conditional-format payloads while updating and deleting entries', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.x14DataValidations.push(
			{
				index: 0,
				sqref: 'A2',
				type: 'list',
				formula1: 'List!$A$2:$A$3',
				preservedAttributes: { customFlag: 'keep' },
				preservedChildXml: ['<x14ac:metadata flag="keep"><x14ac:item val="1"/></x14ac:metadata>'],
			},
			{
				index: 1,
				sqref: 'B1',
				type: 'whole',
				deleted: true,
			},
		)
		sheet.x14ConditionalFormats.push(
			{
				index: 0,
				sqref: 'C2',
				formulas: ['A2>0'],
				type: 'expression',
				preservedRuleAttributes: {
					activePresent: '1',
					'xr:uid': '{CF-KEEP}',
				},
				preservedRuleChildXml: [
					'<x14:extLst><x14:ext uri="{cf-child}"><x14ac:metadata flag="keep"/></x14:ext></x14:extLst>',
				],
			},
			{
				index: 1,
				sqref: 'D1',
				formulas: ['B1>0'],
				type: 'expression',
				deleted: true,
			},
		)
		sheet.preservedExtLst = `<extLst xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression" activePresent="1" xr:uid="{CF-KEEP}"><xm:f>A1&gt;0</xm:f><x14:extLst><x14:ext uri="{cf-child}"><x14ac:metadata flag="keep"/></x14:ext></x14:extLst></x14:cfRule><xm:sqref>C1</xm:sqref></x14:conditionalFormatting><x14:conditionalFormatting><x14:cfRule type="expression"><xm:f>B1&gt;0</xm:f></x14:cfRule><xm:sqref>D1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings><x14:dataValidations count="2"><x14:dataValidation type="list" customFlag="keep"><x14:formula1><xm:f>List!$A$1:$A$2</xm:f></x14:formula1><x14ac:metadata flag="keep"><x14ac:item val="1"/></x14ac:metadata><xm:sqref>A1</xm:sqref></x14:dataValidation><x14:dataValidation type="whole" customFlag="drop"><xm:sqref>B1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`

		const written = writeXlsx(wb)
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)
		const validationXml = xml.match(/<x14:dataValidations\b[\s\S]*?<\/x14:dataValidations>/)?.[0]
		const conditionalXml = xml.match(
			/<x14:conditionalFormattings\b[\s\S]*?<\/x14:conditionalFormattings>/,
		)?.[0]

		expect(validationXml).toContain('<x14:dataValidations count="1">')
		expect(validationXml).toContain('customFlag="keep"')
		expect(validationXml).toContain(
			'<x14ac:metadata flag="keep"><x14ac:item val="1"/></x14ac:metadata>',
		)
		expect(validationXml).toContain('<xm:f>List!$A$2:$A$3</xm:f>')
		expect(validationXml).toContain('<xm:sqref>A2</xm:sqref>')
		expect(validationXml).not.toContain('customFlag="drop"')
		expect(validationXml).not.toContain('<xm:sqref>B1</xm:sqref>')
		expect(validationXml?.indexOf('<x14:formula1')).toBeLessThan(
			validationXml?.indexOf('<x14ac:metadata') ?? -1,
		)
		expect(validationXml?.indexOf('<x14ac:metadata')).toBeLessThan(
			validationXml?.indexOf('<xm:sqref>A2</xm:sqref>') ?? -1,
		)

		expect(conditionalXml).toContain('activePresent="1"')
		expect(conditionalXml).toContain('xr:uid="{CF-KEEP}"')
		expect(conditionalXml).toContain(
			'<x14:extLst><x14:ext uri="{cf-child}"><x14ac:metadata flag="keep"/></x14:ext></x14:extLst>',
		)
		expect(conditionalXml).toContain('<xm:f>A2&gt;0</xm:f>')
		expect(conditionalXml).toContain('<xm:sqref>C2</xm:sqref>')
		expect(conditionalXml).not.toContain('<xm:sqref>D1</xm:sqref>')
		expect(conditionalXml).not.toContain('<xm:f>B1&gt;0</xm:f>')
		expect(conditionalXml?.indexOf('<xm:f>A2&gt;0</xm:f>')).toBeLessThan(
			conditionalXml?.indexOf('<x14:extLst>') ?? -1,
		)
		expect(conditionalXml?.indexOf('<x14:extLst>')).toBeLessThan(
			conditionalXml?.indexOf('<xm:sqref>C2</xm:sqref>') ?? -1,
		)
	})

	it('rewrites full modeled x14 data-validation entries over preserved XML', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'A2:A4',
			type: 'whole',
			operator: 'between',
			allowBlank: true,
			showInputMessage: false,
			showErrorMessage: true,
			showDropDown: false,
			promptTitle: 'New prompt title',
			prompt: 'New prompt',
			errorTitle: 'New error title',
			error: 'New error',
			errorStyle: 'warning',
			imeMode: 'disabled',
			formula1: '1',
			formula2: '10',
			preservedAttributes: { customFlag: 'new' },
			preservedChildXml: ['<x14ac:metadata flag="new"/>'],
		})
		sheet.preservedExtLst = `<extLst xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1"><x14:dataValidation type="list" allowBlank="0" showInputMessage="1" prompt="Old prompt" customFlag="old"><x14:formula1><xm:f>Old!$A$1:$A$2</xm:f></x14:formula1><x14ac:metadata flag="old"/><xm:sqref>A1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`

		const written = writeXlsx(wb)
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)
		const validationXml = xml.match(/<x14:dataValidation\b[\s\S]*?<\/x14:dataValidation>/)?.[0]

		expect(validationXml).toContain('type="whole"')
		expect(validationXml).toContain('operator="between"')
		expect(validationXml).toContain('allowBlank="1"')
		expect(validationXml).toContain('showInputMessage="0"')
		expect(validationXml).toContain('showErrorMessage="1"')
		expect(validationXml).toContain('showDropDown="0"')
		expect(validationXml).toContain('promptTitle="New prompt title"')
		expect(validationXml).toContain('prompt="New prompt"')
		expect(validationXml).toContain('errorTitle="New error title"')
		expect(validationXml).toContain('error="New error"')
		expect(validationXml).toContain('errorStyle="warning"')
		expect(validationXml).toContain('imeMode="disabled"')
		expect(validationXml).toContain('customFlag="new"')
		expect(validationXml).toContain('<x14:formula1><xm:f>1</xm:f></x14:formula1>')
		expect(validationXml).toContain('<x14:formula2><xm:f>10</xm:f></x14:formula2>')
		expect(validationXml).toContain('<x14ac:metadata flag="new"/>')
		expect(validationXml).toContain('<xm:sqref>A2:A4</xm:sqref>')
		expect(validationXml).not.toContain('Old prompt')
		expect(validationXml).not.toContain('customFlag="old"')
		expect(validationXml).not.toContain('Old!$A$1:$A$2')
		expect(validationXml).not.toContain('<xm:sqref>A1</xm:sqref>')
	})

	it('rewrites full modeled x14 conditional-format entries over preserved XML', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'C2:C4',
			type: 'iconSet',
			priority: 7,
			id: '{NEW-CF}',
			formulas: ['A2>0'],
			preservedRuleAttributes: {
				activePresent: '1',
				'xr:uid': '{NEW-UID}',
			},
			preservedRuleChildXml: ['<x14ac:metadata flag="new"/>'],
			colorScale: {
				cfvo: [{ type: 'formula', value: 'A2' }, { type: 'max' }],
				colors: [{ rgb: 'FF63BE7B' }, { rgb: 'FFFFEB84' }],
				preservedAttributes: { customScale: '1' },
				preservedChildXml: ['<x14ac:metadata flag="scale-new"/>'],
			},
			dataBar: {
				minLength: 5,
				maxLength: 95,
				border: true,
				showValue: false,
				gradient: false,
				direction: 'rightToLeft',
				axisPosition: 'middle',
				negativeBarColorSameAsPositive: false,
				negativeBarBorderColorSameAsPositive: true,
				cfvo: [
					{ type: 'formula', value: 'A2' },
					{ type: 'num', value: '10', gte: false },
				],
				fillColor: { rgb: 'FF638EC6' },
				borderColor: { rgb: 'FF003300' },
				negativeFillColor: { rgb: 'FF00AA00' },
				negativeBorderColor: { theme: 5 },
				axisColor: { auto: true },
			},
			iconSet: {
				iconSet: '3Flags',
				reverse: true,
				showValue: false,
				cfvo: [
					{ type: 'percent', value: '0' },
					{ type: 'percent', value: '50' },
					{ type: 'percent', value: '90' },
				],
				icons: [
					{ iconSet: '3Flags', iconId: 0 },
					{ iconSet: '3Flags', iconId: 1 },
					{ iconSet: '3Flags', iconId: 2 },
				],
			},
		})
		sheet.preservedExtLst = `<extLst xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" priority="1" id="{OLD-CF}" oldFlag="drop"><xm:f>Old!A1&gt;0</xm:f><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:fillColor rgb="FFFF0000"/></x14:dataBar><x14ac:metadata flag="old"/></x14:cfRule><xm:sqref>B1:B3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`

		const written = writeXlsx(wb)
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)
		const ruleXml = xml.match(/<x14:cfRule\b[\s\S]*?<\/x14:cfRule>/)?.[0]

		expect(ruleXml).toContain('type="iconSet"')
		expect(ruleXml).toContain('priority="7"')
		expect(ruleXml).toContain('id="{NEW-CF}"')
		expect(ruleXml).toContain('activePresent="1"')
		expect(ruleXml).toContain('xr:uid="{NEW-UID}"')
		expect(ruleXml).toContain('<xm:f>A2&gt;0</xm:f>')
		expect(ruleXml).toContain(
			'<x14:colorScale customScale="1"><x14:cfvo type="formula"><xm:f>A2</xm:f></x14:cfvo><x14:cfvo type="max"/><x14:color rgb="FF63BE7B"/><x14:color rgb="FFFFEB84"/><x14ac:metadata flag="scale-new"/></x14:colorScale>',
		)
		expect(ruleXml).toContain(
			'<x14:dataBar minLength="5" maxLength="95" border="1" showValue="0" gradient="0" direction="rightToLeft" axisPosition="middle" negativeBarColorSameAsPositive="0" negativeBarBorderColorSameAsPositive="1"><x14:cfvo type="formula"><xm:f>A2</xm:f></x14:cfvo><x14:cfvo type="num" gte="0" val="10"/><x14:fillColor rgb="FF638EC6"/><x14:borderColor rgb="FF003300"/><x14:negativeFillColor rgb="FF00AA00"/><x14:negativeBorderColor theme="5"/><x14:axisColor auto="1"/></x14:dataBar>',
		)
		expect(ruleXml).toContain(
			'<x14:iconSet iconSet="3Flags" showValue="0" reverse="1"><x14:cfvo type="percent" val="0"/><x14:cfvo type="percent" val="50"/><x14:cfvo type="percent" val="90"/><x14:cfIcon iconSet="3Flags" iconId="0"/><x14:cfIcon iconSet="3Flags" iconId="1"/><x14:cfIcon iconSet="3Flags" iconId="2"/></x14:iconSet>',
		)
		expect(ruleXml).toContain('<x14ac:metadata flag="new"/>')
		expect(xml).toContain('<xm:sqref>C2:C4</xm:sqref>')
		expect(ruleXml).not.toContain('oldFlag="drop"')
		expect(ruleXml).not.toContain('{OLD-CF}')
		expect(ruleXml).not.toContain('Old!A1')
		expect(ruleXml).not.toContain('FFFF0000')
		expect(xml).not.toContain('<xm:sqref>B1:B3</xm:sqref>')
	})

	it('does not rewrite non-x14 extension nodes with x14 local names before the worksheet extension', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'A2',
			type: 'list',
			formula1: 'List!$A$2:$A$3',
		})
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'C2',
			formulas: ['A2>0'],
			type: 'expression',
		})
		sheet.preservedExtLst = `<extLst xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><ext uri="{UNRELATED-X15}"><x15:conditionalFormattings><x15:conditionalFormatting><x15:cfRule type="expression"><xm:f>Z1&gt;0</xm:f></x15:cfRule><xm:sqref>Z2</xm:sqref></x15:conditionalFormatting></x15:conditionalFormattings><x15:dataValidations count="1"><x15:dataValidation type="whole"><xm:sqref>Z3</xm:sqref></x15:dataValidation></x15:dataValidations></ext><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression"><xm:f>A1&gt;0</xm:f></x14:cfRule><xm:sqref>C1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings><x14:dataValidations count="1"><x14:dataValidation type="list"><x14:formula1><xm:f>List!$A$1:$A$2</xm:f></x14:formula1><xm:sqref>A1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`

		const written = writeXlsx(wb)
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)
		const unrelatedExt = xml.match(/<ext uri="\{UNRELATED-X15\}">[\s\S]*?<\/ext>/)?.[0]
		const worksheetExt = xml.match(
			/<ext uri="\{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF\}"[\s\S]*?<\/ext>/,
		)?.[0]

		expect(unrelatedExt).toContain('<xm:f>Z1&gt;0</xm:f>')
		expect(unrelatedExt).toContain('<xm:sqref>Z2</xm:sqref>')
		expect(unrelatedExt).toContain('<x15:dataValidations count="1">')
		expect(unrelatedExt).toContain('<xm:sqref>Z3</xm:sqref>')
		expect(unrelatedExt).not.toContain('<xm:f>A2&gt;0</xm:f>')
		expect(unrelatedExt).not.toContain('<xm:sqref>A2</xm:sqref>')
		expect(unrelatedExt).not.toContain('<xm:sqref>C2</xm:sqref>')

		expect(worksheetExt).toContain('<xm:f>A2&gt;0</xm:f>')
		expect(worksheetExt).toContain('<xm:sqref>C2</xm:sqref>')
		expect(worksheetExt).toContain('<x14:dataValidations count="1">')
		expect(worksheetExt).toContain('<xm:f>List!$A$2:$A$3</xm:f>')
		expect(worksheetExt).toContain('<xm:sqref>A2</xm:sqref>')
		expect(worksheetExt).not.toContain('<xm:f>A1&gt;0</xm:f>')
		expect(worksheetExt).not.toContain('<xm:sqref>C1</xm:sqref>')
		expect(worksheetExt).not.toContain('List!$A$1:$A$2')
		expect(xml.indexOf('<ext uri="{UNRELATED-X15}">')).toBeLessThan(
			xml.indexOf('<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"'),
		)
	})

	it('writes generated x14 containers into the worksheet extension and preserves extension order', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.x14DataValidations.push(
			{
				index: 2,
				sqref: 'B3',
				type: 'whole',
			},
			{
				index: 1,
				sqref: 'B2',
				type: 'list',
				formula1: 'List!$A$1:$A$2',
				preservedAttributes: { customFlag: '1' },
				preservedChildXml: ['<x14ac:metadata flag="dv"/>'],
			},
		)
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'C2',
			formulas: ['B2>0'],
			type: 'expression',
			preservedRuleAttributes: { activePresent: '1' },
			preservedRuleChildXml: [
				'<x14:extLst><x14:ext uri="{cf-child}"><x14ac:metadata flag="cf"/></x14:ext></x14:extLst>',
			],
		})
		sheet.preservedExtLst = `<extLst xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><ext uri="{UNRELATED}"><x15:keep/></ext><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x15:before/></ext></extLst>`

		const written = writeXlsx(wb)
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)
		const unrelatedExt = xml.match(/<ext uri="\{UNRELATED\}">[\s\S]*?<\/ext>/)?.[0]
		const dataValidationExt = xml.match(
			/<ext uri="\{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF\}"[\s\S]*?<\/ext>/,
		)?.[0]
		const conditionalFormatExt = xml.match(
			/<ext uri="\{78C0D931-6437-407d-A8EE-F0AAD7539E65\}"[\s\S]*?<\/ext>/,
		)?.[0]

		expect(unrelatedExt).toBe('<ext uri="{UNRELATED}"><x15:keep/></ext>')
		expect(dataValidationExt).toContain(
			'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"',
		)
		expect(dataValidationExt).toContain(
			'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"',
		)
		expect(dataValidationExt).toContain('<x15:before/>')
		expect(dataValidationExt).toContain('<x14:dataValidations count="2">')
		expect(dataValidationExt).toContain('customFlag="1"')
		expect(dataValidationExt).toContain('<x14ac:metadata flag="dv"/>')
		expect(conditionalFormatExt).toContain('<x14:conditionalFormattings>')
		expect(conditionalFormatExt).toContain('activePresent="1"')
		expect(conditionalFormatExt).toContain(
			'<x14:extLst><x14:ext uri="{cf-child}"><x14ac:metadata flag="cf"/></x14:ext></x14:extLst>',
		)
		expect(dataValidationExt?.indexOf('<xm:sqref>B2</xm:sqref>')).toBeLessThan(
			dataValidationExt?.indexOf('<xm:sqref>B3</xm:sqref>') ?? -1,
		)
		expect(xml.indexOf('<ext uri="{UNRELATED}">')).toBeLessThan(
			xml.indexOf('<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"'),
		)
		expect(xml.indexOf('<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"')).toBeLessThan(
			xml.indexOf('<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"'),
		)
	})

	it('expands self-closing x14 containers when generated entries are added', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.x14DataValidations.push({
			index: 0,
			sqref: 'A1',
			type: 'whole',
		})
		sheet.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B1',
			formulas: ['A1>0'],
			type: 'expression',
		})
		sheet.preservedExtLst = `<extLst xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:conditionalFormattings/><x14:dataValidations count="0"/></ext></extLst>`

		const written = writeXlsx(wb)
		expectOk(written)
		const entries = unzipSync(written.value)
		const worksheet = entries['xl/worksheets/sheet1.xml']
		expect(worksheet).toBeDefined()
		if (!worksheet) return
		const xml = new TextDecoder().decode(worksheet)

		expect([...xml.matchAll(/<x14:conditionalFormattings\b/g)]).toHaveLength(1)
		expect([...xml.matchAll(/<x14:dataValidations\b/g)]).toHaveLength(1)
		expect(xml).toContain(
			'<x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression"><xm:f>A1&gt;0</xm:f></x14:cfRule><xm:sqref>B1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings>',
		)
		expect(xml).toContain(
			'<x14:dataValidations count="1"><x14:dataValidation type="whole"><xm:sqref>A1</xm:sqref></x14:dataValidation></x14:dataValidations>',
		)
	})

	it('preserves threaded comments XML through read-write round-trip', () => {
		const threadedCommentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="0" id="tc1" dT="2024-01-01T00:00:00.000">
    <text>Comment text</text>
  </threadedComment>
</ThreadedComments>`

		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })

		const capsules: PreservationCapsule[] = [
			{
				partPath: 'xl/threadedComments/threadedComment1.xml',
				contentType: 'application/vnd.ms-excel.threadedcomments+xml',
				relationships: [],
				content: new TextEncoder().encode(threadedCommentXml),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
			},
		]

		const { bytes, result } = roundTrip(wb, capsules)

		const entries = unzipSync(bytes)
		expect(entries['xl/threadedComments/threadedComment1.xml']).toBeDefined()
		const decoded = new TextDecoder().decode(entries['xl/threadedComments/threadedComment1.xml'])
		expect(decoded).toContain('<ThreadedComments')
		expect(decoded).toContain('threadedComment ref="A1"')
		expect(decoded).toContain('<text>Comment text</text>')

		const tcCapsule = result.capsules.find(
			(c) => c.partPath === 'xl/threadedComments/threadedComment1.xml',
		)
		expect(tcCapsule).toBeDefined()
		expect(result.report.features.find((f) => f.feature === 'preservedThreadedComments')).toEqual(
			expect.objectContaining({
				tier: 'preserved',
				count: 1,
				locations: ['xl/threadedComments/threadedComment1.xml'],
			}),
		)
	})

	it('writes edited threaded comment text while preserving thread attributes', () => {
		const threadedCommentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="0" id="tc1" dT="2024-01-01T00:00:00.000">
    <text>Comment text</text>
  </threadedComment>
  <threadedComment ref="A1" personId="1" id="tc2" parentId="tc1" dT="2024-01-02T00:00:00.000" done="1">
    <text>Reviewed</text>
    <mentions><mention mentionpersonId="0" mentionId="m1" startIndex="0" length="3"/></mentions>
  </threadedComment>
</ThreadedComments>`

		const wb = new Workbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		sheet.threadedComments.push(
			{
				ref: 'A1',
				text: 'Comment text',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				id: 'tc1',
				personId: '0',
				dateTime: '2024-01-01T00:00:00.000',
			},
			{
				ref: 'A1',
				text: 'Reviewed',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				id: 'tc2',
				parentId: 'tc1',
				personId: '1',
				dateTime: '2024-01-02T00:00:00.000',
				done: true,
			},
		)
		const applied = applyOperations(wb, [
			{
				op: 'setThreadedComment',
				sheet: 'Sheet1',
				threadedCommentId: 'tc2',
				text: 'Reviewed & approved',
			},
		])
		expectOk(applied)

		const written = writeXlsx(wb, [
			{
				partPath: 'xl/threadedComments/threadedComment1.xml',
				contentType: 'application/vnd.ms-excel.threadedcomments+xml',
				relationships: [],
				content: new TextEncoder().encode(threadedCommentXml),
				anchor: { kind: 'sheet', sheetId: sheet.id, sheetName: sheet.name },
				relType: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
			},
		])
		expectOk(written)

		const entries = unzipSync(written.value)
		const decoded = new TextDecoder().decode(entries['xl/threadedComments/threadedComment1.xml'])
		expect(decoded).toContain('<text>Comment text</text>')
		expect(decoded).toContain('<text>Reviewed &amp; approved</text>')
		expect(decoded).toContain('parentId="tc1"')
		expect(decoded).toContain('personId="1"')
		expect(decoded).toContain('done="1"')
		expect(decoded).toContain('<mentions>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		const comments = reopened.value.workbook.getSheet('Sheet1')?.threadedComments
		expect(comments?.[1]).toMatchObject({
			id: 'tc2',
			parentId: 'tc1',
			personId: '1',
			text: 'Reviewed & approved',
			done: true,
		})
	})

	it('generates threaded comment persons XML from model author identity', () => {
		const source = commentsAndThreadedCommentsWithoutPersonsWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)
		const sheet = opened.value.workbook.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.threadedComments = sheet.threadedComments.map((comment) => ({
			...comment,
			author: comment.personId === '{person-grace}' ? 'Grace Thread' : 'Ada Thread',
		}))

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: [sheet.name],
		})
		expectOk(written)
		const entries = unzipSync(written.value)
		const personsXml = decodeTestXml(entries['xl/persons/person.xml'])
		expect(personsXml).toContain('id="{person-ada}"')
		expect(personsXml).toContain('displayName="Ada Thread"')
		expect(personsXml).toContain('id="{person-grace}"')
		expect(personsXml).toContain('displayName="Grace Thread"')
		const contentTypesXml = decodeTestXml(entries['[Content_Types].xml'])
		expect(contentTypesXml).toContain('ContentType="application/vnd.ms-excel.person+xml"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.getSheet('Sheet1')?.threadedComments).toEqual(
			sheet.threadedComments,
		)
	})

	it('repairs stale threaded comment person display names on dirty writes', () => {
		const source = commentsAndThreadedCommentsWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)
		const sheet = opened.value.workbook.getSheet('Sheet1')
		expect(sheet).toBeDefined()
		if (!sheet) return
		sheet.threadedComments = sheet.threadedComments.map((comment) =>
			comment.personId === '{person-ada}' ? { ...comment, author: 'Ada Updated' } : comment,
		)

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: [sheet.name],
		})
		expectOk(written)
		const entries = unzipSync(written.value)
		const personsXml = decodeTestXml(entries['xl/persons/person.xml'])
		expect(personsXml).toContain('id="{person-ada}" displayName="Ada Updated"')
		expect(personsXml).toContain('id="{person-grace}" displayName="Grace Thread"')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.getSheet('Sheet1')?.threadedComments).toEqual(
			sheet.threadedComments,
		)
	})

	it('patches threaded comment text without changing thread identity or persons sidecar', () => {
		const source = commentsAndThreadedCommentsWorkbook()
		const sourceEntries = unzipSync(source)
		const opened = readXlsx(source)
		expectOk(opened)

		const applied = applyOperations(opened.value.workbook, [
			{
				op: 'setThreadedComment',
				sheet: 'Sheet1',
				threadedCommentId: '{reply-thread}',
				text: 'Thread reply & approved',
			},
		])
		expectOk(applied)

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)

		const entries = unzipSync(written.value)
		expect(entries['xl/comments1.xml']).toEqual(sourceEntries['xl/comments1.xml'])
		expect(entries['xl/drawings/vmlDrawing1.vml']).toEqual(
			sourceEntries['xl/drawings/vmlDrawing1.vml'],
		)
		expect(entries['xl/persons/person.xml']).toEqual(sourceEntries['xl/persons/person.xml'])

		const threadedXml = decodeTestXml(entries['xl/threadedComments/threadedComment1.xml'])
		expect(threadedXml).toContain('<text>Thread root</text>')
		expect(threadedXml).toContain('<text>Thread reply &amp; approved</text>')
		expect(threadedXml).not.toContain('<text>Thread reply</text>')
		expect(threadedXml).toContain(
			'<threadedComment ref="D4" personId="{person-ada}" id="{root-thread}" dT="2024-03-01T10:11:12.000Z">',
		)
		expect(threadedXml).toContain(
			'<threadedComment ref="D4" personId="{person-grace}" id="{reply-thread}" parentId="{root-thread}" dT="2024-03-02T10:11:12.000Z" done="1">',
		)
		expect(threadedXml).toContain(
			'<mentions><mention mentionpersonId="{person-grace}" startIndex="0" length="6"/></mentions>',
		)
		expect(threadedXml).toContain(
			'<extLst><ext uri="{reply-ext}"><futureThreadMetadata preserved="1"/></ext></extLst>',
		)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.getSheet('Sheet1')?.threadedComments).toEqual([
			expect.objectContaining({
				id: '{root-thread}',
				personId: '{person-ada}',
				dateTime: '2024-03-01T10:11:12.000Z',
				text: 'Thread root',
			}),
			expect.objectContaining({
				id: '{reply-thread}',
				parentId: '{root-thread}',
				personId: '{person-grace}',
				dateTime: '2024-03-02T10:11:12.000Z',
				done: true,
				text: 'Thread reply & approved',
			}),
		])
	})

	it('writes shifted threaded comment refs from the current model', () => {
		const source = commentsAndThreadedCommentsWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)
		const applied = applyOperations(opened.value.workbook, [
			{ op: 'insertRows', sheet: 'Sheet1', at: 0, count: 1 },
		])
		expectOk(applied)

		const sheet = opened.value.workbook.getSheet('Sheet1')
		expect(sheet?.threadedComments).toEqual([
			expect.objectContaining({
				ref: 'D5',
				id: '{root-thread}',
				personId: '{person-ada}',
				dateTime: '2024-03-01T10:11:12.000Z',
			}),
			expect.objectContaining({
				ref: 'D5',
				id: '{reply-thread}',
				parentId: '{root-thread}',
				personId: '{person-grace}',
				done: true,
			}),
		])

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)
		const entries = unzipSync(written.value)
		const threadedXml = new TextDecoder().decode(
			entries['xl/threadedComments/threadedComment1.xml'],
		)
		expect(threadedXml).toContain('ref="D5"')
		expect(threadedXml).not.toContain('ref="D4"')
		expect(threadedXml).toContain('parentId="{root-thread}"')
		expect(threadedXml).toContain('done="1"')
		expect(threadedXml).toContain(
			'<mentions><mention mentionpersonId="{person-grace}" startIndex="0" length="6"/></mentions>',
		)
		expect(threadedXml).toContain(
			'<extLst><ext uri="{reply-ext}"><futureThreadMetadata preserved="1"/></ext></extLst>',
		)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.getSheet('Sheet1')?.threadedComments).toEqual(
			sheet?.threadedComments,
		)
	})

	it('writes cloned threaded comments created by comment copy operations', () => {
		const source = commentsAndThreadedCommentsWorkbook()
		const opened = readXlsx(source)
		expectOk(opened)
		const applied = applyOperations(opened.value.workbook, [
			{ op: 'copyRange', sheet: 'Sheet1', source: 'D4', target: 'E4', mode: 'comments' },
		])
		expectOk(applied)

		const sheet = opened.value.workbook.getSheet('Sheet1')
		expect(sheet?.threadedComments).toEqual([
			expect.objectContaining({ ref: 'D4', id: '{root-thread}' }),
			expect.objectContaining({
				ref: 'D4',
				id: '{reply-thread}',
				parentId: '{root-thread}',
			}),
			expect.objectContaining({ ref: 'E4', id: '{root-thread}-copy' }),
			expect.objectContaining({
				ref: 'E4',
				id: '{reply-thread}-copy',
				parentId: '{root-thread}-copy',
			}),
		])

		const written = writeXlsx(opened.value.workbook, opened.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)
		const entries = unzipSync(written.value)
		const threadedXml = new TextDecoder().decode(
			entries['xl/threadedComments/threadedComment1.xml'],
		)
		expect(threadedXml).toContain('id="{root-thread}-copy"')
		expect(threadedXml).toContain('parentId="{root-thread}-copy"')
		expect(threadedXml).toContain('ref="E4"')
		expect(threadedXml).toMatch(
			/id="\{root-thread\}-copy"[\s\S]*<mentions><mention mentionpersonId="\{person-grace\}" startIndex="0" length="6"\/><\/mentions>/,
		)
		expect(threadedXml).toMatch(
			/id="\{reply-thread\}-copy"[\s\S]*<extLst><ext uri="\{reply-ext\}"><futureThreadMetadata preserved="1"\/><\/ext><\/extLst>/,
		)

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.getSheet('Sheet1')?.threadedComments).toEqual(
			sheet?.threadedComments,
		)
	})

	it('produces a valid ZIP file', () => {
		const wb = new Workbook()
		wb.addSheet('Empty')

		const written = writeXlsx(wb)
		expectOk(written)

		const entries = unzipSync(written.value)
		expect(entries['[Content_Types].xml']).toBeDefined()
		expect(entries['_rels/.rels']).toBeDefined()
		expect(entries['xl/workbook.xml']).toBeDefined()
		expect(entries['xl/_rels/workbook.xml.rels']).toBeDefined()
		expect(entries['xl/styles.xml']).toBeDefined()
		expect(entries['xl/worksheets/sheet1.xml']).toBeDefined()
	})

	it('streaming write produces valid XLSX that round-trips (50K rows)', async () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		const rows = 50_000
		const cols = 5
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				sheet.cells.set(r, c, {
					value: numberValue(r * cols + c + 1),
					formula: null,
					styleId: S0,
				})
			}
		}

		const written = await writeXlsxStreaming(wb, undefined, { streaming: true })
		expectOk(written)

		const entries = unzipSync(written.value)
		expect(entries['xl/worksheets/sheet1.xml']).toBeDefined()
		const sheetXml = new TextDecoder().decode(entries['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain('<sheetData>')
		expect(sheetXml).toContain(`<row r="${rows}"`)

		const read = readXlsx(written.value)
		expectOk(read)
		const readSheet = read.value.workbook.sheets[0]
		expect(readSheet?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 1 })
		expect(readSheet?.cells.get(rows - 1, cols - 1)?.value).toEqual({
			kind: 'number',
			value: rows * cols,
		})
	}, 90_000)

	it('streaming write plan emits worksheet XML through streamingBuild', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		for (let r = 0; r < 10_000; r++) {
			sheet.cells.set(r, 0, {
				value: numberValue(r + 1),
				formula: null,
				styleId: S0,
			})
		}

		const plan = planWriteXlsx(wb, undefined, { streaming: true })
		expectOk(plan)
		const sheetPart = plan.value.descriptors.find(
			(entry) => entry.path === 'xl/worksheets/sheet1.xml',
		)
		expect(sheetPart?.streamingBuild).toBeDefined()
		expect(plan.value.parts.has('xl/worksheets/sheet1.xml')).toBe(false)
	})

	it('emits a stable structure fingerprint for synthetic workbooks', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Report')
		const percentId = wb.styles.register({
			font: { bold: true },
			numberFormat: '0.0%',
		})
		sheet.cells.set(0, 0, { value: stringValue('Revenue'), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(0.25), formula: null, styleId: percentId })
		sheet.cells.set(1, 1, { value: numberValue(0.5), formula: 'B1*2', styleId: percentId })
		sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } })
		wb.definedNames.set('RevenuePct', 'Report!$B$1')
		wb.calcSettings = {
			...wb.calcSettings,
			calcMode: 'manual',
			fullCalcOnLoad: true,
		}

		const written = writeXlsx(wb)
		expectOk(written)

		const fingerprint = fingerprintXlsx(written.value)
		expect(fingerprint.partPaths).toEqual([
			'[Content_Types].xml',
			'_rels/.rels',
			'docProps/app.xml',
			'docProps/core.xml',
			'xl/_rels/workbook.xml.rels',
			'xl/sharedStrings.xml',
			'xl/styles.xml',
			'xl/workbook.xml',
			'xl/worksheets/sheet1.xml',
		])
		expect(fingerprint.workbook?.tagCounts).toMatchObject({
			workbook: 1,
			sheets: 1,
			sheet: 1,
			definedNames: 1,
			definedName: 1,
			calcPr: 1,
		})
		expect(fingerprint.workbookRels?.tagCounts).toMatchObject({
			Relationships: 1,
			Relationship: 3,
		})
		expect(fingerprint.styles?.tagCounts).toMatchObject({
			styleSheet: 1,
			fonts: 1,
			fills: 1,
			borders: 1,
			numFmts: 1,
			numFmt: 1,
			cellXfs: 1,
			xf: 2,
		})
		expect(fingerprint.sheets).toHaveLength(1)
		expect(fingerprint.sheets[0]).toEqual(
			expect.objectContaining({
				path: 'xl/worksheets/sheet1.xml',
				xml: expect.objectContaining({
					tagCounts: expect.objectContaining({
						worksheet: 1,
						sheetData: 1,
						row: 2,
						c: 3,
						f: 1,
						mergeCells: 1,
						mergeCell: 1,
					}),
				}),
			}),
		)
	})

	it('does not request full recalculation on load for clean generated calc settings', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Calc')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(2), formula: 'A1*2', styleId: S0 })
		wb.calcSettings = {
			...wb.calcSettings,
			calcMode: 'auto',
			fullCalcOnLoad: false,
			calcCompleted: true,
			calcOnSave: true,
			forceFullCalc: false,
		}

		const written = writeXlsx(wb)
		expectOk(written)

		const zip = unzipSync(written.value)
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		expect(workbookXml).not.toContain('fullCalcOnLoad="1"')
		expect(workbookXml).toContain('calcCompleted="1"')
		expect(workbookXml).toContain('calcOnSave="1"')
		expect(workbookXml).toContain('forceFullCalc="0"')
	})

	it('drops calcChain and marks workbook stale after formula-affecting edits', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcMode="manual" fullCalcOnLoad="0" calcCompleted="1" calcOnSave="0" calcId="191029"/>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><f>A1*2</f><v>2</v></c>
    </row>
  </sheetData>
</worksheet>`,
			'xl/calcChain.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)

		const applied = applyOperations(source.value.workbook, [
			{ op: 'setCells', sheet: 'Calc', updates: [{ ref: 'A1', value: 3 }] },
		])
		expectOk(applied)

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Calc'],
			calcStateDirty: true,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		expect(zip['xl/calcChain.xml']).toBeUndefined()
		expect(workbookRels).not.toContain('relationships/calcChain')
		expect(contentTypes).not.toContain('calcChain+xml')
		expect(workbookXml).toContain('fullCalcOnLoad="1"')
		expect(workbookXml).toContain('calcCompleted="0"')
		expect(workbookXml).toContain('forceFullCalc="1"')
	})

	it('preserves calcChain for value edits that do not change formula topology', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcMode="manual" fullCalcOnLoad="0" calcCompleted="1" calcOnSave="0" calcId="191029"/>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><f>A1*2</f><v>2</v></c>
    </row>
  </sheetData>
</worksheet>`,
			'xl/calcChain.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const applied = applyOperations(source.value.workbook, [
			{ op: 'setCells', sheet: 'Calc', updates: [{ ref: 'A1', value: 3 }] },
		])
		expectOk(applied)

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Calc'],
			calcStateDirty: true,
			calcChainDirty: false,
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		const workbookRels = new TextDecoder().decode(
			zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array(),
		)
		const contentTypes = new TextDecoder().decode(zip['[Content_Types].xml'] ?? new Uint8Array())
		expect(zip['xl/calcChain.xml']).toBeDefined()
		expect(workbookRels).toContain('relationships/calcChain')
		expect(contentTypes).toContain('calcChain+xml')
		expect(workbookXml).toContain('fullCalcOnLoad="1"')
		expect(workbookXml).toContain('calcCompleted="0"')
	})

	it('preserves calcPr fidelity on clean round-trip', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Calc" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcMode="manual" fullCalcOnLoad="1" calcCompleted="0" calcOnSave="0" forceFullCalc="1" calcId="191029"/>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		})

		const source = readXlsx(sourceBytes)
		expectOk(source)

		const written = writeXlsx(source.value.workbook, source.value.capsules)
		expectOk(written)

		const zip = unzipSync(written.value)
		const workbookXml = new TextDecoder().decode(zip['xl/workbook.xml'] ?? new Uint8Array())
		expect(workbookXml).toContain('calcMode="manual"')
		expect(workbookXml).toContain('fullCalcOnLoad="1"')
		expect(workbookXml).toContain('calcCompleted="0"')
		expect(workbookXml).toContain('calcOnSave="0"')
		expect(workbookXml).toContain('forceFullCalc="1"')
		expect(workbookXml).toContain('calcId="191029"')
	})

	it('preserves sheetView attributes (zoomScale, zoomScaleNormal, zoomScaleSheetLayoutView, showGridLines, showFormulas, rightToLeft, tabSelected, view, topLeftCell) on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('View')
		sheet.cells.set(0, 0, { value: stringValue('hi'), formula: null, styleId: S0 })
		sheet.sheetView = {
			zoomScale: 125,
			zoomScaleNormal: 100,
			zoomScaleSheetLayoutView: 214,
			showGridLines: false,
			showFormulas: true,
			rightToLeft: true,
			tabSelected: true,
			view: 'pageLayout',
			topLeftCell: 'E1',
		}

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		if (!s) throw new Error('Expected round-tripped workbook to contain a sheet')
		expect(s.sheetView).toEqual({
			zoomScale: 125,
			zoomScaleNormal: 100,
			zoomScaleSheetLayoutView: 214,
			showGridLines: false,
			showFormulas: true,
			rightToLeft: true,
			tabSelected: true,
			view: 'pageLayout',
			topLeftCell: 'E1',
		})
	})

	it('preserves plain sheetView presence on regenerated sheets', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="View" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['View'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<sheetView workbookViewId="0">')
		expect(sheetXml).toContain('<selection activeCell="A1" sqref="A1"/>')
	})

	it('preserves multiple sheetView selections on regenerated sheets', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('View')
		sheet.cells.set(0, 0, { value: stringValue('hi'), formula: null, styleId: S0 })
		sheet.frozenRows = 1
		sheet.frozenCols = 1
		sheet.preservedPaneAttributes = {
			xSplit: '1',
			ySplit: '1',
			state: 'frozen',
			topLeftCell: 'B2',
			activePane: 'bottomRight',
		}
		sheet.preservedSheetViewSelections = [
			{ pane: 'topRight', activeCell: 'B1', sqref: 'B1' },
			{ pane: 'bottomLeft', activeCell: 'A2', sqref: 'A2' },
			{ pane: 'bottomRight', activeCell: 'B2', activeCellId: '0', sqref: 'B2:C4' },
		]

		const { bytes, result } = roundTrip(wb)
		const zip = unzipSync(bytes)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain(
			'<pane xSplit="1" ySplit="1" state="frozen" topLeftCell="B2" activePane="bottomRight"/>',
		)
		expect(sheetXml).toContain('<selection pane="topRight" activeCell="B1" sqref="B1"/>')
		expect(sheetXml).toContain('<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>')
		expect(sheetXml).toContain(
			'<selection pane="bottomRight" activeCell="B2" activeCellId="0" sqref="B2:C4"/>',
		)
		const s = result.workbook.sheets[0]
		expect(s?.preservedSheetViewSelections).toEqual([
			{ pane: 'topRight', activeCell: 'B1', sqref: 'B1' },
			{ pane: 'bottomLeft', activeCell: 'A2', sqref: 'A2' },
			{ pane: 'bottomRight', activeCell: 'B2', activeCellId: '0', sqref: 'B2:C4' },
		])
	})

	it('preserves blank physical cells on regenerated sheets without hydrating semantic values', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Blank" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C1"/>
  <sheetData><row r="1"><c r="A1"/><c r="B1" s="1"/><c r="C1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		expect(sheet.cells.get(0, 0)).toBeUndefined()
		expect(sheet.cells.get(0, 1)).toBeUndefined()
		sheet.cells.set(0, 2, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Blank'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<dimension ref="A1:C1"/>')
		expect(sheetXml).toContain('<c r="A1"/>')
		expect(sheetXml).toContain('<c r="B1" s="1"/>')
		expect(sheetXml).toContain('<c r="C1"><v>2</v></c>')
	})

	it('preserves raw sheetView attributes on regenerated sheets', () => {
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="View" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0" showGridLines="true" showZeros="false" showRowColHeaders="true" colorId="64" topLeftCell="C3"/></sheetViews>
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['View'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('showGridLines="true"')
		expect(sheetXml).toContain('showZeros="false"')
		expect(sheetXml).toContain('showRowColHeaders="true"')
		expect(sheetXml).toContain('colorId="64"')
		expect(sheetXml).toContain('topLeftCell="C3"')
	})

	it('preserves sheetPr tabColor and sheetFormatPr on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Colored')
		sheet.cells.set(0, 0, { value: stringValue('hi'), formula: null, styleId: S0 })
		sheet.codeName = 'SheetCode1'
		sheet.filterMode = true
		sheet.enableFormatConditionsCalculation = false
		sheet.tabColor = { rgb: 'FF0000FF', theme: 4, tint: -0.25 }
		sheet.pageSetupPr = { fitToPage: true, autoPageBreaks: false }
		sheet.sheetFormatPr = {
			baseColWidth: 12,
			defaultRowHeight: 14.5,
			defaultColWidth: 10.0,
			zeroHeight: true,
			dyDescent: 0.25,
		}

		const { result, bytes } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		if (!s) throw new Error('Expected round-tripped workbook to contain a sheet')
		expect(s.codeName).toBe('SheetCode1')
		expect(s.filterMode).toBe(true)
		expect(s.enableFormatConditionsCalculation).toBe(false)
		expect(s.tabColor).toEqual({ rgb: 'FF0000FF', theme: 4, tint: -0.25 })
		expect(s.pageSetupPr).toEqual({ fitToPage: true, autoPageBreaks: false })
		expect(s.sheetFormatPr).toEqual({
			baseColWidth: 12,
			defaultRowHeight: 14.5,
			defaultColWidth: 10,
			zeroHeight: true,
			dyDescent: 0.25,
		})
		const sheetXml = new TextDecoder().decode(
			unzipSync(bytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array(),
		)
		expect(sheetXml).toContain(
			'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"',
		)
		expect(sheetXml).toContain(
			'<sheetPr codeName="SheetCode1" filterMode="1" enableFormatConditionsCalculation="0">',
		)
		expect(sheetXml).toContain('<pageSetUpPr fitToPage="1" autoPageBreaks="0"/>')
		expect(sheetXml).toContain(
			'<sheetFormatPr baseColWidth="12" defaultRowHeight="14.5" defaultColWidth="10" zeroHeight="1" x14ac:dyDescent="0.25"/>',
		)
	})

	it('preserves pageSetup printerSettings relationship on regenerated sheets', () => {
		const printerSettings = 'printer-settings'
		const sourceBytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Print" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <pageSetup orientation="portrait" horizontalDpi="300" verticalDpi="300" r:id="rIdPrinter"/>
</worksheet>`,
			'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPrinter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>
</Relationships>`,
			'xl/printerSettings/printerSettings1.bin': printerSettings,
		})
		const source = readXlsx(sourceBytes)
		expectOk(source)
		const sheet = source.value.workbook.sheets[0]
		if (!sheet) throw new Error('Expected source workbook to contain a sheet')
		expect(sheet.pageSetup).toEqual({
			orientation: 'portrait',
			horizontalDpi: 300,
			verticalDpi: 300,
			printerSettingsRelId: 'rIdPrinter',
		})
		sheet.cells.set(0, 0, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(source.value.workbook, source.value.capsules, {
			dirtySheetNames: ['Print'],
		})
		expectOk(written)

		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		const sheetRelsXml = new TextDecoder().decode(
			zip['xl/worksheets/_rels/sheet1.xml.rels'] ?? new Uint8Array(),
		)
		expect(sheetXml).toContain(
			'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
		)
		expect(sheetXml).toContain(
			'<pageSetup orientation="portrait" horizontalDpi="300" verticalDpi="300" r:id="rIdPrinter"/>',
		)
		expect(sheetRelsXml).toContain('Id="rIdPrinter"')
		expect(sheetRelsXml).toContain(
			'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings"',
		)
		expect(sheetRelsXml).toContain('Target="../printerSettings/printerSettings1.bin"')
		expect(new TextDecoder().decode(zip['xl/printerSettings/printerSettings1.bin'])).toBe(
			printerSettings,
		)
	})

	it('preserves row and column outline metadata on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Outline')
		sheet.cells.set(0, 0, { value: stringValue('hi'), formula: null, styleId: S0 })
		sheet.outlinePr = { summaryBelow: false, summaryRight: false, showOutlineSymbols: true }
		sheet.sheetFormatPr = { outlineLevelRow: 1, outlineLevelCol: 2 }
		sheet.rowDefs.set(1, { hidden: true, outlineLevel: 1 })
		sheet.rowDefs.set(2, { collapsed: true })
		sheet.colDefs.push({ min: 0, max: 0, hidden: true, outlineLevel: 2 })
		sheet.colDefs.push({ min: 1, max: 1, collapsed: true })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		if (!s) throw new Error('Expected round-tripped workbook to contain a sheet')
		expect(s.outlinePr).toEqual({
			summaryBelow: false,
			summaryRight: false,
			showOutlineSymbols: true,
		})
		expect(s.sheetFormatPr).toEqual({ outlineLevelRow: 1, outlineLevelCol: 2 })
		expect(s.rowDefs.get(1)).toEqual({ hidden: true, outlineLevel: 1 })
		expect(s.rowDefs.get(2)).toEqual({ collapsed: true })
		expect(s.colDefs).toContainEqual({ min: 0, max: 0, hidden: true, outlineLevel: 2 })
		expect(s.colDefs).toContainEqual({ min: 1, max: 1, collapsed: true })
	})

	it('preserves extLst blocks through round-trip when present', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('ExtLst')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		sheet.preservedExtLst = '<extLst><ext uri="{test}"><x14:sparklines/></ext></extLst>'

		const written = writeXlsx(wb)
		if (!written.ok) throw new Error(written.error.message)
		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain('<extLst>')
		expect(sheetXml).toContain('x14:sparklines')

		const read = readXlsx(written.value)
		if (!read.ok) throw new Error(read.error.message)
		expect(read.value.workbook.sheets[0]?.preservedExtLst).toContain('<extLst>')
	})

	it('preserves sparkline extLst XML through read-write round-trip', () => {
		const sparklineExtLst = `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
    <x14:sparklineGroup>
      <x14:f>Sheet1!A1:A5</x14:f>
      <x14:sparklines>
        <x14:sparkline f="Sheet1!B1:B5"/>
      </x14:sparklines>
    </x14:sparklineGroup>
  </x14:sparklineGroups>
</ext></extLst>`

		const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
${sparklineExtLst}
</worksheet>`

		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': sheetXml,
		})

		const read = readXlsx(bytes)
		if (!read.ok) throw new Error(read.error.message)
		const preserved = read.value.workbook.sheets[0]?.preservedExtLst
		expect(preserved).toBeDefined()
		expect(preserved).toContain('x14:sparklineGroups')
		expect(preserved).toContain('x14:sparkline')

		const written = writeXlsx(read.value.workbook, read.value.capsules)
		if (!written.ok) throw new Error(written.error.message)
		const zip = unzipSync(written.value)
		const writtenSheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
		expect(writtenSheetXml).toContain('<extLst>')
		expect(writtenSheetXml).toContain('x14:sparklineGroups')
		expect(writtenSheetXml).toContain('Sheet1!A1:A5')
	})

	it('writes edited sparkline group source and display flags into preserved extension XML', () => {
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetData/>
  <extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
    <x14:sparklineGroups>
      <x14:sparklineGroup type='line' markers='1' high='1' displayXAxis='1'>
        <x14:sparklines><x14:sparkline><xm:f>Data!B2:B4</xm:f><xm:sqref>D2:D4</xm:sqref></x14:sparkline></x14:sparklines>
      </x14:sparklineGroup>
    </x14:sparklineGroups>
  </ext></extLst>
</worksheet>`,
		})
		const read = readXlsx(bytes)
		expectOk(read)
		const applied = applyOperations(read.value.workbook, [
			{
				op: 'setSparklineGroup',
				sheet: 'Data',
				groupIndex: 0,
				range: 'Data!C2:C4',
				locationRange: 'E2:E4',
				type: 'column',
				markers: false,
				highPoint: false,
				displayXAxis: false,
			},
		])
		expectOk(applied)

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('type="column"')
		expect(sheetXml).toContain('markers="0"')
		expect(sheetXml).toContain('high="0"')
		expect(sheetXml).toContain('displayXAxis="0"')
		expect(sheetXml).not.toContain("markers='1' markers=")
		expect(sheetXml).not.toContain("displayXAxis='1' displayXAxis=")
		expect(sheetXml).toContain('<xm:f>Data!C2:C4</xm:f>')
		expect(sheetXml).toContain('<xm:sqref>E2:E4</xm:sqref>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.sparklineGroups[0]).toMatchObject({
			type: 'column',
			markers: false,
			highPoint: false,
			displayXAxis: false,
			range: 'Data!C2:C4',
			locationRange: 'E2:E4',
		})
	})

	it('preserves advanced filters in custom sheet views on clean and dirty writes', () => {
		const read = readXlsx(advancedFilterSparklineWorkbook())
		expectOk(read)
		expect(read.value.workbook.sheets[0]?.preservedCustomSheetViews).toContain('<customSheetViews>')

		const cleanWritten = writeXlsx(read.value.workbook, read.value.capsules)
		expectOk(cleanWritten)
		const cleanZip = unzipSync(cleanWritten.value)
		const cleanSheetXml = new TextDecoder().decode(
			cleanZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array(),
		)
		expect(cleanSheetXml).toContain('<customSheetViews>')
		expect(cleanSheetXml).toContain('name="WestOnly"')
		const cleanReopened = readXlsx(cleanWritten.value)
		expectOk(cleanReopened)
		expect(cleanReopened.value.workbook.sheets[0]?.advancedFilters[0]).toMatchObject({
			viewName: 'WestOnly',
			guid: '{11111111-1111-1111-1111-111111111111}',
			ref: 'A1:C20',
			filterColumnCount: 1,
			sortConditionCount: 1,
		})

		const applied = applyOperations(read.value.workbook, [
			{
				op: 'setAdvancedFilter',
				sheet: 'Data',
				filterIndex: 0,
				range: 'A1:D20',
				column: 0,
				values: ['East', 'North'],
				sortRef: 'A2:D20',
				sortBy: 'B2:B20',
				descending: false,
			},
		])
		expectOk(applied)

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: applied.value.sheetsModified,
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		expect(sheetXml).toContain('<customSheetViews>')
		expect(sheetXml).toContain('name="WestOnly"')
		expect(sheetXml).toContain('guid="{11111111-1111-1111-1111-111111111111}"')
		expect(sheetXml).toContain('<autoFilter ref="A1:D20">')
		expect(sheetXml).toContain('<filter val="East"/>')
		expect(sheetXml).toContain('<filter val="North"/>')
		expect(sheetXml).toContain('<sortState ref="A2:D20">')
		expect(sheetXml).toContain('<sortCondition ref="B2:B20" descending="0"/>')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.advancedFilters[0]).toMatchObject({
			viewName: 'WestOnly',
			guid: '{11111111-1111-1111-1111-111111111111}',
			ref: 'A1:D20',
			filterColumnCount: 1,
			sortConditionCount: 1,
			autoFilter: {
				ref: 'A1:D20',
				columns: [{ colId: 0, kind: 'filters', values: ['East', 'North'] }],
				sortState: {
					ref: 'A2:D20',
					conditions: [{ ref: 'B2:B20', descending: false }],
				},
			},
		})
	})

	it('removes cleared advanced filters from preserved custom sheet views on dirty writes', () => {
		const read = readXlsx(advancedFilterSparklineWorkbook())
		expectOk(read)
		const sheet = read.value.workbook.sheets[0]
		if (!sheet) throw new Error('expected sheet')
		const filter = sheet.advancedFilters[0]
		if (!filter) throw new Error('expected advanced filter')
		const { autoFilter: _autoFilter, ref: _ref, ...clearedFilter } = filter
		sheet.advancedFilters[0] = {
			...clearedFilter,
			filterColumnCount: 0,
			sortConditionCount: 0,
		}

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: [sheet.name],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
		const customSheetViewsXml = sheetXml.match(/<customSheetViews>[\s\S]*<\/customSheetViews>/)?.[0]
		expect(customSheetViewsXml).toContain('name="WestOnly"')
		expect(customSheetViewsXml).not.toContain('<autoFilter')

		const reopened = readXlsx(written.value)
		expectOk(reopened)
		expect(reopened.value.workbook.sheets[0]?.advancedFilters).toHaveLength(0)
	})

	it('keeps worksheet and nested rule extLst payloads in their owning scopes', () => {
		const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData>
  <x:conditionalFormatting sqref="A1">
    <x:cfRule type="dataBar" priority="1"><x:extLst><x:ext uri="{nested}"/></x:extLst></x:cfRule>
  </x:conditionalFormatting>
  <x:extLst><x:ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{worksheet}"><x14:conditionalFormattings/></x:ext></x:extLst>
</x:worksheet>`
		const bytes = makeXlsx({
			'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
			'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
			'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
			'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
			'xl/worksheets/sheet1.xml': sheetXml,
		})

		const read = readXlsx(bytes)
		expectOk(read)
		const preserved = read.value.workbook.sheets[0]?.preservedExtLst
		expect(preserved).toContain('uri="{worksheet}"')
		expect(preserved).toContain('x14:conditionalFormattings')
		expect(preserved).not.toContain('uri="{nested}"')
		expect(
			read.value.workbook.sheets[0]?.conditionalFormats[0]?.rules[0]?.preservedRuleChildXml,
		).toEqual(['<extLst><ext uri="{nested}"/></extLst>'])

		const written = writeXlsx(read.value.workbook, read.value.capsules, {
			dirtySheetNames: ['Sheet1'],
		})
		expectOk(written)
		const zip = unzipSync(written.value)
		const writtenSheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
		expect(writtenSheetXml).toContain('<extLst>')
		expect(writtenSheetXml).toContain('x14:conditionalFormattings')
		expect(writtenSheetXml).toContain(
			'<cfRule type="dataBar" priority="1"><extLst><ext uri="{nested}"/></extLst></cfRule>',
		)
	})

	it('emits dimension element for non-empty sheets', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Dim')
		sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		sheet.cells.set(9, 4, { value: numberValue(2), formula: null, styleId: S0 })

		const written = writeXlsx(wb)
		if (!written.ok) throw new Error(written.error.message)
		const zip = unzipSync(written.value)
		const sheetXml = new TextDecoder().decode(zip['xl/worksheets/sheet1.xml'])
		expect(sheetXml).toContain('<dimension ref="A1:E10"/>')
	})

	it('preserves full ignoredError attributes on round-trip', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Errors')
		sheet.cells.set(0, 0, { value: stringValue('123'), formula: null, styleId: S0 })
		sheet.ignoredErrors.push({
			sqref: 'A1:B2',
			numberStoredAsText: true,
			formula: true,
			evalError: true,
		})

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		if (!s) throw new Error('Expected round-tripped workbook to contain a sheet')
		expect(s.ignoredErrors).toEqual([
			{
				sqref: 'A1:B2',
				numberStoredAsText: true,
				formula: true,
				evalError: true,
			},
		])
	})

	describe('round-trip fidelity', () => {
		it('round-trips mixed data types (numbers, strings, booleans, dates, errors)', () => {
			const wb = new Workbook()
			const dateFmtId = wb.styles.register({ numberFormat: 'yyyy-mm-dd' })
			const sheet = wb.addSheet('Mixed')
			sheet.cells.set(0, 0, { value: numberValue(42), formula: null, styleId: S0 })
			sheet.cells.set(0, 1, { value: stringValue('text'), formula: null, styleId: S0 })
			sheet.cells.set(0, 2, { value: booleanValue(true), formula: null, styleId: S0 })
			sheet.cells.set(0, 3, { value: booleanValue(false), formula: null, styleId: S0 })
			sheet.cells.set(0, 4, {
				value: { kind: 'date', serial: 45292 },
				formula: null,
				styleId: dateFmtId,
			})
			sheet.cells.set(0, 5, {
				value: { kind: 'error', value: '#DIV/0!' },
				formula: null,
				styleId: S0,
			})

			const { result } = roundTrip(wb)
			const s = result.workbook.sheets[0]
			expect(s?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 42 })
			expect(s?.cells.get(0, 1)?.value).toEqual({ kind: 'string', value: 'text' })
			expect(s?.cells.get(0, 2)?.value).toEqual({ kind: 'boolean', value: true })
			expect(s?.cells.get(0, 3)?.value).toEqual({ kind: 'boolean', value: false })
			expect(s?.cells.get(0, 4)?.value).toEqual({ kind: 'date', serial: 45292 })
			expect(s?.cells.get(0, 5)?.value).toEqual({ kind: 'error', value: '#DIV/0!' })
		})

		it('round-trips formulas with various complexity', () => {
			const wb = new Workbook()
			const sheet = wb.addSheet('Formulas')
			sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: S0 })
			sheet.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: S0 })
			sheet.cells.set(1, 0, { value: numberValue(30), formula: 'A1+B1', styleId: S0 })
			sheet.cells.set(1, 1, { value: numberValue(600), formula: 'A1*B1*2', styleId: S0 })
			sheet.cells.set(2, 0, { value: numberValue(3), formula: 'SUM(A1:B2)', styleId: S0 })
			sheet.cells.set(2, 1, {
				value: numberValue(2),
				formula: 'IF(A1>B1,"big","small")',
				styleId: S0,
			})

			const { result } = roundTrip(wb)
			const s = result.workbook.sheets[0]
			expect(s?.cells.get(1, 0)?.formula).toBe('A1+B1')
			expect(s?.cells.get(1, 1)?.formula).toBe('A1*B1*2')
			expect(s?.cells.get(2, 0)?.formula).toBe('SUM(A1:B2)')
			expect(s?.cells.get(2, 1)?.formula).toBe('IF(A1>B1,"big","small")')
		})

		it('round-trips styled cells (bold, italic, colors, number formats)', () => {
			const wb = new Workbook()
			const boldId = wb.styles.register({ font: { bold: true } })
			const italicId = wb.styles.register({ font: { italic: true } })
			const colorId = wb.styles.register({
				font: { color: { kind: 'rgb', rgb: 'FF0000FF' } },
			})
			const pctId = wb.styles.register({ numberFormat: '0.00%' })
			const sheet = wb.addSheet('Styled')
			sheet.cells.set(0, 0, { value: stringValue('Bold'), formula: null, styleId: boldId })
			sheet.cells.set(0, 1, { value: stringValue('Italic'), formula: null, styleId: italicId })
			sheet.cells.set(0, 2, { value: stringValue('Blue'), formula: null, styleId: colorId })
			sheet.cells.set(0, 3, { value: numberValue(0.75), formula: null, styleId: pctId })

			const { result } = roundTrip(wb)
			const s = result.workbook.sheets[0]
			const styles = result.workbook.styles
			expect(styles.get(s?.cells.get(0, 0)?.styleId ?? S0)?.font?.bold).toBe(true)
			expect(styles.get(s?.cells.get(0, 1)?.styleId ?? S0)?.font?.italic).toBe(true)
			expect(styles.get(s?.cells.get(0, 2)?.styleId ?? S0)?.font?.color).toEqual({
				kind: 'rgb',
				rgb: 'FF0000FF',
			})
			expect(styles.get(s?.cells.get(0, 3)?.styleId ?? S0)?.numberFormat).toBe('0.00%')
		})

		it('round-trips merged cells', () => {
			const wb = new Workbook()
			const sheet = wb.addSheet('Merged')
			sheet.cells.set(0, 0, { value: stringValue('Title'), formula: null, styleId: S0 })
			sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 0, col: 2 } })
			sheet.merges.push({ start: { row: 2, col: 0 }, end: { row: 4, col: 1 } })

			const { result } = roundTrip(wb)
			const s = result.workbook.sheets[0]
			expect(s?.merges).toHaveLength(2)
			expect(s?.merges[0]).toEqual({ start: { row: 0, col: 0 }, end: { row: 0, col: 2 } })
			expect(s?.merges[1]).toEqual({ start: { row: 2, col: 0 }, end: { row: 4, col: 1 } })
		})

		it('round-trips multiple sheets with cross-sheet references', () => {
			const wb = new Workbook()
			const data = wb.addSheet('Data')
			const summary = wb.addSheet('Summary')
			data.cells.set(0, 0, { value: numberValue(100), formula: null, styleId: S0 })
			data.cells.set(1, 0, { value: numberValue(200), formula: null, styleId: S0 })
			summary.cells.set(0, 0, {
				value: numberValue(300),
				formula: 'SUM(Data!A1:A2)',
				styleId: S0,
			})

			const { result } = roundTrip(wb)
			expect(result.workbook.sheets).toHaveLength(2)
			expect(result.workbook.sheets[0]?.name).toBe('Data')
			expect(result.workbook.sheets[1]?.name).toBe('Summary')
			expect(result.workbook.sheets[1]?.cells.get(0, 0)?.formula).toBe('SUM(Data!A1:A2)')
			expect(result.workbook.sheets[1]?.cells.get(0, 0)?.value).toEqual({
				kind: 'number',
				value: 300,
			})
		})
	})
})
