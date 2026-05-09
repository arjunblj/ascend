import { describe, expect, test } from 'bun:test'
import { makeXlsx } from '../../test/helpers.ts'
import { readXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

describe('threaded comment inventory', () => {
	test('parses threaded comments and person metadata while preserving source parts', () => {
		const result = readXlsx(threadedCommentWorkbook())
		expectOk(result)

		const sheet = result.value.workbook.sheets[0]
		expect(sheet?.threadedComments).toEqual([
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
		expect(
			result.value.report.features.find(
				(feature) => feature.feature === 'preservedThreadedComments',
			),
		).toMatchObject({
			tier: 'preserved',
			locations: ['xl/threadedComments/threadedComment1.xml'],
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
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
