import { describe, expect, test } from 'bun:test'
import { zipSync } from 'fflate'
import { readXlsx } from './index.ts'

const CONTENT_TYPES = `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`

const ROOT_RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const WORKBOOK_RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`

const WORKBOOK_XML = `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`

const SHEET_XML = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>`

function makeMinimalXlsx(): Uint8Array {
	return zipSync({
		'[Content_Types].xml': str(CONTENT_TYPES),
		'_rels/.rels': str(ROOT_RELS),
		'xl/_rels/workbook.xml.rels': str(WORKBOOK_RELS),
		'xl/workbook.xml': str(WORKBOOK_XML),
		'xl/worksheets/sheet1.xml': str(SHEET_XML),
	})
}

function str(value: string): Uint8Array {
	return new TextEncoder().encode(value)
}

function rng(seed: number): () => number {
	let state = seed >>> 0
	return () => {
		state = (Math.imul(1664525, state) + 1013904223) >>> 0
		return state / 0x1_0000_0000
	}
}

function mutate(base: Uint8Array, random: () => number): Uint8Array {
	const bytes = new Uint8Array(base)
	const mode = Math.floor(random() * 4)
	if (mode === 0) {
		const flips = 1 + Math.floor(random() * 8)
		for (let i = 0; i < flips; i++) {
			const idx = Math.floor(random() * bytes.length)
			bytes[idx] = bytes[idx] ^ (1 + Math.floor(random() * 255))
		}
		return bytes
	}
	if (mode === 1) {
		const cut = Math.max(1, Math.floor(random() * bytes.length))
		return bytes.subarray(0, cut)
	}
	if (mode === 2) {
		const duplicated = new Uint8Array(bytes.length + 32)
		duplicated.set(bytes, 0)
		for (let i = bytes.length; i < duplicated.length; i++)
			duplicated[i] = Math.floor(random() * 256)
		return duplicated
	}
	const start = Math.floor(random() * Math.max(1, bytes.length - 1))
	const length = Math.max(1, Math.floor(random() * Math.min(64, bytes.length - start)))
	bytes.fill(0, start, start + length)
	return bytes
}

describe('readXlsx fuzz harness', () => {
	test('handles randomized corpus mutations without throwing', () => {
		const base = makeMinimalXlsx()
		const random = rng(0xdecafbad)
		let okCount = 0
		let errCount = 0

		for (let i = 0; i < 250; i++) {
			const mutated = mutate(base, random)
			let threw = false
			try {
				const result = readXlsx(mutated)
				if (result.ok) okCount++
				else errCount++
			} catch {
				threw = true
			}
			expect(threw).toBe(false)
		}

		// Fuzz corpus should include both valid-ish and invalid mutations.
		expect(okCount).toBeGreaterThan(0)
		expect(errCount).toBeGreaterThan(0)
	})
})
