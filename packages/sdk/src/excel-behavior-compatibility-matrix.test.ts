import { describe, expect, test } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'

const REPO_ROOT = new URL('../../../', import.meta.url)
const MATRIX_PATH = 'docs/EXCEL_BEHAVIOR_COMPATIBILITY_MATRIX.md'

const EXPECTED_SURFACES = [
	'Existing cells, sheet names, dimensions, styles, merges, panes, row/column sizing, workbook views, doc properties',
	'Formula text, shared formulas, legacy arrays, data tables, formula binding metadata, calc settings and calcChain invalidation',
	'Tables, structured references, defined names, validations, conditional formatting, filters, hyperlinks, comments, sheet/workbook/protected-range protection metadata',
	'Charts, chart sheets, drawings, images, VML, media, visual sidecars',
	'Pivot tables, pivot caches, slicers, timelines, data model, GETPIVOTDATA-facing output',
	'Macros, Excel 4 macro sheets, ActiveX, form controls, custom UI callbacks, vendor security parts',
	'External workbook links, workbook connections, query tables, Power Query mashups',
	'Digital signatures, arbitrary unknown package parts, encrypted workbooks, malformed packages',
]

const CLASSIFICATIONS = new Set([
	'fully support',
	'preserve unchanged',
	'report honestly',
	'gate explicitly',
])

const REQUIRED_COMMANDS = [
	'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
	'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
	'bun test packages/sdk/src/open-plan.test.ts',
	'bun test packages/core/src/protection.test.ts packages/engine/src/operations.test.ts packages/sdk/src/ops-schema.test.ts packages/sdk/src/sdk.test.ts',
	'bun test fixtures/corpus/feature-contract.test.ts',
	'bun run fixtures/corpus/audit.ts',
	'bun run fixtures/benchmarks/package-action-proof.ts --json',
	'bun test fixtures/xlsx/roundtrip-fidelity.test.ts',
	'bun run fixtures/benchmarks/real-workbook.ts research/excel-corpus/NYC_311_SR_2010-2020-sample-1M.xlsx --step open-metadata',
	'bun test fixtures/benchmarks/upstream-real-workbooks.test.ts',
]

const REQUIRED_SOURCE_URLS = [
	'https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain',
	'https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e',
	'https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1',
	'https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet',
	'https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653',
	'https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/15197a02-26f7-4a37-ab9c-7de4d4894ea5',
	'https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/3146988c-a3c7-488f-a84c-0cc115665cc3',
	'https://support.microsoft.com/en-gb/office/change-macro-security-settings-in-excel-a97c09d2-c082-46b8-b19f-e8621e8fe373',
	'https://support.microsoft.com/en-us/office/enable-or-disable-activex-settings-in-office-files-f1303e08-a3f8-41c5-a17e-b0b8898743ed',
	'https://openpyxl.readthedocs.io/en/3.1/tutorial.html',
	'https://docs.sheetjs.com/docs/api/write-options/',
]

describe('Excel behavior compatibility matrix', () => {
	test('keeps the ranked ladder tied to measured workbook evidence', () => {
		const markdown = readRepoFile(MATRIX_PATH)
		const rows = matrixRows(markdown)
		const claimRows = releaseClaimRows(markdown)

		expect(rows.map((row) => row.surface)).toEqual(EXPECTED_SURFACES)
		expect(rows.map((row) => row.rank)).toEqual([1, 2, 3, 4, 5, 6, 7, 8])
		expect(new Set(rows.map((row) => row.classification))).toEqual(CLASSIFICATIONS)
		expect(markdown).toContain('224 tracked public XLSX/XLSM fixtures scanned')
		expect(markdown).toContain('1 public unknown-path-family fixture')
		expect(markdown).toContain('no public signed-workbook candidate found')
		expect(markdown).toContain('does not echo the password in returned plan JSON')
		expect(markdown).toContain('sheet/workbook/protected-range plaintext operation input')
		expect(markdown).toContain('without storing the plaintext')

		for (const row of rows) {
			expect(row.open).not.toBe('')
			expect(row.inspect).not.toBe('')
			expect(row.editSafely).not.toBe('')
			expect(row.saveReopen).not.toBe('')
			expect(row.verify).not.toBe('')
			expect(row.nextReleaseGap).toMatch(
				/fixture|Excel|workbook|unsupported|claim|refresh|release/i,
			)
			for (const path of proofPaths(row.evidence)) {
				expect(existsSync(new URL(path, REPO_ROOT)), `${row.surface}: ${path}`).toBe(true)
			}
		}

		for (const command of REQUIRED_COMMANDS) expect(markdown).toContain(`\`${command}\``)
		for (const url of REQUIRED_SOURCE_URLS) expect(markdown).toContain(url)

		expect(claimRows.map((row) => row.rank)).toEqual([1, 2, 3, 4, 5, 6, 7, 8])
		for (const row of claimRows) {
			expect(row.claim).not.toBe('')
			expect(row.evidenceWeHave).toMatch(/proof|test|audit|fixture|roundtrip|scan|journal/i)
			expect(row.evidenceMissing).toMatch(/missing|needed|need|incomplete|unsupported|broader/i)
			expect(row.qssContrast).toMatch(/QSS/)
			expect(row.allowedWording).not.toBe('')
			expect(row.forbiddenWording).toMatch(/"[^"]+"/)
			expect(row.nextOwnerAction).toMatch(/owner/)
		}
		expect(claimRows[7]).toMatchObject({
			allowedWording: expect.stringContaining('encrypted workbook password handling'),
			forbiddenWording: expect.stringContaining('password recovery'),
			nextOwnerAction: expect.stringContaining('generated signed/malformed edge packages'),
		})
	})
})

interface MatrixRow {
	readonly rank: number
	readonly surface: string
	readonly classification: string
	readonly open: string
	readonly inspect: string
	readonly editSafely: string
	readonly saveReopen: string
	readonly verify: string
	readonly evidence: string
	readonly nextReleaseGap: string
}

interface ReleaseClaimRow {
	readonly rank: number
	readonly claim: string
	readonly evidenceWeHave: string
	readonly evidenceMissing: string
	readonly qssContrast: string
	readonly allowedWording: string
	readonly forbiddenWording: string
	readonly nextOwnerAction: string
}

function matrixRows(markdown: string): MatrixRow[] {
	const section = markdown.split('## Ladder')[1]?.split('## Release Claim Decisions')[0]
	if (!section) return []
	return section
		.split('\n')
		.filter((line) => line.startsWith('| ') && !line.includes(' --- '))
		.slice(1)
		.map((line) => {
			const cells = line.split('|').map((cell) => cell.trim())
			return {
				rank: Number.parseInt(cells[1] ?? '', 10),
				surface: cells[2] ?? '',
				classification: cells[3] ?? '',
				open: cells[4] ?? '',
				inspect: cells[5] ?? '',
				editSafely: cells[6] ?? '',
				saveReopen: cells[7] ?? '',
				verify: cells[8] ?? '',
				evidence: cells[9] ?? '',
				nextReleaseGap: cells[10] ?? '',
			}
		})
}

function releaseClaimRows(markdown: string): ReleaseClaimRow[] {
	const section = markdown.split('## Release Claim Decisions')[1]?.split('## Source Boundary')[0]
	if (!section) return []
	return section
		.split('\n')
		.filter((line) => line.startsWith('| ') && !line.includes(' --- '))
		.slice(1)
		.map((line) => {
			const cells = line.split('|').map((cell) => cell.trim())
			return {
				rank: Number.parseInt(cells[1] ?? '', 10),
				claim: cells[2] ?? '',
				evidenceWeHave: cells[3] ?? '',
				evidenceMissing: cells[4] ?? '',
				qssContrast: cells[5] ?? '',
				allowedWording: cells[6] ?? '',
				forbiddenWording: cells[7] ?? '',
				nextOwnerAction: cells[8] ?? '',
			}
		})
}

function proofPaths(evidence: string): string[] {
	return [...evidence.matchAll(/`([^`]+)`/g)]
		.map((match) => match[1])
		.filter((path) => /^(packages|fixtures)\//.test(path))
}

function readRepoFile(path: string): string {
	return readFileSync(new URL(path, REPO_ROOT), 'utf-8')
}
