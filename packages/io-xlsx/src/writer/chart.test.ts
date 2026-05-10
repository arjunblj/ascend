import { describe, expect, test } from 'bun:test'
import { updateChartXml } from './chart.ts'

describe('updateChartXml', () => {
	test('preserves unchanged chart formula XML with quotes and entities', () => {
		const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:f>&#x27;LY &amp; Forecast&#x27;!$B$1</c:f></c:strRef></c:tx><c:cat><c:strRef><c:f>&#x27;LY &amp; Forecast&#x27;!$A$2:$A$4</c:f></c:strRef></c:cat><c:xVal><c:strRef><c:f>Stale!$A$2:$A$4</c:f></c:strRef></c:xVal><c:val><c:numRef><c:f>&apos;LY &amp; Forecast&apos;!$B$2:$B$4</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`

		const updated = updateChartXml(xml, {
			partPath: 'xl/charts/chart1.xml',
			series: [
				{
					nameRef: "'LY & Forecast'!$B$1",
					categoryRef: "'LY & Forecast'!$A$2:$A$4",
					valueRef: "'LY & Forecast'!$B$2:$B$4",
				},
			],
		})

		expect(updated).toBe(xml)
	})

	test('updates changed chart formulas', () => {
		const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:f>&#x27;LY &amp; Forecast&#x27;!$B$1</c:f></c:strRef></c:tx><c:cat><c:strRef><c:f>&#x27;LY &amp; Forecast&#x27;!$A$2:$A$4</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>&apos;LY &amp; Forecast&apos;!$B$2:$B$4</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`

		const updated = updateChartXml(xml, {
			partPath: 'xl/charts/chart1.xml',
			series: [
				{
					nameRef: "'CY & Forecast'!$C$1",
					categoryRef: "'CY & Forecast'!$A$2:$A$10",
					valueRef: "'CY & Forecast'!$C$2:$C$10",
				},
			],
		})

		expect(updated).toContain('<c:f>&#x27;CY &amp; Forecast&#x27;!$C$1</c:f>')
		expect(updated).toContain('<c:f>&#x27;CY &amp; Forecast&#x27;!$A$2:$A$10</c:f>')
		expect(updated).toContain('<c:f>&#x27;CY &amp; Forecast&#x27;!$C$2:$C$10</c:f>')
		expect(updated).not.toContain('&#x27;LY &amp; Forecast&#x27;')
		expect(updated).not.toContain('&apos;LY &amp; Forecast&apos;')
	})
})
