import { describe, expect, test } from 'bun:test'
import { parseChartXml } from './charts.ts'

describe('chart inventory', () => {
	test('parses chart metadata across namespace prefixes', () => {
		expect(
			parseChartXml(
				`<?xml version="1.0"?>
<chart:chartSpace xmlns:chart="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:text="http://schemas.openxmlformats.org/drawingml/2006/main">
  <chart:chart>
    <chart:title><chart:tx><chart:rich><text:p><text:r><text:t>Revenue &amp; margin</text:t></text:r></text:p></chart:rich></chart:tx></chart:title>
    <chart:plotArea>
      <chart:barChart>
        <chart:ser>
          <chart:tx><chart:strRef><chart:f>Data!$B$1</chart:f></chart:strRef></chart:tx>
          <chart:cat><chart:strRef><chart:f>Data!$A$2:$A$4</chart:f></chart:strRef></chart:cat>
          <chart:val><chart:numRef><chart:f>Data!$B$2:$B$4</chart:f></chart:numRef></chart:val>
        </chart:ser>
      </chart:barChart>
    </chart:plotArea>
  </chart:chart>
</chart:chartSpace>`,
				'xl/charts/chart1.xml',
				'Data',
			),
		).toEqual({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Data',
			chartType: 'barChart',
			title: 'Revenue & margin',
			series: [
				{
					nameRef: 'Data!$B$1',
					categoryRef: 'Data!$A$2:$A$4',
					valueRef: 'Data!$B$2:$B$4',
				},
			],
		})
	})
})
