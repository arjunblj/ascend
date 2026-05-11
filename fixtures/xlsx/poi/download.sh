#!/bin/sh
set -eu

BASE_URL="https://raw.githubusercontent.com/apache/poi/refs/heads/trunk/test-data/spreadsheet"
SCRIPT_DIR="$(CDPATH= cd -- "$(dirname "$0")" && pwd)"

FILES="
shared_formulas.xlsx
TestShiftRowSharedFormula.xlsx
MatrixFormulaEvalTestData.xlsx
ConditionalFormattingSamples.xlsx
WithConditionalFormatting.xlsx
NewStyleConditionalFormattings.xlsx
DataValidationEvaluations.xlsx
DataValidations-49244.xlsx
styles.xlsx
NumberFormatTests.xlsx
DateFormatTests.xlsx
Themes.xlsx
InlineString.xlsx
noSharedStringTable.xlsx
StructuredReferences.xlsx
Tables.xlsx
sheetProtection_allLocked.xlsx
workbookProtection_workbook_structure_protected.xlsx
55906-MultiSheetRefs.xlsx
SimpleStrict.xlsx
SimpleWithComments.xlsx
FormulaEvalTestData_Copy.xlsx
FormulaSheetRange.xlsx
NewlineInFormulas.xlsx
atp.xlsx
Booleans.xlsx
Formatting.xlsx
TextFormatTests.xlsx
GeneralFormatTests.xlsx
ElapsedFormatTests.xlsx
TwoSheetsOneHidden.xlsx
GroupTest.xlsx
50784-font_theme_colours.xlsx
50786-indexed_colours.xlsx
48495.xlsx
SampleSS.xlsx
WithVariousData.xlsx
WithChart.xlsx
WithDrawing.xlsx
ShrinkToFit.xlsx
BrNotClosed.xlsx
sheetProtection_not_protected.xlsx
123233_charts.xlsx
45540_classic_Footer.xlsx
45540_classic_Header.xlsx
45540_form_Footer.xlsx
45540_form_Header.xlsx
50755_workday_formula_example.xlsx
50846-border_colours.xlsx
50867_with_table.xlsx
55406_Conditional_formatting_sample.xlsx
60255_extra_drawingparts.xlsx
61060-conditional-number-formatting.xlsx
62629_toMerge.xlsx
AmpersandHeader.xlsx
chart_sheet.xlsx
chartTitle_noTitle.xlsx
chartTitle_withTitle.xlsx
chartTitle_withTitleFormula.xlsx
comments.xlsx
commentTest.xlsx
conditional_formatting_cell_is.xlsx
conditional_formatting_multiple_ranges.xlsx
conditional_formatting_with_formula_on_second_sheet.xlsx
customIndexedColors.xlsx
CustomXMLMapping-singleattributenamespace.xlsx
CustomXMLMappings-complex-type.xlsx
CustomXmlMappings-inverse-order.xlsx
CustomXMLMappings.xlsx
DataTableCities.xlsx
DataValidationListTooLong.xlsx
dataValidationTableRange.xlsx
DateFormatNumberTests.xlsx
decimal-format.xlsx
evaluate_formula_with_structured_table_references.xlsx
ExcelPivotTableSample.xlsx
ExcelTables.xlsx
FillWithoutColor.xlsx
FormatChoiceTests.xlsx
FormatConditionTests.xlsx
FormatKM.xlsx
formula-eval.xlsx
HeaderFooterComplexFormats.xlsx
headerFooterTest.xlsx
link-external-workbook-a.xlsx
link-external-workbook-b.xlsx
NumberFormatApproxTests.xlsx
poc-shared-strings.xlsx
sample.strict.xlsx
SampleSS.strict.xlsx
sharedhyperlink.xlsx
SheetTabColors.xlsx
simple-table-named-range.xlsx
SimpleScatterChart.xlsx
SingleCellTable.xlsx
style-alternate-content.xlsx
table-sample.xlsx
tableStyle.xlsx
TablesWithDifferentHeaders.xlsx
test_conditional_formatting.xlsx
testSharedFormulasRangeSetBlankBug.xlsx
testSharedFormulasSetBlank.xlsx
Themes2.xlsx
WithChartSheet.xlsx
WithTable.xlsx
WithThreeCharts.xlsx
WithTwoCharts.xlsx
workbookProtection_not_protected.xlsx
workbookProtection_workbook_revision_protected.xlsx
workbookProtection_workbook_windows_protected.xlsx
workbookProtection_worksheet_protected.xlsx
workbookProtection-sheet_password-2013.xlsx
workbookProtection-workbook_password_user_range-2010.xlsx
workbookProtection-workbook_password-2013.xlsx
"

SHEETJS_BASE_URL="https://oss.sheetjs.com/test_files"
SHEETJS_FILES="
formula_stress_test.xlsx
merge_cells.xlsx
AutoFilter.xlsx
named_ranges_2011.xlsx
"

for file in $FILES; do
	if [ -f "$SCRIPT_DIR/$file" ]; then
		printf 'skip  %s\n' "$file"
		continue
	fi
	printf 'fetch %s\n' "$file"
	curl -fsSL "$BASE_URL/$file" -o "$SCRIPT_DIR/$file"
done

for file in $SHEETJS_FILES; do
	if [ -f "$SCRIPT_DIR/$file" ]; then
		printf 'skip  %s\n' "$file"
		continue
	fi
	printf 'fetch %s (SheetJS)\n' "$file"
	curl -fsSL "$SHEETJS_BASE_URL/$file" -o "$SCRIPT_DIR/$file"
done

printf 'Downloaded fixtures into %s\n' "$SCRIPT_DIR"
