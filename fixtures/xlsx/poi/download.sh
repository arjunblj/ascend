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
