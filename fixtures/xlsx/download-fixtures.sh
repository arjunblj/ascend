#!/bin/sh
# Download external XLSX test fixtures for integration testing.
# Run manually: ./fixtures/xlsx/download-fixtures.sh
# Downloaded files go to fixtures/xlsx/external/ (gitignored).

set -eu

SCRIPT_DIR="$(CDPATH= cd -- "$(dirname "$0")" && pwd)"
OUT_DIR="$SCRIPT_DIR/external"
mkdir -p "$OUT_DIR"

# Apache POI - key test files from test-data/spreadsheet
POI_BASE="https://raw.githubusercontent.com/apache/poi/refs/heads/trunk/test-data/spreadsheet"
POI_FILES="
styles.xlsx
SimpleWithComments.xlsx
shared_formulas.xlsx
55906-MultiSheetRefs.xlsx
StructuredReferences.xlsx
NumberFormatTests.xlsx
ConditionalFormattingSamples.xlsx
DataValidationEvaluations.xlsx
sheetProtection_allLocked.xlsx
workbookProtection_workbook_structure_protected.xlsx
SimpleStrict.xlsx
"

# SheetJS test_files
SHEETJS_BASE="https://oss.sheetjs.com/test_files"
SHEETJS_FILES="
merge_cells.xlsx
formula_stress_test.xlsx
named_ranges_2011.xlsx
AutoFilter.xlsx
"

# data.gov - small sample XLSX (catalog metadata)
# URLs may change; add more as needed
DATAGOV_FILES="
"

for file in $POI_FILES; do
	[ -z "$file" ] && continue
	if [ -f "$OUT_DIR/poi-$file" ]; then
		printf 'skip  poi-%s\n' "$file"
		continue
	fi
	printf 'fetch poi-%s\n' "$file"
	curl -fsSL "$POI_BASE/$file" -o "$OUT_DIR/poi-$file" || true
done

for file in $SHEETJS_FILES; do
	[ -z "$file" ] && continue
	if [ -f "$OUT_DIR/sheetjs-$file" ]; then
		printf 'skip  sheetjs-%s\n' "$file"
		continue
	fi
	printf 'fetch sheetjs-%s\n' "$file"
	curl -fsSL "$SHEETJS_BASE/$file" -o "$OUT_DIR/sheetjs-$file" || true
done

printf 'Downloaded fixtures into %s\n' "$OUT_DIR"
