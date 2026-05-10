#!/bin/sh
set -eu

BASE_URL="https://raw.githubusercontent.com/exceljs/exceljs/master/spec/integration/data"
SCRIPT_DIR="$(CDPATH= cd -- "$(dirname "$0")" && pwd)"

FILES="
1904.xlsx
bogus-defined-name.xlsx
chart-sheet.xlsx
dateIssue.xlsx
fibonacci.xlsx
formulas.xlsx
many-columns.xlsx
shared_string_with_escape.xlsx
test-issue-1364.xlsx
test-issue-1575.xlsx
test-issue-163.xlsx
test-issue-1669.xlsx
test-issue-176.xlsx
test-issue-1842.xlsx
test-issue-623.xlsx
test-pr-1204.xlsx
test-pr-1220.xlsx
test-pr-567.xlsx
test-pr-728.xlsx
test-row-styles.xlsx
"

for file in $FILES; do
	if [ -f "$SCRIPT_DIR/$file" ]; then
		printf 'skip  %s\n' "$file"
		continue
	fi
	printf 'fetch %s\n' "$file"
	curl -fsSL "$BASE_URL/$file" -o "$SCRIPT_DIR/$file"
done

printf 'Downloaded ExcelJS fixtures into %s\n' "$SCRIPT_DIR"
