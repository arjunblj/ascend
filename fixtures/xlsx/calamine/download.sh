#!/bin/sh
set -eu

BASE_URL="https://raw.githubusercontent.com/tafia/calamine/master/tests"
SCRIPT_DIR="$(CDPATH= cd -- "$(dirname "$0")" && pwd)"

FILES="
any_sheets.xlsx
column_row_ranges.xlsx
date.xlsx
date_1904.xlsx
date_iso.xlsx
empty_s_attribute.xlsx
empty_shared_string.xlsx
empty_shared_string_value.xlsx
empty_sheet.xlsx
encoded_entities.xlsx
errors.xlsx
formula.issue.xlsx
has_x000D_.xlsx
has_x000D_inline.xlsx
header-row.xlsx
inlineStr_with_value.xlsx
inventory-table.xlsx
issue127.xlsx
issue221.xlsm
issue252.xlsx
issue281.xlsm
issue3.xlsm
issue438.xlsx
issue446.xlsx
issue9.xlsx
issue_174.xlsx
issue_261.xlsx
issue_261_fixed_by_excel.xlsx
issue_391.xlsx
issue_419.xlsx
issue_530.xlsx
issue_553.xlsx
issue_565_multi_axis_shared.xlsx
issue_567_absolute_shared.xlsx
issues.xlsx
merge_cells.xlsx
merged_range.xlsx
no-header.xlsx
non_monotonic_si.xlsx
pass_protected.xlsx
picture.xlsx
pivots.xlsx
richtext-namespaced.xlsx
rph.xlsx
shared_formula_reversed.xlsx
string-ref.xlsx
table-multiple.xlsx
table_with_absolute_paths.xlsx
table_with_insertrow_attribute.xlsx
temperature-in-middle.xlsx
temperature-table.xlsx
temperature.xlsx
vba.xlsm
"

for file in $FILES; do
	if [ -f "$SCRIPT_DIR/$file" ]; then
		printf 'skip  %s\n' "$file"
		continue
	fi
	printf 'fetch %s\n' "$file"
	curl -fsSL "$BASE_URL/$file" -o "$SCRIPT_DIR/$file"
done

printf 'Downloaded Calamine fixtures into %s\n' "$SCRIPT_DIR"
