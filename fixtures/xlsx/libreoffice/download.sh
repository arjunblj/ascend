#!/usr/bin/env sh
set -eu

base_url="https://raw.githubusercontent.com/LibreOffice/core/master/sc/qa/unit/data/xlsx"
dir="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"

download() {
	file="$1"
	curl -fsSL "$base_url/$file" -o "$dir/$file"
}

download "129969-min.xlsx"
download "CalcThemeTest.xlsx"
download "MissingPathExternal.xlsx"
download "PivotTable_CachedDefinitionAndDataInSync.xlsx"
download "PivotTable_CachedDefinitionAndDataNotInSync_SheetColumnsRemoved_WithCacheData.xlsx"
download "ProtecteSheet1234Pass.xlsx"
download "Sparklines.xlsx"
download "TableEmptyHeaders.xlsx"
download "TableStyleTest.xlsx"
download "Test_ThemeColor_Text_Background_Border.xlsx"
download "activex_checkbox.xlsx"
download "autofilter-colors.xlsx"
download "autofilter.xlsx"
download "textLengthDataValidity.xlsx"
download "textbox-hyperlink.xlsx"
download "totalsRowFunction.xlsx"
download "totalsRowShown.xlsx"
download "universal-content-strict.xlsx"
download "universal-content.xlsx"
download "user_defined_function.xlsx"
download "value-in-column-2000.xlsx"
download "writingMode.xlsx"
