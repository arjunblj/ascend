#!/usr/bin/env sh
set -eu

base_url="https://raw.githubusercontent.com/ClosedXML/ClosedXML/develop"

download() {
	source_path="$1"
	target="$2"
	printf 'Downloading %s\n' "$target"
	curl -fsSL "$base_url/$source_path" -o "$target"
}

cd "$(dirname "$0")"

download "ClosedXML.Tests/Resource/Examples/AutoFilter/CustomAutoFilter.xlsx" "AutoFilter_CustomAutoFilter.xlsx"
download "ClosedXML.Tests/Resource/Examples/Comments/AddingComments.xlsx" "Comments_AddingComments.xlsx"
download "ClosedXML.Tests/Resource/Examples/ConditionalFormatting/CFDataBars.xlsx" "ConditionalFormatting_CFDataBars.xlsx"
download "ClosedXML.Tests/Resource/Examples/ConditionalFormatting/CFIconSet.xlsx" "ConditionalFormatting_CFIconSet.xlsx"
download "ClosedXML.Tests/Resource/Examples/ImageHandling/ImageAnchors.xlsx" "ImageHandling_ImageAnchors.xlsx"
download "ClosedXML.Tests/Resource/Examples/Misc/DataValidation.xlsx" "Misc_DataValidation.xlsx"
download "ClosedXML.Tests/Resource/Examples/Misc/Formulas.xlsx" "Misc_Formulas.xlsx"
download "ClosedXML.Tests/Resource/Examples/Misc/FormulasWithEvaluation.xlsx" "Misc_FormulasWithEvaluation.xlsx"
download "ClosedXML.Tests/Resource/Examples/Misc/Hyperlinks.xlsx" "Misc_Hyperlinks.xlsx"
download "ClosedXML.Tests/Resource/Examples/Misc/MergeCells.xlsx" "Misc_MergeCells.xlsx"
download "ClosedXML.Tests/Resource/Examples/Misc/ShiftingFormulas.xlsx" "Misc_ShiftingFormulas.xlsx"
download "ClosedXML.Tests/Resource/Examples/Misc/SheetProtection.xlsx" "Misc_SheetProtection.xlsx"
download "ClosedXML.Tests/Resource/Examples/PivotTables/PivotTables.xlsx" "PivotTables_PivotTables.xlsx"
download "ClosedXML.Tests/Resource/Examples/Ranges/DefinedNames.xlsx" "Ranges_DefinedNames.xlsx"
download "ClosedXML.Tests/Resource/Examples/Ranges/SortExample.xlsx" "Ranges_SortExample.xlsx"
download "ClosedXML.Tests/Resource/Examples/Sparklines/SampleSparklines.xlsx" "Sparklines_SampleSparklines.xlsx"
download "ClosedXML.Tests/Resource/Examples/Styles/StyleNumberFormat.xlsx" "Styles_StyleNumberFormat.xlsx"
download "ClosedXML.Tests/Resource/Examples/Styles/UsingRichText.xlsx" "Styles_UsingRichText.xlsx"
download "ClosedXML.Tests/Resource/Examples/Tables/UsingTables.xlsx" "Tables_UsingTables.xlsx"
download "ClosedXML.Tests/Resource/Examples/Tables/ResizingTables.xlsx" "Tables_ResizingTables.xlsx"
download "ClosedXML.Tests/Resource/Other/Charts/PreserveCharts/inputfile.xlsx" "Other_Charts_PreserveCharts_inputfile.xlsx"
download "ClosedXML.Tests/Resource/Other/ExternalLinks/WorkbookWithExternalLink.xlsx" "Other_ExternalLinks_WorkbookWithExternalLink.xlsx"
download "ClosedXML.Tests/Resource/Other/Formulas/ArrayFormula.xlsx" "Other_Formulas_ArrayFormula.xlsx"
download "ClosedXML.Tests/Resource/Other/Formulas/BooleanFormulaValues.xlsx" "Other_Formulas_BooleanFormulaValues.xlsx"
download "ClosedXML.Tests/Resource/Other/Formulas/DataTableFormula-Excel-Input.xlsx" "Other_Formulas_DataTableFormula-Excel-Input.xlsx"
download "ClosedXML.Tests/Resource/Other/PivotTableReferenceFiles/ChartsheetAndPivotTable.xlsx" "Other_PivotTableReferenceFiles_ChartsheetAndPivotTable.xlsx"
