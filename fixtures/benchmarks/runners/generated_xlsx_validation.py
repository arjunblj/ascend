from __future__ import annotations

import hashlib
import io
import re
import zipfile
from typing import Any
from xml.etree import ElementTree


GENERATED_VALUE_WORKLOADS = {
    "dense-values",
    "mixed-10pct-text",
    "mixed-50pct-text",
    "mixed-closedxml-10text-5number",
    "plain-text",
    "string-heavy",
    "sparse-wide",
    "styles-heavy",
}


def can_validate_generated_workload(workload: str) -> bool:
    return workload in GENERATED_VALUE_WORKLOADS


def column_name(index: int) -> str:
    name = ""
    current = index
    while current > 0:
        current, rem = divmod(current - 1, 26)
        name = chr(65 + rem) + name
    return name or "A"


def column_index(name: str) -> int:
    value = 0
    for char in name:
        if char < "A" or char > "Z":
            break
        value = value * 26 + (ord(char) - 64)
    return value


def hash_lines(lines: list[str]) -> str:
    digest = hashlib.sha256()
    for line in sorted(lines):
        digest.update(f"{len(line)}:".encode("utf-8"))
        digest.update(line.encode("utf-8"))
        digest.update(b"\n")
    return digest.hexdigest()


class OrderedHasher:
    def __init__(self) -> None:
        self._digest = hashlib.sha256()

    def add(self, line: str) -> None:
        self._digest.update(f"{len(line)}:".encode("utf-8"))
        self._digest.update(line.encode("utf-8"))
        self._digest.update(b"\n")

    def hexdigest(self) -> str:
        return self._digest.hexdigest()


def generated_cell_count(workload: str, rows: int, cols: int) -> int:
    if workload != "sparse-wide":
        return rows * cols
    count = 0
    for row in range(rows):
        for col in range(cols):
            if workload_value(workload, row, col, cols) is not None:
                count += 1
    return count


def workload_value(workload: str, row: int, col: int, cols: int) -> str | int | None:
    if workload == "dense-values":
        return row * cols + col
    if workload == "mixed-10pct-text":
        key = row * cols + col
        return f"text-{key:08d}" if key % 10 == 0 else key
    if workload == "mixed-50pct-text":
        key = row * cols + col
        return f"text-{key:08d}" if key % 2 == 0 else key
    if workload == "mixed-closedxml-10text-5number":
        return "Hello world" if col < 10 else col - 10
    if workload == "plain-text":
        return f"text-{row * cols + col:08d}"
    if workload == "string-heavy":
        key = row * cols + col
        branch = col % 5
        if branch == 0:
            return f"sku-{key:08d}"
        if branch == 1:
            return f"region-{(row % 17) + 1}"
        if branch == 2:
            return f"customer-{row % 997}-segment-{col % 13}"
        if branch == 3:
            return f"note row {row} col {col} token {key % 104729}"
        return f"status-open-{key % 31}" if key % 2 == 0 else f"status-closed-{key % 29}"
    if workload == "styles-heavy":
        return (row + 1) * (col + 1)
    if workload == "sparse-wide":
        if col == 0:
            return row
        if col == cols - 1:
            return f"edge-{row}-{cols}"
        if (row * 31 + col * 17) % 97 == 0:
            return row * cols + col
        return None
    return None


def scalar_payload(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return f"b:{str(value).lower()}"
    if isinstance(value, (int, float)):
        number = float(value)
        if number == 0:
            return "n:0"
        if number.is_integer():
            return f"n:{int(number)}"
        return f"n:{format(number, '.15g')}"
    return f"s:{value}"


def expected_ordered_values_hash(workload: str, rows: int, cols: int) -> str:
    digest = OrderedHasher()
    column_names = [column_name(col + 1) for col in range(cols)]
    for row in range(rows):
        for col in range(cols):
            payload = scalar_payload(workload_value(workload, row, col, cols))
            if payload is None:
                continue
            digest.add(f"Data!{column_names[col]}{row + 1}\t{payload}")
    return digest.hexdigest()


def generated_write_assertions(
    data: bytes,
    *,
    workload: str,
    rows: int,
    cols: int,
    runner_version: str,
    expected_ordered_hash: str | None = None,
) -> dict[str, str | int | bool | None]:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            sheet_names = workbook_sheet_names(archive)
            shared_strings = workbook_shared_strings(archive)
            values = worksheet_value_summary(archive, shared_strings, rows, cols)
            expected_cells = generated_cell_count(workload, rows, cols)
            result: dict[str, str | int | bool | None] = {
                "runnerVersion": runner_version,
                "workload": workload,
                "bytes": len(data),
                "sheetCount": len(sheet_names),
                "sheetNamesHash": hash_lines([f"{index}:{name}" for index, name in enumerate(sheet_names)]),
                "cellCount": values["cellCount"],
                "expectedCellCount": expected_cells,
                "cellCountMatches": values["cellCount"] == expected_cells,
                "semanticCellValuesHash": values["orderedSemanticCellValuesHash"],
                "orderedSemanticCellValuesHash": values["orderedSemanticCellValuesHash"],
                "semanticCellRefsHash": values["orderedSemanticCellRefsHash"],
                "orderedSemanticCellRefsHash": values["orderedSemanticCellRefsHash"],
                "usedRangeCount": 1,
                "firstUsedRange": values["firstUsedRange"],
                "firstPhysicalUsedRange": values["firstUsedRange"],
                "usedRangesHash": values["usedRangesHash"],
                "physicalUsedRangesHash": values["usedRangesHash"],
                "physicalCellCount": values["cellCount"],
                "formulaTextHash": hash_lines([]),
                "generatedWriteZipValidation": True,
            }
            if expected_ordered_hash is not None:
                result["expectedOrderedSemanticCellValuesHash"] = expected_ordered_hash
                result["semanticCellValuesHashMatches"] = (
                    values["orderedSemanticCellValuesHash"] == expected_ordered_hash
                )
            return result
    except zipfile.BadZipFile as error:
        return {
            "runnerVersion": runner_version,
            "workload": workload,
            "bytes": len(data),
            "reopenOk": False,
            "reopenError": str(error),
            "cellCountMatches": False,
            "semanticCellValuesHashMatches": False,
        }


def feature_rich_assertions(data: bytes, *, rows: int, cols: int) -> dict[str, str | int | bool]:
    expected = expected_feature_rich_contract(rows, cols)
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            names = archive.namelist()
            workbook_xml = read_zip_text(archive, "xl/workbook.xml")
            sheet_xml = read_zip_text(archive, "xl/worksheets/sheet1.xml")
            sheet_rels_xml = read_zip_text(archive, "xl/worksheets/_rels/sheet1.xml.rels")
            comments_xml = first_existing_text(
                archive,
                [name for name in names if re.match(r"^xl/(?:comments\d*|comments/.+)\.xml$", name)],
            )
            defined_name_ref = defined_name_value(workbook_xml, "FeatureRange")
            hyperlink = first_element_attrs(sheet_xml, "hyperlink", "ref", "A1")
            hyperlink_rel_id = hyperlink.get("r:id", hyperlink.get("id", ""))
            hyperlink_target = relationship_target(sheet_rels_xml, hyperlink_rel_id)
            data_validation = element_with_attrs_and_body(
                sheet_xml, "dataValidation", {"sqref": expected["validationRef"]}
            )
            conditional_formatting = element_with_attrs_and_body(
                sheet_xml, "conditionalFormatting", {"sqref": expected["conditionalFormatRef"]}
            )
            cf_rule = (
                first_element_attrs(conditional_formatting[1], "cfRule")
                if conditional_formatting is not None
                else {}
            )
            comment = comment_entry(comments_xml, expected["commentRef"])
            defined_name_matches = normalize_formula_ref(defined_name_ref) == expected["featureRange"]
            hyperlink_matches = (
                hyperlink.get("ref") == expected["hyperlinkRef"]
                and (
                    hyperlink.get("display") is None
                    or hyperlink.get("display") == expected["hyperlinkDisplay"]
                )
                and hyperlink.get("tooltip") == expected["hyperlinkTooltip"]
                and hyperlink_target == expected["hyperlinkTarget"]
            )
            comment_matches = (
                comment["ref"] == expected["commentRef"]
                and comment["author"] == expected["commentAuthor"]
                and comment["text"] == expected["commentText"]
            )
            data_validation_matches = (
                data_validation is not None
                and data_validation[0].get("type") == "list"
                and data_validation[0].get("allowBlank") == "1"
                and data_validation[0].get("showInputMessage") == "1"
                and first_tag_text(data_validation[1], "formula1") == expected["validationFormula"]
            )
            conditional_formatting_matches = (
                conditional_formatting is not None
                and cf_rule.get("type") == "cellIs"
                and cf_rule.get("operator") == "greaterThan"
                and first_tag_text(conditional_formatting[1], "formula")
                == expected["conditionalFormula"]
            )
            semantic_matches = (
                defined_name_matches
                and hyperlink_matches
                and comment_matches
                and data_validation_matches
                and conditional_formatting_matches
            )
            return {
                "commentPartCount": sum(
                    1
                    for name in names
                    if re.match(r"^xl/(?:comments\d*|comments/.+)\.xml$", name)
                ),
                "vmlDrawingPartCount": sum(
                    1 for name in names if name.startswith("xl/drawings/") and name.endswith(".vml")
                ),
                "worksheetHyperlinkCount": sheet_xml.count("<hyperlink "),
                "worksheetDataValidationCount": sheet_xml.count("<dataValidation "),
                "worksheetConditionalFormattingCount": sheet_xml.count("<conditionalFormatting"),
                "definedNameCount": workbook_xml.count("<definedName "),
                "featureRichSemanticMatches": semantic_matches,
                "featureRichDefinedNameMatches": defined_name_matches,
                "featureRichHyperlinkMatches": hyperlink_matches,
                "featureRichCommentMatches": comment_matches,
                "featureRichDataValidationMatches": data_validation_matches,
                "featureRichConditionalFormattingMatches": conditional_formatting_matches,
                "featureRichDefinedNameRef": normalize_formula_ref(defined_name_ref),
                "featureRichHyperlinkTarget": hyperlink_target,
                "featureRichCommentText": comment["text"],
                "featureRichDataValidationFormula": (
                    first_tag_text(data_validation[1], "formula1") if data_validation else ""
                ),
                "featureRichConditionalFormula": (
                    first_tag_text(conditional_formatting[1], "formula")
                    if conditional_formatting
                    else ""
                ),
            }
    except zipfile.BadZipFile:
        return empty_feature_rich_assertions()


def empty_feature_rich_assertions() -> dict[str, str | int | bool]:
    return {
        "commentPartCount": 0,
        "vmlDrawingPartCount": 0,
        "worksheetHyperlinkCount": 0,
        "worksheetDataValidationCount": 0,
        "worksheetConditionalFormattingCount": 0,
        "definedNameCount": 0,
        "featureRichSemanticMatches": False,
        "featureRichDefinedNameMatches": False,
        "featureRichHyperlinkMatches": False,
        "featureRichCommentMatches": False,
        "featureRichDataValidationMatches": False,
        "featureRichConditionalFormattingMatches": False,
        "featureRichDefinedNameRef": "",
        "featureRichHyperlinkTarget": "",
        "featureRichCommentText": "",
        "featureRichDataValidationFormula": "",
        "featureRichConditionalFormula": "",
    }


def expected_feature_rich_contract(rows: int, cols: int) -> dict[str, str]:
    return {
        "featureRange": f"Data!$A$1:${column_name(cols)}${max(1, rows)}",
        "hyperlinkRef": "A1",
        "hyperlinkDisplay": "Ascend",
        "hyperlinkTarget": "https://example.com/ascend",
        "hyperlinkTooltip": "Open Ascend",
        "commentRef": "B2",
        "commentAuthor": "Ascend",
        "commentText": "Review",
        "validationRef": f"C2:C{max(2, rows)}",
        "validationFormula": '"Q1,Q2,Q3"',
        "conditionalFormatRef": f"A1:A{max(1, rows)}",
        "conditionalFormula": "0",
    }


def read_zip_text(archive: zipfile.ZipFile, name: str) -> str:
    try:
        return archive.read(name).decode("utf-8")
    except KeyError:
        return ""


def first_existing_text(archive: zipfile.ZipFile, names: list[str]) -> str:
    for name in names:
        text = read_zip_text(archive, name)
        if text:
            return text
    return ""


def parse_attrs(attrs_text: str) -> dict[str, str]:
    return {
        match.group(1): decode_xml_text(match.group(2))
        for match in re.finditer(r'([A-Za-z_][\w:.-]*)="([^"]*)"', attrs_text)
    }


def first_element_attrs(
    xml: str, tag_name: str, attr_name: str | None = None, attr_value: str | None = None
) -> dict[str, str]:
    for match in re.finditer(fr"<{tag_name}\b([^>]*)>", xml):
        attrs = parse_attrs(match.group(1))
        if attr_name is None or attrs.get(attr_name) == attr_value:
            return attrs
    return {}


def element_with_attrs_and_body(
    xml: str, tag_name: str, expected_attrs: dict[str, str]
) -> tuple[dict[str, str], str] | None:
    for match in re.finditer(fr"<{tag_name}\b([^>]*)>([\s\S]*?)</{tag_name}>", xml):
        attrs = parse_attrs(match.group(1))
        if all(attrs.get(key) == value for key, value in expected_attrs.items()):
            return attrs, match.group(2)
    return None


def relationship_target(rels_xml: str, rel_id: str) -> str:
    if not rel_id:
        return ""
    for match in re.finditer(r"<Relationship\b([^>]*)/?>", rels_xml):
        attrs = parse_attrs(match.group(1))
        if attrs.get("Id") == rel_id:
            return attrs.get("Target", "")
    return ""


def defined_name_value(workbook_xml: str, name: str) -> str:
    for match in re.finditer(r"<definedName\b([^>]*)>([\s\S]*?)</definedName>", workbook_xml):
        attrs = parse_attrs(match.group(1))
        if attrs.get("name") == name:
            return decode_xml_text(match.group(2))
    return ""


def comment_entry(comments_xml: str, expected_ref: str) -> dict[str, str]:
    authors = [
        decode_xml_text(match.group(1))
        for match in re.finditer(r"<author>([\s\S]*?)</author>", comments_xml)
    ]
    for match in re.finditer(r"<comment\b([^>]*)>([\s\S]*?)</comment>", comments_xml):
        attrs = parse_attrs(match.group(1))
        ref = attrs.get("ref", "")
        if ref != expected_ref:
            continue
        try:
            author = authors[int(attrs.get("authorId", ""))]
        except (ValueError, IndexError):
            author = ""
        text = "".join(
            decode_xml_text(text_match.group(1))
            for text_match in re.finditer(r"<t(?:\s[^>]*)?>([\s\S]*?)</t>", match.group(2))
        )
        return {"ref": ref, "author": author, "text": text}
    return {"ref": "", "author": "", "text": ""}


def first_tag_text(xml: str, tag_name: str) -> str:
    match = re.search(fr"<{tag_name}>([\s\S]*?)</{tag_name}>", xml)
    return decode_xml_text(match.group(1)) if match else ""


def decode_xml_text(text: str) -> str:
    return (
        text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


def normalize_formula_ref(value: str) -> str:
    return value[1:] if value.startswith("=") else value


def workbook_sheet_names(archive: zipfile.ZipFile) -> list[str]:
    with archive.open("xl/workbook.xml") as handle:
        tree = ElementTree.iterparse(handle, events=("start",))
        return [
            element.attrib.get("name", "")
            for _, element in tree
            if strip_namespace(element.tag) == "sheet" and element.attrib.get("name")
        ]


def workbook_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        handle = archive.open("xl/sharedStrings.xml")
    except KeyError:
        return []
    values: list[str] = []
    current: list[str] = []
    in_string = False
    with handle:
        for event, element in ElementTree.iterparse(handle, events=("start", "end")):
            tag = strip_namespace(element.tag)
            if event == "start" and tag == "si":
                current = []
                in_string = True
            elif event == "end" and in_string and tag == "t":
                current.append(element.text or "")
            elif event == "end" and tag == "si":
                values.append("".join(current))
                current = []
                in_string = False
                element.clear()
    return values


def worksheet_value_summary(
    archive: zipfile.ZipFile,
    shared_strings: list[str],
    expected_rows: int,
    expected_cols: int,
) -> dict[str, str | int | None]:
    sheet_name = next(name for name in archive.namelist() if re.match(r"^xl/worksheets/sheet\d+\.xml$", name))
    ordered_refs = OrderedHasher()
    ordered_values = OrderedHasher()
    column_names = [column_name(col + 1) for col in range(expected_cols)]
    cell_count = 0
    min_row = 10**18
    min_col = 10**18
    max_row = 0
    max_col = 0
    row_index = 0
    implicit_row = 0
    implicit_col = 0
    in_cell = False
    in_value = False
    in_inline_text = False
    cell_type = ""
    cell_col = 0
    cell_value: list[str] = []
    inline_value: list[str] = []
    with archive.open(sheet_name) as handle:
        for event, element in ElementTree.iterparse(handle, events=("start", "end")):
            tag = strip_namespace(element.tag)
            if event == "start" and tag == "row":
                row_index = int(element.attrib.get("r") or implicit_row + 1)
                implicit_row = row_index
                implicit_col = 0
            elif event == "start" and tag == "c":
                in_cell = True
                cell_type = element.attrib.get("t", "")
                cell_value = []
                inline_value = []
                cell_col = column_index(cell_ref_letters(element.attrib.get("r", "")))
                if cell_col <= 0:
                    cell_col = implicit_col + 1
                implicit_col = cell_col
            elif event == "start" and in_cell and tag == "v":
                in_value = True
            elif event == "end" and in_cell and tag == "v":
                if in_value:
                    cell_value.append(element.text or "")
                in_value = False
            elif event == "start" and in_cell and tag == "t":
                in_inline_text = True
            elif event == "end" and in_cell and tag == "t":
                if in_inline_text:
                    inline_value.append(element.text or "")
                in_inline_text = False
            elif event == "end" and tag == "c":
                payload = generated_cell_payload(cell_type, "".join(cell_value), "".join(inline_value), shared_strings)
                if payload is not None:
                    cell_count += 1
                    min_row = min(min_row, row_index)
                    min_col = min(min_col, cell_col)
                    max_row = max(max_row, row_index)
                    max_col = max(max_col, cell_col)
                    col_name = column_names[cell_col - 1] if 0 < cell_col <= len(column_names) else column_name(cell_col)
                    ref = f"Data!{col_name}{row_index}"
                    ordered_refs.add(ref)
                    ordered_values.add(f"{ref}\t{payload}")
                in_cell = False
                element.clear()
            elif event == "end" and tag == "row":
                element.clear()
    used_range = (
        "Data!empty"
        if cell_count == 0
        else f"Data!{column_name(min_col)}{min_row}:{column_name(max_col)}{max_row}"
    )
    return {
        "cellCount": cell_count,
        "firstUsedRange": used_range,
        "usedRangesHash": hash_lines([used_range]),
        "orderedSemanticCellRefsHash": ordered_refs.hexdigest(),
        "orderedSemanticCellValuesHash": ordered_values.hexdigest(),
    }


def generated_cell_payload(
    cell_type: str,
    raw: str,
    inline: str,
    shared_strings: list[str],
) -> str | None:
    if cell_type == "b":
        normalized = raw.lower()
        return f"b:{str(normalized in {'1', 'true'}).lower()}"
    if cell_type == "inlineStr":
        return f"s:{inline}"
    if cell_type == "s":
        try:
            return f"s:{shared_strings[int(raw)]}"
        except (IndexError, ValueError):
            return None
    if cell_type == "str":
        return f"s:{raw}"
    if raw == "":
        return None
    return f"n:{canonical_number_string(raw)}"


def canonical_number_string(raw: str) -> str:
    try:
        value = float(raw)
    except ValueError:
        return raw
    if value == 0:
        return "0"
    if value.is_integer():
        return str(int(value))
    return f"{value:.15g}"


def cell_ref_letters(ref: str) -> str:
    match = re.match(r"([A-Z]+)", ref)
    return match.group(1) if match else ""


def strip_namespace(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
