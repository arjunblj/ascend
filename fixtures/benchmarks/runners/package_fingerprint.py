from __future__ import annotations

import hashlib
import posixpath
import re
import zipfile
from io import BytesIO

STRICT_REL_PREFIX = "http://purl.oclc.org/ooxml/officeDocument/relationships/"
TRANSITIONAL_REL_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def hash_lines(lines: list[str]) -> str:
    hasher = hashlib.sha256()
    for line in sorted(lines):
        encoded = line.encode("utf-8")
        hasher.update(str(len(line)).encode("utf-8"))
        hasher.update(b":")
        hasher.update(encoded)
        hasher.update(b"\n")
    return hasher.hexdigest()


def roundtrip_package_assertions(data: bytes) -> dict[str, str | int]:
    fingerprint = extract_package_fingerprint(data)
    return {
        "roundtripPackagePartCount": fingerprint["partCount"],
        "roundtripPackagePartNamesHash": fingerprint["partNamesHash"],
        "roundtripPackageContentTypeCount": fingerprint["contentTypeCount"],
        "roundtripPackageContentTypesHash": fingerprint["contentTypesHash"],
        "roundtripPackageRelationshipCount": fingerprint["relationshipCount"],
        "roundtripPackageRelationshipGraphHash": fingerprint["relationshipGraphHash"],
        "roundtripPreservedPartCount": fingerprint["preservedPartCount"],
        "roundtripPreservedPartNamesHash": fingerprint["preservedPartNamesHash"],
        "roundtripPreservedPartContentHash": fingerprint["preservedPartContentHash"],
    }


def roundtrip_feature_assertions(data: bytes) -> dict[str, str | int]:
    summary = extract_feature_summary(data)
    return {
        "roundtripTablePartCount": summary["tablePartCount"],
        "roundtripChartPartCount": summary["chartPartCount"],
        "roundtripChartExPartCount": summary["chartExPartCount"],
        "roundtripDrawingPartCount": summary["drawingPartCount"],
        "roundtripVmlDrawingPartCount": summary["vmlDrawingPartCount"],
        "roundtripPivotTablePartCount": summary["pivotTablePartCount"],
        "roundtripPivotCachePartCount": summary["pivotCachePartCount"],
        "roundtripSlicerPartCount": summary["slicerPartCount"],
        "roundtripCommentPartCount": summary["commentPartCount"],
        "roundtripThreadedCommentPartCount": summary["threadedCommentPartCount"],
        "roundtripMediaPartCount": summary["mediaPartCount"],
        "roundtripExternalLinkPartCount": summary["externalLinkPartCount"],
        "roundtripConnectionPartCount": summary["connectionPartCount"],
        "roundtripCustomXmlPartCount": summary["customXmlPartCount"],
        "roundtripWorksheetHyperlinkCount": summary["worksheetHyperlinkCount"],
        "roundtripWorksheetDataValidationCount": summary["worksheetDataValidationCount"],
        "roundtripWorksheetConditionalFormattingCount": summary[
            "worksheetConditionalFormattingCount"
        ],
        "roundtripDefinedNameCount": summary["definedNameCount"],
        "roundtripFeaturePartNamesHash": summary["featurePartNamesHash"],
        "roundtripFeatureInventoryHash": summary["featureInventoryHash"],
    }


def extract_package_fingerprint(data: bytes) -> dict[str, str | int]:
    with zipfile.ZipFile(BytesIO(data)) as archive:
        part_paths = sorted(
            info.filename for info in archive.infolist() if not info.filename.endswith("/")
        )
        content_type_lines = content_type_fingerprint_lines(read_text(archive, "[Content_Types].xml"))
        relationship_lines: list[str] = []
        for rels_path in [path for path in part_paths if is_relationship_part(path)]:
            rels_xml = read_text(archive, rels_path)
            if rels_xml is None:
                continue
            source_part = source_part_for_relationships(rels_path)
            target_modes = relationship_target_modes(rels_xml)
            for relationship in parse_relationships(rels_xml):
                target_mode = target_modes.get(relationship["id"], "Internal")
                resolved_target = (
                    relationship["target"]
                    if target_mode == "External"
                    else resolve_path(source_part, relationship["target"])
                )
                relationship_lines.append(
                    "\t".join(
                        [
                            source_part or "/",
                            relationship["type"],
                            target_mode,
                            relationship["target"],
                            resolved_target,
                        ]
                    )
                )
        preserved_part_paths = [path for path in part_paths if is_preserved_non_cell_part(path)]
        preserved_content_lines = [
            f"{path}\t{hashlib.sha256(archive.read(path)).hexdigest()}"
            for path in preserved_part_paths
        ]
        return {
            "partCount": len(part_paths),
            "partNamesHash": hash_lines(part_paths),
            "contentTypeCount": len(content_type_lines),
            "contentTypesHash": hash_lines(content_type_lines),
            "relationshipCount": len(relationship_lines),
            "relationshipGraphHash": hash_lines(relationship_lines),
            "preservedPartCount": len(preserved_part_paths),
            "preservedPartNamesHash": hash_lines(preserved_part_paths),
            "preservedPartContentHash": hash_lines(preserved_content_lines),
        }


def extract_feature_summary(data: bytes) -> dict[str, str | int]:
    with zipfile.ZipFile(BytesIO(data)) as archive:
        part_paths = sorted(
            info.filename for info in archive.infolist() if not info.filename.endswith("/")
        )
        workbook_xml = read_text(archive, "xl/workbook.xml") or ""
        workbook_rels_xml = read_text(archive, "xl/_rels/workbook.xml.rels") or ""
        workbook_rels = parse_relationships(workbook_rels_xml)
        worksheet_paths = [
            resolve_path("xl/workbook.xml", relationship["target"])
            for relationship in workbook_rels
            if relationship["type"]
            == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        ]
        feature_part_lines = [
            f"{kind}\t{path}"
            for path in part_paths
            if (kind := classify_feature_part(path)) is not None
        ]
        worksheet_feature_lines: list[str] = []
        for path in worksheet_paths:
            xml = read_text(archive, path)
            if xml is None:
                continue
            worksheet_feature_lines.extend(worksheet_feature_entries(path, xml))
        defined_name_lines = defined_name_entries(workbook_xml)
        shared_string_feature_lines = shared_string_feature_entries(
            read_text(archive, "xl/sharedStrings.xml") or ""
        )
        calc_chain_feature_lines = calc_chain_feature_entries(
            read_text(archive, "xl/calcChain.xml") or ""
        )
        inventory_lines = (
            feature_part_lines
            + worksheet_feature_lines
            + defined_name_lines
            + shared_string_feature_lines
            + calc_chain_feature_lines
        )
        return {
            "tablePartCount": count_paths(part_paths, r"^xl/tables/.+\.xml$"),
            "chartPartCount": count_paths(part_paths, r"^xl/charts/.+\.xml$"),
            "chartExPartCount": count_paths(part_paths, r"^xl/chartEx/.+\.xml$"),
            "drawingPartCount": count_paths(part_paths, r"^xl/drawings/.+\.xml$"),
            "vmlDrawingPartCount": count_paths(part_paths, r"^xl/drawings/.+\.vml$"),
            "pivotTablePartCount": count_paths(part_paths, r"^xl/pivotTables/.+\.xml$"),
            "pivotCachePartCount": count_paths(part_paths, r"^xl/pivotCache/.+\.xml$"),
            "slicerPartCount": count_paths(part_paths, r"^xl/(?:slicers|slicerCaches)/.+\.xml$"),
            "commentPartCount": count_paths(part_paths, r"^xl/(?:comments\d*|comments/.+)\.xml$"),
            "threadedCommentPartCount": count_paths(
                part_paths, r"^xl/threadedComments/.+\.xml$"
            ),
            "mediaPartCount": count_paths(part_paths, r"^xl/media/.+"),
            "externalLinkPartCount": count_paths(part_paths, r"^xl/externalLinks/.+\.xml$"),
            "connectionPartCount": sum(1 for path in part_paths if path == "xl/connections.xml"),
            "customXmlPartCount": count_paths(part_paths, r"^customXml/.+"),
            "worksheetHyperlinkCount": sum(
                1 for line in worksheet_feature_lines if line.startswith("worksheet-hyperlink\t")
            ),
            "worksheetDataValidationCount": sum(
                1
                for line in worksheet_feature_lines
                if line.startswith("worksheet-data-validation\t")
            ),
            "worksheetConditionalFormattingCount": sum(
                1
                for line in worksheet_feature_lines
                if line.startswith("worksheet-conditional-formatting\t")
            ),
            "definedNameCount": len(defined_name_lines),
            "featurePartNamesHash": hash_lines(feature_part_lines),
            "featureInventoryHash": hash_lines(inventory_lines),
        }


def count_paths(paths: list[str], pattern: str) -> int:
    return sum(1 for path in paths if re.match(pattern, path))


def classify_feature_part(path: str) -> str | None:
    if re.match(r"^xl/tables/.+\.xml$", path):
        return "table-part"
    if re.match(r"^xl/charts/.+\.xml$", path):
        return "chart-part"
    if re.match(r"^xl/chartEx/.+\.xml$", path):
        return "chart-ex-part"
    if re.match(r"^xl/drawings/.+\.xml$", path):
        return "drawing-part"
    if re.match(r"^xl/drawings/.+\.vml$", path):
        return "vml-drawing-part"
    if re.match(r"^xl/pivotTables/.+\.xml$", path):
        return "pivot-table-part"
    if re.match(r"^xl/pivotCache/.+\.xml$", path):
        return "pivot-cache-part"
    if re.match(r"^xl/(?:slicers|slicerCaches)/.+\.xml$", path):
        return "slicer-part"
    if re.match(r"^xl/(?:comments\d*|comments/.+)\.xml$", path):
        return "comment-part"
    if re.match(r"^xl/threadedComments/.+\.xml$", path):
        return "threaded-comment-part"
    if re.match(r"^xl/media/.+", path):
        return "media-part"
    if re.match(r"^xl/externalLinks/.+\.xml$", path):
        return "external-link-part"
    if path == "xl/connections.xml":
        return "connection-part"
    if path == "xl/calcChain.xml":
        return "calc-chain-part"
    if re.match(r"^customXml/.+", path):
        return "custom-xml-part"
    return None


def worksheet_feature_entries(path: str, xml: str) -> list[str]:
    lines: list[str] = []
    for match in re.finditer(r"<mergeCell\b([^>]*)/?>", xml):
        lines.append(f"worksheet-merge-cell\t{path}\t{feature_ref(match.group(1) or '')}")
    for match in re.finditer(r"<autoFilter\b([^>]*?)(?:/?>)", xml):
        lines.append(f"worksheet-auto-filter\t{path}\t{feature_ref(match.group(1) or '')}")
    for match in re.finditer(r"<sortState\b([^>]*?)(?:/?>)", xml):
        lines.append(
            f"worksheet-sort-state\t{path}\t{canonical_attributes(match.group(1) or '')}"
        )
    for match in re.finditer(r"<sheetProtection\b([^>]*)/?>", xml):
        lines.append(
            f"worksheet-sheet-protection\t{path}\t{canonical_attributes(match.group(1) or '')}"
        )
    for match in re.finditer(r"<sheetView\b([^>]*?)(?:/?>)", xml):
        lines.append(
            f"worksheet-sheet-view\t{path}\t{canonical_attributes(match.group(1) or '')}"
        )
    for match in re.finditer(r"<hyperlink\b([^>]*)/?>", xml):
        lines.append(f"worksheet-hyperlink\t{path}\t{feature_ref(match.group(1) or '')}")
    for match in re.finditer(r"<dataValidation\b([^>]*)/?>", xml):
        lines.append(f"worksheet-data-validation\t{path}\t{feature_sqref(match.group(1) or '')}")
    for match in re.finditer(r"<conditionalFormatting\b([^>]*)>", xml):
        lines.append(
            f"worksheet-conditional-formatting\t{path}\t{feature_sqref(match.group(1) or '')}"
        )
    return lines


def shared_string_feature_entries(xml: str) -> list[str]:
    lines: list[str] = []
    for index, match in enumerate(re.finditer(r"<si\b[^>]*>([\s\S]*?)</si>", xml)):
        inner = match.group(1) or ""
        if re.search(r"<r\b", inner):
            lines.append(f"shared-string-rich-text\t{index}\t{hash_text(inner)}")
    return lines


def calc_chain_feature_entries(xml: str) -> list[str]:
    lines: list[str] = []
    for match in re.finditer(r"<c\b([^>]*)/?>", xml):
        lines.append(f"calc-chain-cell\t{canonical_attributes(match.group(1) or '')}")
    return lines


def defined_name_entries(workbook_xml: str) -> list[str]:
    lines: list[str] = []
    defined_name_re = re.compile(
        r"<\s*(?:[A-Za-z_][\w.-]*:)?definedName\b([^>]*)>"
        r"([\s\S]*?)"
        r"</\s*(?:[A-Za-z_][\w.-]*:)?definedName\s*>"
    )
    for match in defined_name_re.finditer(workbook_xml):
        attrs = parse_xml_attributes(match.group(1) or "")
        lines.append(
            "\t".join(
                [
                    "defined-name",
                    attrs.get("name", ""),
                    attrs.get("localSheetId", ""),
                    decode_xml_text(match.group(2) or ""),
                ]
            )
        )
    return lines


def feature_ref(attrs_text: str) -> str:
    return parse_xml_attributes(attrs_text).get("ref", "")


def feature_sqref(attrs_text: str) -> str:
    return parse_xml_attributes(attrs_text).get("sqref", "")


def canonical_attributes(attrs_text: str) -> str:
    attrs = parse_xml_attributes(attrs_text)
    return ";".join(f"{key}={attrs[key]}" for key in sorted(attrs))


def hash_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def read_text(archive: zipfile.ZipFile, path: str) -> str | None:
    try:
        return archive.read(path).decode("utf-8")
    except KeyError:
        return None


def content_type_fingerprint_lines(xml: str | None) -> list[str]:
    if xml is None:
        return []
    lines: list[str] = []
    for match in re.finditer(r"<(Default|Override)\b([^>]*)/?>", xml):
        kind = match.group(1)
        attrs = parse_xml_attributes(match.group(2) or "")
        if kind == "Default":
            extension = attrs.get("Extension")
            content_type = attrs.get("ContentType")
            if extension and content_type:
                lines.append(f"Default\t{extension}\t{content_type}")
        else:
            part_name = attrs.get("PartName")
            content_type = attrs.get("ContentType")
            if part_name and content_type:
                lines.append(f"Override\t{part_name.lstrip('/')}\t{content_type}")
    return lines


def relationship_target_modes(xml: str) -> dict[str, str]:
    modes: dict[str, str] = {}
    for match in re.finditer(r"<Relationship\b([^>]*)/>", xml):
        attrs = parse_xml_attributes(match.group(1) or "")
        relationship_id = attrs.get("Id")
        if relationship_id:
            modes[relationship_id] = attrs.get("TargetMode", "Internal")
    return modes


def parse_relationships(xml: str) -> list[dict[str, str]]:
    relationships: list[dict[str, str]] = []
    for match in re.finditer(r"<Relationship\b([^>]*)/>", xml):
        attrs = parse_xml_attributes(match.group(1) or "")
        relationship_id = attrs.get("Id")
        relationship_type = attrs.get("Type")
        target = attrs.get("Target")
        if relationship_id and relationship_type and target:
            relationships.append(
                {"id": relationship_id, "type": normalize_relationship_type(relationship_type), "target": target}
            )
    return relationships


def normalize_relationship_type(relationship_type: str) -> str:
    if not relationship_type.startswith(STRICT_REL_PREFIX):
        return relationship_type
    suffix = relationship_type[len(STRICT_REL_PREFIX) :]
    if suffix == "sheetMetadata":
        return relationship_type
    if suffix == "extendedProperties":
        return f"{TRANSITIONAL_REL_PREFIX}extended-properties"
    return f"{TRANSITIONAL_REL_PREFIX}{suffix}"


def parse_xml_attributes(attrs: str) -> dict[str, str]:
    return {
        match.group(1): decode_xml_text(match.group(2))
        for match in re.finditer(r'([A-Za-z_][\w:.-]*)="([^"]*)"', attrs)
    }


def decode_xml_text(value: str) -> str:
    return (
        value.replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
    )


def is_relationship_part(path: str) -> bool:
    return path == "_rels/.rels" or path.endswith(".rels")


def source_part_for_relationships(rels_path: str) -> str:
    if rels_path == "_rels/.rels":
        return ""
    marker = "/_rels/"
    index = rels_path.rfind(marker)
    if index < 0 or not rels_path.endswith(".rels"):
        return ""
    return f"{rels_path[:index]}/{rels_path[index + len(marker):-len('.rels')]}"


def resolve_path(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base, target)).lstrip("./")


def is_preserved_non_cell_part(path: str) -> bool:
    if (
        path == "[Content_Types].xml"
        or path == "_rels/.rels"
        or path.endswith(".rels")
        or path == "xl/workbook.xml"
        or path == "xl/sharedStrings.xml"
        or path == "xl/calcChain.xml"
        or re.match(r"^xl/worksheets/sheet\d+\.xml$", path)
        or re.match(r"^docProps/(?:app|core)\.xml$", path)
    ):
        return False
    return True
