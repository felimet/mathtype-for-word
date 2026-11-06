from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
DISPLAY_STYLE_PREFIX = "MTDisplayEquation"


def qn(local: str) -> str:
    return f"{{{W}}}{local}"


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def ensure_child(parent: etree._Element, tag: str, order: list[str]) -> etree._Element:
    existing = parent.find(f"w:{tag}", NS)
    if existing is not None:
        return existing
    child = etree.Element(qn(tag))
    rank = order.index(tag)
    for index, sibling in enumerate(parent):
        local = etree.QName(sibling).localname
        if local in order and order.index(local) > rank:
            parent.insert(index, child)
            return child
    parent.append(child)
    return child


PPR_ORDER = [
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
]

RPR_ORDER = [
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
]


def is_inline_math_paragraph(paragraph: etree._Element) -> bool:
    if not paragraph.findall(".//w:object", NS):
        return False
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    style_name = style.get(qn("val"), "") if style is not None else ""
    return not style_name.startswith(DISPLAY_STYLE_PREFIX)


def style_line_spacing(styles_xml: bytes) -> dict[str, int]:
    root = etree.fromstring(styles_xml)
    values: dict[str, int] = {}
    for style in root.xpath(".//w:style", namespaces=NS):
        style_id = style.get(qn("styleId"), "")
        spacing = style.find("./w:pPr/w:spacing", NS)
        if style_id and spacing is not None and spacing.get(qn("line")):
            values[style_id] = int(spacing.get(qn("line")))
    return values


def effective_line_spacing(
    paragraph: etree._Element, style_spacing: dict[str, int]
) -> int:
    spacing = paragraph.find("./w:pPr/w:spacing", NS)
    if spacing is not None and spacing.get(qn("line")):
        return int(spacing.get(qn("line")))
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    style_id = style.get(qn("val"), "") if style is not None else ""
    return style_spacing.get(style_id, 240)


def patch_document(
    document_xml: bytes, styles_xml: bytes
) -> tuple[bytes, dict]:
    root = etree.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("document.xml has no w:body")
    style_spacing = style_line_spacing(styles_xml)

    where_changed: list[int] = []
    inline_changed: list[int] = []
    run_position_changed = 0
    line_spacing_relaxed = 0

    for index, paragraph in enumerate(body.xpath(".//w:p", namespaces=NS), 1):
        text = paragraph_text(paragraph)
        ppr = paragraph.find("w:pPr", NS)
        if ppr is None:
            ppr = etree.Element(qn("pPr"))
            paragraph.insert(0, ppr)

        if text.startswith("其中，"):
            indent = ensure_child(ppr, "ind", PPR_ORDER)
            indent.set(qn("firstLine"), "0")
            indent.attrib.pop(qn("firstLineChars"), None)
            where_changed.append(index)

        if not is_inline_math_paragraph(paragraph):
            continue

        spacing = ensure_child(ppr, "spacing", PPR_ORDER)
        spacing.set(qn("line"), str(effective_line_spacing(paragraph, style_spacing)))
        if spacing.get(qn("lineRule")) != "atLeast":
            line_spacing_relaxed += 1
        spacing.set(qn("lineRule"), "atLeast")

        text_alignment = ensure_child(ppr, "textAlignment", PPR_ORDER)
        text_alignment.set(qn("val"), "center")

        for run in paragraph.findall(".//w:r", NS):
            if run.find("w:object", NS) is None:
                continue
            rpr = run.find("w:rPr", NS)
            if rpr is None:
                rpr = etree.Element(qn("rPr"))
                run.insert(0, rpr)
            position = ensure_child(rpr, "position", RPR_ORDER)
            if position.get(qn("val")) != "0":
                run_position_changed += 1
            position.set(qn("val"), "0")
        inline_changed.append(index)

    report = {
        "where_paragraph_count": len(where_changed),
        "where_paragraph_indexes": where_changed,
        "inline_math_paragraph_count": len(inline_changed),
        "inline_math_paragraph_indexes": inline_changed,
        "inline_object_run_positions_normalized": run_position_changed,
        "inline_paragraphs_line_spacing_relaxed": line_spacing_relaxed,
    }
    return (
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
        report,
    )


def inspect_document(document_xml: bytes, styles_xml: bytes) -> dict:
    root = etree.fromstring(document_xml)
    style_spacing = style_line_spacing(styles_xml)
    where_nonexact: list[int] = []
    inline_not_center: list[int] = []
    inline_position_nonzero: list[int] = []
    inline_not_at_least: list[int] = []

    for index, paragraph in enumerate(
        root.xpath(".//w:body//w:p", namespaces=NS), 1
    ):
        text = paragraph_text(paragraph)
        ppr = paragraph.find("w:pPr", NS)
        indent = ppr.find("w:ind", NS) if ppr is not None else None
        if text.startswith("其中，"):
            first = indent.get(qn("firstLine")) if indent is not None else None
            first_chars = (
                indent.get(qn("firstLineChars")) if indent is not None else None
            )
            if first not in (None, "0") or first_chars not in (None, "0"):
                where_nonexact.append(index)

        if not is_inline_math_paragraph(paragraph):
            continue
        alignment = ppr.find("w:textAlignment", NS) if ppr is not None else None
        if alignment is None or alignment.get(qn("val")) != "center":
            inline_not_center.append(index)
        spacing = ppr.find("w:spacing", NS) if ppr is not None else None
        expected_line = effective_line_spacing(paragraph, style_spacing)
        if (
            spacing is None
            or spacing.get(qn("lineRule")) != "atLeast"
            or int(spacing.get(qn("line"), "0")) < expected_line
        ):
            inline_not_at_least.append(index)
        for run in paragraph.findall(".//w:r", NS):
            if run.find("w:object", NS) is None:
                continue
            position = run.find("w:rPr/w:position", NS)
            if position is None or position.get(qn("val")) not in (None, "0"):
                inline_position_nonzero.append(index)
                break

    return {
        "where_paragraphs_with_indent": where_nonexact,
        "inline_paragraphs_not_center": inline_not_center,
        "inline_paragraphs_with_nonstandard_position": inline_position_nonzero,
        "inline_paragraphs_without_minimum_line_spacing": inline_not_at_least,
        "ok": not where_nonexact
        and not inline_not_center
        and not inline_position_nonzero
        and not inline_not_at_least,
    }


def rewrite_docx(input_path: Path, output_path: Path) -> tuple[dict, dict, dict]:
    with ZipFile(input_path) as source:
        document_xml = source.read("word/document.xml")
        styles_xml = source.read("word/styles.xml")
        patched_xml, report = patch_document(document_xml, styles_xml)
        validation = inspect_document(patched_xml, styles_xml)
        if not validation["ok"]:
            raise ValueError(f"Post-patch validation failed: {validation}")

        ole_before = {
            name: hashlib.sha256(source.read(name)).hexdigest()
            for name in source.namelist()
            if name.startswith("word/embeddings/")
        }
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            prefix=f".{output_path.stem}.",
            suffix=".tmp.docx",
            dir=output_path.parent,
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
        try:
            with ZipFile(temporary, "w", ZIP_DEFLATED) as target:
                for item in source.infolist():
                    payload = (
                        patched_xml
                        if item.filename == "word/document.xml"
                        else source.read(item.filename)
                    )
                    target.writestr(item, payload)
            os.replace(temporary, output_path)
        finally:
            temporary.unlink(missing_ok=True)

    with ZipFile(output_path) as output:
        ole_after = {
            name: hashlib.sha256(output.read(name)).hexdigest()
            for name in output.namelist()
            if name.startswith("word/embeddings/")
        }
    if ole_before != ole_after:
        raise ValueError("OLE embedded parts changed during inline-math normalization")
    return report, validation, {"ole_parts_unchanged": True, "ole_count": len(ole_after)}


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Normalize Chinese inline MathType paragraph alignment without rebuilding OLE objects."
    )
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path, nargs="?")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()

    input_path = args.input.resolve()
    if args.check:
        with ZipFile(input_path) as archive:
            result = inspect_document(
                archive.read("word/document.xml"),
                archive.read("word/styles.xml"),
            )
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if not result["ok"]:
            raise SystemExit(1)
        return

    if args.output is None:
        parser.error("output is required unless --check is used")
    output_path = args.output.resolve()
    if input_path == output_path:
        parser.error("input and output must differ; use a staged file and atomic replacement")

    report, validation, preservation = rewrite_docx(input_path, output_path)
    print(
        json.dumps(
            {
                "input": str(input_path),
                "output": str(output_path),
                "changes": report,
                "validation": validation,
                "preservation": preservation,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
