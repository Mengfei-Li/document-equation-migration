from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
}

_LINE_FLOOR = 360


def _w(tag: str) -> str:
    return f"{{{NS['w']}}}{tag}"


def _parse_int(value: str | None, *, default: int) -> int:
    if value is None:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def _parse_style_pt(style: str, key: str) -> float | None:
    match = re.search(rf"(?:^|;)\s*{re.escape(key)}\s*:\s*([0-9.]+)\s*pt\b", style or "", re.I)
    if match is None:
        return None
    return float(match.group(1))


def collect_source_paragraph_max_heights(document_root: ET.Element) -> dict[int, float]:
    paragraph_heights: dict[int, float] = {}
    for paragraph_index, paragraph in enumerate(document_root.findall(".//w:p", NS), start=1):
        heights: list[float] = []
        for shape in paragraph.findall(".//v:shape", NS):
            height = _parse_style_pt(shape.attrib.get("style", ""), "height")
            if height is not None:
                heights.append(height)
        if heights:
            paragraph_heights[paragraph_index] = max(heights)
    return paragraph_heights


def load_source_paragraph_max_heights(input_docx: Path) -> dict[int, float]:
    with zipfile.ZipFile(input_docx) as zf:
        document_xml = zf.read("word/document.xml")
    document_root = ET.fromstring(document_xml)
    return collect_source_paragraph_max_heights(document_root)


def _ensure_spacing(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find(_w("pPr"))
    if ppr is None:
        ppr = ET.Element(_w("pPr"))
        paragraph.insert(0, ppr)
    spacing = ppr.find(_w("spacing"))
    if spacing is None:
        spacing = ET.Element(_w("spacing"))
        ppr.append(spacing)
    return spacing


def apply_layout_preservation(
    document_root: ET.Element,
    *,
    replaced_records: list[dict[str, object]],
    source_paragraph_max_heights: dict[int, float],
    factor: float = 1.01375,
) -> dict[str, object]:
    if factor < 1.0:
        raise ValueError(f"layout preservation factor must be >= 1.0, got {factor!r}")

    paragraphs = document_root.findall(".//w:p", NS)
    adjusted_records: list[dict[str, object]] = []
    adjusted_lines: list[int] = []
    paragraph_indices = sorted(
        {
            int(item["paragraph_index"])
            for item in replaced_records
            if int(item.get("paragraph_index", 0)) in source_paragraph_max_heights
        }
    )

    for paragraph_index in paragraph_indices:
        if paragraph_index < 1 or paragraph_index > len(paragraphs):
            continue
        paragraph = paragraphs[paragraph_index - 1]
        spacing = _ensure_spacing(paragraph)
        previous_line = _parse_int(spacing.attrib.get(_w("line")), default=_LINE_FLOOR)
        source_height_pt = source_paragraph_max_heights[paragraph_index]
        base_line = max(_LINE_FLOOR, previous_line, int(round(source_height_pt * 20)))
        adjusted_line = base_line
        if adjusted_line > _LINE_FLOOR:
            adjusted_line = max(_LINE_FLOOR, int(round(adjusted_line * factor)))
        spacing.set(_w("line"), str(adjusted_line))
        spacing.set(_w("lineRule"), "auto")
        adjusted_lines.append(adjusted_line)
        adjusted_records.append(
            {
                "paragraph_index": paragraph_index,
                "source_max_height_pt": source_height_pt,
                "previous_line": previous_line,
                "base_line": base_line,
                "adjusted_line": adjusted_line,
            }
        )

    return {
        "enabled": True,
        "factor": factor,
        "line_floor": _LINE_FLOOR,
        "adjusted_paragraph_count": len(adjusted_records),
        "line_min": min(adjusted_lines) if adjusted_lines else _LINE_FLOOR,
        "line_max": max(adjusted_lines) if adjusted_lines else _LINE_FLOOR,
        "line_mean": (sum(adjusted_lines) / len(adjusted_lines)) if adjusted_lines else float(_LINE_FLOOR),
        "paragraphs": adjusted_records,
    }
