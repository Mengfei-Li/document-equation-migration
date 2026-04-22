from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from document_equation_migration.mathtype_layout import (
    apply_layout_preservation,
    collect_source_paragraph_max_heights,
    load_source_paragraph_max_heights,
)


NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_V = "urn:schemas-microsoft-com:vml"


SOURCE_XML = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{NS_W}" xmlns:v="{NS_V}">
  <w:body>
    <w:p>
      <w:r><w:object><v:shape style="width:42pt;height:18pt" /></w:object></w:r>
    </w:p>
    <w:p>
      <w:r><w:object><v:shape style="width:56pt;height:24pt" /></w:object></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:line="390" w:lineRule="auto" /></w:pPr>
      <w:r><w:object><v:shape style="width:50pt;height:18pt" /></w:object></w:r>
    </w:p>
  </w:body>
</w:document>
"""


TARGET_XML = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{NS_W}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto" /></w:pPr><m:oMath /></w:p>
    <w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto" /></w:pPr><m:oMath /></w:p>
    <w:p><w:pPr><w:spacing w:line="390" w:lineRule="auto" /></w:pPr><m:oMath /></w:p>
  </w:body>
</w:document>
"""


def test_collect_source_paragraph_max_heights() -> None:
    root = ET.fromstring(SOURCE_XML)
    heights = collect_source_paragraph_max_heights(root)
    assert heights == {1: 18.0, 2: 24.0, 3: 18.0}


def test_load_source_paragraph_max_heights_from_docx(tmp_path: Path) -> None:
    docx_path = tmp_path / "sample.docx"
    with zipfile.ZipFile(docx_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", SOURCE_XML)

    heights = load_source_paragraph_max_heights(docx_path)

    assert heights == {1: 18.0, 2: 24.0, 3: 18.0}


def test_apply_layout_preservation_uses_source_heights_and_factor() -> None:
    target_root = ET.fromstring(TARGET_XML)
    summary = apply_layout_preservation(
        target_root,
        replaced_records=[
            {"paragraph_index": 1, "inserted_tag": "oMath"},
            {"paragraph_index": 2, "inserted_tag": "oMath"},
            {"paragraph_index": 3, "inserted_tag": "oMath"},
        ],
        source_paragraph_max_heights={1: 18.0, 2: 24.0, 3: 18.0},
        factor=1.01375,
    )

    paragraphs = target_root.findall(f".//{{{NS_W}}}p")
    spacing_1 = paragraphs[0].find(f".//{{{NS_W}}}spacing")
    spacing_2 = paragraphs[1].find(f".//{{{NS_W}}}spacing")
    spacing_3 = paragraphs[2].find(f".//{{{NS_W}}}spacing")

    assert spacing_1.attrib[f"{{{NS_W}}}line"] == "360"
    assert spacing_2.attrib[f"{{{NS_W}}}line"] == "487"
    assert spacing_3.attrib[f"{{{NS_W}}}line"] == "395"
    assert summary["adjusted_paragraph_count"] == 3
    assert summary["line_min"] == 360
    assert summary["line_max"] == 487
