from pathlib import Path

from ensure_mathtype_display_mode import ensure_display_mode


def test_adds_block_display_mode_to_mtef3_without_equation_options(tmp_path: Path) -> None:
    xml_path = tmp_path / "mtef3.xml"
    xml_path.write_text(
        """<?xml version="1.0"?>
<root>
  <mtef>
    <mtef_version>3</mtef_version>
    <platform>1</platform>
    <product>1</product>
    <product_version>3</product_version>
    <product_subversion>10</product_subversion>
    <slot />
  </mtef>
</root>
""",
        encoding="utf-8",
    )

    assert ensure_display_mode(xml_path) is True
    text = xml_path.read_text(encoding="utf-8")

    assert "<equation_options>block</equation_options>" in text
    assert text.index("<product_subversion>10</product_subversion>") < text.index(
        "<equation_options>block</equation_options>"
    )


def test_leaves_existing_display_mode_unchanged(tmp_path: Path) -> None:
    xml_path = tmp_path / "mtef3-inline.xml"
    original = """<?xml version="1.0"?>
<root><mtef><mtef_version>3</mtef_version><equation_options>inline</equation_options></mtef></root>
"""
    xml_path.write_text(original, encoding="utf-8")

    assert ensure_display_mode(xml_path) is False
    assert xml_path.read_text(encoding="utf-8") == original


def test_leaves_mtef5_without_display_mode_unchanged(tmp_path: Path) -> None:
    xml_path = tmp_path / "mtef5.xml"
    original = """<?xml version="1.0"?>
<root><mtef><mtef_version>5</mtef_version><slot /></mtef></root>
"""
    xml_path.write_text(original, encoding="utf-8")

    assert ensure_display_mode(xml_path) is False
    assert xml_path.read_text(encoding="utf-8") == original
