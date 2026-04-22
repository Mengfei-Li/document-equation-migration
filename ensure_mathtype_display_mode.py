"""Normalize MathType intermediate XML before MathML conversion.

Older MTEF3 payloads decoded by the third-party MathType parser can omit the
``equation_options`` element that the downstream MathML XSLT uses to choose an
entry template. This helper adds a conservative display-mode default only for
that missing MTEF3 case.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from xml.etree import ElementTree as ET


def _text(parent: ET.Element, child_name: str) -> str:
    child = parent.find(child_name)
    return "" if child is None or child.text is None else child.text.strip()


def ensure_display_mode(xml_path: Path, default: str = "block") -> bool:
    """Add ``equation_options`` to MTEF3 XML when it is missing.

    Returns ``True`` when the file was changed.
    """

    tree = ET.parse(xml_path)
    root = tree.getroot()
    mtef = root.find(".//mtef")
    if mtef is None:
        return False

    if mtef.find("equation_options") is not None:
        return False

    if _text(mtef, "mtef_version") != "3":
        return False

    equation_options = ET.Element("equation_options")
    equation_options.text = default

    insert_at = 0
    for index, child in enumerate(list(mtef)):
        insert_at = index + 1
        if child.tag == "product_subversion":
            break

    mtef.insert(insert_at, equation_options)
    tree.write(xml_path, encoding="utf-8", xml_declaration=True)
    return True


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Add a missing MathType display-mode marker to MTEF3 XML."
    )
    parser.add_argument("xml_path", help="Intermediate XML file produced from MathType OLE.")
    parser.add_argument(
        "--default",
        choices=("block", "inline"),
        default="block",
        help="Display mode to add when MTEF3 XML omits equation_options.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    changed = ensure_display_mode(Path(args.xml_path), default=args.default)
    print("updated" if changed else "unchanged")


if __name__ == "__main__":
    main()
