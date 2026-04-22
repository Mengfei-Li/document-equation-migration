import argparse
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

SRC_DIR = Path(__file__).resolve().parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from document_equation_migration.mathtype_layout import (  # noqa: E402
    apply_layout_preservation,
    load_source_paragraph_max_heights,
)

sys.stdout.reconfigure(encoding="utf-8")


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def local_name(tag: str) -> str:
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def load_omml(omml_path: Path):
    raw = omml_path.read_text(encoding="utf-8").strip()
    if not raw:
        raise RuntimeError(f"OMML file is empty: {omml_path}")
    return ET.fromstring(raw)


def clone_element(node: ET.Element):
    return ET.fromstring(ET.tostring(node, encoding="utf-8"))


def collect_run_text(run: ET.Element):
    text_parts = []
    for node in run.iter():
        tag = local_name(node.tag)
        if tag == "t" and node.text:
            text_parts.append(node.text)
        elif tag == "tab":
            text_parts.append("\t")
        elif tag == "br":
            text_parts.append("\n")
    return "".join(text_parts)


def paragraph_inline_context(paragraph: ET.Element, run_index: int):
    runs = [child for child in list(paragraph) if local_name(child.tag) == "r"]
    left_text = "".join(collect_run_text(item) for item in runs[: run_index - 1]).strip()
    right_text = "".join(collect_run_text(item) for item in runs[run_index:]).strip()
    return bool(left_text or right_text)


def normalize_omml_root(omml_node: ET.Element, inline_context: bool):
    tag = local_name(omml_node.tag)

    if inline_context:
        if tag == "oMathPara":
            children = [child for child in list(omml_node) if local_name(child.tag) == "oMath"]
            if len(children) != 1:
                raise RuntimeError("inline OMML root is oMathPara but does not contain exactly one oMath child")
            return clone_element(children[0])
        return clone_element(omml_node)

    if tag == "oMath":
        wrapper = ET.Element(f"{{{NS['m']}}}oMathPara")
        wrapper.append(clone_element(omml_node))
        return wrapper

    return clone_element(omml_node)


def patch_root_namespaces(document_xml: Path, original_xml_text: str):
    output_text = document_xml.read_text(encoding="utf-8")
    root_match = re.search(r"<w:document\b[^>]*>", output_text)
    if root_match is None:
        return

    root_tag = root_match.group(0)
    original_root_match = re.search(r"<w:document\b[^>]*>", original_xml_text)
    if original_root_match is None:
        return

    original_root_tag = original_root_match.group(0)
    needed_prefixes = ["w14", "w15", "wp14"]

    additions = []
    for prefix in needed_prefixes:
        if f'xmlns:{prefix}=' in root_tag:
            continue
        original_decl = re.search(rf'xmlns:{prefix}="([^"]+)"', original_root_tag)
        if original_decl is None:
            continue
        additions.append(f' xmlns:{prefix}="{original_decl.group(1)}"')

    if not additions:
        return

    patched_root = root_tag[:-1] + "".join(additions) + ">"
    output_text = output_text.replace(root_tag, patched_root, 1)
    document_xml.write_text(output_text, encoding="utf-8")


def should_replace(ole_target: str, replacements: set[str]) -> bool:
    if not ole_target:
        return False
    name = Path(ole_target).name
    stem = Path(name).stem
    return name in replacements or stem in replacements


def replace_document_xml(document_xml: Path, rels_xml: Path, omml_dir: Path, replacements: set[str]):
    original_xml_text = document_xml.read_text(encoding="utf-8")
    doc_tree = ET.parse(document_xml)
    doc_root = doc_tree.getroot()
    rel_tree = ET.parse(rels_xml)
    rel_root = rel_tree.getroot()

    rel_map = {}
    for rel in rel_root:
        rel_map[rel.attrib.get("Id", "")] = rel.attrib.get("Target", "")

    replaced = []
    paragraphs = doc_root.findall(".//w:p", NS)
    for paragraph_index, paragraph in enumerate(paragraphs, start=1):
        runs = [child for child in list(paragraph) if local_name(child.tag) == "r"]
        for run_index, run in enumerate(runs, start=1):
            children = list(run)
            for child in children:
                if local_name(child.tag) != "object":
                    continue

                ole_node = child.find("o:OLEObject", NS)
                if ole_node is None:
                    continue

                rid = ole_node.attrib.get(f"{{{NS['r']}}}id", "")
                ole_target = rel_map.get(rid, "")
                if not should_replace(ole_target, replacements):
                    continue

                ole_stem = Path(ole_target).stem
                omml_path = omml_dir / f"{ole_stem}.omml.xml"
                if not omml_path.exists():
                    raise FileNotFoundError(f"Missing OMML file: {omml_path}")

                extra_children = [
                    local_name(node.tag) for node in children if local_name(node.tag) not in {"rPr", "object"}
                ]
                if extra_children:
                    raise RuntimeError(
                        f"Run contains mixed content and will not be replaced as a whole: {ole_target}, extra nodes={extra_children}"
                    )

                inline_context = paragraph_inline_context(paragraph, run_index)
                omml_node = normalize_omml_root(load_omml(omml_path), inline_context)
                paragraph_child_index = list(paragraph).index(run)
                paragraph.remove(run)
                paragraph.insert(paragraph_child_index, omml_node)
                replaced.append(
                    {
                        "ole_target": ole_target,
                        "ole_stem": ole_stem,
                        "paragraph_index": paragraph_index,
                        "run_index": run_index,
                        "inline_context": inline_context,
                        "inserted_tag": local_name(omml_node.tag),
                    }
                )

    return doc_tree, original_xml_text, replaced


def rezip_directory(source_dir: Path, output_docx: Path):
    if output_docx.exists():
        output_docx.unlink()
    with zipfile.ZipFile(output_docx, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(source_dir.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(source_dir))


def discover_replacements(omml_dir: Path, tokens: list[str]) -> set[str]:
    if tokens:
        return set(tokens)
    return {path.stem.replace(".omml", "") for path in omml_dir.glob("*.omml.xml")}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Replace MathType OLE instances with OMML in a DOCX.")
    parser.add_argument("input_docx", help="Source DOCX path.")
    parser.add_argument("omml_dir", help="Directory containing generated .omml.xml files.")
    parser.add_argument("output_docx", help="Output DOCX path.")
    parser.add_argument("replacements", nargs="*", help="Optional subset of oleObject stems to replace.")
    parser.add_argument(
        "--preserve-mathtype-layout",
        action="store_true",
        help="Apply guarded experimental line-height compensation from source OLE/VML heights.",
    )
    parser.add_argument(
        "--mathtype-layout-factor",
        type=float,
        default=1.01375,
        help="Multiplier for layout-preservation line values greater than 360.",
    )
    return parser


def main():
    args = build_parser().parse_args()

    input_docx = Path(args.input_docx).resolve()
    omml_dir = Path(args.omml_dir).resolve()
    output_docx = Path(args.output_docx).resolve()
    replacements = discover_replacements(omml_dir, args.replacements)
    source_paragraph_max_heights: dict[int, float] = {}
    if args.preserve_mathtype_layout:
        source_paragraph_max_heights = load_source_paragraph_max_heights(input_docx)

    with tempfile.TemporaryDirectory(prefix="docx_omml_replace_") as tmp:
        temp_dir = Path(tmp)
        with zipfile.ZipFile(input_docx) as zf:
            zf.extractall(temp_dir)

        doc_tree, original_xml_text, replaced = replace_document_xml(
            temp_dir / "word" / "document.xml",
            temp_dir / "word" / "_rels" / "document.xml.rels",
            omml_dir,
            replacements,
        )
        layout_summary: dict[str, object] | None = None
        if args.preserve_mathtype_layout:
            layout_summary = apply_layout_preservation(
                doc_tree.getroot(),
                replaced_records=replaced,
                source_paragraph_max_heights=source_paragraph_max_heights,
                factor=args.mathtype_layout_factor,
            )
        document_xml_path = temp_dir / "word" / "document.xml"
        doc_tree.write(document_xml_path, encoding="utf-8", xml_declaration=True)
        patch_root_namespaces(document_xml_path, original_xml_text)
        rezip_directory(temp_dir, output_docx)

    summary = {
        "input_docx": str(input_docx),
        "output_docx": str(output_docx),
        "replaced_count": len(replaced),
        "replaced": replaced,
        "layout_preservation": layout_summary
        if args.preserve_mathtype_layout
        else {"enabled": False, "factor": args.mathtype_layout_factor},
    }
    summary_path = output_docx.with_suffix(output_docx.suffix + ".replace_summary.json")
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
