import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


sys.stdout.reconfigure(encoding="utf-8")


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def local_name(tag: str) -> str:
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def parse_xml(zf: zipfile.ZipFile, name: str):
    return ET.fromstring(zf.read(name))


def build_rel_map(zf: zipfile.ZipFile):
    rels_root = parse_xml(zf, "word/_rels/document.xml.rels")
    rel_map = {}
    for rel in rels_root.findall("rel:Relationship", NS):
        rel_map[rel.attrib["Id"]] = {
            "type": rel.attrib.get("Type", ""),
            "target": rel.attrib.get("Target", ""),
        }
    return rel_map


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
        elif tag == "oMath":
            text_parts.append("[OMML]")
        elif tag == "object":
            text_parts.append("[OLE]")
    return "".join(text_parts)


def paragraph_runs(paragraph: ET.Element):
    return [child for child in list(paragraph) if local_name(child.tag) == "r"]


def paragraph_direct_omml_count(paragraph: ET.Element):
    return sum(1 for child in list(paragraph) if local_name(child.tag) in {"oMath", "oMathPara"})


def paragraph_text(paragraph: ET.Element):
    parts = []
    for child in list(paragraph):
        tag = local_name(child.tag)
        if tag == "r":
            parts.append(collect_run_text(child))
        elif tag in {"oMath", "oMathPara"}:
            parts.append("[OMML]")
    return "".join(parts)


def map_docx(docx_path: Path):
    with zipfile.ZipFile(docx_path) as zf:
        document_root = parse_xml(zf, "word/document.xml")
        rel_map = build_rel_map(zf)

    body = document_root.find("w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml is missing w:body")

    paragraphs = [node for node in body.iter() if local_name(node.tag) == "p"]
    rows = []
    sequence = 0
    omml_run_child_count = 0
    omml_paragraph_child_count = 0

    for paragraph_index, paragraph in enumerate(paragraphs, start=1):
        runs = paragraph_runs(paragraph)
        para_text = paragraph_text(paragraph)
        omml_paragraph_child_count += paragraph_direct_omml_count(paragraph)

        for run_index, run in enumerate(runs, start=1):
            run_text = collect_run_text(run)
            objects = [child for child in list(run) if local_name(child.tag) == "object"]
            omml_nodes = [child for child in list(run) if local_name(child.tag) == "oMath"]
            omml_run_child_count += len(omml_nodes)

            for obj in objects:
                sequence += 1
                ole_node = obj.find("o:OLEObject", NS)
                image_node = obj.find(".//v:imagedata", NS)
                rid = ole_node.attrib.get(f"{{{NS['r']}}}id", "") if ole_node is not None else ""
                image_rid = image_node.attrib.get(f"{{{NS['r']}}}id", "") if image_node is not None else ""
                rel = rel_map.get(rid, {})
                image_rel = rel_map.get(image_rid, {})

                left_text = "".join(collect_run_text(item) for item in runs[: run_index - 1])
                right_text = "".join(collect_run_text(item) for item in runs[run_index:])

                rows.append(
                    {
                        "sequence": sequence,
                        "paragraph_index": paragraph_index,
                        "run_index": run_index,
                        "run_text": run_text,
                        "paragraph_text": para_text,
                        "text_before": left_text[-60:],
                        "text_after": right_text[:60],
                        "object_id": ole_node.attrib.get("ObjectID", "") if ole_node is not None else "",
                        "shape_id": ole_node.attrib.get("ShapeID", "") if ole_node is not None else "",
                        "prog_id": ole_node.attrib.get("ProgID", "") if ole_node is not None else "",
                        "ole_rid": rid,
                        "ole_target": rel.get("target", ""),
                        "image_rid": image_rid,
                        "image_target": image_rel.get("target", ""),
                    }
                )

    return {
        "docx_path": str(docx_path),
        "paragraph_count": len(paragraphs),
        "ole_count": len(rows),
        "omml_run_child_count": omml_run_child_count,
        "omml_paragraph_child_count": omml_paragraph_child_count,
        "items": rows,
    }


def main():
    if len(sys.argv) not in {2, 3}:
        print("Usage: python docx_math_object_map.py <input.docx> [output.json]")
        sys.exit(1)

    docx_path = Path(sys.argv[1]).resolve()
    output_path = Path(sys.argv[2]).resolve() if len(sys.argv) == 3 else docx_path.with_suffix(".ole_map.json")
    result = map_docx(docx_path)
    output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
