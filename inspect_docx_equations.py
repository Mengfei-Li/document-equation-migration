import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


sys.stdout.reconfigure(encoding="utf-8")


BASE_DIR = Path(__file__).resolve().parent
SAMPLE_DIR = BASE_DIR / "sample_unzipped"
OUTPUT_JSON = BASE_DIR / "docx_inspection.json"
OUTPUT_TXT = BASE_DIR / "docx_inspection.txt"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

KEYWORDS = [
    b"MathType",
    b"Equation.DSMT",
    b"Design Science",
    b"MTEF",
    b"Equation Native",
    b"Wiris",
    b"MathML",
    b"LaTeX",
]


def parse_xml_from_zip(zf: zipfile.ZipFile, inner_path: str):
    try:
        data = zf.read(inner_path)
    except KeyError:
        return None
    return ET.fromstring(data)


def rels_map(zf: zipfile.ZipFile):
    root = parse_xml_from_zip(zf, "word/_rels/document.xml.rels")
    mapping = {}
    if root is None:
        return mapping
    for rel in root.findall("rel:Relationship", NS):
        mapping[rel.attrib.get("Id")] = {
            "type": rel.attrib.get("Type"),
            "target": rel.attrib.get("Target"),
        }
    return mapping


def summarize_bin(data: bytes):
    found = []
    for keyword in KEYWORDS:
        if keyword in data:
            found.append(keyword.decode("ascii", errors="ignore"))
    ascii_hits = re.findall(rb"[ -~]{8,}", data)
    ascii_preview = []
    for hit in ascii_hits:
        text = hit.decode("ascii", errors="ignore")
        if any(k.lower() in text.lower() for k in ["math", "equation", "design", "latex", "wiris"]):
            ascii_preview.append(text[:200])
        if len(ascii_preview) >= 8:
            break
    return {
        "size_bytes": len(data),
        "found_keywords": found,
        "ascii_preview": ascii_preview,
    }


def inspect_docx(path: Path):
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        document_root = parse_xml_from_zip(zf, "word/document.xml")
        rels = rels_map(zf)

        omath_count = 0
        omathpara_count = 0
        ole_nodes = []
        embed_targets = []
        if document_root is not None:
            omath_count = len(document_root.findall(".//m:oMath", NS))
            omathpara_count = len(document_root.findall(".//m:oMathPara", NS))
            for node in document_root.findall(".//o:OLEObject", NS):
                rid = node.attrib.get(f"{{{NS['r']}}}id")
                embed_targets.append(
                    {
                        "prog_id": node.attrib.get("ProgID"),
                        "shape_id": node.attrib.get("ShapeID"),
                        "draw_aspect": node.attrib.get("DrawAspect"),
                        "relationship_id": rid,
                        "relationship": rels.get(rid),
                    }
                )
            for node in document_root.findall(".//w:object", NS):
                ole_nodes.append(ET.tostring(node, encoding="unicode")[:500])

        embeddings = sorted([name for name in names if name.startswith("word/embeddings/")])
        embedding_summaries = []
        for name in embeddings[:12]:
            data = zf.read(name)
            embedding_summaries.append({"name": name, **summarize_bin(data)})

        return {
            "file_name": path.name,
            "path": str(path),
            "entry_count": len(names),
            "embedding_count": len(embeddings),
            "first_embeddings": embedding_summaries,
            "omath_count": omath_count,
            "omath_para_count": omathpara_count,
            "ole_object_count": len(embed_targets),
            "ole_objects_preview": embed_targets[:20],
            "word_object_nodes_preview": ole_nodes[:8],
        }


def main():
    results = []
    for path in sorted(SAMPLE_DIR.glob("*.docx")):
        results.append(inspect_docx(path))

    OUTPUT_JSON.write_text(
        json.dumps(results, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    lines = []
    for item in results:
        lines.append(f"File: {item['file_name']}")
        lines.append(f"Path: {item['path']}")
        lines.append(f"Package entry count: {item['entry_count']}")
        lines.append(f"OMML formula count: {item['omath_count']}")
        lines.append(f"OMML paragraph count: {item['omath_para_count']}")
        lines.append(f"OLE formula object count: {item['ole_object_count']}")
        lines.append(f"Embedded object count: {item['embedding_count']}")
        lines.append("OLE object preview:")
        for ole in item["ole_objects_preview"][:10]:
            lines.append(
                "  - "
                + json.dumps(ole, ensure_ascii=False)
            )
        lines.append("Embedded object binary preview:")
        for emb in item["first_embeddings"]:
            lines.append(
                "  - "
                + json.dumps(emb, ensure_ascii=False)
            )
        lines.append("")

    OUTPUT_TXT.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    main()
