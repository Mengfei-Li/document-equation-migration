import hashlib
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


DETECTOR_VERSION = "0.1.0"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

NS_TO_PREFIX = {uri: prefix for prefix, uri in NS.items()}

STORY_TYPE_BY_PART = {
    "word/document.xml": "main",
    "word/footnotes.xml": "footnote",
    "word/endnotes.xml": "endnote",
    "word/comments.xml": "comment",
}


def local_name(tag: str) -> str:
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def qualified_name(tag: str) -> str:
    if tag.startswith("{"):
        uri, local = tag[1:].split("}", 1)
        prefix = NS_TO_PREFIX.get(uri)
        if prefix:
            return f"{prefix}:{local}"
        return local
    return tag


def parse_xml_from_zip(zf: zipfile.ZipFile, inner_path: str) -> ET.Element | None:
    try:
        data = zf.read(inner_path)
    except KeyError:
        return None
    return ET.fromstring(data)


def sha256_hex(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def story_type_for_part(part_name: str) -> str:
    if part_name in STORY_TYPE_BY_PART:
        return STORY_TYPE_BY_PART[part_name]
    if part_name.startswith("word/header") and part_name.endswith(".xml"):
        return "header"
    if part_name.startswith("word/footer") and part_name.endswith(".xml"):
        return "footer"
    return "other"


def iter_story_parts(zf: zipfile.ZipFile) -> list[str]:
    names = set(zf.namelist())
    ordered = []
    for name in (
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
    ):
        if name in names:
            ordered.append(name)

    for prefix in ("word/header", "word/footer"):
        ordered.extend(
            sorted(
                name
                for name in names
                if name.startswith(prefix) and name.endswith(".xml")
            )
        )

    return ordered


def has_math_properties(zf: zipfile.ZipFile) -> bool:
    root = parse_xml_from_zip(zf, "word/settings.xml")
    if root is None:
        return False
    return root.find(".//m:mathPr", NS) is not None


def build_xpath(path_tokens: list[tuple[str, int]]) -> str:
    return "/" + "/".join(f"{name}[{index}]" for name, index in path_tokens)


def make_formula_record(
    *,
    formula_id: str,
    part_name: str,
    story_type: str,
    paragraph_index: int | None,
    run_index: int | None,
    object_sequence: int,
    node: ET.Element,
    path_tokens: list[tuple[str, int]],
    has_math_pr: bool,
) -> dict:
    node_bytes = ET.tostring(node, encoding="utf-8")
    container_element = qualified_name(node.tag)
    display_mode = "display" if local_name(node.tag) == "oMathPara" else "inline"
    storage_kind = "omml-display" if display_mode == "display" else "omml-inline"
    risk_flags = []
    risk_level = "low"
    if story_type != "main":
        risk_flags.append("story-part-nonmain")
        risk_level = "medium"

    evidence_sources = [part_name]
    if has_math_pr:
        evidence_sources.append("word/settings.xml")

    return {
        "formula_id": formula_id,
        "source_family": "omml-native",
        "source_role": "native-source",
        "doc_part_path": part_name,
        "story_type": story_type,
        "storage_kind": storage_kind,
        "relationship_id": None,
        "embedding_target": None,
        "preview_target": None,
        "paragraph_index": paragraph_index,
        "run_index": run_index,
        "object_sequence": object_sequence,
        "canonical_mathml_status": "unverified",
        "omml_status": "available",
        "latex_status": "not-applicable",
        "risk_level": risk_level,
        "risk_flags": risk_flags,
        "failure_mode": None,
        "confidence": 0.99,
        "xpath": build_xpath(path_tokens),
        "provenance": {
            "prog_id_raw": None,
            "field_code_raw": None,
            "ole_stream_names": [],
            "raw_payload_status": "present",
            "raw_payload_sha256": sha256_hex(node_bytes),
            "transform_chain": None,
            "generator_id": None,
            "evidence_sources": evidence_sources,
        },
        "omml": {
            "container_element": container_element,
            "display_mode": display_mode,
            "has_mathPr": has_math_pr,
            "raw_omml_sha256": sha256_hex(node_bytes),
            "generated_from_formula_id": None,
            "math_child_count": len(node.findall(".//m:oMath", NS))
            if local_name(node.tag) == "oMathPara"
            else 1,
        },
    }


def scan_story_part(
    root: ET.Element,
    part_name: str,
    has_math_pr: bool,
    start_sequence: int,
) -> tuple[list[dict], int]:
    story_type = story_type_for_part(part_name)
    results: list[dict] = []
    sequence = start_sequence

    root_path = [(qualified_name(root.tag), 1)]
    paragraph_counter = 0

    def walk_container(
        node: ET.Element,
        path_tokens: list[tuple[str, int]],
        paragraph_index: int | None,
        run_counter: list[int],
        current_run_index: int | None,
    ) -> None:
        nonlocal sequence

        sibling_counts: dict[str, int] = {}
        for child in list(node):
            child_local = local_name(child.tag)
            sibling_counts[child_local] = sibling_counts.get(child_local, 0) + 1
            child_path = path_tokens + [
                (qualified_name(child.tag), sibling_counts[child_local])
            ]

            if child_local in {"oMath", "oMathPara"}:
                sequence += 1
                results.append(
                    make_formula_record(
                        formula_id=f"omml-native-{sequence:04d}",
                        part_name=part_name,
                        story_type=story_type,
                        paragraph_index=paragraph_index,
                        run_index=current_run_index,
                        object_sequence=sequence,
                        node=child,
                        path_tokens=child_path,
                        has_math_pr=has_math_pr,
                    )
                )
                continue

            if child_local == "r":
                run_counter[0] += 1
                walk_container(
                    child,
                    child_path,
                    paragraph_index,
                    run_counter,
                    run_counter[0],
                )
                continue

            walk_container(
                child,
                child_path,
                paragraph_index,
                run_counter,
                current_run_index,
            )

    sibling_counts: dict[str, int] = {}
    for child in list(root):
        child_local = local_name(child.tag)
        sibling_counts[child_local] = sibling_counts.get(child_local, 0) + 1
        child_path = root_path + [(qualified_name(child.tag), sibling_counts[child_local])]
        if child_local == "body":
            body_sibling_counts: dict[str, int] = {}
            for body_child in list(child):
                body_local = local_name(body_child.tag)
                body_sibling_counts[body_local] = body_sibling_counts.get(body_local, 0) + 1
                body_child_path = child_path + [
                    (qualified_name(body_child.tag), body_sibling_counts[body_local])
                ]
                if body_local == "p":
                    paragraph_counter += 1
                    walk_container(
                        body_child,
                        body_child_path,
                        paragraph_counter,
                        [0],
                        None,
                    )
                else:
                    walk_container(
                        body_child,
                        body_child_path,
                        paragraph_counter if paragraph_counter else None,
                        [0],
                        None,
                    )
            continue

        if child_local == "p":
            paragraph_counter += 1
            walk_container(child, child_path, paragraph_counter, [0], None)
            continue

        walk_container(
            child,
            child_path,
            paragraph_counter if paragraph_counter else None,
            [0],
            None,
        )

    return results, sequence


def detect_omml_native(docx_path: str | Path) -> dict:
    path = Path(docx_path).resolve()
    formulas: list[dict] = []

    with zipfile.ZipFile(path) as zf:
        math_pr_present = has_math_properties(zf)
        sequence = 0
        for part_name in iter_story_parts(zf):
            root = parse_xml_from_zip(zf, part_name)
            if root is None:
                continue
            story_formulas, sequence = scan_story_part(
                root=root,
                part_name=part_name,
                has_math_pr=math_pr_present,
                start_sequence=sequence,
            )
            formulas.extend(story_formulas)

    source_counts = {}
    if formulas:
        source_counts["omml-native"] = len(formulas)

    return {
        "document": {
            "input_path": str(path),
            "container_format": path.suffix.lstrip(".").lower(),
            "detector_version": DETECTOR_VERSION,
        },
        "source_counts": source_counts,
        "formulas": formulas,
    }
