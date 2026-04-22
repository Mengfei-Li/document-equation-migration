import hashlib
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import olefile


DETECTOR_VERSION = "0.1.0"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

NS_TO_PREFIX = {uri: prefix for prefix, uri in NS.items()}

STORY_TYPE_BY_PART = {
    "word/document.xml": "main",
    "word/footnotes.xml": "footnote",
    "word/endnotes.xml": "endnote",
    "word/comments.xml": "comment",
}

MTEF_VERSION_PATTERN = re.compile(rb"MTEF\s*([0-9]+)|MTEF([0-9]+)")
APPLICATION_KEY_PATTERN = re.compile(rb"DSMT[0-9]+")
PRODUCT_VERSION_PATTERN = re.compile(rb"ProductVersion=([0-9]+)")
PRODUCT_SUBVERSION_PATTERN = re.compile(rb"ProductSubVersion=([0-9]+)")


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


def rels_path_for_part(part_name: str) -> str:
    parent = posixpath.dirname(part_name)
    leaf = posixpath.basename(part_name)
    if not parent:
        return f"_rels/{leaf}.rels"
    return f"{parent}/_rels/{leaf}.rels"


def resolve_rel_target(part_name: str, target: str | None) -> str | None:
    if not target:
        return None
    if "://" in target:
        return target
    parent = posixpath.dirname(part_name)
    normalized = posixpath.normpath(posixpath.join(parent, target))
    return normalized.lstrip("./")


def build_rel_map(zf: zipfile.ZipFile, part_name: str) -> dict[str, dict]:
    rels_root = parse_xml_from_zip(zf, rels_path_for_part(part_name))
    if rels_root is None:
        return {}

    rel_map = {}
    for rel in rels_root.findall("rel:Relationship", NS):
        rel_id = rel.attrib.get("Id")
        if not rel_id:
            continue
        raw_target = rel.attrib.get("Target", "")
        rel_map[rel_id] = {
            "type": rel.attrib.get("Type", ""),
            "target": raw_target,
            "resolved_target": resolve_rel_target(part_name, raw_target),
        }
    return rel_map


def build_xpath(path_tokens: list[tuple[str, int]]) -> str:
    return "/" + "/".join(f"{name}[{index}]" for name, index in path_tokens)


def build_path_index(root: ET.Element) -> dict[int, list[tuple[str, int]]]:
    path_index: dict[int, list[tuple[str, int]]] = {}

    def walk(node: ET.Element, path_tokens: list[tuple[str, int]]) -> None:
        path_index[id(node)] = path_tokens
        sibling_counts: dict[str, int] = {}
        for child in list(node):
            child_name = qualified_name(child.tag)
            sibling_counts[child_name] = sibling_counts.get(child_name, 0) + 1
            child_path = path_tokens + [(child_name, sibling_counts[child_name])]
            walk(child, child_path)

    walk(root, [(qualified_name(root.tag), 1)])
    return path_index


def collect_run_text(run: ET.Element) -> str:
    text_parts = []
    for node in run.iter():
        tag = local_name(node.tag)
        if tag == "t" and node.text:
            text_parts.append(node.text)
        elif tag == "tab":
            text_parts.append("\t")
        elif tag == "br":
            text_parts.append("\n")
        elif tag == "object":
            text_parts.append("[OLE]")
    return "".join(text_parts)


def paragraph_runs(paragraph: ET.Element) -> list[ET.Element]:
    return [child for child in list(paragraph) if local_name(child.tag) == "r"]


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(collect_run_text(run) for run in paragraph_runs(paragraph))


def paragraph_field_code(paragraph: ET.Element) -> str | None:
    field_parts = []
    for node in paragraph.iter():
        if local_name(node.tag) == "instrText" and node.text:
            field_parts.append(node.text)
    field_code = " ".join(part.strip() for part in field_parts if part.strip()).strip()
    return field_code or None


def add_marker(markers: list[str], marker: str) -> None:
    if marker not in markers:
        markers.append(marker)


def inspect_embedded_payload(
    zf: zipfile.ZipFile,
    embedding_target: str | None,
) -> dict:
    result = {
        "raw_payload_status": "missing",
        "raw_payload_sha256": None,
        "is_ole": False,
        "stream_names": [],
        "equation_native_stream_exists": False,
        "equation_native_size_bytes": None,
        "mtef_version": None,
        "application_key": None,
        "product_version": None,
        "product_subversion": None,
        "marker_hits": [],
        "parse_error": None,
    }

    if not embedding_target:
        return result

    try:
        payload = zf.read(embedding_target)
    except KeyError:
        return result

    result["raw_payload_sha256"] = sha256_hex(payload)
    result["is_ole"] = olefile.isOleFile(payload)
    marker_source = payload[:8192]

    if not result["is_ole"]:
        result["raw_payload_status"] = "corrupt"
    else:
        try:
            ole = olefile.OleFileIO(payload)
            try:
                stream_paths = ole.listdir()
                result["stream_names"] = ["/".join(path) for path in stream_paths]

                equation_native_path = next(
                    (
                        path
                        for path in stream_paths
                        if path and path[-1].lower() == "equation native"
                    ),
                    None,
                )
                if equation_native_path is not None:
                    equation_native_bytes = ole.openstream(equation_native_path).read()
                    result["equation_native_stream_exists"] = True
                    result["equation_native_size_bytes"] = len(equation_native_bytes)
                    marker_source = equation_native_bytes

                comp_obj_path = next(
                    (
                        path
                        for path in stream_paths
                        if path and path[-1].lower().endswith("compobj")
                    ),
                    None,
                )
                if comp_obj_path is not None:
                    comp_obj_bytes = ole.openstream(comp_obj_path).read()
                    marker_source += b"\x00" + comp_obj_bytes[:1024]
            finally:
                ole.close()
            result["raw_payload_status"] = "present"
        except Exception as exc:  # pragma: no cover - defensive branch
            result["raw_payload_status"] = "corrupt"
            result["parse_error"] = str(exc)

    if b"Equation Native" in marker_source:
        add_marker(result["marker_hits"], "Equation Native")
    if b"MathType EF" in marker_source:
        add_marker(result["marker_hits"], "MathType EF")

    mtef_match = MTEF_VERSION_PATTERN.search(marker_source)
    if mtef_match:
        version_text = mtef_match.group(1) or mtef_match.group(2)
        result["mtef_version"] = int(version_text.decode("ascii"))
        add_marker(result["marker_hits"], f"MTEF{result['mtef_version']}")

    application_key_match = APPLICATION_KEY_PATTERN.search(marker_source)
    if application_key_match:
        result["application_key"] = application_key_match.group().decode("ascii")
        add_marker(result["marker_hits"], result["application_key"])

    product_version_match = PRODUCT_VERSION_PATTERN.search(marker_source)
    if product_version_match:
        result["product_version"] = int(product_version_match.group(1).decode("ascii"))

    product_subversion_match = PRODUCT_SUBVERSION_PATTERN.search(marker_source)
    if product_subversion_match:
        result["product_subversion"] = int(
            product_subversion_match.group(1).decode("ascii")
        )

    return result


def candidate_is_mathtype(
    prog_id_raw: str | None,
    field_code_raw: str | None,
    payload_info: dict,
) -> tuple[bool, list[str]]:
    basis = []
    upper_prog_id = (prog_id_raw or "").upper()
    upper_field_code = (field_code_raw or "").upper()

    if upper_prog_id.startswith("EQUATION.DSMT"):
        basis.append("prog-id")
    if "EQUATION.DSMT" in upper_field_code:
        basis.append("field-code")
    if payload_info["equation_native_stream_exists"]:
        basis.append("equation-native-stream")
    if payload_info["application_key"]:
        basis.append("application-key")
    if "MathType EF" in payload_info["marker_hits"]:
        basis.append("mathtype-ef-marker")

    matched = (
        "prog-id" in basis
        or "field-code" in basis
        or (
            "equation-native-stream" in basis
            and (
                "application-key" in basis or "mathtype-ef-marker" in basis
            )
        )
    )
    return matched, basis


def build_confidence(story_type: str, payload_info: dict, basis: list[str]) -> float:
    confidence = 0.78
    if "prog-id" in basis:
        confidence += 0.14
    if "field-code" in basis:
        confidence += 0.03
    if "equation-native-stream" in basis:
        confidence += 0.03
    if "application-key" in basis:
        confidence += 0.01
    if "mathtype-ef-marker" in basis:
        confidence += 0.01
    if story_type != "main":
        confidence -= 0.03
    if payload_info["raw_payload_status"] != "present":
        confidence -= 0.12
    return round(max(0.0, min(confidence, 0.99)), 2)


def make_formula_record(
    *,
    formula_id: str,
    part_name: str,
    story_type: str,
    paragraph_index: int,
    run_index: int,
    object_sequence: int,
    xpath: str,
    relationship_id: str | None,
    embedding_target: str | None,
    preview_target: str | None,
    paragraph_text_raw: str,
    text_before: str,
    text_after: str,
    prog_id_raw: str | None,
    field_code_raw: str | None,
    payload_info: dict,
    evidence_sources: list[str],
    basis: list[str],
) -> dict:
    risk_flags = []
    risk_level = "low"
    failure_mode = None

    if story_type != "main":
        risk_flags.append("story-part-nonmain")
        risk_level = "medium"

    if payload_info["raw_payload_status"] != "present":
        risk_flags.append("invalid-ole-payload")
        risk_level = "high"
        failure_mode = "invalid-ole-payload"
    elif not payload_info["equation_native_stream_exists"]:
        risk_flags.append("missing-equation-native-stream")
        risk_level = "high"
        failure_mode = "missing-equation-native-stream"

    canonical_mathml_status = (
        "unverified" if payload_info["equation_native_stream_exists"] else "missing"
    )

    return {
        "formula_id": formula_id,
        "source_family": "mathtype-ole",
        "source_role": "native-source",
        "doc_part_path": part_name,
        "story_type": story_type,
        "storage_kind": "ole-embedded",
        "relationship_id": relationship_id,
        "embedding_target": embedding_target,
        "preview_target": preview_target,
        "paragraph_index": paragraph_index,
        "run_index": run_index,
        "object_sequence": object_sequence,
        "canonical_mathml_status": canonical_mathml_status,
        "omml_status": "missing",
        "latex_status": "missing",
        "risk_level": risk_level,
        "risk_flags": risk_flags,
        "failure_mode": failure_mode,
        "confidence": build_confidence(story_type, payload_info, basis),
        "xpath": xpath,
        "paragraph_text": paragraph_text_raw,
        "text_before": text_before,
        "text_after": text_after,
        "provenance": {
            "prog_id_raw": prog_id_raw,
            "field_code_raw": field_code_raw,
            "ole_stream_names": payload_info["stream_names"],
            "raw_payload_status": payload_info["raw_payload_status"],
            "raw_payload_sha256": payload_info["raw_payload_sha256"],
            "transform_chain": None,
            "generator_id": None,
            "evidence_sources": evidence_sources,
        },
        "mathtype": {
            "equation_native_stream_exists": payload_info["equation_native_stream_exists"],
            "equation_native_size_bytes": payload_info["equation_native_size_bytes"],
            "mtef_version": payload_info["mtef_version"],
            "application_key": payload_info["application_key"],
            "product_version": payload_info["product_version"],
            "product_subversion": payload_info["product_subversion"],
        },
    }


def scan_story_part(
    *,
    zf: zipfile.ZipFile,
    root: ET.Element,
    part_name: str,
    rel_map: dict[str, dict],
    start_sequence: int,
) -> tuple[list[dict], int]:
    story_type = story_type_for_part(part_name)
    path_index = build_path_index(root)
    paragraphs = [node for node in root.iter() if local_name(node.tag) == "p"]
    results = []
    sequence = start_sequence

    for paragraph_index, paragraph in enumerate(paragraphs, start=1):
        runs = paragraph_runs(paragraph)
        para_text = paragraph_text(paragraph)
        field_code_raw = paragraph_field_code(paragraph)
        run_texts = [collect_run_text(run) for run in runs]

        for run_index, run in enumerate(runs, start=1):
            objects = [child for child in list(run) if local_name(child.tag) == "object"]
            if not objects:
                continue

            text_before = "".join(run_texts[: run_index - 1])[-60:]
            text_after = "".join(run_texts[run_index:])[:60]

            for object_node in objects:
                ole_node = object_node.find("o:OLEObject", NS)
                if ole_node is None:
                    ole_node = object_node.find(".//o:OLEObject", NS)
                if ole_node is None:
                    continue

                image_node = object_node.find(".//v:imagedata", NS)
                relationship_id = ole_node.attrib.get(f"{{{NS['r']}}}id")
                preview_relationship_id = (
                    image_node.attrib.get(f"{{{NS['r']}}}id")
                    if image_node is not None
                    else None
                )
                relationship = rel_map.get(relationship_id or "", {})
                preview_relationship = rel_map.get(preview_relationship_id or "", {})

                embedding_target = relationship.get("resolved_target")
                preview_target = preview_relationship.get("resolved_target")
                payload_info = inspect_embedded_payload(zf, embedding_target)
                prog_id_raw = ole_node.attrib.get("ProgID")

                matched, basis = candidate_is_mathtype(
                    prog_id_raw=prog_id_raw,
                    field_code_raw=field_code_raw,
                    payload_info=payload_info,
                )
                if not matched:
                    continue

                sequence += 1
                evidence_sources = [part_name]
                rels_path = rels_path_for_part(part_name)
                if parse_xml_from_zip(zf, rels_path) is not None:
                    evidence_sources.append(rels_path)
                if embedding_target:
                    evidence_sources.append(embedding_target)

                node_path = path_index.get(id(ole_node)) or path_index[id(object_node)]
                results.append(
                    make_formula_record(
                        formula_id=f"mathtype-ole-{sequence:04d}",
                        part_name=part_name,
                        story_type=story_type,
                        paragraph_index=paragraph_index,
                        run_index=run_index,
                        object_sequence=sequence,
                        xpath=build_xpath(node_path),
                        relationship_id=relationship_id,
                        embedding_target=embedding_target,
                        preview_target=preview_target,
                        paragraph_text_raw=para_text,
                        text_before=text_before,
                        text_after=text_after,
                        prog_id_raw=prog_id_raw,
                        field_code_raw=field_code_raw,
                        payload_info=payload_info,
                        evidence_sources=evidence_sources,
                        basis=basis,
                    )
                )

    return results, sequence


def detect_mathtype_ole(docx_path: str | Path) -> dict:
    path = Path(docx_path).resolve()
    formulas = []
    input_bytes = path.read_bytes()

    with zipfile.ZipFile(path) as zf:
        sequence = 0
        for part_name in iter_story_parts(zf):
            root = parse_xml_from_zip(zf, part_name)
            if root is None:
                continue
            part_formulas, sequence = scan_story_part(
                zf=zf,
                root=root,
                part_name=part_name,
                rel_map=build_rel_map(zf, part_name),
                start_sequence=sequence,
            )
            formulas.extend(part_formulas)

    source_counts = {}
    if formulas:
        source_counts["mathtype-ole"] = len(formulas)

    return {
        "document": {
            "input_path": str(path),
            "input_sha256": sha256_hex(input_bytes),
            "container_format": path.suffix.lstrip(".").lower(),
            "detector_version": DETECTOR_VERSION,
        },
        "source_counts": source_counts,
        "formulas": formulas,
    }
