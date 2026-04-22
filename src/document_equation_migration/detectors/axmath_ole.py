import hashlib
import re
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

try:
    import olefile
except ImportError:  # pragma: no cover - optional runtime dependency
    olefile = None


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

AXMATH_PROGID = "equation.axmath"
AXMATH_FIELD_CODE = "embed equation.axmath"
KNOWN_NON_AXMATH_PROGIDS = ("equation.dsmt", "equation.3")
AXMATH_ARTIFACT_MARKERS = {
    "AxMath.dotm": "axmath.dotm",
    "AMDisplayEquation": "amdisplayequation",
    "AMObj.afx": "amobj.afx",
}
EXPORT_CHANNELS = ["latex", "mathml"]
BASE_RISK_FLAGS = [
    "export-assisted-route",
    "native-static-parse-unverified",
    "mathml-export-unverified",
]
CUSTOM_SYMBOL_MARKERS = (
    "customsymbol",
    "custom symbol",
    "usersymbol",
    "user symbol",
)


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


def parse_xml_from_zip(zf: zipfile.ZipFile, inner_path: str) -> ET.Element | None:
    try:
        data = zf.read(inner_path)
    except KeyError:
        return None
    return ET.fromstring(data)


def build_xpath(path_tokens: list[tuple[str, int]]) -> str:
    return "/" + "/".join(f"{name}[{index}]" for name, index in path_tokens)


def rels_path_for_part(part_name: str) -> str:
    part = PurePosixPath(part_name)
    return str(part.parent / "_rels" / f"{part.name}.rels")


def resolve_relationship_target(part_name: str, target: str | None) -> str | None:
    if not target:
        return None
    target_path = PurePosixPath(target)
    if target_path.is_absolute():
        return target_path.as_posix().lstrip("/")
    resolved = PurePosixPath(part_name).parent / target_path
    return resolved.as_posix()


def build_rel_map(zf: zipfile.ZipFile, part_name: str) -> tuple[dict[str, dict], str | None]:
    rels_path = rels_path_for_part(part_name)
    root = parse_xml_from_zip(zf, rels_path)
    if root is None:
        return {}, None

    mapping = {}
    for rel in root.findall("rel:Relationship", NS):
        mapping[rel.attrib.get("Id", "")] = {
            "type": rel.attrib.get("Type"),
            "target": resolve_relationship_target(part_name, rel.attrib.get("Target")),
        }
    return mapping, rels_path


def printable_ascii_segments(data: bytes, limit: int = 12) -> list[str]:
    segments = []
    for hit in re.findall(rb"[ -~]{6,}", data):
        text = hit.decode("ascii", errors="ignore").strip()
        if text:
            segments.append(text[:200])
        if len(segments) >= limit:
            break
    return segments


def payload_probe(
    zf: zipfile.ZipFile,
    embedding_target: str | None,
) -> dict:
    if not embedding_target:
        return {
            "raw_payload_status": "unknown",
            "raw_payload_sha256": None,
            "ole_stream_names": [],
            "payload_ascii_preview": [],
        }

    names = set(zf.namelist())
    if embedding_target not in names:
        return {
            "raw_payload_status": "missing",
            "raw_payload_sha256": None,
            "ole_stream_names": [],
            "payload_ascii_preview": [],
        }

    data = zf.read(embedding_target)
    stream_names: list[str] = []
    if olefile is not None and olefile.isOleFile(data=data):
        ole = olefile.OleFileIO(data)
        try:
            stream_names = ["/".join(parts) for parts in ole.listdir()]
        finally:
            ole.close()

    return {
        "raw_payload_status": "present",
        "raw_payload_sha256": sha256_hex(data),
        "ole_stream_names": stream_names,
        "payload_ascii_preview": printable_ascii_segments(data),
    }


def collect_field_codes(paragraph: ET.Element) -> list[str]:
    codes = []
    for node in paragraph.findall(".//w:instrText", NS):
        text = (node.text or "").strip()
        if text:
            codes.append(text)
    return codes


def paragraph_style_id(paragraph: ET.Element) -> str | None:
    style = paragraph.find("w:pPr/w:pStyle", NS)
    if style is None:
        return None
    return style.attrib.get(f"{{{NS['w']}}}val")


def scan_package_artifacts(zf: zipfile.ZipFile) -> list[str]:
    names = set(zf.namelist())
    found: set[str] = set()

    for name in names:
        basename = PurePosixPath(name).name.casefold()
        for artifact_name, marker in AXMATH_ARTIFACT_MARKERS.items():
            if basename == marker:
                found.add(artifact_name)

    styles_root = parse_xml_from_zip(zf, "word/styles.xml")
    if styles_root is not None:
        for style in styles_root.findall(".//w:style", NS):
            style_id = style.attrib.get(f"{{{NS['w']}}}styleId", "")
            if style_id.casefold() == AXMATH_ARTIFACT_MARKERS["AMDisplayEquation"]:
                found.add("AMDisplayEquation")
                break

    for member_name in (
        "word/settings.xml",
        "word/styles.xml",
        "word/document.xml",
        "word/comments.xml",
    ):
        if member_name not in names:
            continue
        text = zf.read(member_name).decode("utf-8", errors="ignore")
        for artifact_name, marker in AXMATH_ARTIFACT_MARKERS.items():
            if marker in text.casefold():
                found.add(artifact_name)

    return sorted(found)


def payload_mentions_axmath(payload_ascii_preview: list[str]) -> bool:
    return any(
        AXMATH_PROGID in item.casefold() or "axmath" in item.casefold()
        for item in payload_ascii_preview
    )


def custom_symbol_present(payload_ascii_preview: list[str]) -> bool:
    return any(
        marker in item.casefold()
        for item in payload_ascii_preview
        for marker in CUSTOM_SYMBOL_MARKERS
    )


def choose_field_code(field_codes: list[str]) -> str | None:
    for code in field_codes:
        if AXMATH_FIELD_CODE in code.casefold():
            return code
    return None


def build_formula_record(
    *,
    zf: zipfile.ZipFile,
    part_name: str,
    rels_path: str | None,
    story_type: str,
    paragraph_index: int | None,
    run_index: int | None,
    object_sequence: int,
    object_node: ET.Element,
    path_tokens: list[tuple[str, int]],
    ole_node: ET.Element | None,
    rel_map: dict[str, dict],
    package_artifacts: list[str],
    paragraph_field_codes: list[str],
    paragraph_style: str | None,
) -> dict | None:
    prog_id_raw = None if ole_node is None else ole_node.attrib.get("ProgID") or None
    if prog_id_raw and any(prog_id_raw.casefold().startswith(prefix) for prefix in KNOWN_NON_AXMATH_PROGIDS):
        return None

    field_code_raw = choose_field_code(paragraph_field_codes)

    rid = None if ole_node is None else ole_node.attrib.get(f"{{{NS['r']}}}id")
    relation = rel_map.get(rid or "", {})
    embedding_target = relation.get("target")

    image_node = object_node.find(".//v:imagedata", NS)
    image_rid = None if image_node is None else image_node.attrib.get(f"{{{NS['r']}}}id")
    image_relation = rel_map.get(image_rid or "", {})
    preview_target = image_relation.get("target")

    payload = payload_probe(zf, embedding_target)
    payload_is_axmath = payload_mentions_axmath(payload["payload_ascii_preview"])
    has_ole_storage = embedding_target is not None
    style_is_axmath = (paragraph_style or "").casefold() == AXMATH_ARTIFACT_MARKERS["AMDisplayEquation"]
    addin_artifacts = sorted(
        set(package_artifacts) | ({"AMDisplayEquation"} if style_is_axmath else set())
    )

    prog_id_match = (prog_id_raw or "").casefold() == AXMATH_PROGID
    field_code_match = field_code_raw is not None

    score = 0
    if prog_id_match:
        score += 2
    if field_code_match:
        score += 2
    if payload_is_axmath:
        score += 2
    if has_ole_storage:
        score += 1
    if addin_artifacts:
        score += 1

    if score < 3 or not (prog_id_match or field_code_match or payload_is_axmath):
        return None

    if prog_id_match and has_ole_storage:
        confidence = 0.99
    elif field_code_match and payload_is_axmath:
        confidence = 0.97
    elif field_code_match and has_ole_storage and addin_artifacts:
        confidence = 0.95
    else:
        confidence = 0.93

    risk_flags = list(BASE_RISK_FLAGS)
    if story_type != "main":
        risk_flags.append("story-part-nonmain")
    if not prog_id_match:
        risk_flags.append("prog-id-missing")
    if not payload_is_axmath:
        risk_flags.append("payload-marker-missing")
    if preview_target:
        risk_flags.append("preview-image-present")
    if custom_symbol_present(payload["payload_ascii_preview"]):
        risk_flags.append("custom-symbols-present")

    risk_level = "medium"
    if story_type != "main" or not prog_id_match or not has_ole_storage:
        risk_level = "high"

    evidence_sources = [part_name]
    if rels_path is not None:
        evidence_sources.append(rels_path)
    if embedding_target:
        evidence_sources.append(embedding_target)
    if "AMDisplayEquation" in addin_artifacts:
        evidence_sources.append("word/styles.xml")

    return {
        "formula_id": f"axmath-ole-{object_sequence:04d}",
        "source_family": "axmath-ole",
        "source_role": "native-source",
        "doc_part_path": part_name,
        "story_type": story_type,
        "storage_kind": "ole-embedded" if has_ole_storage else "ole-object",
        "relationship_id": rid,
        "embedding_target": embedding_target,
        "preview_target": preview_target,
        "paragraph_index": paragraph_index,
        "run_index": run_index,
        "object_sequence": object_sequence,
        "canonical_mathml_status": "unverified",
        "omml_status": "not-applicable",
        "latex_status": "available",
        "risk_level": risk_level,
        "risk_flags": risk_flags,
        "failure_mode": None,
        "confidence": confidence,
        "xpath": build_xpath(path_tokens),
        "provenance": {
            "prog_id_raw": prog_id_raw,
            "field_code_raw": field_code_raw,
            "ole_stream_names": payload["ole_stream_names"],
            "raw_payload_status": payload["raw_payload_status"],
            "raw_payload_sha256": payload["raw_payload_sha256"],
            "transform_chain": None,
            "generator_id": None,
            "evidence_sources": evidence_sources,
        },
        "axmath": {
            "word_addin_artifacts": addin_artifacts,
            "export_channels": list(EXPORT_CHANNELS),
            "export_route_verified": False,
            "automation_mode": "ui-export-required",
            "custom_symbol_present": custom_symbol_present(payload["payload_ascii_preview"]),
            "payload_ascii_preview": payload["payload_ascii_preview"],
        },
    }


def scan_story_part(
    *,
    zf: zipfile.ZipFile,
    part_name: str,
    package_artifacts: list[str],
    start_sequence: int,
) -> tuple[list[dict], int]:
    root = parse_xml_from_zip(zf, part_name)
    if root is None:
        return [], start_sequence

    story_type = story_type_for_part(part_name)
    rel_map, rels_path = build_rel_map(zf, part_name)
    results: list[dict] = []
    sequence = start_sequence
    paragraph_counter = 0
    root_path = [(qualified_name(root.tag), 1)]

    def walk_container(
        node: ET.Element,
        path_tokens: list[tuple[str, int]],
        paragraph_index: int | None,
        run_counter: list[int],
        current_run_index: int | None,
        field_codes: list[str],
        paragraph_style: str | None,
    ) -> None:
        nonlocal sequence, paragraph_counter

        sibling_counts: dict[str, int] = {}
        for child in list(node):
            child_local = local_name(child.tag)
            sibling_counts[child_local] = sibling_counts.get(child_local, 0) + 1
            child_path = path_tokens + [
                (qualified_name(child.tag), sibling_counts[child_local])
            ]

            next_paragraph_index = paragraph_index
            next_run_counter = run_counter
            next_run_index = current_run_index
            next_field_codes = field_codes
            next_paragraph_style = paragraph_style

            if child_local == "p":
                paragraph_counter += 1
                next_paragraph_index = paragraph_counter
                next_run_counter = [0]
                next_run_index = None
                next_field_codes = collect_field_codes(child)
                next_paragraph_style = paragraph_style_id(child)

            elif child_local == "r":
                next_run_counter[0] += 1
                next_run_index = next_run_counter[0]

            elif child_local == "object":
                sequence += 1
                ole_node = child.find("o:OLEObject", NS)
                record = build_formula_record(
                    zf=zf,
                    part_name=part_name,
                    rels_path=rels_path,
                    story_type=story_type,
                    paragraph_index=paragraph_index,
                    run_index=current_run_index,
                    object_sequence=sequence,
                    object_node=child,
                    path_tokens=child_path,
                    ole_node=ole_node,
                    rel_map=rel_map,
                    package_artifacts=package_artifacts,
                    paragraph_field_codes=field_codes,
                    paragraph_style=paragraph_style,
                )
                if record is not None:
                    results.append(record)

            walk_container(
                child,
                child_path,
                next_paragraph_index,
                next_run_counter,
                next_run_index,
                next_field_codes,
                next_paragraph_style,
            )

    walk_container(
        root,
        root_path,
        None,
        [0],
        None,
        [],
        None,
    )
    return results, sequence


def detect_axmath_ole(docx_path: str | Path) -> dict:
    path = Path(docx_path).resolve()
    formulas: list[dict] = []

    with zipfile.ZipFile(path) as zf:
        package_artifacts = scan_package_artifacts(zf)
        sequence = 0
        for part_name in iter_story_parts(zf):
            story_formulas, sequence = scan_story_part(
                zf=zf,
                part_name=part_name,
                package_artifacts=package_artifacts,
                start_sequence=sequence,
            )
            formulas.extend(story_formulas)

    source_counts = {}
    if formulas:
        source_counts["axmath-ole"] = len(formulas)

    return {
        "document": {
            "input_path": str(path),
            "container_format": path.suffix.lstrip(".").lower(),
            "detector_version": DETECTOR_VERSION,
        },
        "source_counts": source_counts,
        "formulas": formulas,
    }
