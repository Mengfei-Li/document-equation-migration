from __future__ import annotations

import hashlib
import io
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

try:
    import olefile
except ImportError:  # pragma: no cover - exercised only when dependency is missing
    olefile = None


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

OLE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
CONTROL_STREAMS = {"\x01CompObj", "\x01Ole", "\x03ObjInfo"}
FIELD_CODE_EQ3_RE = re.compile(r"\bEMBED\s+Equation(?:\s|$|\\)", re.IGNORECASE)
FIELD_CODE_VENDOR_RE = re.compile(r"\bEMBED\s+Equation\.(?:DSMT\d*|AxMath)\b", re.IGNORECASE)
ASCII_MARKER_RE = re.compile(rb"[ -~]{6,}")
DEFAULT_HEADER_SIZE = 28


def _parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def _story_type_for_part(part_path: str) -> str:
    if part_path == "word/document.xml":
        return "main"
    if part_path.startswith("word/header"):
        return "header"
    if part_path.startswith("word/footer"):
        return "footer"
    if part_path.endswith("footnotes.xml"):
        return "footnote"
    if part_path.endswith("endnotes.xml"):
        return "endnote"
    if part_path.endswith("comments.xml"):
        return "comment"
    return "other"


def _rels_path_for_part(part_path: str) -> str:
    parent, name = part_path.rsplit("/", 1)
    return f"{parent}/_rels/{name}.rels"


def _resolve_target(part_path: str, target: str | None) -> str | None:
    if not target:
        return None
    base_dir = posixpath.dirname(part_path)
    normalized = posixpath.normpath(posixpath.join(base_dir, target))
    return normalized.lstrip("/")


def _read_relationships(zf: zipfile.ZipFile, part_path: str) -> dict[str, dict[str, str]]:
    rels_path = _rels_path_for_part(part_path)
    try:
        root = _parse_xml(zf.read(rels_path))
    except KeyError:
        return {}

    rels: dict[str, dict[str, str]] = {}
    for rel in root.findall("rel:Relationship", NS):
        rel_id = rel.attrib.get("Id")
        if not rel_id:
            continue
        rels[rel_id] = {
            "type": rel.attrib.get("Type", ""),
            "target": _resolve_target(part_path, rel.attrib.get("Target")),
        }
    return rels


def _iter_story_parts(zf: zipfile.ZipFile) -> list[str]:
    story_parts: list[str] = []
    priority_parts = [
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
    ]

    names = set(zf.namelist())
    for part in priority_parts:
        if part in names:
            story_parts.append(part)

    for name in sorted(names):
        if name.startswith("word/header") and name.endswith(".xml"):
            story_parts.append(name)
        elif name.startswith("word/footer") and name.endswith(".xml"):
            story_parts.append(name)

    return story_parts


def _normalize_whitespace(text: str | None) -> str | None:
    if text is None:
        return None
    compact = " ".join(text.split())
    return compact or None


def _collect_field_code(paragraph: ET.Element) -> str | None:
    chunks: list[str] = []
    for fld_simple in paragraph.findall(".//w:fldSimple", NS):
        instr = fld_simple.attrib.get("instr")
        if instr:
            chunks.append(instr)
    for instr_text in paragraph.findall(".//w:instrText", NS):
        if instr_text.text:
            chunks.append(instr_text.text)
    return _normalize_whitespace(" ".join(chunks)) if chunks else None


def _ascii_markers(data: bytes) -> list[str]:
    markers: list[str] = []
    for match in ASCII_MARKER_RE.findall(data):
        text = match.decode("ascii", errors="ignore").strip()
        lowered = text.lower()
        if any(token in lowered for token in ("equation", "mathtype", "axmath", "mtef", "dsmt")):
            markers.append(text[:120])
        if len(markers) >= 12:
            break
    return markers


def probe_eqnolefilehdr(raw_payload: bytes, header_size: int = DEFAULT_HEADER_SIZE) -> dict[str, Any]:
    result: dict[str, Any] = {
        "header_detected": False,
        "native_header_size_bytes": None,
        "mtef_version": None,
        "mtef_platform": None,
        "mtef_generating_product": None,
        "mtef_product_version": None,
        "mtef_product_subversion": None,
    }

    if len(raw_payload) < header_size + 5:
        return result

    probe = raw_payload[header_size : header_size + 5]
    mtef_version, platform, generating_product, product_version, product_subversion = probe

    result.update(
        {
            "native_header_size_bytes": header_size,
            "mtef_version": mtef_version,
            "mtef_platform": platform,
            "mtef_generating_product": generating_product,
            "mtef_product_version": product_version,
            "mtef_product_subversion": product_subversion,
        }
    )

    if mtef_version == 3 and generating_product == 1:
        result["header_detected"] = True

    return result


def _read_ole_streams(data: bytes) -> list[dict[str, Any]]:
    if olefile is None or not data:
        return []

    buffer = io.BytesIO(data)
    if not olefile.isOleFile(buffer):
        return []

    streams: list[dict[str, Any]] = []
    with olefile.OleFileIO(io.BytesIO(data)) as ole:
        for stream_path in ole.listdir():
            stream_name = "/".join(stream_path)
            stream_data = ole.openstream(stream_path).read()
            streams.append(
                {
                    "name": stream_name,
                    "data": stream_data,
                    "ascii_markers": _ascii_markers(stream_data),
                }
            )
    return streams


def _find_payload_probe(data: bytes, streams: list[dict[str, Any]]) -> tuple[dict[str, Any], bytes | None, str | None]:
    candidates: list[tuple[dict[str, Any], bytes, str]] = []

    for stream in streams:
        stream_name = stream["name"]
        if stream_name in CONTROL_STREAMS:
            continue
        probe = probe_eqnolefilehdr(stream["data"])
        if probe["header_detected"]:
            candidates.append((probe, stream["data"], stream_name))

    if candidates:
        probe, payload, stream_name = candidates[0]
        return probe, payload, stream_name

    if streams:
        for stream in streams:
            if stream["name"] not in CONTROL_STREAMS:
                return probe_eqnolefilehdr(stream["data"]), stream["data"], stream["name"]

    return probe_eqnolefilehdr(data), data if data else None, None


def _build_formula_id(sequence: int) -> str:
    return f"equation-editor-3-ole-{sequence:04d}"


def _determine_source_role(raw_payload_status: str, preview_target: str | None) -> str:
    if raw_payload_status in {"missing", "corrupt"} and preview_target:
        return "preview-only"
    return "native-source"


def _determine_route(header_detected: bool, source_role: str) -> str:
    if source_role == "preview-only":
        return "preview-only"
    if header_detected:
        return "mtef-v3-mainline"
    return "mtef-v3-mainline"


def _determine_risk_level(
    header_detected: bool,
    raw_payload_status: str,
    conflicting_vendor_signal: bool,
    source_role: str,
) -> str:
    if source_role == "preview-only":
        return "manual-review"
    if conflicting_vendor_signal or raw_payload_status == "corrupt":
        return "high"
    if header_detected:
        return "low"
    return "medium"


def _detect_from_embedding(
    embedding_data: bytes | None,
    *,
    formula_id: str,
    doc_part_path: str,
    story_type: str,
    relationship_id: str | None,
    embedding_target: str | None,
    preview_target: str | None,
    paragraph_index: int,
    run_index: int,
    object_sequence: int,
    prog_id_raw: str | None,
    field_code_raw: str | None,
) -> dict[str, Any] | None:
    prog_id = prog_id_raw or ""
    field_code = field_code_raw or ""

    prog_id_lower = prog_id.lower()
    conflicting_vendor_signal = bool(
        prog_id_lower.startswith("equation.dsmt")
        or prog_id_lower.startswith("equation.axmath")
        or FIELD_CODE_VENDOR_RE.search(field_code)
    )
    if conflicting_vendor_signal:
        return None

    streams = _read_ole_streams(embedding_data or b"")
    probe, payload, payload_stream_name = _find_payload_probe(embedding_data or b"", streams)

    ascii_markers = _ascii_markers(embedding_data or b"")
    for stream in streams:
        ascii_markers.extend(stream["ascii_markers"])
    deduped_markers = list(dict.fromkeys(ascii_markers))

    raw_payload_status = "missing"
    if embedding_data:
        raw_payload_status = "present"
    if embedding_data and payload is None:
        raw_payload_status = "corrupt"

    header_detected = bool(probe["header_detected"])
    prog_id_match = prog_id.lower() == "equation.3"
    field_code_match = bool(FIELD_CODE_EQ3_RE.search(field_code)) and not conflicting_vendor_signal

    if not any((prog_id_match, field_code_match, header_detected)):
        return None

    confidence = 0.0
    if prog_id_match:
        confidence += 0.55
    if field_code_match:
        confidence += 0.2
    if header_detected:
        confidence += 0.2
    if any("equation.3" in marker.lower() for marker in deduped_markers):
        confidence += 0.05
    confidence = round(min(confidence, 0.99), 2)

    source_role = _determine_source_role(raw_payload_status, preview_target)
    selected_route = _determine_route(header_detected, source_role)
    risk_flags: list[str] = []
    if prog_id_match:
        risk_flags.append("eq3-prog-id")
    if field_code_match:
        risk_flags.append("eq3-field-code")
    if header_detected:
        risk_flags.append("eqnolefilehdr-present")
    if source_role == "preview-only":
        risk_flags.extend(["missing-native-payload", "preview-only"])
    if not header_detected and raw_payload_status == "present":
        risk_flags.append("header-unverified")

    risk_level = _determine_risk_level(
        header_detected=header_detected,
        raw_payload_status=raw_payload_status,
        conflicting_vendor_signal=conflicting_vendor_signal,
        source_role=source_role,
    )

    return {
        "formula_id": formula_id,
        "source_family": "equation-editor-3-ole",
        "source_role": source_role,
        "doc_part_path": doc_part_path,
        "story_type": story_type,
        "storage_kind": "ole-embedded",
        "relationship_id": relationship_id,
        "embedding_target": embedding_target,
        "preview_target": preview_target,
        "paragraph_index": paragraph_index,
        "run_index": run_index,
        "object_sequence": object_sequence,
        "canonical_mathml_status": "unverified",
        "omml_status": "missing",
        "latex_status": "missing",
        "risk_level": risk_level,
        "risk_flags": risk_flags,
        "failure_mode": "missing-native-payload" if source_role == "preview-only" else None,
        "confidence": confidence,
        "provenance": {
            "prog_id_raw": prog_id_raw,
            "field_code_raw": field_code_raw,
            "ole_stream_names": [stream["name"] for stream in streams],
            "raw_payload_status": raw_payload_status,
            "raw_payload_sha256": hashlib.sha256(embedding_data).hexdigest() if embedding_data else None,
            "transform_chain": [],
            "generator_id": None,
            "evidence_sources": [item for item in [doc_part_path, embedding_target, preview_target] if item],
            "ascii_markers": deduped_markers,
            "payload_stream_name": payload_stream_name,
        },
        "source_specific": {
            "equation_editor_3": {
                "class_id_raw": "{0002CE02-0000-0000-C000-000000000046}" if prog_id_match or field_code_match else None,
                "native_header_size_bytes": probe["native_header_size_bytes"],
                "mtef_platform": probe["mtef_platform"],
                "mtef_generating_product": probe["mtef_generating_product"],
                "mtef_version": probe["mtef_version"],
                "selected_route": selected_route,
            }
        },
    }


def detect_equation_editor_3_ole(docx_path: str | Path) -> list[dict[str, Any]]:
    path = Path(docx_path)
    results: list[dict[str, Any]] = []
    sequence = 0

    with zipfile.ZipFile(path) as zf:
        for part_path in _iter_story_parts(zf):
            part_root = _parse_xml(zf.read(part_path))
            rels = _read_relationships(zf, part_path)
            story_type = _story_type_for_part(part_path)

            for paragraph_index, paragraph in enumerate(part_root.findall(".//w:p", NS)):
                field_code_raw = _collect_field_code(paragraph)
                for run_index, run in enumerate(paragraph.findall(".//w:r", NS)):
                    objects = run.findall("./w:object", NS)
                    if not objects:
                        continue

                    for obj in objects:
                        ole_node = obj.find(".//o:OLEObject", NS)
                        if ole_node is None:
                            continue

                        sequence += 1
                        formula_id = _build_formula_id(sequence)

                        relationship_id = ole_node.attrib.get(f"{{{NS['r']}}}id")
                        preview_node = obj.find(".//v:imagedata", NS)
                        preview_relationship_id = None
                        if preview_node is not None:
                            preview_relationship_id = preview_node.attrib.get(f"{{{NS['r']}}}id")

                        embedding_target = None
                        if relationship_id and rels.get(relationship_id, {}).get("type") == OLE_REL_TYPE:
                            embedding_target = rels[relationship_id].get("target")
                        preview_target = rels.get(preview_relationship_id, {}).get("target") if preview_relationship_id else None

                        embedding_data = None
                        if embedding_target and embedding_target in zf.namelist():
                            embedding_data = zf.read(embedding_target)

                        record = _detect_from_embedding(
                            embedding_data,
                            formula_id=formula_id,
                            doc_part_path=part_path,
                            story_type=story_type,
                            relationship_id=relationship_id,
                            embedding_target=embedding_target,
                            preview_target=preview_target,
                            paragraph_index=paragraph_index,
                            run_index=run_index,
                            object_sequence=sequence,
                            prog_id_raw=ole_node.attrib.get("ProgID"),
                            field_code_raw=field_code_raw,
                        )
                        if record is not None:
                            results.append(record)

    return results
