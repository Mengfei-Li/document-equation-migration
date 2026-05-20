from __future__ import annotations

import hashlib
from xml.etree import ElementTree as ET


CANONICAL_EVIDENCE_SCHEMA_VERSION = "2026-05-05"

CANONICAL_EVIDENCE_SOURCE_FAMILIES = (
    "mathtype-ole",
    "omml-native",
    "odf-native",
    "libreoffice-transformed",
    "equation-editor-3-ole",
    "axmath-ole",
)

CANONICALIZATION_SUMMARY_REQUIRED_FIELDS = (
    "expected_formula_count",
    "canonical_mathml_count",
    "unsupported_fragment_count",
    "formula_count_parity",
    "canonical_mathml_dir",
    "source_to_canonical_provenance",
    "property_summary",
    "unsupported_fragments",
)

CANONICAL_PROVENANCE_REQUIRED_FIELDS = (
    "formula_id",
    "canonical_artifact_path",
    "canonical_sha256",
    "preservation_status",
    "property_signals",
)

CANONICAL_BLOCKER_RECORD_REQUIRED_FIELDS = (
    "artifact_type",
    "source_family",
    "canonical_target",
    "status",
    "required_evidence",
    "next_ready_condition",
)

CANONICAL_VALIDATION_EVIDENCE_REQUIRED_FIELDS = (
    "artifact_type",
    "source_family",
    "status",
)

CANONICAL_FORMULA_COUNT_PARITY_VALUES = (
    "passed",
    "mismatch",
    "unsupported-fragments",
    "failed",
)

CANONICAL_CLAIM_BOUNDARY_FIELDS = (
    "conversion_claim",
    "limited_conversion_claim",
    "general_converter_claim",
    "deliverability_claim",
    "word_visual_fill_back_claim",
    "claim_boundary",
)


def canonical_evidence_schema() -> dict[str, object]:
    return {
        "version": CANONICAL_EVIDENCE_SCHEMA_VERSION,
        "source_families": list(CANONICAL_EVIDENCE_SOURCE_FAMILIES),
        "canonicalization_summary": {
            "required_fields": list(CANONICALIZATION_SUMMARY_REQUIRED_FIELDS),
            "formula_count_parity_values": list(CANONICAL_FORMULA_COUNT_PARITY_VALUES),
        },
        "source_to_canonical_provenance": {
            "required_fields": list(CANONICAL_PROVENANCE_REQUIRED_FIELDS),
        },
        "blocker_record": {
            "required_fields": list(CANONICAL_BLOCKER_RECORD_REQUIRED_FIELDS),
        },
        "validation_evidence": {
            "required_fields": list(CANONICAL_VALIDATION_EVIDENCE_REQUIRED_FIELDS),
        },
        "claim_boundary_fields": list(CANONICAL_CLAIM_BOUNDARY_FIELDS),
    }


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def mathml_property_signals(root: ET.Element) -> dict[str, object]:
    nodes = list(root.iter())
    return {
        "root_attributes": dict(root.attrib),
        "root_display": root.attrib.get("display", ""),
        "mathml_attribute_count": sum(len(node.attrib) for node in nodes),
        "has_semantics": any(local_name(node.tag) == "semantics" for node in nodes),
        "has_annotation": any(local_name(node.tag) == "annotation" for node in nodes),
        "has_mfrac_linethickness": any(
            "linethickness" in node.attrib
            for node in nodes
            if local_name(node.tag) == "mfrac"
        ),
        "has_mfrac_bevelled": any(
            node.attrib.get("bevelled") == "true"
            for node in nodes
            if local_name(node.tag) == "mfrac"
        ),
        "has_mfenced_separators": any(
            "separators" in node.attrib
            for node in nodes
            if local_name(node.tag) == "mfenced"
        ),
        "has_movablelimits": any("movablelimits" in node.attrib for node in nodes),
        "has_mathvariant": any("mathvariant" in node.attrib for node in nodes),
        "has_accent": any(node.attrib.get("accent") == "true" for node in nodes),
        "has_accentunder": any(node.attrib.get("accentunder") == "true" for node in nodes),
    }


def property_summary(items: list[dict[str, object]]) -> dict[str, object]:
    property_keys = (
        "has_semantics",
        "has_annotation",
        "has_mfrac_linethickness",
        "has_mfrac_bevelled",
        "has_mfenced_separators",
        "has_movablelimits",
        "has_mathvariant",
        "has_accent",
        "has_accentunder",
    )
    signals = [item.get("property_signals", {}) for item in items]
    root_display_values = sorted(
        {
            str(signal.get("root_display"))
            for signal in signals
            if isinstance(signal, dict) and signal.get("root_display")
        }
    )
    return {
        "mathml_attribute_count": sum(
            int(signal.get("mathml_attribute_count", 0))
            for signal in signals
            if isinstance(signal, dict)
        ),
        "root_display_values": root_display_values,
        "signal_counts": {
            key: sum(1 for signal in signals if isinstance(signal, dict) and signal.get(key))
            for key in property_keys
        },
    }
