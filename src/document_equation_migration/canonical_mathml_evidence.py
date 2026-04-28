from __future__ import annotations

import hashlib
from xml.etree import ElementTree as ET


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

