from __future__ import annotations

import json
import re
import sys
from pathlib import Path

try:
    from .odf_native import ODF_NS, detect_odf_native, load_odf_package, local_name
except ImportError:
    from odf_native import ODF_NS, detect_odf_native, load_odf_package, local_name


sys.stdout.reconfigure(encoding="utf-8")


META_NAME_ATTR = f"{{{ODF_NS['meta']}}}name"
BRIDGE_FIELD_ALIASES = {
    "original_origin": {"ag_original_origin", "original_origin", "libreoffice_original_origin"},
    "conversion_mode": {"ag_conversion_mode", "conversion_mode", "libreoffice_conversion_mode"},
    "input_filter": {"ag_input_filter", "input_filter", "libreoffice_input_filter"},
    "profile_isolated": {"ag_profile_isolated", "profile_isolated", "libreoffice_profile_isolated"},
}


def normalize_meta_name(name: str) -> str:
    normalized = re.sub(r"[^a-z0-9]+", "_", name.strip().lower())
    return normalized.strip("_")


def parse_office_meta(path: str | Path) -> dict:
    package = load_odf_package(path)
    meta_root = package["xml_roots"].get("meta.xml")
    generator = ""
    user_defined: dict[str, str] = {}

    if meta_root is not None:
        for node in meta_root.iter():
            tag = local_name(node.tag)
            if tag == "generator" and not generator:
                generator = (node.text or "").strip()
            elif tag == "user-defined":
                name = node.attrib.get(META_NAME_ATTR) or node.attrib.get("name") or ""
                value = (node.text or "").strip()
                if name:
                    user_defined[normalize_meta_name(name)] = value

    return {
        "package": package,
        "generator": generator,
        "user_defined": user_defined,
    }


def meta_value(user_defined: dict[str, str], key: str) -> str:
    for alias in BRIDGE_FIELD_ALIASES[key]:
        if alias in user_defined and user_defined[alias]:
            return user_defined[alias]
    return ""


def parse_bool(text: str) -> bool | None:
    normalized = text.strip().lower()
    if normalized in {"true", "1", "yes"}:
        return True
    if normalized in {"false", "0", "no"}:
        return False
    return None


def is_libreoffice_generator(text: str) -> bool:
    lowered = text.lower()
    return "libreoffice" in lowered or "collabora office" in lowered


def extract_producer_version(generator: str) -> str:
    match = re.search(r"LibreOffice/([0-9][^ $/]*)", generator)
    if match:
        return match.group(1)
    return "unknown"


def detect_libreoffice_transformed(path: str | Path) -> dict:
    native_result = detect_odf_native(path)
    meta = parse_office_meta(path)
    generator = meta["generator"]
    user_defined = meta["user_defined"]

    original_origin = meta_value(user_defined, "original_origin")
    conversion_mode = meta_value(user_defined, "conversion_mode")
    input_filter = meta_value(user_defined, "input_filter")
    raw_profile_isolated = meta_value(user_defined, "profile_isolated")
    profile_isolated = parse_bool(raw_profile_isolated)

    has_bridge_provenance = any(
        value
        for value in (
            original_origin,
            conversion_mode,
            input_filter,
            raw_profile_isolated,
        )
    )

    if not is_libreoffice_generator(generator) or not has_bridge_provenance:
        return {
            "input_path": native_result["input_path"],
            "container_format": native_result["container_format"],
            "formula_count": 0,
            "source_counts": {"libreoffice-transformed": 0},
            "formulas": [],
        }

    formulas: list[dict] = []
    for index, item in enumerate(native_result["formulas"], start=1):
        risk_flags = ["transformed-source", "libreoffice-bridge"]
        if original_origin:
            risk_flags.append(f"original-origin:{original_origin}")

        formulas.append(
            {
                "formula_id": f"libreoffice-transformed-{index:04d}",
                "source_family": "libreoffice-transformed",
                "source_role": "transformed-source",
                "doc_part_path": item["doc_part_path"],
                "story_type": item.get("story_type", "main"),
                "storage_kind": item["storage_kind"],
                "embedding_target": item.get("embedding_target"),
                "canonical_mathml_status": item["canonical_mathml_status"],
                "omml_status": "not-applicable",
                "latex_status": "not-applicable",
                "risk_level": "high",
                "risk_flags": risk_flags,
                "failure_mode": "",
                "confidence": 0.9,
                "provenance": {
                    "transform_chain": ["libreoffice-filter"],
                    "evidence_sources": item["provenance"]["evidence_sources"] + ["meta.xml"],
                    "generator_raw": generator,
                },
                "libreoffice": {
                    "producer_version": extract_producer_version(generator),
                    "conversion_mode": conversion_mode or "unknown",
                    "input_filter": input_filter or "unknown",
                    "profile_isolated": profile_isolated,
                    "original_origin": original_origin or "unknown",
                },
            }
        )
        if not item.get("embedding_target"):
            formulas[-1].pop("embedding_target")

    return {
        "input_path": native_result["input_path"],
        "container_format": native_result["container_format"],
        "formula_count": len(formulas),
        "source_counts": {"libreoffice-transformed": len(formulas)},
        "formulas": formulas,
    }


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: python libreoffice_transformed.py <input.odt|input.odf|input.fodt>")
        sys.exit(1)

    result = detect_libreoffice_transformed(Path(sys.argv[1]).resolve())
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
