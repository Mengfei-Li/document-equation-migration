from __future__ import annotations

import json
import re
import sys
from pathlib import Path

try:
    from .odf_native import (
        collect_odf_native_formulas,
        has_libreoffice_bridge_provenance,
        load_odf_package,
        meta_value,
        parse_office_meta,
    )
except ImportError:
    from odf_native import (
        collect_odf_native_formulas,
        has_libreoffice_bridge_provenance,
        load_odf_package,
        meta_value,
        parse_office_meta,
    )


sys.stdout.reconfigure(encoding="utf-8")


BRIDGE_FIELD_ALIASES = {
    "original_origin": {"ag_original_origin", "original_origin", "libreoffice_original_origin"},
    "conversion_mode": {"ag_conversion_mode", "conversion_mode", "libreoffice_conversion_mode"},
    "input_filter": {"ag_input_filter", "input_filter", "libreoffice_input_filter"},
    "profile_isolated": {"ag_profile_isolated", "profile_isolated", "libreoffice_profile_isolated"},
}


def parse_bool(text: str) -> bool | None:
    normalized = text.strip().lower()
    if normalized in {"true", "1", "yes"}:
        return True
    if normalized in {"false", "0", "no"}:
        return False
    return None

def extract_producer_version(generator: str) -> str:
    match = re.search(r"LibreOffice/([0-9][^ $/]*)", generator)
    if match:
        return match.group(1)
    return "unknown"


def detect_libreoffice_transformed(path: str | Path) -> dict:
    package = load_odf_package(path)
    meta = parse_office_meta(package)
    generator = str(meta["generator"])
    user_defined = meta["user_defined"]

    original_origin = meta_value(user_defined, "original_origin")
    conversion_mode = meta_value(user_defined, "conversion_mode")
    input_filter = meta_value(user_defined, "input_filter")
    raw_profile_isolated = meta_value(user_defined, "profile_isolated")
    profile_isolated = parse_bool(raw_profile_isolated)
    if not has_libreoffice_bridge_provenance(meta):
        return {
            "input_path": str(package["path"]),
            "container_format": package["container_format"],
            "formula_count": 0,
            "source_counts": {"libreoffice-transformed": 0},
            "formulas": [],
        }

    native_formulas = collect_odf_native_formulas(package, suppress_bridge_provenance=False)
    formulas: list[dict] = []
    for index, item in enumerate(native_formulas, start=1):
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
                    "generator_id": "libreoffice",
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
        "input_path": str(package["path"]),
        "container_format": package["container_format"],
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
