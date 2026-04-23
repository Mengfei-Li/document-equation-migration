from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


sys.stdout.reconfigure(encoding="utf-8")


MATHML_NS = "http://www.w3.org/1998/Math/MathML"
ODF_NS = {
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
}
XLINK_HREF = f"{{{ODF_NS['xlink']}}}href"
ODF_FORMULA_MIMETYPE = "application/vnd.oasis.opendocument.formula"
ODF_TEXT_MIMETYPE = "application/vnd.oasis.opendocument.text"
META_NAME_ATTR = f"{{{ODF_NS['meta']}}}name"
BRIDGE_FIELD_ALIASES = {
    "original_origin": {"ag_original_origin", "original_origin", "libreoffice_original_origin"},
    "conversion_mode": {"ag_conversion_mode", "conversion_mode", "libreoffice_conversion_mode"},
    "input_filter": {"ag_input_filter", "input_filter", "libreoffice_input_filter"},
    "profile_isolated": {"ag_profile_isolated", "profile_isolated", "libreoffice_profile_isolated"},
}


def local_name(tag: str) -> str:
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def qname(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def parse_xml_bytes(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def normalize_meta_name(name: str) -> str:
    normalized = "".join(char.lower() if char.isalnum() else "_" for char in name.strip())
    return normalized.strip("_")


def is_mathml_math(node: ET.Element | None) -> bool:
    return node is not None and node.tag == qname(MATHML_NS, "math")


def normalize_href(href: str) -> str:
    normalized = href.replace("\\", "/").split("#", 1)[0].strip()
    while normalized.startswith("./"):
        normalized = normalized[2:]
    return normalized.rstrip("/")


def member_from_href(href: str) -> str | None:
    normalized = normalize_href(href)
    if not normalized:
        return None
    if normalized.endswith(".xml"):
        return normalized
    return f"{normalized}/content.xml"


def guess_container_format(path: Path, mimetype: str) -> str:
    suffix = path.suffix.lower()
    if mimetype == ODF_FORMULA_MIMETYPE:
        return "odf"
    if mimetype == ODF_TEXT_MIMETYPE:
        return "odt"
    if suffix == ".fodt":
        return "fodt"
    if suffix == ".odt":
        return "odt"
    if suffix == ".odf":
        return "odf"
    if suffix:
        return suffix.lstrip(".")
    return "unknown"


def load_odf_package(path: str | Path) -> dict:
    path = Path(path)
    xml_roots: dict[str, ET.Element] = {}
    mimetype = ""
    names: set[str] = set()

    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            if "mimetype" in names:
                mimetype = zf.read("mimetype").decode("utf-8", errors="ignore").strip()
            for name in sorted(names):
                if name.endswith(".xml"):
                    xml_roots[name] = parse_xml_bytes(zf.read(name))
    else:
        xml_roots["content.xml"] = parse_xml_bytes(path.read_bytes())

    return {
        "path": path,
        "mimetype": mimetype,
        "container_format": guess_container_format(path, mimetype),
        "names": names,
        "xml_roots": xml_roots,
    }


def parse_office_meta(package: dict) -> dict[str, object]:
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
        "generator": generator,
        "user_defined": user_defined,
    }


def meta_value(user_defined: dict[str, str], key: str) -> str:
    for alias in BRIDGE_FIELD_ALIASES[key]:
        if alias in user_defined and user_defined[alias]:
            return user_defined[alias]
    return ""


def is_libreoffice_generator(text: str) -> bool:
    lowered = text.lower()
    return "libreoffice" in lowered or "collabora office" in lowered


def has_libreoffice_bridge_provenance(meta: dict[str, object]) -> bool:
    user_defined = meta["user_defined"]
    return bool(
        is_libreoffice_generator(str(meta["generator"]))
        and any(meta_value(user_defined, key) for key in BRIDGE_FIELD_ALIASES)
    )


def _build_record(
    *,
    index: int,
    member_path: str,
    storage_kind: str,
    embedding_target: str | None,
    evidence_sources: list[str],
) -> dict:
    record = {
        "formula_id": f"odf-native-{index:04d}",
        "source_family": "odf-native",
        "source_role": "native-source",
        "doc_part_path": member_path,
        "story_type": "main",
        "storage_kind": storage_kind,
        "embedding_target": embedding_target,
        "canonical_mathml_status": "available",
        "omml_status": "not-applicable",
        "latex_status": "not-applicable",
        "risk_level": "low",
        "risk_flags": [],
        "failure_mode": "",
        "confidence": 0.99,
        "provenance": {
            "transform_chain": [],
            "evidence_sources": evidence_sources,
        },
    }
    if not embedding_target:
        record.pop("embedding_target")
    return record


def collect_odf_native_formulas(
    package: dict,
    *,
    suppress_bridge_provenance: bool = True,
) -> list[dict]:
    if suppress_bridge_provenance and has_libreoffice_bridge_provenance(parse_office_meta(package)):
        return []

    formulas: list[dict] = []
    content_root = package["xml_roots"].get("content.xml")
    if is_mathml_math(content_root):
        formulas.append(
            _build_record(
                index=1,
                member_path="content.xml",
                storage_kind="odf-formula-root",
                embedding_target=None,
                evidence_sources=[str(package["path"]), "content.xml"],
            )
        )
    elif content_root is not None:
        for node in content_root.iter():
            if node.tag != qname(ODF_NS["draw"], "object"):
                continue

            inline_math = next((child for child in node.iter() if child.tag == qname(MATHML_NS, "math")), None)
            href = node.attrib.get(XLINK_HREF, "")
            subdocument_member = member_from_href(href) if href else None

            if inline_math is not None:
                formulas.append(
                    _build_record(
                        index=len(formulas) + 1,
                        member_path="content.xml",
                        storage_kind="odf-draw-object-inline",
                        embedding_target=href or None,
                        evidence_sources=[str(package["path"]), "content.xml"],
                    )
                )
                continue

            if subdocument_member and is_mathml_math(package["xml_roots"].get(subdocument_member)):
                formulas.append(
                    _build_record(
                        index=len(formulas) + 1,
                        member_path=subdocument_member,
                        storage_kind="odf-draw-object-subdocument",
                        embedding_target=href,
                        evidence_sources=[str(package["path"]), "content.xml", subdocument_member],
                    )
                )

    return formulas


def detect_odf_native(path: str | Path) -> dict:
    package = load_odf_package(path)
    formulas = collect_odf_native_formulas(package)

    return {
        "input_path": str(package["path"]),
        "container_format": package["container_format"],
        "formula_count": len(formulas),
        "source_counts": {"odf-native": len(formulas)},
        "formulas": formulas,
    }


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: python odf_native.py <input.odf|input.odt|input.fodt>")
        sys.exit(1)

    result = detect_odf_native(Path(sys.argv[1]).resolve())
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
