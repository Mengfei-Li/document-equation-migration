import hashlib
import json
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from document_equation_migration.equation3_mtef import (
    Equation3MtefError,
    convert_equation_native_stream_to_mathml,
    local_name,
)


FIXTURE_DIR = Path(__file__).resolve().parent / "fixtures" / "equation_editor_3_ole" / "apache_poi"


def _read_hex_fixture(file_name: str) -> bytes:
    text = (FIXTURE_DIR / file_name).read_text(encoding="ascii")
    return bytes.fromhex("".join(text.split()))


def _source_manifest() -> dict[str, object]:
    return json.loads((FIXTURE_DIR / "SOURCES.json").read_text(encoding="utf-8"))


def test_apache_poi_fixture_manifest_records_minimal_public_controls() -> None:
    manifest = _source_manifest()

    assert manifest["source_project"] == "Apache POI"
    assert manifest["source_commit"] == "e6a04b49211e23c704fcdbe524d99d2f4486b083"
    assert manifest["license"] == "Apache-2.0"
    assert "full source .doc files are not vendored" in manifest["fixture_policy"]
    assert len(manifest["fixtures"]) == 3


@pytest.mark.parametrize(
    "fixture_id",
    [
        "apache_poi_bug61268_formula0001",
        "apache_poi_bug61268_formula0003",
    ],
)
def test_apache_poi_equation3_native_controls_convert_to_canonical_mathml(fixture_id: str) -> None:
    manifest = _source_manifest()
    item = next(entry for entry in manifest["fixtures"] if entry["fixture_id"] == fixture_id)
    payload = _read_hex_fixture(item["file_name"])

    assert hashlib.sha256(payload).hexdigest() == item["equation_native_sha256"]
    assert item["source_doc_sha256"]
    assert item["source_stream"].endswith("/Equation Native")

    result = convert_equation_native_stream_to_mathml(payload)
    root = ET.fromstring(result.mathml_text)

    assert local_name(root.tag) == "math"
    assert hashlib.sha256(result.mathml_text.encode("utf-8")).hexdigest() == item["canonical_sha256"]
    assert result.mtef_version == 3
    assert result.product == 1
    assert result.template_selector_counts == item["template_selector_counts"]
    assert "".join(root.itertext()) == item["mathml_text"]


def test_apache_poi_unsupported_selector43v2_control_stays_blocked() -> None:
    manifest = _source_manifest()
    item = next(
        entry
        for entry in manifest["fixtures"]
        if entry["fixture_id"] == "apache_poi_bug50936_1_formula0013_selector43v2"
    )
    payload = _read_hex_fixture(item["file_name"])

    assert hashlib.sha256(payload).hexdigest() == item["equation_native_sha256"]
    assert item["expected_status"] == "unsupported-selector-43-variation-2"

    with pytest.raises(Equation3MtefError, match="Unsupported template selector=43 variation=2"):
        convert_equation_native_stream_to_mathml(payload)
