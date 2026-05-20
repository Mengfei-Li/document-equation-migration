import base64
import hashlib
import json
import zipfile
from pathlib import Path

import olefile

from document_equation_migration.detectors.mathtype_ole import detect_mathtype_ole


FIXTURE_ROOT = Path(__file__).resolve().parent / "fixtures" / "mathtype_ole"
LIVE_CONTROL_ROOT = FIXTURE_ROOT / "live_control"
PAYLOAD_PATH = LIVE_CONTROL_ROOT / "word" / "embeddings" / "oleObject1.bin.b64"
SOURCES_PATH = LIVE_CONTROL_ROOT / "SOURCES.json"
EXPECTED_PAYLOAD_SHA256 = (
    "9f53c650efc68c5c94952892a5432a7bbc6966558a5cc7de6f7c0581ead14d4e"
)
EXPECTED_EQUATION_NATIVE_SHA256 = (
    "0d6a8f914733f9f89aa64a6babba78016f90264deaabafb751a8cd1acf5853f8"
)


def decode_live_control_payload() -> bytes:
    return base64.b64decode(PAYLOAD_PATH.read_text(encoding="ascii"))


def load_live_control_sources() -> dict[str, object]:
    return json.loads(SOURCES_PATH.read_text(encoding="utf-8"))


def build_docx_from_fixture(tmp_path: Path, fixture_dir: Path) -> Path:
    output_path = tmp_path / f"{fixture_dir.name}.docx"
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(fixture_dir.rglob("*")):
            if not file_path.is_file() or file_path.name in {"README.md", "SOURCES.json"}:
                continue

            relative_path = file_path.relative_to(fixture_dir).as_posix()
            if file_path.suffix == ".b64":
                arcname = relative_path[: -len(".b64")]
                zf.writestr(arcname, base64.b64decode(file_path.read_text(encoding="ascii")))
                continue

            zf.write(file_path, relative_path)
    return output_path


def test_live_control_sources_manifest_records_provenance_and_claim_boundary() -> None:
    manifest = load_live_control_sources()
    fixtures = manifest["fixtures"]
    item = fixtures[0]

    assert manifest["artifact_type"] == "mathtype-live-control-fixture-sources"
    assert manifest["source_project"] == "transpect/mathtype-extension"
    assert manifest["source_commit"] == "c1f788c7857802193220d370894ae52a2ce40d6c"
    assert manifest["license"] == "MIT"
    assert "Jure-MathType-MIT-LICENSE.txt" in manifest["notice"]
    assert "generated DOCX" in manifest["fixture_policy"]
    assert "converter prerequisites" in " ".join(manifest["safety"])

    claim_boundary = manifest["claim_boundary"]
    assert claim_boundary["public_fixture_claim"] == "integrity-detection-temporary-packaging-control"
    assert claim_boundary["default_ci_runs_external_converter"] is False
    assert claim_boundary["requires_external_tools_for_conversion"] is True
    assert claim_boundary["canonical_mathml_claim"] is False
    assert claim_boundary["production_ready_claim"] is False
    assert claim_boundary["lossless_claim"] is False
    assert claim_boundary["pixel_identity_claim"] is False
    assert claim_boundary["universal_mathtype_support_claim"] is False
    assert claim_boundary["general_live_conversion_claim"] is False

    assert len(fixtures) == 1
    assert item["fixture_id"] == "transpect_mathtype5_equation1_live_control"
    assert item["stored_file"] == "word/embeddings/oleObject1.bin.b64"
    assert item["source_path"] == "ruby/mathtype-0.0.7.5/spec/fixtures/input/mathtype5/equation1.bin"
    assert item["decoded_payload_sha256"] == EXPECTED_PAYLOAD_SHA256
    assert item["decoded_payload_size_bytes"] == 3584
    assert item["equation_native_stream_sha256"] == EXPECTED_EQUATION_NATIVE_SHA256
    assert item["equation_native_stream_size_bytes"] == 366
    assert item["mtef_version"] == 5
    assert item["conversion_status"] == "external-tool-gated"


def test_live_control_payload_matches_audited_source() -> None:
    manifest = load_live_control_sources()
    item = manifest["fixtures"][0]
    payload = decode_live_control_payload()

    assert len(payload) == item["decoded_payload_size_bytes"]
    assert hashlib.sha256(payload).hexdigest() == item["decoded_payload_sha256"]
    assert olefile.isOleFile(payload)

    with olefile.OleFileIO(payload) as ole:
        stream_names = {"/".join(path) for path in ole.listdir()}
        equation_native = ole.openstream("Equation Native").read()

    assert set(item["expected_ole_streams"]).issubset(stream_names)
    assert hashlib.sha256(equation_native).hexdigest() == item["equation_native_stream_sha256"]
    assert len(equation_native) == item["equation_native_stream_size_bytes"]


def test_live_control_fixture_builds_detectable_temporary_docx(tmp_path: Path) -> None:
    manifest = load_live_control_sources()
    item = manifest["fixtures"][0]
    docx_path = build_docx_from_fixture(tmp_path, LIVE_CONTROL_ROOT)

    result = detect_mathtype_ole(docx_path)

    assert result["source_counts"] == {item["expected_source_family"]: item["expected_detector_count"]}
    formula = result["formulas"][0]
    assert formula["source_family"] == item["expected_source_family"]
    assert formula["story_type"] == "main"
    assert formula["risk_level"] == "low"
    assert formula["provenance"]["raw_payload_sha256"] == item["decoded_payload_sha256"]
    assert formula["mathtype"]["equation_native_stream_exists"] is True
    assert formula["mathtype"]["equation_native_size_bytes"] == item["equation_native_stream_size_bytes"]
    assert formula["mathtype"]["mtef_version"] is None
