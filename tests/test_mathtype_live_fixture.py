import base64
import hashlib
import zipfile
from pathlib import Path

import olefile

from document_equation_migration.detectors.mathtype_ole import detect_mathtype_ole


FIXTURE_ROOT = Path(__file__).resolve().parent / "fixtures" / "mathtype_ole"
LIVE_CONTROL_ROOT = FIXTURE_ROOT / "live_control"
PAYLOAD_PATH = LIVE_CONTROL_ROOT / "word" / "embeddings" / "oleObject1.bin.b64"
EXPECTED_PAYLOAD_SHA256 = (
    "9f53c650efc68c5c94952892a5432a7bbc6966558a5cc7de6f7c0581ead14d4e"
)


def decode_live_control_payload() -> bytes:
    return base64.b64decode(PAYLOAD_PATH.read_text(encoding="ascii"))


def build_docx_from_fixture(tmp_path: Path, fixture_dir: Path) -> Path:
    output_path = tmp_path / f"{fixture_dir.name}.docx"
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(fixture_dir.rglob("*")):
            if not file_path.is_file() or file_path.name == "README.md":
                continue

            relative_path = file_path.relative_to(fixture_dir).as_posix()
            if file_path.suffix == ".b64":
                arcname = relative_path[: -len(".b64")]
                zf.writestr(arcname, base64.b64decode(file_path.read_text(encoding="ascii")))
                continue

            zf.write(file_path, relative_path)
    return output_path


def test_live_control_payload_matches_audited_source() -> None:
    payload = decode_live_control_payload()

    assert len(payload) == 3584
    assert hashlib.sha256(payload).hexdigest() == EXPECTED_PAYLOAD_SHA256
    assert olefile.isOleFile(payload)

    with olefile.OleFileIO(payload) as ole:
        stream_names = {"/".join(path) for path in ole.listdir()}

    assert "Equation Native" in stream_names


def test_live_control_fixture_builds_detectable_temporary_docx(tmp_path: Path) -> None:
    docx_path = build_docx_from_fixture(tmp_path, LIVE_CONTROL_ROOT)

    result = detect_mathtype_ole(docx_path)

    assert result["source_counts"] == {"mathtype-ole": 1}
    formula = result["formulas"][0]
    assert formula["source_family"] == "mathtype-ole"
    assert formula["story_type"] == "main"
    assert formula["risk_level"] == "low"
    assert formula["provenance"]["raw_payload_sha256"] == EXPECTED_PAYLOAD_SHA256
    assert formula["mathtype"]["equation_native_stream_exists"] is True
    assert formula["mathtype"]["equation_native_size_bytes"] > 0
