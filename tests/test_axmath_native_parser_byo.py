import json
import re
import zipfile
from io import BytesIO
from pathlib import Path

from document_equation_migration import axmath_native_parser as axmath_native_parser_module
from document_equation_migration.axmath_native_parser import (
    AXMATH_BYO_CONTENTS_REPORT_ARTIFACT_TYPE,
    AXMATH_BYO_LOCAL_REPORT_ARTIFACT_TYPE,
    build_byo_axmath_contents_report,
    build_byo_axmath_local_input_report,
)
from document_equation_migration.cli import main


FALSE_CLAIMS = {
    "native_parser_claim": False,
    "conversion_claim": False,
    "export_success_claim": False,
    "public_fixture_eligibility": False,
    "score_10_claim": False,
    "universal_axmath_support": False,
    "public_native_parser_binding": False,
    "semantic_decode_claim": False,
    "canonical_mathml_claim": False,
}


def _synthetic_contents(total_length: int) -> bytes:
    return bytes(total_length)


def _assert_false_claim_boundaries(report: dict[str, object]) -> None:
    claim_boundaries = report["claim_boundaries"]
    assert isinstance(claim_boundaries, dict)
    for claim_name, expected_value in FALSE_CLAIMS.items():
        assert claim_boundaries[claim_name] is expected_value


def _assert_no_private_material(report: dict[str, object]) -> None:
    payload = json.dumps(report, sort_keys=True)
    windows_home_forward = "C:" + "/" + "Users"
    windows_home_backslash = "C:" + "\\" + "Users"
    mathml_root = "<" + "math"
    mathml_mrow = "<" + "mrow"
    assert windows_home_forward not in payload
    assert windows_home_backslash not in payload
    assert "raw_payload" not in payload
    assert "source_payload" not in payload
    assert mathml_root not in payload
    assert mathml_mrow not in payload
    assert not re.search(r"[A-Fa-f0-9]{64}", payload)
    assert report["payload_material_included"] is False
    assert report["source_bytes_retained"] is False
    assert report["hex_included"] is False
    assert report["base64_included"] is False
    assert report["digest_included"] is False
    assert report["formula_body_included"] is False
    assert report["canonical_mathml_included"] is False


def test_byo_contents_report_accepts_synthetic_grid_without_payload_material() -> None:
    report = build_byo_axmath_contents_report(
        _synthetic_contents(573),
        input_name=r"C:\private\sample.Contents",
    )

    assert report["artifact_type"] == AXMATH_BYO_CONTENTS_REPORT_ARTIFACT_TYPE
    assert report["status"] == "structural_scan_passed"
    assert report["input_kind"] == "axmath-contents-bytestream"
    assert report["input_name"] == "sample.Contents"
    assert report["input_path_recorded"] is False
    assert report["local_byo_structural_decoder_available"] is True
    assert report["semantic_decode_supported"] is False
    assert report["parser_capability"] == "bounded_structural_grid_13_plus_20n_only"
    assert report["conversion_attempted"] is False

    structural_scan = report["structural_scan"]
    assert isinstance(structural_scan, dict)
    assert structural_scan["status"] == "bounded_structural_grid_scanned"
    assert structural_scan["total_length"] == 573
    assert structural_scan["record_cell_count"] == 28
    assert structural_scan["semantic_record_cell_supported_count"] == 0
    assert structural_scan["semantic_record_cell_unknown_count"] == 28

    _assert_false_claim_boundaries(report)
    _assert_no_private_material(report)


def test_byo_contents_report_rejects_non_integral_grid_without_payload_material() -> None:
    report = build_byo_axmath_contents_report(
        _synthetic_contents(574),
        input_name="/tmp/private/bad.bin",
    )

    assert report["status"] == "structural_scan_rejected"
    assert report["input_name"] == "bad.bin"
    structural_scan = report["structural_scan"]
    assert isinstance(structural_scan, dict)
    assert structural_scan["status"] == "bounded_structural_grid_rejected"
    assert structural_scan["error_code"] == "malformed_length_non_integral_record_grid"
    assert structural_scan["total_length"] == 574
    assert structural_scan["remainder_after_preface"] == 1

    _assert_false_claim_boundaries(report)
    _assert_no_private_material(report)


def test_cli_inspect_axmath_contents_writes_public_safe_report(tmp_path: Path) -> None:
    contents_path = tmp_path / "sample.Contents"
    report_path = tmp_path / "out" / "axmath-byo-report.json"
    contents_path.write_bytes(_synthetic_contents(573))

    exit_code = main(
        [
            "inspect-axmath-contents",
            str(contents_path),
            "--output",
            str(report_path),
        ]
    )

    assert exit_code == 0
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["artifact_type"] == AXMATH_BYO_CONTENTS_REPORT_ARTIFACT_TYPE
    assert report["status"] == "structural_scan_passed"
    assert report["input_name"] == "sample.Contents"
    assert report["input_path_recorded"] is False
    assert report["structural_scan"]["record_cell_count"] == 28
    _assert_no_private_material(report)


def test_cli_inspect_axmath_contents_reports_rejection_without_cli_failure(tmp_path: Path) -> None:
    contents_path = tmp_path / "bad.Contents"
    report_path = tmp_path / "out" / "axmath-byo-rejected.json"
    contents_path.write_bytes(_synthetic_contents(574))

    exit_code = main(
        [
            "inspect-axmath-contents",
            str(contents_path),
            "--output",
            str(report_path),
        ]
    )

    assert exit_code == 0
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["status"] == "structural_scan_rejected"
    assert report["input_name"] == "bad.Contents"
    assert report["structural_scan"]["error_code"] == "malformed_length_non_integral_record_grid"
    _assert_no_private_material(report)


def test_byo_local_input_report_accepts_raw_contents_file(tmp_path: Path) -> None:
    contents_path = tmp_path / "raw.Contents"
    contents_path.write_bytes(_synthetic_contents(573))

    report = build_byo_axmath_local_input_report(contents_path)

    assert report["artifact_type"] == AXMATH_BYO_LOCAL_REPORT_ARTIFACT_TYPE
    assert report["status"] == "structural_scan_passed"
    assert report["input_kind"] == "axmath-contents-bytestream"
    assert report["input_name"] == "raw.Contents"
    assert report["stream_count"] == 1
    assert report["stream_reports"][0]["status"] == "structural_scan_passed"
    assert report["stream_reports"][0]["structural_scan"]["record_cell_count"] == 28
    _assert_false_claim_boundaries(report)
    _assert_no_private_material(report)


class _FakeOle:
    def __init__(self, *_args: object, **_kwargs: object) -> None:
        self.closed = False

    def exists(self, stream_name: str) -> bool:
        return stream_name == "Contents"

    def openstream(self, stream_name: str) -> BytesIO:
        assert stream_name == "Contents"
        return BytesIO(_synthetic_contents(573))

    def close(self) -> None:
        self.closed = True

    def __enter__(self) -> "_FakeOle":
        return self

    def __exit__(self, *_args: object) -> None:
        self.close()


class _FakeOlefileModule:
    OleFileIO = _FakeOle

    @staticmethod
    def isOleFile(*args: object, data: bytes | None = None) -> bool:
        if data is not None:
            return data == b"FAKE-OLE"
        if args and isinstance(args[0], str):
            return Path(args[0]).read_bytes() == b"FAKE-OLE"
        return bool(args and args[0] == b"FAKE-OLE")


def test_byo_local_input_report_extracts_docx_ole_contents_without_retaining_payload(
    monkeypatch: object,
    tmp_path: Path,
) -> None:
    monkeypatch.setattr(axmath_native_parser_module, "olefile", _FakeOlefileModule)
    docx_path = tmp_path / "sample.docx"
    with zipfile.ZipFile(docx_path, "w") as zf:
        zf.writestr("word/embeddings/oleObject1.bin", b"FAKE-OLE")

    report = build_byo_axmath_local_input_report(docx_path)

    assert report["artifact_type"] == AXMATH_BYO_LOCAL_REPORT_ARTIFACT_TYPE
    assert report["status"] == "structural_scan_passed"
    assert report["input_kind"] == "docx-package"
    assert report["input_name"] == "sample.docx"
    assert report["stream_count"] == 1
    stream_report = report["stream_reports"][0]
    assert stream_report["input_name"] == "oleObject1.bin:Contents"
    assert stream_report["container_member_name"] == "word/embeddings/oleObject1.bin"
    assert stream_report["structural_scan"]["record_cell_count"] == 28
    _assert_no_private_material(report)


def test_cli_inspect_axmath_local_writes_public_safe_report(tmp_path: Path) -> None:
    contents_path = tmp_path / "local.Contents"
    report_path = tmp_path / "out" / "axmath-local-report.json"
    contents_path.write_bytes(_synthetic_contents(573))

    exit_code = main(
        [
            "inspect-axmath-local",
            str(contents_path),
            "--output",
            str(report_path),
        ]
    )

    assert exit_code == 0
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["artifact_type"] == AXMATH_BYO_LOCAL_REPORT_ARTIFACT_TYPE
    assert report["status"] == "structural_scan_passed"
    assert report["input_kind"] == "axmath-contents-bytestream"
    assert report["stream_count"] == 1
    _assert_no_private_material(report)
