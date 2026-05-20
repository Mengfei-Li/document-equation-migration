"""Public-safe bounded structural scanner for AxMath Contents byte grids.

This module intentionally exposes only the BPSC-002 length-grid contract:
13 preface bytes followed by zero or more 20-byte record cells. It does not
decode, transform, or retain caller-supplied bytes.
"""

from __future__ import annotations

from pathlib import Path, PurePosixPath
import zipfile
from dataclasses import dataclass
from typing import Any

try:
    import olefile
except ImportError:  # pragma: no cover - optional runtime dependency
    olefile = None


PREFACE_LENGTH = 13
RECORD_CELL_SIZE = 20
UNKNOWN_SEMANTIC_FAMILY_STATUS = "unknown_semantic_family"
AXMATH_BYO_CONTENTS_REPORT_ARTIFACT_TYPE = "axmath-byo-contents-structural-report"
AXMATH_BYO_LOCAL_REPORT_ARTIFACT_TYPE = "axmath-byo-local-input-structural-report"


_FALSE_CLAIMS = {
    "native_parser_claim": False,
    "conversion_claim": False,
    "export_success_claim": False,
    "public_fixture_eligibility": False,
    "score_10_claim": False,
    "universal_axmath_support": False,
}
_BYO_CLAIM_BOUNDARIES = {
    **_FALSE_CLAIMS,
    "public_native_parser_binding": False,
    "semantic_decode_claim": False,
    "canonical_mathml_claim": False,
}


def _public_input_name(input_name: str | None) -> str:
    if input_name is None:
        return ""
    normalized = str(input_name).strip().replace("\\", "/")
    if not normalized:
        return ""
    return normalized.rsplit("/", 1)[-1]


def _base_local_report(
    *,
    artifact_type: str,
    input_kind: str,
    input_name: str | None,
) -> dict[str, Any]:
    return {
        "artifact_type": artifact_type,
        "input_kind": input_kind,
        "input_name": _public_input_name(input_name),
        "input_path_recorded": False,
        "payload_material_included": False,
        "source_bytes_retained": False,
        "hex_included": False,
        "base64_included": False,
        "digest_included": False,
        "formula_body_included": False,
        "canonical_mathml_included": False,
        "conversion_attempted": False,
        "local_byo_structural_decoder_available": True,
        "semantic_decode_supported": False,
        "parser_capability": "bounded_structural_grid_13_plus_20n_only",
        "claim_boundaries": dict(_BYO_CLAIM_BOUNDARIES),
    }


def _span_to_dict(span: tuple[int, int] | None) -> dict[str, int] | None:
    if span is None:
        return None
    return {"start_offset": span[0], "end_offset_exclusive": span[1]}


@dataclass(frozen=True)
class AxMathBoundedStructuralParseResult:
    """Metadata for one validated 13 + 20n AxMath structural grid."""

    total_length: int
    record_cell_count: int
    preface_length: int = PREFACE_LENGTH
    record_cell_size: int = RECORD_CELL_SIZE
    preface_span: tuple[int, int] = (0, PREFACE_LENGTH)
    first_record_span: tuple[int, int] | None = None
    last_record_span: tuple[int, int] | None = None
    semantic_record_cell_supported_count: int = 0
    semantic_record_family_supported_count: int = 0
    semantic_record_family_unknown_count: int | None = None
    semantic_status: str = UNKNOWN_SEMANTIC_FAMILY_STATUS
    native_parser_claim: bool = False
    conversion_claim: bool = False
    export_success_claim: bool = False
    public_fixture_eligibility: bool = False
    score_10_claim: bool = False
    universal_axmath_support: bool = False

    @property
    def semantic_record_cell_unknown_count(self) -> int:
        return self.record_cell_count

    @property
    def remainder_after_preface(self) -> int:
        return 0

    @property
    def unaccounted_grid_length(self) -> int:
        return 0

    @property
    def conversion_attempted(self) -> bool:
        return False

    def record_span(self, record_index: int) -> tuple[int, int]:
        if record_index < 0 or record_index >= self.record_cell_count:
            raise IndexError("record_index out of range for bounded AxMath grid")
        start_offset = self.preface_length + (self.record_cell_size * record_index)
        return (start_offset, start_offset + self.record_cell_size)

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": "bounded_structural_grid_scanned",
            "total_length": self.total_length,
            "preface_length": self.preface_length,
            "record_cell_size": self.record_cell_size,
            "record_cell_count": self.record_cell_count,
            "preface_span": _span_to_dict(self.preface_span),
            "first_record_span": _span_to_dict(self.first_record_span),
            "last_record_span": _span_to_dict(self.last_record_span),
            "remainder_after_preface": self.remainder_after_preface,
            "unaccounted_grid_length": self.unaccounted_grid_length,
            "semantic_record_cell_supported_count": self.semantic_record_cell_supported_count,
            "semantic_record_cell_unknown_count": self.semantic_record_cell_unknown_count,
            "semantic_record_family_supported_count": self.semantic_record_family_supported_count,
            "semantic_record_family_unknown_count": self.semantic_record_family_unknown_count,
            "semantic_status": self.semantic_status,
            "conversion_attempted": self.conversion_attempted,
            **_FALSE_CLAIMS,
        }


class AxMathBoundedStructuralParseError(ValueError):
    """Rejected BPSC-002 structural scan without exposing caller bytes."""

    def __init__(
        self,
        error_code: str,
        message: str,
        *,
        total_length: int | None = None,
        remainder_after_preface: int | None = None,
    ) -> None:
        super().__init__(f"{error_code}: {message}")
        self.error_code = error_code
        self.total_length = total_length
        self.preface_length = PREFACE_LENGTH
        self.record_cell_size = RECORD_CELL_SIZE
        self.remainder_after_preface = remainder_after_preface
        self.semantic_record_cell_supported_count = 0
        self.semantic_record_family_supported_count = 0
        self.conversion_attempted = False
        for claim_name, claim_value in _FALSE_CLAIMS.items():
            setattr(self, claim_name, claim_value)

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": "bounded_structural_grid_rejected",
            "error_code": self.error_code,
            "total_length": self.total_length,
            "preface_length": self.preface_length,
            "record_cell_size": self.record_cell_size,
            "remainder_after_preface": self.remainder_after_preface,
            "semantic_record_cell_supported_count": self.semantic_record_cell_supported_count,
            "semantic_record_family_supported_count": self.semantic_record_family_supported_count,
            "conversion_attempted": self.conversion_attempted,
            **_FALSE_CLAIMS,
        }


def scan_record_grid_13_plus_20n(
    contents: bytes | bytearray | memoryview,
) -> AxMathBoundedStructuralParseResult:
    """Scan one bytes-like AxMath Contents value for the 13 + 20n grid shape."""

    try:
        view = memoryview(contents)
    except TypeError as exc:
        raise AxMathBoundedStructuralParseError(
            "input_not_bytes_like",
            "contents must be a bytes-like object",
        ) from exc

    total_length = view.nbytes
    if total_length < PREFACE_LENGTH:
        raise AxMathBoundedStructuralParseError(
            "malformed_length_before_preface",
            "contents length is shorter than the fixed AxMath preface",
            total_length=total_length,
        )

    remainder_after_preface = (total_length - PREFACE_LENGTH) % RECORD_CELL_SIZE
    if remainder_after_preface != 0:
        raise AxMathBoundedStructuralParseError(
            "malformed_length_non_integral_record_grid",
            "contents length does not match the 13 + 20n record grid",
            total_length=total_length,
            remainder_after_preface=remainder_after_preface,
        )

    record_cell_count = (total_length - PREFACE_LENGTH) // RECORD_CELL_SIZE
    first_record_span = None
    last_record_span = None
    if record_cell_count:
        first_record_span = (PREFACE_LENGTH, PREFACE_LENGTH + RECORD_CELL_SIZE)
        last_start = PREFACE_LENGTH + (RECORD_CELL_SIZE * (record_cell_count - 1))
        last_record_span = (last_start, last_start + RECORD_CELL_SIZE)

    return AxMathBoundedStructuralParseResult(
        total_length=total_length,
        record_cell_count=record_cell_count,
        first_record_span=first_record_span,
        last_record_span=last_record_span,
    )


def build_byo_axmath_contents_report(
    contents: bytes | bytearray | memoryview,
    *,
    input_name: str | None = None,
) -> dict[str, Any]:
    """Build a public-safe local report for caller-supplied AxMath Contents bytes.

    The report is meant for bring-your-own local validation. It records only
    structural metadata and explicit claim boundaries; it intentionally omits
    source bytes, hex/base64 encodings, digests, MathML bodies, and paths.
    """

    report = _base_local_report(
        artifact_type=AXMATH_BYO_CONTENTS_REPORT_ARTIFACT_TYPE,
        input_kind="axmath-contents-bytestream",
        input_name=input_name,
    )
    report["notes"] = [
        "BYO/local report only; no AxMath payload fixture is included or retained.",
        "This scanner validates the public 13 + 20n structural grid only.",
        "It does not emit canonical MathML, convert formulas, or claim universal AxMath support.",
    ]
    try:
        structural_scan = scan_record_grid_13_plus_20n(contents).to_dict()
    except AxMathBoundedStructuralParseError as exc:
        report.update(
            {
                "status": "structural_scan_rejected",
                "structural_scan": exc.to_dict(),
            }
        )
        return report

    report.update(
        {
            "status": "structural_scan_passed",
            "structural_scan": structural_scan,
        }
    )
    return report


def _ole_contents_from_bytes(data: bytes) -> bytes | None:
    if olefile is None:
        return None
    if not olefile.isOleFile(data=data):
        return None
    ole = olefile.OleFileIO(data)
    try:
        if not ole.exists("Contents"):
            return None
        return ole.openstream("Contents").read()
    finally:
        ole.close()


def _ole_contents_from_path(path: Path) -> bytes | None:
    if olefile is None:
        return None
    if not olefile.isOleFile(str(path)):
        return None
    with olefile.OleFileIO(str(path)) as ole:
        if not ole.exists("Contents"):
            return None
        return ole.openstream("Contents").read()


def _docx_contents_reports(path: Path) -> list[dict[str, Any]]:
    reports: list[dict[str, Any]] = []
    with zipfile.ZipFile(path) as zf:
        for member_name in sorted(zf.namelist()):
            normalized = PurePosixPath(member_name)
            if normalized.parent.as_posix() != "word/embeddings":
                continue
            contents = _ole_contents_from_bytes(zf.read(member_name))
            if contents is None:
                continue
            report = build_byo_axmath_contents_report(
                contents,
                input_name=f"{normalized.name}:Contents",
            )
            report["container_member_name"] = normalized.as_posix()
            reports.append(report)
    return reports


def _local_report_status(stream_reports: list[dict[str, Any]], warnings: list[str]) -> str:
    if warnings and not stream_reports:
        return "inspection_unavailable"
    if not stream_reports:
        return "no_contents_stream_found"
    statuses = {str(report.get("status", "")) for report in stream_reports}
    if statuses == {"structural_scan_passed"}:
        return "structural_scan_passed"
    if statuses == {"structural_scan_rejected"}:
        return "structural_scan_rejected"
    return "structural_scan_mixed"


def build_byo_axmath_local_input_report(path: str | Path) -> dict[str, Any]:
    """Inspect a caller-supplied local DOCX/OLE/Contents input without retaining bytes."""

    input_path = Path(path)
    warnings: list[str] = []
    stream_reports: list[dict[str, Any]]
    input_kind = "axmath-contents-bytestream"

    if zipfile.is_zipfile(input_path):
        input_kind = "docx-package"
        if olefile is None:
            stream_reports = []
            warnings.append("olefile dependency unavailable; DOCX OLE embeddings were not inspected.")
        else:
            stream_reports = _docx_contents_reports(input_path)
    else:
        ole_contents = _ole_contents_from_path(input_path)
        if ole_contents is not None:
            input_kind = "ole-cfb-storage"
            stream_reports = [
                build_byo_axmath_contents_report(
                    ole_contents,
                    input_name=f"{input_path.name}:Contents",
                )
            ]
        else:
            stream_reports = [
                build_byo_axmath_contents_report(
                    input_path.read_bytes(),
                    input_name=input_path.name,
                )
            ]

    report = _base_local_report(
        artifact_type=AXMATH_BYO_LOCAL_REPORT_ARTIFACT_TYPE,
        input_kind=input_kind,
        input_name=input_path.name,
    )
    report.update(
        {
            "status": _local_report_status(stream_reports, warnings),
            "stream_count": len(stream_reports),
            "stream_reports": stream_reports,
            "warnings": warnings,
            "notes": [
                "BYO/local report only; local input bytes are read for inspection but not retained in the report.",
                "DOCX/OLE handling extracts top-level AxMath Contents streams when available.",
                "Each Contents stream is checked only against the public 13 + 20n structural grid.",
                "No canonical MathML, conversion output, source bytes, digests, or public fixture material is emitted.",
            ],
        }
    )
    return report


__all__ = [
    "AXMATH_BYO_CONTENTS_REPORT_ARTIFACT_TYPE",
    "AXMATH_BYO_LOCAL_REPORT_ARTIFACT_TYPE",
    "AxMathBoundedStructuralParseError",
    "AxMathBoundedStructuralParseResult",
    "build_byo_axmath_local_input_report",
    "build_byo_axmath_contents_report",
    "scan_record_grid_13_plus_20n",
]
