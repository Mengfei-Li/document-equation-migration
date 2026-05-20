from __future__ import annotations

from collections import Counter
from collections.abc import Mapping, Sequence
from dataclasses import dataclass


AXMATH_SOURCE_FAMILY = "axmath-ole"

_FORBIDDEN_FIELD_FRAGMENTS = (
    "base64",
    "bitmap",
    "body",
    "bytes",
    "clipboard",
    "contents_hex",
    "decoded_formula",
    "formula_text",
    "hex_payload",
    "image",
    "latex",
    "mathml",
    "mml",
    "ocr",
    "omml",
    "payload",
    "pixel",
    "preview",
    "raw",
    "rendered",
    "screenshot",
    "svg",
    "text_body",
    "wmf",
)

_ALLOWED_RECORD_FIELDS = {
    "candidate_boundary_count",
    "candidate_count_attempt_count",
    "candidate_count_relation_fail_count",
    "candidate_count_relation_pass_count",
    "candidate_count_widths",
    "duplicate_invariance_status",
    "final_metric_equality_group_id",
    "final_metric_equality_group_member_rows",
    "full_stream_equality_group_id",
    "full_stream_equality_group_member_rows",
    "metadata_only_row_covered",
    "object_id",
    "object_ordinal",
    "parse_status",
    "relation_check_summary",
    "row_id",
    "shared_prefix_boundary_confirmed",
    "shared_prefix_candidate_length",
    "shared_prefix_candidate_offset",
    "shared_prefix_mismatch_count",
    "stream_length",
    "stream_locator_id",
    "tail_byte_category_histogram",
    "tail_category_max_run_length",
    "tail_category_run_count",
    "tail_length",
    "tail_start_candidate_offset",
    "tail_token_class_id",
    "tail_window_class",
    "trailer_candidate_offset",
    "trailer_length",
    "unique_stream_id",
    "unsupported_boundary",
}

_CLAIM_FIELDS = (
    "conversion_claim",
    "native_parser_claim",
    "export_success_claim",
    "public_fixture_eligibility",
    "semantic_review_claim",
    "visual_equivalence_claim",
)

_FALSE_CLAIMS = {field: False for field in _CLAIM_FIELDS}

_STATUS_MAP = {
    "candidate_count_relation_pass": {
        "structural_status": "metadata_relation_consistent",
        "body_decode_status": "not_attempted",
        "blocker_ids": (),
    },
    "candidate_count_relation_fail": {
        "structural_status": "candidate_count_relation_absent",
        "body_decode_status": "blocked_count_relation_hypothesis_only",
        "blocker_ids": ("candidate_count_relation_absent",),
    },
    "opaque_unsupported": {
        "structural_status": "opaque_tail_boundary",
        "body_decode_status": "blocked_opaque_tail",
        "blocker_ids": ("opaque_tail_nonblocking_boundary",),
    },
}


class AxMathStructuralMetadataError(ValueError):
    """Raised when metadata-only records include formula body or payload material."""

    def __init__(self, message: str, *, forbidden_fields: Sequence[str] = ()) -> None:
        super().__init__(message)
        self.forbidden_fields = tuple(forbidden_fields)


@dataclass(frozen=True, slots=True)
class AxMathStructuralMetadataClassification:
    row_id: object
    tail_token_class_id: object
    tail_window_class: object
    tail_length: object
    tail_start_candidate_offset: object
    trailer_length: object
    parse_status: str
    structural_status: str
    body_decode_status: str
    unsupported_boundary: object
    blocker_ids: tuple[str, ...]
    metadata_only_row_covered: bool
    metadata: dict[str, object]
    ignored_fields: tuple[str, ...]

    def to_dict(self) -> dict[str, object]:
        return {
            "source_family": AXMATH_SOURCE_FAMILY,
            "row_id": self.row_id,
            "tail_token_class_id": self.tail_token_class_id,
            "tail_window_class": self.tail_window_class,
            "tail_length": self.tail_length,
            "tail_start_candidate_offset": self.tail_start_candidate_offset,
            "trailer_length": self.trailer_length,
            "parse_status": self.parse_status,
            "structural_status": self.structural_status,
            "body_decode_status": self.body_decode_status,
            "unsupported_boundary": self.unsupported_boundary,
            "blocker_ids": list(self.blocker_ids),
            "metadata_only_row_covered": self.metadata_only_row_covered,
            **_FALSE_CLAIMS,
            "metadata": dict(self.metadata),
            "ignored_fields": list(self.ignored_fields),
        }


@dataclass(frozen=True, slots=True)
class AxMathStructuralMetadataRowsClassification:
    source_family: str
    status: str
    record_count: int
    records: tuple[dict[str, object], ...]
    structural_status_counts: dict[str, int]
    body_decode_status_counts: dict[str, int]
    blocker_ids: tuple[str, ...]
    claim_boundaries: dict[str, bool]
    ignored_fields: tuple[str, ...]

    def to_dict(self) -> dict[str, object]:
        return {
            "source_family": self.source_family,
            "status": self.status,
            "record_count": self.record_count,
            "records": list(self.records),
            "structural_status_counts": dict(self.structural_status_counts),
            "body_decode_status_counts": dict(self.body_decode_status_counts),
            "blocker_ids": list(self.blocker_ids),
            "claim_boundaries": dict(self.claim_boundaries),
            "ignored_fields": list(self.ignored_fields),
        }


def contains_axmath_structural_metadata(value: object) -> bool:
    if isinstance(value, Mapping):
        return any(key in value for key in ("parse_status", "rows", "records"))
    return isinstance(value, Sequence) and not isinstance(value, str | bytes | bytearray)


def classify_axmath_structural_metadata(
    record: Mapping[str, object],
) -> AxMathStructuralMetadataClassification:
    """Classify one public-safe AxMath metadata row without decoding formula bodies."""
    if not isinstance(record, Mapping):
        raise TypeError("AxMath structural metadata record must be a mapping.")

    forbidden_fields = _find_forbidden_fields(record)
    if forbidden_fields:
        raise AxMathStructuralMetadataError(
            "AxMath structural metadata cannot include payload/body fields.",
            forbidden_fields=forbidden_fields,
        )

    metadata: dict[str, object] = {}
    ignored_fields: list[str] = []
    for key, value in record.items():
        key_text = str(key)
        if key_text in _CLAIM_FIELDS:
            continue
        if key_text not in _ALLOWED_RECORD_FIELDS:
            ignored_fields.append(key_text)
            continue
        metadata[key_text] = _sanitize_value(value, path=key_text)

    parse_status = _parse_status(record)
    status_contract = _STATUS_MAP.get(
        parse_status,
        {
            "structural_status": "metadata_status_unknown",
            "body_decode_status": "blocked_unrecognized_metadata_status",
            "blocker_ids": ("unrecognized_metadata_status",),
        },
    )

    unsupported_boundary = metadata.get("unsupported_boundary")
    if parse_status == "opaque_unsupported" and not unsupported_boundary:
        unsupported_boundary = "opaque_tail"

    return AxMathStructuralMetadataClassification(
        row_id=metadata.get("row_id"),
        tail_token_class_id=metadata.get("tail_token_class_id"),
        tail_window_class=metadata.get("tail_window_class"),
        tail_length=metadata.get("tail_length"),
        tail_start_candidate_offset=metadata.get("tail_start_candidate_offset"),
        trailer_length=metadata.get("trailer_length"),
        parse_status=parse_status,
        structural_status=str(status_contract["structural_status"]),
        body_decode_status=str(status_contract["body_decode_status"]),
        unsupported_boundary=unsupported_boundary,
        blocker_ids=tuple(str(item) for item in status_contract["blocker_ids"]),
        metadata_only_row_covered=_metadata_only_row_covered(record),
        metadata=metadata,
        ignored_fields=tuple(sorted(ignored_fields)),
    )


def classify_axmath_structural_metadata_rows(
    rows: Sequence[Mapping[str, object]],
) -> AxMathStructuralMetadataRowsClassification:
    """Classify a batch of metadata-only rows and preserve blocker boundaries."""
    if isinstance(rows, str | bytes | bytearray):
        raise TypeError("AxMath structural metadata rows must be a sequence of mappings.")

    classifications = tuple(classify_axmath_structural_metadata(row).to_dict() for row in rows)
    structural_counts = Counter(str(item["structural_status"]) for item in classifications)
    body_counts = Counter(str(item["body_decode_status"]) for item in classifications)
    blocker_ids = tuple(
        dict.fromkeys(
            blocker_id
            for item in classifications
            for blocker_id in item.get("blocker_ids", [])
            if isinstance(blocker_id, str)
        )
    )
    ignored_fields = tuple(
        sorted(
            {
                ignored
                for item in classifications
                for ignored in item.get("ignored_fields", [])
                if isinstance(ignored, str)
            }
        )
    )

    return AxMathStructuralMetadataRowsClassification(
        source_family=AXMATH_SOURCE_FAMILY,
        status=_batch_status(len(classifications), blocker_ids),
        record_count=len(classifications),
        records=classifications,
        structural_status_counts=dict(structural_counts),
        body_decode_status_counts=dict(body_counts),
        blocker_ids=blocker_ids,
        claim_boundaries=dict(_FALSE_CLAIMS),
        ignored_fields=ignored_fields,
    )


def _parse_status(record: Mapping[str, object]) -> str:
    return str(record.get("parse_status") or "").strip().lower().replace("-", "_")


def _metadata_only_row_covered(record: Mapping[str, object]) -> bool:
    value = record.get("metadata_only_row_covered")
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes"}:
            return True
        if lowered in {"0", "false", "no"}:
            return False
    return True


def _batch_status(row_count: int, blocker_ids: tuple[str, ...]) -> str:
    if row_count == 0:
        return "metadata-only-no-records"
    if blocker_ids:
        return "metadata-only-blockers-preserved"
    return "metadata-only-structural-status-recorded"


def _find_forbidden_fields(value: object, *, path: str = "") -> tuple[str, ...]:
    found: list[str] = []
    if isinstance(value, Mapping):
        for key, child in value.items():
            key_text = str(key)
            child_path = f"{path}.{key_text}" if path else key_text
            if _is_forbidden_key(key_text):
                found.append(child_path)
            found.extend(_find_forbidden_fields(child, path=child_path))
    elif isinstance(value, Sequence) and not isinstance(value, str | bytes | bytearray):
        for index, child in enumerate(value):
            child_path = f"{path}[{index}]" if path else f"[{index}]"
            found.extend(_find_forbidden_fields(child, path=child_path))
    elif isinstance(value, bytes | bytearray | memoryview):
        found.append(path or "byte-value")
    return tuple(found)


def _is_forbidden_key(key: str) -> bool:
    normalized = key.strip().lower().replace("-", "_")
    return any(fragment in normalized for fragment in _FORBIDDEN_FIELD_FRAGMENTS)


def _sanitize_value(value: object, *, path: str) -> object:
    if isinstance(value, Mapping):
        return {
            str(key): _sanitize_value(child, path=f"{path}.{key}")
            for key, child in value.items()
            if not _is_forbidden_key(str(key))
        }
    if isinstance(value, Sequence) and not isinstance(value, str | bytes | bytearray):
        return [_sanitize_value(item, path=f"{path}[]") for item in value]
    if isinstance(value, bytes | bytearray | memoryview):
        raise AxMathStructuralMetadataError(
            "AxMath structural metadata cannot include byte payload values.",
            forbidden_fields=(path,),
        )
    return value
