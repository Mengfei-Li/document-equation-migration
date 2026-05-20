from __future__ import annotations

from collections import Counter
from collections.abc import Mapping, Sequence
from dataclasses import dataclass


AXMATH_SOURCE_FAMILY = "axmath-ole"
AXMATH_CONTROLLED_ACCEPTANCE_KEY = "axmath_controlled_acceptance"
SCHEMA_ID = "axmath.controlled_acceptance.public_contract.v1"
PRIVATE_V2_DECISION_ENUM = "axmath_angle_direct_native_acceptance_refresh_passed"

RESULT_CLASS_EXACT = "exact"
RESULT_CLASS_EXACT_RECAPTURE = "exact_recapture"
RESULT_CLASS_SEMANTIC_EQUIVALENT = "semantic_equivalent_rewrite_not_exact_source"
RESULT_CLASS_UNSUPPORTED = "unsupported"

DEFAULT_PRIVATE_V2_RESULT_CLASS_COUNTS = {
    RESULT_CLASS_EXACT: 334,
    RESULT_CLASS_EXACT_RECAPTURE: 21,
    RESULT_CLASS_SEMANTIC_EQUIVALENT: 0,
}
DEFAULT_PRIVATE_V2_DIAGNOSTIC_COUNTS = {
    RESULT_CLASS_EXACT_RECAPTURE: 21,
}
DEFAULT_PRIVATE_V2_CONTROLLED_SAMPLE_COUNT = 355

_CONTROLLED_RESULT_CLASSES = {
    RESULT_CLASS_EXACT,
    RESULT_CLASS_EXACT_RECAPTURE,
    RESULT_CLASS_SEMANTIC_EQUIVALENT,
    RESULT_CLASS_UNSUPPORTED,
}

_FALSE_PUBLIC_CLAIMS = {
    "native_parser_claim": False,
    "conversion_claim": False,
    "export_success_claim": False,
    "public_fixture_eligibility": False,
    "public_native_parser_binding": False,
    "universal_axmath_support": False,
}

_FORBIDDEN_FIELD_FRAGMENTS = (
    "base64",
    "body",
    "bytes",
    "clipboard",
    "contents",
    "decoded",
    "docx",
    "formula",
    "hex",
    "image",
    "latex",
    "mathml",
    "mml",
    "ocr",
    "ole",
    "omml",
    "payload",
    "pixel",
    "preview",
    "raw",
    "render",
    "screenshot",
    "svg",
    "wmf",
)


class AxMathControlledAcceptanceContractError(ValueError):
    """Raised when a public contract attempts to carry private formula material."""

    def __init__(self, message: str, *, forbidden_fields: Sequence[str] = ()) -> None:
        super().__init__(message)
        self.forbidden_fields = tuple(forbidden_fields)


@dataclass(frozen=True, slots=True)
class AxMathControlledAcceptanceResult:
    result_class: str
    accepted: bool
    diagnostic_ids: tuple[str, ...]
    blocker_ids: tuple[str, ...]
    public_release_status: str
    claim_boundaries: dict[str, bool]

    def to_dict(self) -> dict[str, object]:
        return {
            "source_family": AXMATH_SOURCE_FAMILY,
            "result_class": self.result_class,
            "accepted": self.accepted,
            "diagnostic_ids": list(self.diagnostic_ids),
            "blocker_ids": list(self.blocker_ids),
            "public_release_status": self.public_release_status,
            "claim_boundaries": dict(self.claim_boundaries),
        }


@dataclass(frozen=True, slots=True)
class AxMathControlledAcceptanceContract:
    schema_id: str
    source_family: str
    evidence_scope: str
    decision_enum: str
    controlled_sample_count: int
    accepted_count: int
    remaining_fail_closed_count: int
    result_class_counts: dict[str, int]
    diagnostic_counts: dict[str, int]
    artifact_boundaries: dict[str, bool]
    claim_boundaries: dict[str, bool]
    notes: tuple[str, ...]

    def to_dict(self) -> dict[str, object]:
        return {
            "schema_id": self.schema_id,
            "source_family": self.source_family,
            "evidence_scope": self.evidence_scope,
            "decision_enum": self.decision_enum,
            "controlled_sample_count": self.controlled_sample_count,
            "accepted_count": self.accepted_count,
            "remaining_fail_closed_count": self.remaining_fail_closed_count,
            "result_class_counts": dict(self.result_class_counts),
            "diagnostic_counts": dict(self.diagnostic_counts),
            "artifact_boundaries": dict(self.artifact_boundaries),
            "claim_boundaries": dict(self.claim_boundaries),
            "notes": list(self.notes),
        }


@dataclass(frozen=True, slots=True)
class AxMathControlledAcceptanceRowsClassification:
    source_family: str
    schema_id: str
    row_count: int
    accepted_count: int
    remaining_fail_closed_count: int
    result_class_counts: dict[str, int]
    diagnostic_counts: dict[str, int]
    records: tuple[dict[str, object], ...]
    claim_boundaries: dict[str, bool]

    def to_dict(self) -> dict[str, object]:
        return {
            "source_family": self.source_family,
            "schema_id": self.schema_id,
            "row_count": self.row_count,
            "accepted_count": self.accepted_count,
            "remaining_fail_closed_count": self.remaining_fail_closed_count,
            "result_class_counts": dict(self.result_class_counts),
            "diagnostic_counts": dict(self.diagnostic_counts),
            "records": list(self.records),
            "claim_boundaries": dict(self.claim_boundaries),
        }


def default_axmath_controlled_acceptance_contract() -> AxMathControlledAcceptanceContract:
    """Return the public-safe private-v2 controlled acceptance contract."""

    return AxMathControlledAcceptanceContract(
        schema_id=SCHEMA_ID,
        source_family=AXMATH_SOURCE_FAMILY,
        evidence_scope="private-controlled-v2-contract-only",
        decision_enum=PRIVATE_V2_DECISION_ENUM,
        controlled_sample_count=DEFAULT_PRIVATE_V2_CONTROLLED_SAMPLE_COUNT,
        accepted_count=DEFAULT_PRIVATE_V2_CONTROLLED_SAMPLE_COUNT,
        remaining_fail_closed_count=0,
        result_class_counts=dict(DEFAULT_PRIVATE_V2_RESULT_CLASS_COUNTS),
        diagnostic_counts=dict(DEFAULT_PRIVATE_V2_DIAGNOSTIC_COUNTS),
        artifact_boundaries={
            "private_source_artifacts_included": False,
            "private_binary_material_included": False,
            "private_formula_bodies_included": False,
            "generated_canonical_bodies_included": False,
            "public_fixture_material_included": False,
        },
        claim_boundaries=dict(_FALSE_PUBLIC_CLAIMS),
        notes=(
            "Records private controlled-v2 acceptance after direct-native recapture removed semantic-equivalent rows.",
            "Does not publish a public native parser binding or public fixture corpus.",
            "Does not change the AxMath export-assisted delivery route.",
        ),
    )


def normalize_axmath_controlled_acceptance_contract(
    value: object | None,
) -> AxMathControlledAcceptanceContract:
    """Normalize optional route metadata into a no-payload acceptance contract."""

    if value is None:
        return default_axmath_controlled_acceptance_contract()
    if not isinstance(value, Mapping):
        raise TypeError("AxMath controlled acceptance contract must be a mapping.")

    forbidden_fields = _find_forbidden_fields(value)
    if forbidden_fields:
        raise AxMathControlledAcceptanceContractError(
            "AxMath controlled acceptance contract cannot include private formula material.",
            forbidden_fields=forbidden_fields,
        )

    result_counts = _counts_from_mapping(
        value.get("result_class_counts"),
        default=DEFAULT_PRIVATE_V2_RESULT_CLASS_COUNTS,
    )
    diagnostic_counts = _counts_from_mapping(
        value.get("diagnostic_counts"),
        default=DEFAULT_PRIVATE_V2_DIAGNOSTIC_COUNTS,
    )
    accepted_count = _int_value(value.get("accepted_count"), default=sum(result_counts.values()))
    remaining_fail_closed_count = _int_value(value.get("remaining_fail_closed_count"), default=0)
    controlled_sample_count = _int_value(
        value.get("controlled_sample_count"),
        default=accepted_count + remaining_fail_closed_count,
    )
    _validate_contract_counts(
        result_counts=result_counts,
        diagnostic_counts=diagnostic_counts,
        accepted_count=accepted_count,
        remaining_fail_closed_count=remaining_fail_closed_count,
        controlled_sample_count=controlled_sample_count,
    )

    base = default_axmath_controlled_acceptance_contract()
    return AxMathControlledAcceptanceContract(
        schema_id=str(value.get("schema_id") or base.schema_id),
        source_family=str(value.get("source_family") or AXMATH_SOURCE_FAMILY),
        evidence_scope=str(value.get("evidence_scope") or base.evidence_scope),
        decision_enum=str(value.get("decision_enum") or base.decision_enum),
        controlled_sample_count=controlled_sample_count,
        accepted_count=accepted_count,
        remaining_fail_closed_count=remaining_fail_closed_count,
        result_class_counts=result_counts,
        diagnostic_counts=diagnostic_counts,
        artifact_boundaries=dict(base.artifact_boundaries),
        claim_boundaries=dict(_FALSE_PUBLIC_CLAIMS),
        notes=tuple(str(item) for item in value.get("notes", base.notes))
        if isinstance(value.get("notes", base.notes), Sequence)
        and not isinstance(value.get("notes", base.notes), str | bytes | bytearray)
        else base.notes,
    )


def classify_axmath_controlled_acceptance_result(
    record: Mapping[str, object],
) -> AxMathControlledAcceptanceResult:
    """Classify one public-safe controlled acceptance record without formula material."""

    forbidden_fields = _find_forbidden_fields(record)
    if forbidden_fields:
        raise AxMathControlledAcceptanceContractError(
            "AxMath controlled acceptance record cannot include private formula material.",
            forbidden_fields=forbidden_fields,
        )

    result_class = str(record.get("result_class") or RESULT_CLASS_UNSUPPORTED).strip()
    diagnostic = str(record.get("diagnostic_id") or "").strip()
    if result_class not in _CONTROLLED_RESULT_CLASSES:
        result_class = RESULT_CLASS_UNSUPPORTED
        diagnostic = diagnostic or "unknown_controlled_acceptance_result_class"

    accepted = result_class != RESULT_CLASS_UNSUPPORTED
    diagnostic_ids = tuple(item for item in (diagnostic,) if item)
    blocker_ids: tuple[str, ...] = () if accepted else ("controlled_acceptance_unsupported",)
    return AxMathControlledAcceptanceResult(
        result_class=result_class,
        accepted=accepted,
        diagnostic_ids=diagnostic_ids,
        blocker_ids=blocker_ids,
        public_release_status="contract-only-no-public-artifacts",
        claim_boundaries=dict(_FALSE_PUBLIC_CLAIMS),
    )


def classify_axmath_controlled_acceptance_rows(
    rows: Sequence[Mapping[str, object]],
) -> AxMathControlledAcceptanceRowsClassification:
    """Classify public-safe acceptance rows and preserve public claim boundaries."""

    if isinstance(rows, str | bytes | bytearray):
        raise TypeError("AxMath controlled acceptance rows must be a sequence of mappings.")

    records = tuple(classify_axmath_controlled_acceptance_result(row).to_dict() for row in rows)
    result_class_counts = Counter(str(record["result_class"]) for record in records)
    diagnostic_counts = Counter(
        str(diagnostic_id)
        for record in records
        for diagnostic_id in record.get("diagnostic_ids", [])
        if isinstance(diagnostic_id, str)
    )
    accepted_count = sum(1 for record in records if record["accepted"] is True)
    return AxMathControlledAcceptanceRowsClassification(
        source_family=AXMATH_SOURCE_FAMILY,
        schema_id=SCHEMA_ID,
        row_count=len(records),
        accepted_count=accepted_count,
        remaining_fail_closed_count=len(records) - accepted_count,
        result_class_counts=dict(result_class_counts),
        diagnostic_counts=dict(diagnostic_counts),
        records=records,
        claim_boundaries=dict(_FALSE_PUBLIC_CLAIMS),
    )


def _int_value(value: object, *, default: int) -> int:
    if isinstance(value, bool):
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return default
        try:
            return int(text)
        except ValueError:
            return default
    return default


def _counts_from_mapping(value: object, *, default: Mapping[str, int]) -> dict[str, int]:
    if not isinstance(value, Mapping):
        return dict(default)
    counts = {str(key): _int_value(child, default=0) for key, child in value.items()}
    if not counts:
        return dict(default)
    unknown = sorted(set(counts) - _CONTROLLED_RESULT_CLASSES)
    if unknown:
        raise AxMathControlledAcceptanceContractError(
            f"Unknown AxMath controlled acceptance result class: {unknown[0]}"
        )
    if any(count < 0 for count in counts.values()):
        raise AxMathControlledAcceptanceContractError(
            "AxMath controlled acceptance counts must be non-negative."
        )
    return counts


def _validate_contract_counts(
    *,
    result_counts: Mapping[str, int],
    diagnostic_counts: Mapping[str, int],
    accepted_count: int,
    remaining_fail_closed_count: int,
    controlled_sample_count: int,
) -> None:
    if accepted_count < 0 or remaining_fail_closed_count < 0 or controlled_sample_count < 0:
        raise AxMathControlledAcceptanceContractError(
            "AxMath controlled acceptance aggregate counts must be non-negative."
        )
    if accepted_count != sum(
        count
        for result_class, count in result_counts.items()
        if result_class != RESULT_CLASS_UNSUPPORTED
    ):
        raise AxMathControlledAcceptanceContractError(
            "AxMath controlled acceptance accepted_count does not match result_class_counts."
        )
    unsupported_count = result_counts.get(RESULT_CLASS_UNSUPPORTED, 0)
    if remaining_fail_closed_count != unsupported_count:
        raise AxMathControlledAcceptanceContractError(
            "AxMath controlled acceptance remaining_fail_closed_count does not match unsupported count."
        )
    if controlled_sample_count != accepted_count + remaining_fail_closed_count:
        raise AxMathControlledAcceptanceContractError(
            "AxMath controlled acceptance sample count does not match accepted plus fail-closed counts."
        )
    if any(count < 0 for count in diagnostic_counts.values()):
        raise AxMathControlledAcceptanceContractError(
            "AxMath controlled acceptance diagnostic counts must be non-negative."
        )


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


__all__ = [
    "AXMATH_CONTROLLED_ACCEPTANCE_KEY",
    "AxMathControlledAcceptanceContract",
    "AxMathControlledAcceptanceContractError",
    "AxMathControlledAcceptanceResult",
    "classify_axmath_controlled_acceptance_result",
    "classify_axmath_controlled_acceptance_rows",
    "default_axmath_controlled_acceptance_contract",
    "normalize_axmath_controlled_acceptance_contract",
]
