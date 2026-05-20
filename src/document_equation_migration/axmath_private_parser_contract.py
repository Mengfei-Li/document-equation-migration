from __future__ import annotations

import json
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any


AXMATH_PRIVATE_PARSER_CONTRACT_SCHEMA_ID = "axmath.private_parser_contract.no_payload.v1"
AXMATH_SOURCE_FAMILY = "axmath-ole"
CANONICAL_RULE_UNWRAP_SINGLE_CHILD_MROW_V1 = "unwrap_single_child_mrow_v1"

EXPECTED_PRIVATE_RUN_IDS = (
    "run_20260519_0038_private_canonical_acceptance_gate",
    "run_20260519_0039_original_payload_fail_closed_gate",
    "run_20260519_0040_private_parser_contract_package",
)

FALSE_PUBLIC_CLAIMS = (
    "private_exact_source_completion",
    "original_controlled_corpus_source_free_completion",
    "source_free_native_parser_completion",
    "broad_native_parser_completion",
    "exact_source_completion",
    "conversion_claim",
    "native_parser_claim",
    "export_success_claim",
    "public_fixture_eligibility",
    "public_conversion_support",
    "universal_axmath_support",
)

PRIVATE_BOUNDARY_FIELDS = (
    "binary_material_included",
    "canonical_markup_included",
    "document_artifacts_included",
    "source_text_included",
    "derived_digests_included",
    "display_media_included",
)

REQUIRED_FORBIDDEN_EVIDENCE_SOURCES = (
    "raw Contents bytes",
    "private DOCX bytes",
    "OLE storage bytes",
    "formula body text",
    "MathML, LaTeX, OMML, SVG, or markup formula bodies",
    "payload-derived hex, base64, or digest material",
)

_FORBIDDEN_FIELD_FRAGMENTS = (
    "base64",
    "body",
    "bytes",
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
    "preview",
    "raw",
    "render",
    "screenshot",
    "svg",
    "wmf",
)

_ALLOWED_PRIVATE_MATERIAL_KEYS = frozenset({"forbidden_evidence_sources"})
_LONG_HEX_RE = re.compile(r"[0-9a-f]{32,}")
_USER_PROFILE_MARKERS = ("c:" + "/users/", "c:" + "\\" + "users" + "\\")
_MATH_TAG_MARKERS = ("<" + "math", "</" + "math")


class AxMathPrivateParserContractError(ValueError):
    """Raised when an AxMath no-payload parser contract is inconsistent or unsafe."""

    def __init__(self, message: str, *, violations: Sequence[str] = ()) -> None:
        super().__init__(message)
        self.violations = tuple(violations)


@dataclass(frozen=True, slots=True)
class AxMathPrivateParserContract:
    schema_id: str
    source_family: str
    private_run_ids: tuple[str, ...]
    clean_row_count: int
    clean_canonical_source_free_count: int
    clean_exact_replay_diagnostic_count: int
    clean_recovered_sample_ids: tuple[str, ...]
    original_row_count: int
    original_priority_v2_parsed_count: int
    original_priority_v2_unsupported_count: int
    original_source_free_promotable_guarded_rows: int
    original_unsupported_reason_counts: dict[str, int]
    private_clean_canonical_completion: bool
    claim_boundaries: dict[str, bool]

    def to_dict(self) -> dict[str, object]:
        return {
            "schema_id": self.schema_id,
            "source_family": self.source_family,
            "private_run_ids": list(self.private_run_ids),
            "clean_surface": {
                "row_count": self.clean_row_count,
                "canonical_source_free_count": self.clean_canonical_source_free_count,
                "exact_replay_diagnostic_count": self.clean_exact_replay_diagnostic_count,
                "recovered_sample_ids": list(self.clean_recovered_sample_ids),
            },
            "original_surface": {
                "row_count": self.original_row_count,
                "priority_v2_parsed_count": self.original_priority_v2_parsed_count,
                "priority_v2_unsupported_count": self.original_priority_v2_unsupported_count,
                "source_free_promotable_guarded_rows": self.original_source_free_promotable_guarded_rows,
                "unsupported_reason_counts": dict(self.original_unsupported_reason_counts),
            },
            "private_clean_canonical_completion": self.private_clean_canonical_completion,
            "claim_boundaries": dict(self.claim_boundaries),
        }

    def summary(self) -> dict[str, object]:
        return {
            "source_family": self.source_family,
            "schema_id": self.schema_id,
            "clean_canonical_source_free": (
                self.clean_canonical_source_free_count,
                self.clean_row_count,
            ),
            "clean_exact_diagnostic": (
                self.clean_exact_replay_diagnostic_count,
                self.clean_row_count,
            ),
            "original_priority_v2": (
                self.original_priority_v2_parsed_count,
                self.original_row_count,
            ),
            "original_fail_closed": (
                self.original_priority_v2_unsupported_count,
                self.original_row_count,
            ),
            "public_claims_false": all(not self.claim_boundaries[claim] for claim in FALSE_PUBLIC_CLAIMS),
        }


def load_axmath_private_parser_contract(path: str | Path) -> AxMathPrivateParserContract:
    value = json.loads(Path(path).read_text(encoding="utf-8"))
    return normalize_axmath_private_parser_contract(value)


def normalize_axmath_private_parser_contract(value: object) -> AxMathPrivateParserContract:
    if not isinstance(value, Mapping):
        raise AxMathPrivateParserContractError("AxMath parser contract must be a mapping.")

    _reject_private_material(value)
    _validate_top_level_identity(value)
    _validate_provenance(value)
    _validate_material_boundaries(value)

    clean = _mapping_at(value, "surfaces.clean_canonical_surface")
    original = _mapping_at(value, "surfaces.original_lossy_surface")
    claims = _mapping_at(value, "claim_freeze")

    _validate_clean_surface(clean)
    _validate_original_surface(original)
    _validate_claims(claims)
    _validate_forbidden_evidence_sources(value)

    return AxMathPrivateParserContract(
        schema_id=str(value["schema_id"]),
        source_family=str(value["source_family"]),
        private_run_ids=tuple(
            str(item) for item in _mapping_at(value, "provenance")["derived_from_private_run_ids"]
        ),
        clean_row_count=_int_at(clean, "row_count"),
        clean_canonical_source_free_count=_int_at(clean, "canonical_source_free_count"),
        clean_exact_replay_diagnostic_count=_int_at(clean, "exact_replay_diagnostic_count"),
        clean_recovered_sample_ids=tuple(str(item) for item in clean["recovered_sample_ids"]),
        original_row_count=_int_at(original, "row_count"),
        original_priority_v2_parsed_count=_int_at(original, "priority_v2_parsed_count"),
        original_priority_v2_unsupported_count=_int_at(original, "priority_v2_unsupported_count"),
        original_source_free_promotable_guarded_rows=_int_at(
            original,
            "source_free_promotable_guarded_rows",
        ),
        original_unsupported_reason_counts={
            str(key): _int_value(child, path=f"unsupported_reason_counts.{key}")
            for key, child in _mapping_at(original, "unsupported_reason_counts").items()
        },
        private_clean_canonical_completion=claims["private_clean_canonical_completion"] is True,
        claim_boundaries={claim: claims[claim] is True for claim in FALSE_PUBLIC_CLAIMS},
    )


def _validate_top_level_identity(value: Mapping[str, object]) -> None:
    if value.get("schema_id") != AXMATH_PRIVATE_PARSER_CONTRACT_SCHEMA_ID:
        raise AxMathPrivateParserContractError("Unexpected AxMath parser contract schema_id.")
    if value.get("source_family") != AXMATH_SOURCE_FAMILY:
        raise AxMathPrivateParserContractError("Unexpected AxMath parser contract source_family.")


def _validate_provenance(value: Mapping[str, object]) -> None:
    provenance = _mapping_at(value, "provenance")
    if provenance.get("metadata_only") is not True:
        raise AxMathPrivateParserContractError("AxMath parser contract must be metadata-only.")
    if provenance.get("public_safe") is not True:
        raise AxMathPrivateParserContractError("AxMath parser contract must be marked public-safe.")
    run_ids = tuple(str(item) for item in _sequence_at(provenance, "derived_from_private_run_ids"))
    if run_ids != EXPECTED_PRIVATE_RUN_IDS:
        raise AxMathPrivateParserContractError("AxMath parser contract provenance run ids changed.")


def _validate_material_boundaries(value: Mapping[str, object]) -> None:
    boundaries = _mapping_at(value, "private_material_boundaries")
    for field in PRIVATE_BOUNDARY_FIELDS:
        if boundaries.get(field) is not False:
            raise AxMathPrivateParserContractError(
                f"AxMath parser contract private material boundary must be false: {field}"
            )


def _validate_clean_surface(clean: Mapping[str, object]) -> None:
    row_count = _int_at(clean, "row_count")
    canonical_count = _int_at(clean, "canonical_source_free_count")
    exact_count = _int_at(clean, "exact_replay_diagnostic_count")
    recovered_ids = tuple(str(item) for item in _sequence_at(clean, "recovered_sample_ids"))
    substitution_counts = _mapping_at(clean, "substitution_counts")

    if row_count != 355 or canonical_count != row_count:
        raise AxMathPrivateParserContractError("Clean canonical source-free counts are inconsistent.")
    if exact_count + len(recovered_ids) != row_count:
        raise AxMathPrivateParserContractError("Clean exact diagnostic counts are inconsistent.")
    if recovered_ids != ("CP2-0067", "CP2-0351"):
        raise AxMathPrivateParserContractError("Clean canonical recovery sample ids changed.")
    if clean.get("canonical_rule") != CANONICAL_RULE_UNWRAP_SINGLE_CHILD_MROW_V1:
        raise AxMathPrivateParserContractError("Clean canonical rule changed.")
    if sum(_int_value(value, path=f"substitution_counts.{key}") for key, value in substitution_counts.items()) != row_count:
        raise AxMathPrivateParserContractError("Clean substitution counts do not sum to row count.")


def _validate_original_surface(original: Mapping[str, object]) -> None:
    row_count = _int_at(original, "row_count")
    parsed_count = _int_at(original, "priority_v2_parsed_count")
    unsupported_count = _int_at(original, "priority_v2_unsupported_count")
    guarded_rows = _int_at(original, "guarded_original_rows")
    promotable_rows = _int_at(original, "source_free_promotable_guarded_rows")
    recaptured_count = _int_at(original, "recaptured_exact_rows_source_free_count")
    reason_counts = {
        str(key): _int_value(child, path=f"unsupported_reason_counts.{key}")
        for key, child in _mapping_at(original, "unsupported_reason_counts").items()
    }

    if row_count != 355 or parsed_count + unsupported_count != row_count:
        raise AxMathPrivateParserContractError("Original priority-v2 counts are inconsistent.")
    if parsed_count != 327 or unsupported_count != 28:
        raise AxMathPrivateParserContractError("Original priority-v2 contract counts changed.")
    if guarded_rows != unsupported_count or promotable_rows != 0:
        raise AxMathPrivateParserContractError("Original fail-closed guarded-row boundary changed.")
    if recaptured_count != 21:
        raise AxMathPrivateParserContractError("Recaptured exact row source-free count changed.")
    if reason_counts != {
        "input_side_conversion_collapse": 18,
        "same_stream_oracle_policy_conflict": 8,
        "script_grouping_policy_conflict": 2,
    }:
        raise AxMathPrivateParserContractError("Original unsupported reason counts changed.")
    if sum(reason_counts.values()) != unsupported_count:
        raise AxMathPrivateParserContractError("Original unsupported reason counts do not sum.")


def _validate_claims(claims: Mapping[str, object]) -> None:
    if claims.get("private_clean_canonical_completion") is not True:
        raise AxMathPrivateParserContractError("Private clean canonical completion must remain true.")
    for claim in FALSE_PUBLIC_CLAIMS:
        if claims.get(claim) is not False:
            raise AxMathPrivateParserContractError(f"Public or broad AxMath claim must remain false: {claim}")


def _validate_forbidden_evidence_sources(value: Mapping[str, object]) -> None:
    sources = tuple(str(item) for item in _sequence_at(value, "forbidden_evidence_sources"))
    for required in REQUIRED_FORBIDDEN_EVIDENCE_SOURCES:
        if required not in sources:
            raise AxMathPrivateParserContractError(
                f"Forbidden evidence source is missing from contract: {required}"
            )


def _reject_private_material(value: object, *, path: str = "") -> None:
    violations: list[str] = []
    _collect_private_material_violations(value, path=path, violations=violations)
    if violations:
        raise AxMathPrivateParserContractError(
            "AxMath parser contract contains private material indicators.",
            violations=violations,
        )


def _collect_private_material_violations(
    value: object,
    *,
    path: str,
    violations: list[str],
) -> None:
    if isinstance(value, Mapping):
        for key, child in value.items():
            key_text = str(key)
            child_path = f"{path}.{key_text}" if path else key_text
            normalized = key_text.strip().lower().replace("-", "_")
            if key_text not in _ALLOWED_PRIVATE_MATERIAL_KEYS and any(
                fragment in normalized for fragment in _FORBIDDEN_FIELD_FRAGMENTS
            ):
                violations.append(child_path)
            _collect_private_material_violations(child, path=child_path, violations=violations)
    elif isinstance(value, Sequence) and not isinstance(value, str | bytes | bytearray):
        for index, child in enumerate(value):
            child_path = f"{path}[{index}]" if path else f"[{index}]"
            _collect_private_material_violations(child, path=child_path, violations=violations)
    elif isinstance(value, str) and not path.endswith("forbidden_evidence_sources"):
        lowered = value.lower()
        if any(marker in lowered for marker in _USER_PROFILE_MARKERS):
            violations.append(path or "string-value")
        if any(marker in lowered for marker in _MATH_TAG_MARKERS):
            violations.append(path or "string-value")
        if _LONG_HEX_RE.fullmatch(lowered):
            violations.append(path or "string-value")
    elif isinstance(value, bytes | bytearray | memoryview):
        violations.append(path or "byte-value")


def _mapping_at(value: Mapping[str, object], dotted_path: str) -> Mapping[str, object]:
    current: object = value
    for part in dotted_path.split("."):
        if not isinstance(current, Mapping) or not isinstance(current.get(part), Mapping):
            raise AxMathPrivateParserContractError(f"Missing mapping: {dotted_path}")
        current = current[part]
    return current


def _sequence_at(value: Mapping[str, object], key: str) -> Sequence[object]:
    child = value.get(key)
    if not isinstance(child, Sequence) or isinstance(child, str | bytes | bytearray):
        raise AxMathPrivateParserContractError(f"Missing sequence: {key}")
    return child


def _int_at(value: Mapping[str, object], key: str) -> int:
    return _int_value(value.get(key), path=key)


def _int_value(value: object, *, path: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise AxMathPrivateParserContractError(f"Expected integer at {path}.")
    if value < 0:
        raise AxMathPrivateParserContractError(f"Expected non-negative integer at {path}.")
    return value


__all__ = [
    "AXMATH_PRIVATE_PARSER_CONTRACT_SCHEMA_ID",
    "AXMATH_SOURCE_FAMILY",
    "AxMathPrivateParserContract",
    "AxMathPrivateParserContractError",
    "load_axmath_private_parser_contract",
    "normalize_axmath_private_parser_contract",
]
