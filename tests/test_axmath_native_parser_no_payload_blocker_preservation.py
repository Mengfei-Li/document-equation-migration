import json
from pathlib import Path
from typing import Any

import pytest

from document_equation_migration.axmath_structural_metadata import (
    AxMathStructuralMetadataError,
    classify_axmath_structural_metadata_rows,
)


FIXTURE_PATH = Path(__file__).parent / "fixtures" / "axmath" / "no-payload-blocker-contracts.json"


def _load_contract() -> dict[str, Any]:
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def _rows() -> list[dict[str, Any]]:
    return list(_load_contract()["fixture_contracts"])


def _row(row_id: int) -> dict[str, Any]:
    matches = [row for row in _rows() if row["row_id"] == row_id]
    assert len(matches) == 1
    return matches[0]


def _rows_by_class() -> dict[str, list[int]]:
    grouped: dict[str, list[int]] = {}
    for row in _rows():
        grouped.setdefault(row["blocker_class"], []).append(row["row_id"])
    return {key: sorted(value) for key, value in grouped.items()}


def _assert_claim_freeze_false(claims: dict[str, Any]) -> None:
    assert claims == {
        "conversion_claim": False,
        "native_parser_claim": False,
        "export_success_claim": False,
        "public_fixture_eligibility": False,
        "score_10_claim": False,
    }


def _assert_no_output_or_payload_keys(value: Any) -> None:
    forbidden_fragments = (
        "base64",
        "body",
        "bytes",
        "contents",
        "decoded",
        "formula",
        "hex",
        "image",
        "latex",
        "mathml",
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
    allowed_keys = {"forbidden_evidence_sources"}
    if isinstance(value, dict):
        for key, child in value.items():
            normalized = key.lower().replace("-", "_")
            assert key in allowed_keys or not any(fragment in normalized for fragment in forbidden_fragments)
            _assert_no_output_or_payload_keys(child)
    elif isinstance(value, list):
        for child in value:
            _assert_no_output_or_payload_keys(child)


def _assert_row_contract(
    row_id: int,
    *,
    blocker_class: str,
    unsupported_boundary: str,
    hypothesis_status: str,
    tail_token_class_id: str,
    stream_length: int,
    tail_length: int,
    candidate_count_relation_pass_count: int,
    candidate_count_relation_fail_count: int,
) -> None:
    row = _row(row_id)
    contract = _load_contract()

    assert row["blocker_class"] == blocker_class
    assert row["unsupported_boundary"] == unsupported_boundary
    assert row["hypothesis_status"] == hypothesis_status
    assert row["tail_token_class_id"] == tail_token_class_id
    assert row["stream_length"] == stream_length
    assert row["tail_length"] == tail_length
    assert row["candidate_count_relation_pass_count"] == candidate_count_relation_pass_count
    assert row["candidate_count_relation_fail_count"] == candidate_count_relation_fail_count
    _assert_claim_freeze_false(contract["claim_freeze"])

    classified = classify_axmath_structural_metadata_rows([row]).to_dict()
    assert classified["record_count"] == 1
    assert classified["claim_boundaries"]["conversion_claim"] is False
    assert classified["claim_boundaries"]["native_parser_claim"] is False
    assert classified["claim_boundaries"]["export_success_claim"] is False
    assert classified["claim_boundaries"]["public_fixture_eligibility"] is False
    assert classified["records"][0]["metadata_only_row_covered"] is True
    assert classified["records"][0]["unsupported_boundary"] == unsupported_boundary

    if blocker_class == "candidate_count_relation_absent":
        assert row["candidate_count_relation_pass_count"] == 0
        assert classified["records"][0]["structural_status"] == "candidate_count_relation_absent"
        assert "candidate_count_relation_absent" in classified["blocker_ids"]
    else:
        assert blocker_class == "opaque_tail"
        assert classified["records"][0]["structural_status"] == "opaque_tail_boundary"
        assert "opaque_tail_nonblocking_boundary" in classified["blocker_ids"]


def test_axmath_native_parser_no_payload_contract_target_row_sets_exact() -> None:
    contract = _load_contract()

    assert len(contract["fixture_contracts"]) == 12
    assert contract["expected_rows"] == {
        "opaque_tail": [24, 25, 26],
        "candidate_count_relation_absent": [5, 7, 9, 11, 12, 13, 14, 15, 18],
    }
    assert _rows_by_class() == contract["expected_rows"]


def test_axmath_native_parser_no_payload_contract_claim_freeze_false() -> None:
    contract = _load_contract()

    _assert_claim_freeze_false(contract["claim_freeze"])
    classified = classify_axmath_structural_metadata_rows(_rows()).to_dict()
    assert classified["claim_boundaries"]["conversion_claim"] is False
    assert classified["claim_boundaries"]["native_parser_claim"] is False
    assert classified["claim_boundaries"]["export_success_claim"] is False
    assert classified["claim_boundaries"]["public_fixture_eligibility"] is False


def test_axmath_native_parser_no_payload_contract_rejects_forbidden_evidence_sources() -> None:
    contract = _load_contract()

    assert "raw Contents bytes" in contract["forbidden_evidence_sources"]
    assert "formula body text" in contract["forbidden_evidence_sources"]
    assert "MathML, LaTeX, OMML, SVG, or markup formula bodies" in contract["forbidden_evidence_sources"]
    with pytest.raises(AxMathStructuralMetadataError):
        classify_axmath_structural_metadata_rows(
            [
                {
                    "row_id": 24,
                    "parse_status": "opaque_unsupported",
                    "raw_contents_bytes": "forbidden",
                }
            ]
        )


def test_axmath_native_parser_no_payload_contract_does_not_implement_body_decode() -> None:
    contract = _load_contract()
    classified = classify_axmath_structural_metadata_rows(_rows()).to_dict()

    assert contract["contract_scope"] == {
        "metadata_only": True,
        "parser_decode_attempted": False,
        "native_parser_support_claimed": False,
        "canonical_output_emitted": False,
    }
    assert classified["status"] == "metadata-only-blockers-preserved"
    assert "metadata_relation_consistent" not in classified["structural_status_counts"]
    _assert_no_output_or_payload_keys(contract["fixture_contracts"])


def test_axmath_native_parser_preserves_row_24_opaque_tail_without_payload() -> None:
    _assert_row_contract(
        24,
        blocker_class="opaque_tail",
        unsupported_boundary="opaque_tail",
        hypothesis_status="opaque_tail_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC014",
        stream_length=5533,
        tail_length=5194,
        candidate_count_relation_pass_count=373,
        candidate_count_relation_fail_count=4103,
    )


def test_axmath_native_parser_preserves_row_25_opaque_tail_without_payload() -> None:
    _assert_row_contract(
        25,
        blocker_class="opaque_tail",
        unsupported_boundary="opaque_tail",
        hypothesis_status="opaque_tail_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC015",
        stream_length=4333,
        tail_length=3994,
        candidate_count_relation_pass_count=260,
        candidate_count_relation_fail_count=3351,
    )


def test_axmath_native_parser_preserves_row_26_opaque_tail_without_payload() -> None:
    _assert_row_contract(
        26,
        blocker_class="opaque_tail",
        unsupported_boundary="opaque_tail",
        hypothesis_status="opaque_tail_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC016",
        stream_length=1853,
        tail_length=1514,
        candidate_count_relation_pass_count=104,
        candidate_count_relation_fail_count=1408,
    )


def test_axmath_native_parser_preserves_row_05_candidate_count_relation_absent_without_payload() -> None:
    _assert_row_contract(
        5,
        blocker_class="candidate_count_relation_absent",
        unsupported_boundary="candidate_count_relation_absent",
        hypothesis_status="relation_absent_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC005",
        stream_length=713,
        tail_length=374,
        candidate_count_relation_pass_count=0,
        candidate_count_relation_fail_count=501,
    )


def test_axmath_native_parser_preserves_row_07_candidate_count_relation_absent_without_payload() -> None:
    _assert_row_contract(
        7,
        blocker_class="candidate_count_relation_absent",
        unsupported_boundary="candidate_count_relation_absent",
        hypothesis_status="relation_absent_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC005",
        stream_length=713,
        tail_length=374,
        candidate_count_relation_pass_count=0,
        candidate_count_relation_fail_count=501,
    )


def test_axmath_native_parser_preserves_row_09_candidate_count_relation_absent_without_payload() -> None:
    _assert_row_contract(
        9,
        blocker_class="candidate_count_relation_absent",
        unsupported_boundary="candidate_count_relation_absent",
        hypothesis_status="relation_absent_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC005",
        stream_length=713,
        tail_length=374,
        candidate_count_relation_pass_count=0,
        candidate_count_relation_fail_count=501,
    )


def test_axmath_native_parser_preserves_row_11_candidate_count_relation_absent_without_payload() -> None:
    _assert_row_contract(
        11,
        blocker_class="candidate_count_relation_absent",
        unsupported_boundary="candidate_count_relation_absent",
        hypothesis_status="relation_absent_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC009",
        stream_length=593,
        tail_length=254,
        candidate_count_relation_pass_count=0,
        candidate_count_relation_fail_count=412,
    )


def test_axmath_native_parser_preserves_row_12_candidate_count_relation_absent_without_payload() -> None:
    _assert_row_contract(
        12,
        blocker_class="candidate_count_relation_absent",
        unsupported_boundary="candidate_count_relation_absent",
        hypothesis_status="relation_absent_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC009",
        stream_length=593,
        tail_length=254,
        candidate_count_relation_pass_count=0,
        candidate_count_relation_fail_count=412,
    )


def test_axmath_native_parser_preserves_row_13_candidate_count_relation_absent_without_payload() -> None:
    _assert_row_contract(
        13,
        blocker_class="candidate_count_relation_absent",
        unsupported_boundary="candidate_count_relation_absent",
        hypothesis_status="relation_absent_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC009",
        stream_length=593,
        tail_length=254,
        candidate_count_relation_pass_count=0,
        candidate_count_relation_fail_count=412,
    )


def test_axmath_native_parser_preserves_row_14_candidate_count_relation_absent_without_payload() -> None:
    _assert_row_contract(
        14,
        blocker_class="candidate_count_relation_absent",
        unsupported_boundary="candidate_count_relation_absent",
        hypothesis_status="relation_absent_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC009",
        stream_length=593,
        tail_length=254,
        candidate_count_relation_pass_count=0,
        candidate_count_relation_fail_count=412,
    )


def test_axmath_native_parser_preserves_row_15_candidate_count_relation_absent_without_payload() -> None:
    _assert_row_contract(
        15,
        blocker_class="candidate_count_relation_absent",
        unsupported_boundary="candidate_count_relation_absent",
        hypothesis_status="relation_absent_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC009",
        stream_length=593,
        tail_length=254,
        candidate_count_relation_pass_count=0,
        candidate_count_relation_fail_count=412,
    )


def test_axmath_native_parser_preserves_row_18_candidate_count_relation_absent_without_payload() -> None:
    _assert_row_contract(
        18,
        blocker_class="candidate_count_relation_absent",
        unsupported_boundary="candidate_count_relation_absent",
        hypothesis_status="relation_absent_blocked_no_payload_no_narrowing",
        tail_token_class_id="TC005",
        stream_length=713,
        tail_length=374,
        candidate_count_relation_pass_count=0,
        candidate_count_relation_fail_count=501,
    )
