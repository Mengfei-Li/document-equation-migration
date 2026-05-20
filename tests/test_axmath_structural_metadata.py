import pytest

from document_equation_migration.axmath_structural_metadata import (
    AxMathStructuralMetadataError,
    classify_axmath_structural_metadata,
    classify_axmath_structural_metadata_rows,
    contains_axmath_structural_metadata,
)


def synthetic_row(parse_status: str, **overrides: object) -> dict[str, object]:
    row: dict[str, object] = {
        "row_id": 1,
        "object_id": "synthetic-object-1",
        "object_ordinal": 1,
        "stream_locator_id": "synthetic-stream-1",
        "stream_length": 573,
        "tail_token_class_id": "TC001",
        "tail_window_class": "small-tail",
        "tail_length": 234,
        "tail_start_candidate_offset": 339,
        "trailer_length": 8,
        "parse_status": parse_status,
        "unsupported_boundary": "none",
        "candidate_boundary_count": 2,
        "candidate_count_attempt_count": 10,
        "candidate_count_relation_pass_count": 2,
        "candidate_count_relation_fail_count": 8,
        "candidate_count_widths": [1, 2],
        "relation_check_summary": {"attempted": True},
        "duplicate_invariance_status": "passed",
        "metadata_only_row_covered": True,
        "conversion_claim": True,
        "native_parser_claim": True,
        "export_success_claim": True,
        "public_fixture_eligibility": True,
    }
    row.update(overrides)
    return row


def test_candidate_relation_pass_maps_to_metadata_relation_consistent() -> None:
    result = classify_axmath_structural_metadata(
        synthetic_row("candidate_count_relation_pass")
    ).to_dict()

    assert result["source_family"] == "axmath-ole"
    assert result["row_id"] == 1
    assert result["structural_status"] == "metadata_relation_consistent"
    assert result["body_decode_status"] == "not_attempted"
    assert result["blocker_ids"] == []
    assert result["metadata_only_row_covered"] is True
    assert result["conversion_claim"] is False
    assert result["native_parser_claim"] is False
    assert result["export_success_claim"] is False
    assert result["public_fixture_eligibility"] is False
    assert result["metadata"]["tail_start_candidate_offset"] == 339
    assert "conversion_claim" not in result["metadata"]


def test_candidate_relation_fail_preserves_hypothesis_only_blocker() -> None:
    result = classify_axmath_structural_metadata(
        synthetic_row(
            "candidate_count_relation_fail",
            row_id=5,
            unsupported_boundary="none",
        )
    ).to_dict()

    assert result["structural_status"] == "candidate_count_relation_absent"
    assert result["body_decode_status"] == "blocked_count_relation_hypothesis_only"
    assert result["blocker_ids"] == ["candidate_count_relation_absent"]
    assert result["conversion_claim"] is False
    assert result["native_parser_claim"] is False


def test_opaque_unsupported_preserves_opaque_tail_boundary() -> None:
    result = classify_axmath_structural_metadata(
        synthetic_row(
            "opaque_unsupported",
            row_id=24,
            tail_token_class_id="TC014",
            tail_window_class="large-tail",
            unsupported_boundary="opaque_tail",
        )
    ).to_dict()

    assert result["structural_status"] == "opaque_tail_boundary"
    assert result["body_decode_status"] == "blocked_opaque_tail"
    assert result["unsupported_boundary"] == "opaque_tail"
    assert result["blocker_ids"] == ["opaque_tail_nonblocking_boundary"]
    assert result["public_fixture_eligibility"] is False


def test_classifies_rows_batch_with_counts_and_frozen_claims() -> None:
    result = classify_axmath_structural_metadata_rows(
        [
            synthetic_row("candidate_count_relation_pass", row_id=1),
            synthetic_row("candidate_count_relation_fail", row_id=5),
            synthetic_row("opaque_unsupported", row_id=24),
        ]
    ).to_dict()

    assert result["status"] == "metadata-only-blockers-preserved"
    assert result["record_count"] == 3
    assert result["structural_status_counts"] == {
        "metadata_relation_consistent": 1,
        "candidate_count_relation_absent": 1,
        "opaque_tail_boundary": 1,
    }
    assert result["blocker_ids"] == [
        "candidate_count_relation_absent",
        "opaque_tail_nonblocking_boundary",
    ]
    assert result["claim_boundaries"] == {
        "conversion_claim": False,
        "native_parser_claim": False,
        "export_success_claim": False,
        "public_fixture_eligibility": False,
        "semantic_review_claim": False,
        "visual_equivalence_claim": False,
    }


def test_rejects_forbidden_payload_or_body_fields() -> None:
    with pytest.raises(AxMathStructuralMetadataError) as exc_info:
        classify_axmath_structural_metadata_rows(
            [
                {
                    "row_id": 1,
                    "raw_contents_bytes": "forbidden",
                    "nested": {"exported_mathml": "forbidden"},
                }
            ]
        )

    assert "raw_contents_bytes" in exc_info.value.forbidden_fields
    assert "nested.exported_mathml" in exc_info.value.forbidden_fields


def test_ignores_unknown_public_safe_fields() -> None:
    result = classify_axmath_structural_metadata(
        synthetic_row(
            "candidate_count_relation_pass",
            unknown_public_safe_counter=7,
        )
    ).to_dict()

    assert contains_axmath_structural_metadata({"rows": []}) is True
    assert result["ignored_fields"] == ["unknown_public_safe_counter"]
    assert "unknown_public_safe_counter" not in result["metadata"]
