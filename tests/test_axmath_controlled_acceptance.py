import pytest

from document_equation_migration.axmath_controlled_acceptance import (
    RESULT_CLASS_EXACT,
    RESULT_CLASS_SEMANTIC_EQUIVALENT,
    RESULT_CLASS_UNSUPPORTED,
    AxMathControlledAcceptanceContractError,
    classify_axmath_controlled_acceptance_rows,
    default_axmath_controlled_acceptance_contract,
    normalize_axmath_controlled_acceptance_contract,
)


def test_default_contract_records_private_v2_counts_without_public_claims() -> None:
    contract = default_axmath_controlled_acceptance_contract().to_dict()

    assert contract["schema_id"] == "axmath.controlled_acceptance.public_contract.v1"
    assert contract["source_family"] == "axmath-ole"
    assert contract["decision_enum"] == "axmath_angle_direct_native_acceptance_refresh_passed"
    assert contract["controlled_sample_count"] == 355
    assert contract["accepted_count"] == 355
    assert contract["remaining_fail_closed_count"] == 0
    assert contract["result_class_counts"] == {
        "exact": 334,
        "exact_recapture": 21,
        "semantic_equivalent_rewrite_not_exact_source": 0,
    }
    assert contract["diagnostic_counts"] == {
        "exact_recapture": 21,
    }
    assert contract["artifact_boundaries"] == {
        "private_source_artifacts_included": False,
        "private_binary_material_included": False,
        "private_formula_bodies_included": False,
        "generated_canonical_bodies_included": False,
        "public_fixture_material_included": False,
    }
    assert contract["claim_boundaries"] == {
        "native_parser_claim": False,
        "conversion_claim": False,
        "export_success_claim": False,
        "public_fixture_eligibility": False,
        "public_native_parser_binding": False,
        "universal_axmath_support": False,
    }


def test_normalize_contract_accepts_public_safe_explicit_counts() -> None:
    contract = normalize_axmath_controlled_acceptance_contract(
        {
            "evidence_scope": "private-controlled-smoke-contract-only",
            "controlled_sample_count": 3,
            "accepted_count": 2,
            "remaining_fail_closed_count": 1,
            "result_class_counts": {
                RESULT_CLASS_EXACT: 1,
                RESULT_CLASS_SEMANTIC_EQUIVALENT: 1,
                RESULT_CLASS_UNSUPPORTED: 1,
            },
            "diagnostic_counts": {
                RESULT_CLASS_SEMANTIC_EQUIVALENT: 1,
            },
            "notes": ["public-safe synthetic aggregate"],
        }
    ).to_dict()

    assert contract["evidence_scope"] == "private-controlled-smoke-contract-only"
    assert contract["controlled_sample_count"] == 3
    assert contract["accepted_count"] == 2
    assert contract["remaining_fail_closed_count"] == 1
    assert contract["result_class_counts"][RESULT_CLASS_UNSUPPORTED] == 1
    assert contract["claim_boundaries"]["public_native_parser_binding"] is False


def test_classify_rows_preserves_semantic_equivalent_diagnostic_and_fail_closed() -> None:
    classified = classify_axmath_controlled_acceptance_rows(
        [
            {"result_class": RESULT_CLASS_EXACT},
            {
                "result_class": RESULT_CLASS_SEMANTIC_EQUIVALENT,
                "diagnostic_id": RESULT_CLASS_SEMANTIC_EQUIVALENT,
            },
            {"result_class": RESULT_CLASS_UNSUPPORTED},
        ]
    ).to_dict()

    assert classified["row_count"] == 3
    assert classified["accepted_count"] == 2
    assert classified["remaining_fail_closed_count"] == 1
    assert classified["result_class_counts"] == {
        RESULT_CLASS_EXACT: 1,
        RESULT_CLASS_SEMANTIC_EQUIVALENT: 1,
        RESULT_CLASS_UNSUPPORTED: 1,
    }
    assert classified["diagnostic_counts"] == {
        RESULT_CLASS_SEMANTIC_EQUIVALENT: 1,
    }
    assert classified["records"][1]["public_release_status"] == "contract-only-no-public-artifacts"
    assert classified["records"][2]["blocker_ids"] == ["controlled_acceptance_unsupported"]
    assert classified["claim_boundaries"]["universal_axmath_support"] is False


def test_rejects_private_or_formula_material_in_contract_input() -> None:
    with pytest.raises(AxMathControlledAcceptanceContractError) as exc_info:
        normalize_axmath_controlled_acceptance_contract(
            {
                "result_class_counts": {RESULT_CLASS_EXACT: 1},
                "accepted_count": 1,
                "controlled_sample_count": 1,
                "raw_contents_bytes": "forbidden",
            }
        )

    assert "raw_contents_bytes" in exc_info.value.forbidden_fields


def test_rejects_private_or_formula_material_in_row_input() -> None:
    with pytest.raises(AxMathControlledAcceptanceContractError) as exc_info:
        classify_axmath_controlled_acceptance_rows(
            [
                {
                    "result_class": RESULT_CLASS_EXACT,
                    "exported_mathml": "forbidden",
                }
            ]
        )

    assert "exported_mathml" in exc_info.value.forbidden_fields


def test_rejects_count_mismatch() -> None:
    with pytest.raises(AxMathControlledAcceptanceContractError):
        normalize_axmath_controlled_acceptance_contract(
            {
                "controlled_sample_count": 2,
                "accepted_count": 2,
                "remaining_fail_closed_count": 0,
                "result_class_counts": {
                    RESULT_CLASS_EXACT: 1,
                    RESULT_CLASS_UNSUPPORTED: 1,
                },
            }
        )
