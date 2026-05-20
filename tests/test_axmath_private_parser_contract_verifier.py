import copy
import json
from pathlib import Path

import pytest

from document_equation_migration.axmath_private_parser_contract import (
    AXMATH_PRIVATE_PARSER_CONTRACT_SCHEMA_ID,
    AXMATH_SOURCE_FAMILY,
    AxMathPrivateParserContractError,
    load_axmath_private_parser_contract,
    normalize_axmath_private_parser_contract,
)


FIXTURE_PATH = (
    Path(__file__).parent / "fixtures" / "axmath" / "private-parser-contract-no-payload.json"
)


def _contract_dict() -> dict[str, object]:
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def test_loads_public_no_payload_contract_fixture() -> None:
    contract = load_axmath_private_parser_contract(FIXTURE_PATH)

    assert contract.schema_id == AXMATH_PRIVATE_PARSER_CONTRACT_SCHEMA_ID
    assert contract.source_family == AXMATH_SOURCE_FAMILY
    assert contract.private_run_ids == (
        "run_20260519_0038_private_canonical_acceptance_gate",
        "run_20260519_0039_original_payload_fail_closed_gate",
        "run_20260519_0040_private_parser_contract_package",
    )
    assert contract.clean_row_count == 355
    assert contract.clean_canonical_source_free_count == 355
    assert contract.clean_exact_replay_diagnostic_count == 353
    assert contract.clean_recovered_sample_ids == ("CP2-0067", "CP2-0351")
    assert contract.original_priority_v2_parsed_count == 327
    assert contract.original_priority_v2_unsupported_count == 28
    assert contract.original_source_free_promotable_guarded_rows == 0
    assert contract.original_unsupported_reason_counts == {
        "input_side_conversion_collapse": 18,
        "same_stream_oracle_policy_conflict": 8,
        "script_grouping_policy_conflict": 2,
    }


def test_summary_is_stable_and_keeps_public_claims_false() -> None:
    summary = load_axmath_private_parser_contract(FIXTURE_PATH).summary()

    assert summary == {
        "source_family": "axmath-ole",
        "schema_id": "axmath.private_parser_contract.no_payload.v1",
        "clean_canonical_source_free": (355, 355),
        "clean_exact_diagnostic": (353, 355),
        "original_priority_v2": (327, 355),
        "original_fail_closed": (28, 355),
        "public_claims_false": True,
    }


def test_rejects_public_claim_drift() -> None:
    value = _contract_dict()
    value["claim_freeze"]["public_conversion_support"] = True  # type: ignore[index]

    with pytest.raises(AxMathPrivateParserContractError) as exc_info:
        normalize_axmath_private_parser_contract(value)

    assert "public_conversion_support" in str(exc_info.value)


def test_rejects_count_drift() -> None:
    value = _contract_dict()
    value["surfaces"]["original_lossy_surface"]["priority_v2_parsed_count"] = 328  # type: ignore[index]

    with pytest.raises(AxMathPrivateParserContractError, match="Original priority-v2 counts"):
        normalize_axmath_private_parser_contract(value)


def test_rejects_forbidden_private_material_keys_and_values() -> None:
    value = _contract_dict()
    value["surfaces"]["clean_canonical_surface"]["raw_contents_bytes"] = "forbidden"  # type: ignore[index]

    with pytest.raises(AxMathPrivateParserContractError) as exc_info:
        normalize_axmath_private_parser_contract(value)

    assert any(item.endswith("raw_contents_bytes") for item in exc_info.value.violations)

    value = _contract_dict()
    value["provenance"]["local_path"] = "C:" + "/Users/" + "example/private.docx"  # type: ignore[index]

    with pytest.raises(AxMathPrivateParserContractError) as path_exc_info:
        normalize_axmath_private_parser_contract(value)

    assert "provenance.local_path" in path_exc_info.value.violations


def test_rejects_missing_forbidden_evidence_source() -> None:
    value = copy.deepcopy(_contract_dict())
    value["forbidden_evidence_sources"] = ["raw Contents bytes"]

    with pytest.raises(AxMathPrivateParserContractError, match="Forbidden evidence source"):
        normalize_axmath_private_parser_contract(value)
