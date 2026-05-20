import json
import re
from pathlib import Path
from typing import Any


FIXTURE_PATH = (
    Path(__file__).parent / "fixtures" / "axmath" / "private-parser-contract-no-payload.json"
)


def _load_contract() -> dict[str, Any]:
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def _assert_count_invariants(contract: dict[str, Any]) -> None:
    clean = contract["surfaces"]["clean_canonical_surface"]
    original = contract["surfaces"]["original_lossy_surface"]

    assert clean["canonical_source_free_count"] == clean["row_count"]
    assert clean["exact_replay_diagnostic_count"] + len(clean["recovered_sample_ids"]) == clean["row_count"]
    assert sum(clean["substitution_counts"].values()) == clean["row_count"]

    assert original["priority_v2_parsed_count"] + original["priority_v2_unsupported_count"] == original[
        "row_count"
    ]
    assert original["priority_v2_unsupported_count"] == original["guarded_original_rows"]
    assert original["source_free_promotable_guarded_rows"] == 0
    assert sum(original["unsupported_reason_counts"].values()) == original["priority_v2_unsupported_count"]


def _assert_no_private_material_fields(value: Any, *, path: str = "") -> None:
    forbidden_key_fragments = (
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
    allowed_keys = {"forbidden_evidence_sources"}

    if isinstance(value, dict):
        for key, child in value.items():
            child_path = f"{path}.{key}" if path else key
            normalized = key.lower().replace("-", "_")
            assert key in allowed_keys or not any(fragment in normalized for fragment in forbidden_key_fragments)
            _assert_no_private_material_fields(child, path=child_path)
    elif isinstance(value, list):
        for index, child in enumerate(value):
            child_path = f"{path}[{index}]"
            _assert_no_private_material_fields(child, path=child_path)
    elif isinstance(value, str) and not path.endswith("forbidden_evidence_sources"):
        math_tag_open = "<" + "math"
        math_tag_close = "</" + "math"
        assert "C:/Users/" not in value
        assert "C:\\Users\\" not in value
        assert math_tag_open not in value.lower()
        assert math_tag_close not in value.lower()
        assert not re.fullmatch(r"[0-9a-f]{32,}", value.lower())


def test_contract_records_latest_private_parser_surfaces_as_metadata() -> None:
    contract = _load_contract()

    assert contract["schema_id"] == "axmath.private_parser_contract.no_payload.v1"
    assert contract["source_family"] == "axmath-ole"
    assert contract["provenance"] == {
        "derived_from_private_run_ids": [
            "run_20260519_0038_private_canonical_acceptance_gate",
            "run_20260519_0039_original_payload_fail_closed_gate",
            "run_20260519_0040_private_parser_contract_package",
        ],
        "metadata_only": True,
        "public_safe": True,
    }
    assert contract["surfaces"]["clean_canonical_surface"] == {
        "row_count": 355,
        "canonical_source_free_count": 355,
        "exact_replay_diagnostic_count": 353,
        "canonical_rule": "unwrap_single_child_mrow_v1",
        "recovered_sample_ids": ["CP2-0067", "CP2-0351"],
        "substitution_counts": {"original_v2_rows": 334, "recaptured_rows": 21},
    }
    assert contract["surfaces"]["original_lossy_surface"] == {
        "row_count": 355,
        "priority_v2_parsed_count": 327,
        "priority_v2_unsupported_count": 28,
        "guarded_original_rows": 28,
        "source_free_promotable_guarded_rows": 0,
        "recaptured_exact_rows_source_free_count": 21,
        "unsupported_reason_counts": {
            "input_side_conversion_collapse": 18,
            "same_stream_oracle_policy_conflict": 8,
            "script_grouping_policy_conflict": 2,
        },
    }


def test_contract_preserves_public_claim_freeze() -> None:
    contract = _load_contract()

    assert contract["contract_scope"] == {
        "scope_id": "private_controlled_corpus_no_payload_contract",
        "private_controlled_corpus_contract": True,
        "runtime_word_axmath_dependency": False,
        "public_conversion_support": False,
        "public_fixture_eligibility": False,
        "broad_native_parser_support": False,
        "universal_axmath_support": False,
    }
    assert contract["claim_freeze"] == {
        "private_clean_canonical_completion": True,
        "private_exact_source_completion": False,
        "original_controlled_corpus_source_free_completion": False,
        "source_free_native_parser_completion": False,
        "broad_native_parser_completion": False,
        "exact_source_completion": False,
        "conversion_claim": False,
        "native_parser_claim": False,
        "export_success_claim": False,
        "public_fixture_eligibility": False,
        "public_conversion_support": False,
        "universal_axmath_support": False,
    }


def test_contract_declares_no_private_material_is_included() -> None:
    contract = _load_contract()

    assert contract["private_material_boundaries"] == {
        "binary_material_included": False,
        "canonical_markup_included": False,
        "document_artifacts_included": False,
        "source_text_included": False,
        "derived_digests_included": False,
        "display_media_included": False,
    }
    assert "raw Contents bytes" in contract["forbidden_evidence_sources"]
    assert "MathML, LaTeX, OMML, SVG, or markup formula bodies" in contract[
        "forbidden_evidence_sources"
    ]
    assert "payload-derived hex, base64, or digest material" in contract["forbidden_evidence_sources"]
    _assert_no_private_material_fields(contract)


def test_contract_count_invariants_are_closed() -> None:
    contract = _load_contract()

    _assert_count_invariants(contract)


def test_contract_does_not_promote_public_conversion_or_parser_support() -> None:
    contract = _load_contract()

    assert contract["claim_freeze"]["private_clean_canonical_completion"] is True
    for claim in (
        "private_exact_source_completion",
        "original_controlled_corpus_source_free_completion",
        "source_free_native_parser_completion",
        "broad_native_parser_completion",
        "conversion_claim",
        "native_parser_claim",
        "public_fixture_eligibility",
        "public_conversion_support",
        "universal_axmath_support",
    ):
        assert contract["claim_freeze"][claim] is False
