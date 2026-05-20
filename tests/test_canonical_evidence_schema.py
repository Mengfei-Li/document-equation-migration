from pathlib import Path

from document_equation_migration.canonical_mathml_evidence import (
    CANONICALIZATION_SUMMARY_REQUIRED_FIELDS,
    CANONICAL_BLOCKER_RECORD_REQUIRED_FIELDS,
    CANONICAL_CLAIM_BOUNDARY_FIELDS,
    CANONICAL_EVIDENCE_SOURCE_FAMILIES,
    CANONICAL_FORMULA_COUNT_PARITY_VALUES,
    CANONICAL_PROVENANCE_REQUIRED_FIELDS,
    CANONICAL_VALIDATION_EVIDENCE_REQUIRED_FIELDS,
)


DOC_PATH = Path(__file__).resolve().parents[1] / "docs" / "canonical-evidence-schema.md"


def _assert_required_fields(payload: dict[str, object], fields: tuple[str, ...]) -> None:
    missing = [field for field in fields if field not in payload]
    assert missing == []


def test_schema_doc_lists_exported_contract_fields() -> None:
    text = DOC_PATH.read_text(encoding="utf-8")

    for source_family in CANONICAL_EVIDENCE_SOURCE_FAMILIES:
        assert f"`{source_family}`" in text

    for field in CANONICALIZATION_SUMMARY_REQUIRED_FIELDS:
        assert f"`{field}`" in text

    for field in CANONICAL_PROVENANCE_REQUIRED_FIELDS:
        assert f"`{field}`" in text

    for field in CANONICAL_BLOCKER_RECORD_REQUIRED_FIELDS:
        assert f"`{field}`" in text

    for field in CANONICAL_VALIDATION_EVIDENCE_REQUIRED_FIELDS:
        assert f"`{field}`" in text

    for field in CANONICAL_CLAIM_BOUNDARY_FIELDS:
        assert f"`{field}`" in text


def test_minimal_canonicalization_summary_satisfies_schema() -> None:
    provenance = {
        "formula_id": "omml-0001",
        "source_omml_path": "normalized/omml-0001.xml",
        "canonical_artifact_path": "canonical-mathml/omml-canonical-0001.xml",
        "source_sha256": "source-hash",
        "canonical_sha256": "canonical-hash",
        "preservation_status": "converted-omml-to-canonical-mathml",
        "property_signals": {
            "root_display": "block",
            "mathml_attribute_count": 1,
            "has_mfrac_linethickness": True,
        },
    }
    summary = {
        "strategy": "internal-basic-omml-to-presentation-mathml",
        "expected_formula_count": 1,
        "canonical_mathml_count": 1,
        "unsupported_fragment_count": 0,
        "formula_count_parity": "passed",
        "canonical_mathml_dir": "canonical-mathml",
        "source_to_canonical_provenance": [provenance],
        "property_summary": {
            "mathml_attribute_count": 1,
            "root_display_values": ["block"],
            "signal_counts": {"has_mfrac_linethickness": 1},
        },
        "unsupported_fragments": [],
    }

    _assert_required_fields(summary, CANONICALIZATION_SUMMARY_REQUIRED_FIELDS)
    _assert_required_fields(provenance, CANONICAL_PROVENANCE_REQUIRED_FIELDS)
    assert summary["formula_count_parity"] in CANONICAL_FORMULA_COUNT_PARITY_VALUES
    assert len(summary["source_to_canonical_provenance"]) == summary["canonical_mathml_count"]


def test_minimal_blocker_record_satisfies_schema() -> None:
    blocker = {
        "artifact_type": "axmath-export-assisted-blocker-record",
        "source_family": "axmath-ole",
        "canonical_target": {
            "source_family": "axmath-ole",
            "target_format": "canonical-mathml",
            "contract_status": "export-gated",
            "conversion_claim": False,
        },
        "status": "blocked-external-tool",
        "required_evidence": [
            "reviewed canonical MathML artifact(s), or LaTeX plus validated MathML conversion",
        ],
        "next_ready_condition": "Provide a reviewed AxMath export bundle before accepting canonical MathML.",
        "export_admissibility": {
            "target_stage": "export-to-canonical-mathml",
        },
    }

    _assert_required_fields(blocker, CANONICAL_BLOCKER_RECORD_REQUIRED_FIELDS)
    assert blocker["canonical_target"]["target_format"] == "canonical-mathml"
    assert blocker["canonical_target"]["conversion_claim"] is False
    assert blocker["required_evidence"]


def test_validation_evidence_keeps_claim_boundary_fields_explicit() -> None:
    validation = {
        "artifact_type": "equation3-validation-evidence",
        "source_family": "equation-editor-3-ole",
        "status": "passed-limited",
        "limited_conversion_claim": True,
        "general_converter_claim": False,
        "deliverability_claim": False,
        "word_visual_fill_back_claim": False,
        "claim_boundary": {
            "not_accepted": [
                "Universal Equation Editor 3.0 support.",
                "Word visual fill-back or DOCX-route deliverability.",
            ],
        },
    }

    _assert_required_fields(validation, CANONICAL_VALIDATION_EVIDENCE_REQUIRED_FIELDS)
    assert set(validation) & set(CANONICAL_CLAIM_BOUNDARY_FIELDS)
    assert validation["limited_conversion_claim"] is True
    assert validation["general_converter_claim"] is False
    assert validation["deliverability_claim"] is False
