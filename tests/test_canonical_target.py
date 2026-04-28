from document_equation_migration.canonical_target import canonical_mathml_contract_for_source_family


def test_canonical_mathml_contract_covers_source_lines() -> None:
    expected_status = {
        "mathtype-ole": "external-tool-gated",
        "omml-native": "implemented-basic",
        "equation-editor-3-ole": "fixture-gated",
        "axmath-ole": "export-gated",
        "odf-native": "implemented",
        "libreoffice-transformed": "bridge-review-gated",
    }

    for source_family, status in expected_status.items():
        contract = canonical_mathml_contract_for_source_family(source_family).to_dict()
        assert contract["source_family"] == source_family
        assert contract["target_format"] == "canonical-mathml"
        assert contract["contract_status"] == status
        assert contract["required_evidence"]


def test_unknown_source_contract_stays_manual_triage() -> None:
    contract = canonical_mathml_contract_for_source_family("unknown-source").to_dict()

    assert contract["target_format"] == "canonical-mathml"
    assert contract["contract_status"] == "manual-triage"
    assert contract["conversion_claim"] is False
