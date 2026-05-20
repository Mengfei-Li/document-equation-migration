from xml.etree import ElementTree as ET

from document_equation_migration.canonical_mathml_evidence import (
    CANONICALIZATION_SUMMARY_REQUIRED_FIELDS,
    CANONICAL_BLOCKER_RECORD_REQUIRED_FIELDS,
    CANONICAL_EVIDENCE_SOURCE_FAMILIES,
    CANONICAL_FORMULA_COUNT_PARITY_VALUES,
    CANONICAL_PROVENANCE_REQUIRED_FIELDS,
    CANONICAL_VALIDATION_EVIDENCE_REQUIRED_FIELDS,
    canonical_evidence_schema,
    mathml_property_signals,
    property_summary,
    sha256_text,
)


def test_sha256_text_hashes_utf8_text() -> None:
    assert sha256_text("x") == "2d711642b726b04401627ca9fbac32f5c8530fb1903cc4db02258717921a4881"


def test_mathml_property_signals_extracts_selected_attributes() -> None:
    payload = """<math xmlns="http://www.w3.org/1998/Math/MathML" display="block">
      <semantics>
        <mfrac linethickness="0"><mi mathvariant="bold">x</mi><mi>y</mi></mfrac>
        <mfenced separators=";"><mi>a</mi></mfenced>
        <mo movablelimits="true">∑</mo>
        <mover accent="true"><mi>z</mi><mo>~</mo></mover>
        <munder accentunder="true"><mi>q</mi><mo>¯</mo></munder>
        <annotation encoding="application/x-tex">x/y</annotation>
      </semantics>
    </math>"""

    signals = mathml_property_signals(ET.fromstring(payload))

    assert signals["root_display"] == "block"
    assert signals["has_semantics"] is True
    assert signals["has_annotation"] is True
    assert signals["has_mfrac_linethickness"] is True
    assert signals["has_mfenced_separators"] is True
    assert signals["has_movablelimits"] is True
    assert signals["has_mathvariant"] is True
    assert signals["has_accent"] is True
    assert signals["has_accentunder"] is True


def test_property_summary_aggregates_signals() -> None:
    items = [
        {
            "property_signals": {
                "root_display": "block",
                "mathml_attribute_count": 3,
                "has_annotation": True,
                "has_mfrac_bevelled": True,
            }
        },
        {
            "property_signals": {
                "root_display": "",
                "mathml_attribute_count": 2,
                "has_annotation": False,
                "has_mathvariant": True,
            }
        },
    ]

    summary = property_summary(items)

    assert summary["mathml_attribute_count"] == 5
    assert summary["root_display_values"] == ["block"]
    assert summary["signal_counts"]["has_annotation"] == 1
    assert summary["signal_counts"]["has_mfrac_bevelled"] == 1
    assert summary["signal_counts"]["has_mathvariant"] == 1


def test_canonical_evidence_schema_exports_current_contract() -> None:
    schema = canonical_evidence_schema()

    assert schema["version"] == "2026-05-05"
    assert schema["source_families"] == list(CANONICAL_EVIDENCE_SOURCE_FAMILIES)
    assert set(schema["source_families"]) == {
        "mathtype-ole",
        "omml-native",
        "odf-native",
        "libreoffice-transformed",
        "equation-editor-3-ole",
        "axmath-ole",
    }
    assert schema["canonicalization_summary"]["required_fields"] == list(
        CANONICALIZATION_SUMMARY_REQUIRED_FIELDS
    )
    assert schema["canonicalization_summary"]["formula_count_parity_values"] == list(
        CANONICAL_FORMULA_COUNT_PARITY_VALUES
    )
    assert schema["source_to_canonical_provenance"]["required_fields"] == list(
        CANONICAL_PROVENANCE_REQUIRED_FIELDS
    )
    assert schema["blocker_record"]["required_fields"] == list(
        CANONICAL_BLOCKER_RECORD_REQUIRED_FIELDS
    )
    assert schema["validation_evidence"]["required_fields"] == list(
        CANONICAL_VALIDATION_EVIDENCE_REQUIRED_FIELDS
    )
