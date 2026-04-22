from document_equation_migration.manifest import DocumentRecord, FormulaRecord, Manifest
from document_equation_migration.routing import build_execution_plan_report, build_routing_report
from document_equation_migration.source_taxonomy import SourceFamily, SourceRole


def make_formula(formula_id: str, source_family: SourceFamily) -> FormulaRecord:
    return FormulaRecord(
        formula_id=formula_id,
        source_family=source_family,
        source_role=SourceRole.NATIVE_SOURCE,
        doc_part_path="word/document.xml",
        story_type="main",
        storage_kind="test",
    )


def test_build_routing_report_orders_by_priority() -> None:
    manifest = Manifest(
        document=DocumentRecord(
            document_id="sample",
            input_path="sample.docx",
            input_sha256="abc",
            container_format="docx",
            detector_version="0.1.0",
        ),
        formulas=[
            make_formula("f1", SourceFamily.AXMATH_OLE),
            make_formula("f2", SourceFamily.OMML_NATIVE),
            make_formula("f3", SourceFamily.UNKNOWN_OLE),
        ],
    )
    manifest.update_source_counts()

    report = build_routing_report(manifest)

    assert report["formula_count"] == 3
    assert report["recommended_sequence"] == ["omml-native", "axmath-ole", "unknown-ole"]
    assert report["manual_review_required"] is True
    assert "axmath-ole" in report["manual_review_reasons"]
    assert "unknown-ole" in report["manual_review_reasons"]


def test_build_routing_report_handles_empty_manifest() -> None:
    manifest = Manifest(
        document=DocumentRecord(
            document_id="empty",
            input_path="empty.docx",
            input_sha256="abc",
            container_format="docx",
            detector_version="0.1.0",
        ),
        formulas=[],
    )
    manifest.update_source_counts()

    report = build_routing_report(manifest)

    assert report["formula_count"] == 0
    assert report["route_plan"] == []
    assert report["recommended_sequence"] == []
    assert report["manual_review_required"] is False


def test_build_execution_plan_report_uses_routing_output() -> None:
    manifest = Manifest(
        document=DocumentRecord(
            document_id="sample",
            input_path="sample.docx",
            input_sha256="abc",
            container_format="docx",
            detector_version="0.1.0",
        ),
        formulas=[
            make_formula("f1", SourceFamily.OMML_NATIVE),
            make_formula("f2", SourceFamily.AXMATH_OLE),
        ],
    )
    manifest.update_source_counts()

    report = build_execution_plan_report(manifest)

    assert report["formula_count"] == 2
    assert report["recommended_sequence"] == ["omml-native", "axmath-ole"]
    assert [item["provider"] for item in report["steps"]] == ["omml", "axmath"]
