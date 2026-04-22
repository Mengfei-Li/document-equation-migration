import json
import zipfile
from pathlib import Path

from document_equation_migration.execution_plan.model import ExecutionAction, ExecutionStep
from document_equation_migration.executor.model import ExecutionContext
from document_equation_migration.executor.omml import execute_omml_step


DOCX_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>
    </w:p>
  </w:body>
</w:document>
"""


def make_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", DOCX_XML)


def make_omml_step(*, requires_manual_review: bool = False) -> ExecutionStep:
    return ExecutionStep(
        source_family="omml-native",
        formula_count=1,
        route_kind="primary-source-first",
        confidence_policy="high",
        requires_manual_review=requires_manual_review,
        provider="omml",
        next_action="run-omml-native-pipeline",
        actions=(
            ExecutionAction(
                action_id="extract-omml",
                description="Extract OMML equations from OOXML document parts.",
            ),
            ExecutionAction(
                action_id="normalize-omml",
                description="Normalize OMML structure for deterministic downstream conversion.",
            ),
            ExecutionAction(
                action_id="render-check",
                description="Create validation plan for render parity checks.",
                blocking=requires_manual_review,
            ),
            ExecutionAction(
                action_id="package-omml-output",
                description="Package normalized OMML output and execution metadata.",
            ),
        ),
    )


def make_context(tmp_path: Path, input_path: Path) -> ExecutionContext:
    return ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(tmp_path / "execution-plan.json"),
        input_path=str(input_path),
        output_dir=str(tmp_path / "out"),
    )


def test_execute_omml_step_writes_manifest_and_validation_plan(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.docx"
    make_docx(input_path)

    reports = execute_omml_step(make_omml_step(), make_context(tmp_path, input_path))

    assert [report.action_id for report in reports] == [
        "extract-omml",
        "normalize-omml",
        "render-check",
        "package-omml-output",
    ]
    assert [report.status for report in reports] == ["completed", "completed", "skipped", "completed"]

    manifest_path = Path(reports[0].output_paths[0])
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["formula_count"] == 1
    assert manifest["items"][0]["kind"] == "oMath"
    assert Path(manifest["items"][0]["extracted_path"]).exists()

    package_report = reports[3]
    validation_target_path = Path(package_report.output_paths[1])
    validation_evidence_path = Path(package_report.output_paths[2])
    assert validation_target_path.exists()
    assert validation_evidence_path.exists()

    validation_evidence = json.loads(validation_evidence_path.read_text(encoding="utf-8"))
    assert validation_evidence["artifact_type"] == "omml-validation-evidence"
    assert validation_evidence["status"] == "evidence-collected"
    assert validation_evidence["gate_status"] == "pending-external-validation"
    assert validation_evidence["artifacts"]["manifest"]["formula_count"] == 1
    assert validation_evidence["artifacts"]["normalization_summary"]["normalized_count"] == 1
    assert validation_evidence["artifacts"]["package_metadata"]["provider"] == "omml"
    assert validation_evidence["artifacts"]["validation_target"]["validation_target_present"] is True
    assert validation_evidence["artifacts"]["validation_target"]["validation_target_docx"] == str(validation_target_path)
    assert validation_evidence["artifacts"]["validation_plan"]["status"] == "pending-external-validation"
    assert validation_evidence["artifacts"]["validation_plan"]["review_mode"] == "spot-check"
    assert {item["id"] for item in validation_evidence["evidence_checks"]} >= {
        "manifest-present",
        "normalization-summary-present",
        "package-metadata-present",
        "validation-target-present",
        "validation-plan-present",
        "word-open",
        "word-export-pdf",
        "render-parity",
    }
    assert any(
        item["id"] == "word-open" and item["status"] == "not-run"
        for item in validation_evidence["evidence_checks"]
    )

    validation_plan_path = Path(reports[2].output_paths[0])
    validation_plan = json.loads(validation_plan_path.read_text(encoding="utf-8"))
    assert validation_plan["artifact_type"] == "omml-validation-plan"
    assert validation_plan["status"] == "pending-external-validation"
    assert validation_plan["review_mode"] == "spot-check"
    assert validation_plan["evidence"]["manifest"]["formula_count"] == 1
    assert {item["id"] for item in validation_plan["planned_checks"]} == {
        "word-open",
        "word-export-pdf",
        "render-parity",
        "formula-count",
    }
    assert all(item["status"] != "completed" for item in validation_plan["planned_checks"])


def test_execute_omml_step_marks_manual_review_validation_gate(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.docx"
    make_docx(input_path)

    reports = execute_omml_step(
        make_omml_step(requires_manual_review=True),
        make_context(tmp_path, input_path),
    )

    render_report = reports[2]
    validation_plan = json.loads(Path(render_report.output_paths[0]).read_text(encoding="utf-8"))
    validation_target_path = Path(reports[3].output_paths[1])
    validation_evidence = json.loads(Path(reports[3].output_paths[2]).read_text(encoding="utf-8"))

    assert render_report.status == "review-gated"
    assert render_report.blocking is True
    assert validation_target_path.exists()
    assert validation_plan["status"] == "manual-review-required"
    assert validation_plan["review_mode"] == "required"
    assert validation_evidence["artifact_type"] == "omml-validation-evidence"
    assert validation_evidence["status"] == "evidence-collected"
    assert validation_evidence["gate_status"] == "manual-review-required"
    assert validation_evidence["artifacts"]["validation_plan"]["status"] == "manual-review-required"
    assert validation_evidence["artifacts"]["validation_plan"]["review_mode"] == "required"
