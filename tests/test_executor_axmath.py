import json
from pathlib import Path

from document_equation_migration.execution_plan.axmath import build_axmath_execution_step
from document_equation_migration.executor.axmath import (
    axmath_export_admissibility_requirements,
    build_axmath_dry_run_reports,
    execute_axmath_step,
)
from document_equation_migration.executor.model import DryRunContext, ExecutionContext


def make_axmath_step(*, requires_manual_review: bool = True):
    return build_axmath_execution_step(
        {
            "source_family": "axmath-ole",
            "formula_count": 2,
            "route_kind": "export-assisted",
            "confidence_policy": "medium",
            "requires_manual_review": requires_manual_review,
            "next_action": "run-axmath-export-assisted-pipeline",
        }
    )


def test_build_axmath_dry_run_reports_describes_export_gate(tmp_path: Path) -> None:
    step = make_axmath_step()
    reports = build_axmath_dry_run_reports(
        step,
        DryRunContext(
            workspace_root=str(tmp_path),
            execution_plan_path="",
            output_dir_hint=str(tmp_path / "out"),
        ),
    )

    assert [report.action_id for report in reports] == [
        "classify-axmath-object",
        "export-assisted-conversion",
        "import-converted-math",
        "manual-spot-check",
    ]
    assert reports[0].status == "export-gate"
    assert reports[1].status == "export-gate"
    assert reports[1].runner == "external-axmath-export"
    assert reports[1].supported is True
    assert reports[2].status == "validation-gated"
    assert reports[3].status == "review-gated"
    assert any("No native AxMath static parser" in note for note in reports[0].notes)
    assert any("export-assisted" in note for note in reports[1].notes)
    assert any("export admissibility checklist" in note for note in reports[1].notes)


def test_execute_axmath_step_blocks_external_export_by_default(tmp_path: Path) -> None:
    step = make_axmath_step()
    reports = execute_axmath_step(
        step,
        ExecutionContext(
            workspace_root=str(tmp_path),
            execution_plan_path="",
            input_path=str(tmp_path / "input.docx"),
            output_dir=str(tmp_path / "out"),
        ),
    )

    statuses = {report.action_id: report.status for report in reports}
    assert statuses["classify-axmath-object"] == "export-gate"
    assert statuses["export-assisted-conversion"] == "blocked-external-tool"
    assert statuses["import-converted-math"] == "skipped-until-export-artifacts"
    assert statuses["manual-spot-check"] == "review-gated"
    assert len(reports[1].output_paths) == 1
    gate_path = Path(reports[1].output_paths[0])
    assert gate_path.name == "blocker-record.json"
    assert gate_path.exists()
    gate = json.loads(gate_path.read_text(encoding="utf-8"))
    assert gate["artifact_type"] == "axmath-export-assisted-blocker-record"
    assert gate["status"] == "blocked-external-tool"
    assert gate["external_export_dependency"]["required"] is True
    assert gate["external_export_dependency"]["allow_external_tools"] is False
    assert gate["export_admissibility"]["target_stage"] == "export-to-canonical-mathml"
    assert {item["id"] for item in gate["export_admissibility"]["accepted_export_channels"]} == {
        "direct-mathml",
        "latex-plus-validated-converter",
    }
    required_property_ids = {
        item["id"] for item in gate["export_admissibility"]["required_candidate_properties"]
    }
    assert required_property_ids == {
        "axmath-identity",
        "export-provenance",
        "canonical-output",
        "semantic-review",
    }
    assert "native static parser claim without verified parser binding" in gate["export_admissibility"][
        "disqualifying_conditions"
    ]
    assert any("Canonical MathML output validates" in item for item in gate["export_admissibility"]["promotion_gate"])
    assert "allow-external-tools" in gate["next_ready_condition"]
    assert "reviewed MathML or LaTeX" in gate["next_ready_condition"]
    assert any("external vendor/export workflow" in note for note in reports[1].notes)
    assert any("does not claim native static parsing" in note for note in reports[1].notes)


def test_execute_axmath_step_stays_validation_gated_when_external_tools_are_allowed(
    tmp_path: Path,
) -> None:
    step = make_axmath_step(requires_manual_review=False)
    reports = execute_axmath_step(
        step,
        ExecutionContext(
            workspace_root=str(tmp_path),
            execution_plan_path="",
            input_path=str(tmp_path / "input.docx"),
            output_dir=str(tmp_path / "out"),
            allow_external_tools=True,
        ),
    )

    statuses = {report.action_id: report.status for report in reports}
    assert statuses["export-assisted-conversion"] == "validation-gated"
    assert statuses["manual-spot-check"] == "validation-gated"
    assert len(reports[1].output_paths) == 1
    gate_path = Path(reports[1].output_paths[0])
    assert gate_path.exists()
    gate = json.loads(gate_path.read_text(encoding="utf-8"))
    assert gate["status"] == "validation-gated"
    assert gate["external_export_dependency"]["allow_external_tools"] is True
    assert gate["export_admissibility"]["target_stage"] == "export-to-canonical-mathml"
    assert "verified AxMath/vendor export workflow" in gate["next_ready_condition"]
    assert any("no verified AxMath CLI binding" in note for note in reports[1].notes)
    assert any("export admissibility requirements" in note for note in reports[1].notes)


def test_axmath_export_admissibility_keeps_native_parser_claim_disallowed() -> None:
    requirements = axmath_export_admissibility_requirements()

    assert requirements["target_stage"] == "export-to-canonical-mathml"
    assert "reviewed AxMath export batch" in requirements["minimum_export_set"]
    assert "native static parser claim without verified parser binding" in requirements[
        "disqualifying_conditions"
    ]
    assert "LaTeX export without a validated LaTeX-to-MathML step" in requirements[
        "disqualifying_conditions"
    ]
    assert any("Manual semantic review" in item for item in requirements["promotion_gate"])
