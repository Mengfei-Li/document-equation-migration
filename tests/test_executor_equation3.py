import json
from pathlib import Path

import pytest

from document_equation_migration.execution_plan.model import ExecutionAction, ExecutionStep
from document_equation_migration.executor.equation3 import (
    build_equation3_dry_run_reports,
    execute_equation3_step,
)
from document_equation_migration.executor.model import DryRunContext, ExecutionContext


def _equation3_step(*, source_family: str = "equation-editor-3-ole") -> ExecutionStep:
    return ExecutionStep(
        source_family=source_family,
        formula_count=2,
        route_kind="primary-candidate",
        confidence_policy="medium",
        requires_manual_review=True,
        provider="equation3",
        next_action="run-equation3-probe-and-conversion",
        actions=(
            ExecutionAction(
                action_id="probe-header-and-classid",
                description="Probe OLE header and ClassID to confirm Equation Editor 3.0 payload.",
            ),
            ExecutionAction(
                action_id="attempt-mtef-conversion",
                description="Attempt MTEF-oriented conversion as primary Equation Editor 3.0 path.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="fallback-manual-triage",
                description="Fallback to manual triage when probe or conversion results are ambiguous.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="word-roundtrip-validation",
                description="Validate converted output through Word roundtrip before delivery.",
                blocking=True,
            ),
        ),
        notes=("Equation3 sample",),
    )


def _dry_run_context(tmp_path: Path) -> DryRunContext:
    return DryRunContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(tmp_path / "execution-plan.json"),
        output_dir_hint=str(tmp_path / "out"),
    )


def _execution_context(tmp_path: Path) -> ExecutionContext:
    return ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(tmp_path / "execution-plan.json"),
        input_path=str(tmp_path / "input.docx"),
        output_dir=str(tmp_path / "out"),
    )


def test_equation3_dry_run_is_provider_binding_not_generic_fallback(tmp_path: Path) -> None:
    reports = build_equation3_dry_run_reports(_equation3_step(), _dry_run_context(tmp_path))

    assert [report.action_id for report in reports] == [
        "probe-header-and-classid",
        "attempt-mtef-conversion",
        "fallback-manual-triage",
        "word-roundtrip-validation",
    ]
    assert reports[0].supported is True
    assert reports[0].status == "ready"
    assert reports[0].runner == "internal-equation3-probe"
    assert reports[0].argv[0] == "probe-equation3-evidence"
    assert reports[1].supported is False
    assert reports[1].status == "manual-gate"
    assert reports[2].status == "manual-gate"
    assert reports[3].status == "review-gated"

    combined_notes = "\n".join("\n".join(report.notes) for report in reports)
    assert "No concrete dry-run binding is registered" not in combined_notes
    assert "generic fallback" not in combined_notes.lower()


def test_equation3_execute_writes_blocker_record_and_keeps_gate_status(tmp_path: Path) -> None:
    reports = execute_equation3_step(_equation3_step(), _execution_context(tmp_path))

    assert [report.status for report in reports] == [
        "review-gated",
        "manual-gate",
        "manual-gate",
        "review-gated",
    ]
    assert reports[0].supported is True
    assert reports[0].runner == "internal-equation3-probe"
    assert reports[1].supported is False
    assert reports[3].runner == "manual-validation"
    assert all(report.status != "completed" for report in reports)
    assert len({report.output_paths for report in reports}) == 1

    blocker_record_path = Path(reports[0].output_paths[0])
    assert blocker_record_path.name == "blocker-record.json"
    assert blocker_record_path.exists()

    blocker_record = json.loads(blocker_record_path.read_text(encoding="utf-8"))
    assert blocker_record["artifact_type"] == "equation3-blocker-record"
    assert blocker_record["provider"] == "equation3"
    assert blocker_record["source_family"] == "equation-editor-3-ole"
    assert blocker_record["status"] == "blocked"
    assert blocker_record["blocking"] is True
    assert blocker_record["conversion_claim"] is False
    assert blocker_record["fixture_status"] == "insufficient"
    assert "deliverable conversion claim" in blocker_record["fixture_gap"]
    assert "stronger Equation Editor 3.0 fixtures" in blocker_record["next_ready_condition"]
    assert blocker_record["probe"]["runner"] == "internal-equation3-probe"
    assert blocker_record["probe"]["signals"] == ["prog-id", "class-id", "eqnolefilehdr", "mtef-v3-header"]
    assert [action["action_id"] for action in blocker_record["actions"]] == [
        "probe-header-and-classid",
        "attempt-mtef-conversion",
        "fallback-manual-triage",
        "word-roundtrip-validation",
    ]
    assert blocker_record["actions"][0]["status"] == "review-gated"
    assert blocker_record["actions"][1]["status"] == "manual-gate"
    assert blocker_record["actions"][3]["supported"] is False

    combined_notes = "\n".join("\n".join(report.notes) for report in reports)
    assert "No payload conversion" in combined_notes
    assert "Blocker record written to" in combined_notes
    assert "deliverable conversion proof" in combined_notes


def test_equation3_provider_rejects_wrong_source_family(tmp_path: Path) -> None:
    step = _equation3_step(source_family="mathtype-ole")

    with pytest.raises(ValueError):
        build_equation3_dry_run_reports(step, _dry_run_context(tmp_path))

    with pytest.raises(ValueError):
        execute_equation3_step(step, _execution_context(tmp_path))
