import json
import subprocess
from pathlib import Path

from document_equation_migration.execution_plan.model import ExecutionAction, ExecutionStep
from document_equation_migration.executor.mathtype import (
    build_mathtype_dry_run_reports,
    execute_mathtype_step,
)
from document_equation_migration.executor.model import DryRunContext, ExecutionContext


def make_mathtype_step(*, metadata: dict[str, object] | None = None) -> ExecutionStep:
    return ExecutionStep(
        source_family="mathtype-ole",
        formula_count=1,
        route_kind="primary-source-first",
        confidence_policy="high",
        requires_manual_review=False,
        provider="mathtype",
        next_action="run-mathtype-source-first-pipeline",
        metadata=dict(metadata or {}),
        actions=(
            ExecutionAction(
                action_id="extract-equation-native",
                description="Extract Equation Native payload from MathType OLE containers.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="mtef-to-mathml",
                description="Decode MTEF from Equation Native and convert it to MathML.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="normalize-mathml",
                description="Normalize MathML structure for deterministic downstream conversion.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="mathml-to-omml",
                description="Convert normalized MathML to OMML for Word-native equations.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="replace-ole-with-omml",
                description="Replace original MathType OLE instances with OMML formulas in DOCX.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="validate-word-output",
                description="Open in Word-compatible flow, verify rendering, and confirm export readiness.",
                blocking=False,
            ),
        ),
        notes=("Provider-local test step.",),
    )


def make_plan(path: Path) -> None:
    path.write_text(
        json.dumps({"input_path": "fixtures/sample.docx"}, indent=2),
        encoding="utf-8",
    )


def test_mathtype_dry_run_preserves_document_pipeline_preview(tmp_path: Path) -> None:
    plan_path = tmp_path / "execution-plan.json"
    make_plan(plan_path)
    context = DryRunContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(plan_path),
        output_dir_hint="out",
    )

    reports = build_mathtype_dry_run_reports(make_mathtype_step(), context)

    assert [report.action_id for report in reports] == [
        "extract-equation-native",
        "mtef-to-mathml",
        "normalize-mathml",
        "mathml-to-omml",
        "replace-ole-with-omml",
        "validate-word-output",
    ]
    assert all(report.supported for report in reports)
    assert all(report.status == "ready" for report in reports)
    assert reports[0].runner == "powershell"
    assert any(item.endswith("run_docx_open_source_pipeline.ps1") for item in reports[0].argv)
    assert "-PreserveMathTypeLayout" not in reports[0].argv
    assert "--preserve-mathtype-layout" not in reports[4].argv
    command_line = reports[0].to_dict()["command_line"]
    assert str(tmp_path / "fixtures" / "sample.docx") in command_line
    assert str(tmp_path / "out" / "sample.omml.docx") in command_line
    assert "document-level pipeline wrapper" in " ".join(reports[0].notes)


def test_mathtype_dry_run_includes_guarded_layout_args_when_enabled(tmp_path: Path) -> None:
    plan_path = tmp_path / "execution-plan.json"
    make_plan(plan_path)
    context = DryRunContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(plan_path),
        output_dir_hint="out",
    )

    reports = build_mathtype_dry_run_reports(
        make_mathtype_step(
            metadata={
                "experimental_options": {
                    "preserve_mathtype_layout": True,
                    "mathtype_layout_factor": 1.02,
                }
            }
        ),
        context,
    )

    assert "-PreserveMathTypeLayout" in reports[0].argv
    assert "-MathTypeLayoutFactor" in reports[0].argv
    assert "1.02" in reports[0].argv
    assert "--preserve-mathtype-layout" in reports[4].argv
    assert "--mathtype-layout-factor" in reports[4].argv
    assert "1.02" in reports[4].argv


def test_mathtype_dry_run_includes_resume_chunk_args_when_enabled(tmp_path: Path) -> None:
    plan_path = tmp_path / "execution-plan.json"
    make_plan(plan_path)
    context = DryRunContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(plan_path),
        output_dir_hint="out",
    )

    reports = build_mathtype_dry_run_reports(
        make_mathtype_step(
            metadata={
                "experimental_options": {
                    "resume_mathtype_pipeline": True,
                    "mathtype_start_index": 216,
                    "mathtype_end_index": 238,
                }
            }
        ),
        context,
    )

    assert "-Resume" in reports[0].argv
    assert "-StartIndex" in reports[0].argv
    assert "216" in reports[0].argv
    assert "-EndIndex" in reports[0].argv
    assert "238" in reports[0].argv
    assert any("resume/chunk" in note for note in reports[0].notes)


def test_mathtype_execute_blocks_external_tools_by_default(tmp_path: Path) -> None:
    context = ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path="",
        input_path=str(tmp_path / "sample.docx"),
        output_dir=str(tmp_path / "execution"),
        allow_external_tools=False,
    )

    reports = execute_mathtype_step(make_mathtype_step(), context)

    blocker_path = tmp_path / "execution" / "mathtype" / "blocker-record.json"
    assert len(reports) == 6
    assert {report.status for report in reports} == {"blocked-external-tool"}
    assert {report.supported for report in reports} == {True}
    assert all(report.stdout_path == "" for report in reports)
    assert all(report.stderr_path == "" for report in reports)
    assert all(report.output_paths == (str(blocker_path),) for report in reports)
    assert blocker_path.exists()

    blocker_record = json.loads(blocker_path.read_text(encoding="utf-8"))
    assert blocker_record["artifact_type"] == "mathtype-blocker-record"
    assert blocker_record["provider"] == "mathtype"
    assert blocker_record["status"] == "blocked-external-tool"
    assert blocker_record["actions"][0]["action_id"] == "extract-equation-native"
    assert blocker_record["required_evidence"]
    assert "allow-external-tools" in blocker_record["next_ready_condition"]

    first_notes = " ".join(reports[0].notes)
    assert "gated by --allow-external-tools" in first_notes
    assert "Run first with --dry-run" in first_notes


def test_mathtype_execute_allowed_uses_single_guarded_pipeline(monkeypatch, tmp_path: Path) -> None:
    calls: list[tuple[tuple[str, ...], str | None]] = []

    def fake_run(argv, *, cwd, text, stdout, stderr, check):
        calls.append((tuple(argv), cwd))
        return subprocess.CompletedProcess(
            args=argv,
            returncode=0,
            stdout="pipeline stdout\n",
            stderr="pipeline stderr\n",
        )

    monkeypatch.setattr("document_equation_migration.executor.mathtype.subprocess.run", fake_run)
    context = ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path="",
        input_path=str(tmp_path / "sample.docx"),
        output_dir=str(tmp_path / "execution"),
        allow_external_tools=True,
    )

    reports = execute_mathtype_step(make_mathtype_step(), context)

    output_root = tmp_path / "execution" / "mathtype"
    evidence_path = output_root / "validation-evidence.json"
    blocker_path = output_root / "blocker-record.json"
    assert len(calls) == 1
    assert calls[0][0][0] == "powershell"
    assert any(item.endswith("run_docx_open_source_pipeline.ps1") for item in calls[0][0])
    assert "-PreserveMathTypeLayout" not in calls[0][0]
    assert "-MathTypeLayoutFactor" not in calls[0][0]
    assert reports[0].status == "completed"
    assert reports[0].stdout_path.endswith("extract-equation-native.stdout.txt")
    assert reports[0].stderr_path.endswith("extract-equation-native.stderr.txt")
    assert Path(reports[0].stdout_path).read_text(encoding="utf-8") == "pipeline stdout\n"
    assert Path(reports[0].stderr_path).read_text(encoding="utf-8") == "pipeline stderr\n"
    assert reports[0].output_paths == (str(evidence_path),)
    assert evidence_path.exists()

    covered = {report.action_id: report for report in reports[1:5]}
    assert {report.status for report in covered.values()} == {"covered-by-document-pipeline"}
    assert all(report.stdout_path == reports[0].stdout_path for report in covered.values())
    assert all(report.stderr_path == reports[0].stderr_path for report in covered.values())
    assert all(report.output_paths == (str(evidence_path),) for report in covered.values())

    validation = reports[-1]
    assert validation.action_id == "validate-word-output"
    assert validation.status == "validation-gated"
    assert validation.runner == "manual-validation"
    assert validation.stdout_path == reports[0].stdout_path
    assert validation.stderr_path == reports[0].stderr_path
    assert validation.output_paths == (str(blocker_path),)
    assert blocker_path.exists()
    assert "validation gate" in " ".join(validation.notes)

    evidence_record = json.loads(evidence_path.read_text(encoding="utf-8"))
    assert evidence_record["artifact_type"] == "mathtype-validation-evidence"
    assert evidence_record["pipeline"]["action_id"] == "extract-equation-native"
    assert evidence_record["validation_gate"]["artifact_path"] == str(blocker_path)
    assert evidence_record["covered_actions"][0]["action_id"] == "mtef-to-mathml"

    blocker_record = json.loads(blocker_path.read_text(encoding="utf-8"))
    assert blocker_record["artifact_type"] == "mathtype-blocker-record"
    assert blocker_record["status"] == "validation-gated"
    assert blocker_record["pipeline_artifact_path"] == str(evidence_path)


def test_mathtype_execute_allowed_passes_guarded_layout_args_when_enabled(monkeypatch, tmp_path: Path) -> None:
    calls: list[tuple[tuple[str, ...], str | None]] = []

    def fake_run(argv, *, cwd, text, stdout, stderr, check):
        calls.append((tuple(argv), cwd))
        return subprocess.CompletedProcess(
            args=argv,
            returncode=0,
            stdout="pipeline stdout\n",
            stderr="pipeline stderr\n",
        )

    monkeypatch.setattr("document_equation_migration.executor.mathtype.subprocess.run", fake_run)
    context = ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path="",
        input_path=str(tmp_path / "sample.docx"),
        output_dir=str(tmp_path / "execution"),
        allow_external_tools=True,
    )

    reports = execute_mathtype_step(
        make_mathtype_step(
            metadata={
                "experimental_options": {
                    "preserve_mathtype_layout": True,
                    "mathtype_layout_factor": 1.02,
                }
            }
        ),
        context,
    )

    assert len(calls) == 1
    assert calls[0][0][0] == "powershell"
    assert "-PreserveMathTypeLayout" in calls[0][0]
    assert "-MathTypeLayoutFactor" in calls[0][0]
    assert "1.02" in calls[0][0]
    assert "-PreserveMathTypeLayout" in reports[0].argv
    assert "-MathTypeLayoutFactor" in reports[0].argv
    assert "1.02" in reports[0].argv


def test_mathtype_execute_allowed_passes_resume_chunk_args_when_enabled(monkeypatch, tmp_path: Path) -> None:
    calls: list[tuple[tuple[str, ...], str | None]] = []

    def fake_run(argv, *, cwd, text, stdout, stderr, check):
        calls.append((tuple(argv), cwd))
        return subprocess.CompletedProcess(
            args=argv,
            returncode=0,
            stdout="pipeline stdout\n",
            stderr="pipeline stderr\n",
        )

    monkeypatch.setattr("document_equation_migration.executor.mathtype.subprocess.run", fake_run)
    context = ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path="",
        input_path=str(tmp_path / "sample.docx"),
        output_dir=str(tmp_path / "execution"),
        allow_external_tools=True,
    )

    reports = execute_mathtype_step(
        make_mathtype_step(
            metadata={
                "experimental_options": {
                    "resume_mathtype_pipeline": True,
                    "mathtype_start_index": 216,
                    "mathtype_end_index": 238,
                }
            }
        ),
        context,
    )

    assert len(calls) == 1
    assert calls[0][0][0] == "powershell"
    assert "-Resume" in calls[0][0]
    assert "-StartIndex" in calls[0][0]
    assert "216" in calls[0][0]
    assert "-EndIndex" in calls[0][0]
    assert "238" in calls[0][0]
    assert "-Resume" in reports[0].argv
