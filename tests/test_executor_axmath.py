import json
from pathlib import Path

from document_equation_migration.execution_plan.axmath import build_axmath_execution_step
from document_equation_migration.executor.axmath import (
    axmath_export_admissibility_requirements,
    build_axmath_dry_run_reports,
    execute_axmath_step,
)
from document_equation_migration.executor.model import DryRunContext, ExecutionContext


def synthetic_structural_metadata() -> dict[str, object]:
    return {
        "source_family": "axmath-ole",
        "metadata_version": "synthetic-no-payload-v1",
        "rows": [
            {
                "row_id": 1,
                "object_id": "synthetic-object-1",
                "stream_locator_id": "locator-1",
                "stream_length": 573,
                "unique_stream_id": "U1",
                "tail_token_class_id": "TT01",
                "tail_window_class": "small-tail",
                "tail_start_candidate_offset": 339,
                "tail_length": 234,
                "trailer_length": 8,
                "parse_status": "candidate_count_relation_fail",
                "unsupported_boundary": "none",
            }
        ],
        "duplicate_invariance_fail_count": 0,
    }


def make_axmath_step(
    *,
    requires_manual_review: bool = True,
    structural_metadata: dict[str, object] | None = None,
    controlled_acceptance: dict[str, object] | None = None,
):
    route_entry = {
        "source_family": "axmath-ole",
        "formula_count": 2,
        "route_kind": "export-assisted",
        "confidence_policy": "medium",
        "requires_manual_review": requires_manual_review,
        "next_action": "run-axmath-export-assisted-pipeline",
    }
    if structural_metadata is not None:
        route_entry["axmath_structural_metadata"] = structural_metadata
    if controlled_acceptance is not None:
        route_entry["axmath_controlled_acceptance"] = controlled_acceptance
    return build_axmath_execution_step(route_entry)


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
    assert any("local BYO Contents inspection" in note for note in reports[0].notes)
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
    assert any("does not claim semantic native parsing" in note for note in reports[1].notes)


def test_structural_metadata_action_reports_metadata_only_and_keeps_export_blocker(
    tmp_path: Path,
) -> None:
    step = make_axmath_step(structural_metadata=synthetic_structural_metadata())
    action_ids = [action.action_id for action in step.actions]

    assert action_ids == [
        "classify-axmath-object",
        "classify-axmath-structural-metadata",
        "export-assisted-conversion",
        "import-converted-math",
        "manual-spot-check",
    ]
    assert step.route_kind == "export-assisted"
    assert step.next_action == "run-axmath-export-assisted-pipeline"

    dry_run_reports = build_axmath_dry_run_reports(
        step,
        DryRunContext(
            workspace_root=str(tmp_path),
            execution_plan_path="",
            output_dir_hint=str(tmp_path / "out"),
        ),
    )
    dry_run_by_id = {report.action_id: report for report in dry_run_reports}
    metadata_dry_run = dry_run_by_id["classify-axmath-structural-metadata"]
    assert metadata_dry_run.status == "metadata-only"
    assert metadata_dry_run.runner == "internal-axmath-structural-metadata"
    assert any("conversion_claim=False" in note for note in metadata_dry_run.notes)
    assert any("native_parser_claim=False" in note for note in metadata_dry_run.notes)

    reports = execute_axmath_step(
        step,
        ExecutionContext(
            workspace_root=str(tmp_path),
            execution_plan_path="",
            input_path=str(tmp_path / "input.docx"),
            output_dir=str(tmp_path / "out"),
        ),
    )
    by_id = {report.action_id: report for report in reports}
    assert by_id["classify-axmath-structural-metadata"].status == "completed"
    assert by_id["classify-axmath-structural-metadata"].runner == "internal-axmath-structural-metadata"
    assert by_id["classify-axmath-structural-metadata"].output_paths == ()
    assert by_id["export-assisted-conversion"].status == "blocked-external-tool"
    assert by_id["import-converted-math"].status == "skipped-until-export-artifacts"

    gate_path = Path(by_id["export-assisted-conversion"].output_paths[0])
    gate = json.loads(gate_path.read_text(encoding="utf-8"))
    structural_metadata = gate["axmath_structural_metadata"]
    assert structural_metadata["status"] == "metadata-only-blockers-preserved"
    assert structural_metadata["record_count"] == 1
    assert structural_metadata["records"][0]["structural_status"] == "candidate_count_relation_absent"
    assert structural_metadata["records"][0]["body_decode_status"] == (
        "blocked_count_relation_hypothesis_only"
    )
    assert structural_metadata["blocker_ids"] == ["candidate_count_relation_absent"]
    assert structural_metadata["claim_boundaries"]["conversion_claim"] is False
    assert structural_metadata["claim_boundaries"]["native_parser_claim"] is False
    assert structural_metadata["claim_boundaries"]["export_success_claim"] is False
    assert structural_metadata["claim_boundaries"]["public_fixture_eligibility"] is False
    assert gate["conversion_claim"] is False
    assert gate["native_parser_claim"] is False
    assert gate["export_success_claim"] is False
    assert gate["public_fixture_eligibility"] is False
    assert gate["status"] == "blocked-external-tool"


def test_controlled_acceptance_action_reports_contract_only_and_keeps_export_blocker(
    tmp_path: Path,
) -> None:
    step = make_axmath_step(controlled_acceptance={})
    action_ids = [action.action_id for action in step.actions]

    assert action_ids == [
        "classify-axmath-object",
        "record-axmath-controlled-acceptance-contract",
        "export-assisted-conversion",
        "import-converted-math",
        "manual-spot-check",
    ]

    dry_run_reports = build_axmath_dry_run_reports(
        step,
        DryRunContext(
            workspace_root=str(tmp_path),
            execution_plan_path="",
            output_dir_hint=str(tmp_path / "out"),
        ),
    )
    dry_run_by_id = {report.action_id: report for report in dry_run_reports}
    acceptance_dry_run = dry_run_by_id["record-axmath-controlled-acceptance-contract"]
    assert acceptance_dry_run.status == "contract-only"
    assert acceptance_dry_run.runner == "internal-axmath-controlled-acceptance"
    assert any("axmath_controlled_sample_count=355" in note for note in acceptance_dry_run.notes)
    assert any("public_native_parser_binding=False" in note for note in acceptance_dry_run.notes)

    reports = execute_axmath_step(
        step,
        ExecutionContext(
            workspace_root=str(tmp_path),
            execution_plan_path="",
            input_path=str(tmp_path / "input.docx"),
            output_dir=str(tmp_path / "out"),
        ),
    )
    by_id = {report.action_id: report for report in reports}
    assert by_id["record-axmath-controlled-acceptance-contract"].status == "completed"
    assert by_id["export-assisted-conversion"].status == "blocked-external-tool"
    assert by_id["import-converted-math"].status == "skipped-until-export-artifacts"

    gate_path = Path(by_id["export-assisted-conversion"].output_paths[0])
    gate = json.loads(gate_path.read_text(encoding="utf-8"))
    controlled_acceptance = gate["axmath_controlled_acceptance"]
    assert controlled_acceptance["decision_enum"] == (
        "axmath_angle_direct_native_acceptance_refresh_passed"
    )
    assert controlled_acceptance["controlled_sample_count"] == 355
    assert controlled_acceptance["accepted_count"] == 355
    assert controlled_acceptance["remaining_fail_closed_count"] == 0
    assert controlled_acceptance["claim_boundaries"]["native_parser_claim"] is False
    assert controlled_acceptance["claim_boundaries"]["public_native_parser_binding"] is False
    assert "public native parser binding" in gate["axmath_controlled_acceptance_boundary"]
    assert gate["status"] == "blocked-external-tool"


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
