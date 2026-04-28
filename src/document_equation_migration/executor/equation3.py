from __future__ import annotations

import json
from pathlib import Path

from ..canonical_target import canonical_mathml_contract_for_source_family
from ..execution_plan.model import ExecutionStep
from .model import ActionExecutionReport, DryRunActionReport, DryRunContext, ExecutionContext


SOURCE_FAMILY = "equation-editor-3-ole"
RUNNER = "internal-equation3-probe"
BLOCKER_RECORD_FILENAME = "blocker-record.json"


def equation3_fixture_admissibility_requirements() -> dict[str, object]:
    return {
        "target_stage": "fixture-backed-canonical-mathml-conversion",
        "minimum_fixture_set": (
            "At least one redistributable or explicitly authorized DOCX containing a real Equation.3 "
            "native OLE payload, not only a preview image or marker payload."
        ),
        "required_candidate_properties": [
            {
                "id": "equation3-identity",
                "description": "The detector record identifies the source as Equation Editor 3.0.",
                "evidence_fields": (
                    "source_family=equation-editor-3-ole",
                    "provenance.prog_id_raw=Equation.3 or field_code_raw contains EMBED Equation",
                    "source_specific.equation_editor_3.class_id_raw",
                ),
            },
            {
                "id": "native-payload",
                "description": "The fixture exposes a readable native OLE payload, not just a rendered preview.",
                "evidence_fields": (
                    "source_role=native-source",
                    "provenance.raw_payload_status=present",
                    "provenance.raw_payload_sha256",
                    "provenance.payload_stream_name",
                ),
            },
            {
                "id": "mtef-v3-header",
                "description": "The native payload has Equation Editor 3.0 / MTEF v3 header evidence.",
                "evidence_fields": (
                    "source_specific.equation_editor_3.native_header_size_bytes",
                    "source_specific.equation_editor_3.mtef_version=3",
                    "source_specific.equation_editor_3.selected_route=mtef-v3-mainline",
                ),
            },
            {
                "id": "canonical-output",
                "description": "A conversion attempt emits valid canonical MathML artifacts with formula-count parity.",
                "evidence_fields": (
                    "canonical-mathml/*.xml",
                    "canonicalization-summary.json",
                    "canonical_mathml_count equals fixture formula count",
                    "unsupported_fragment_count recorded",
                ),
            },
            {
                "id": "provenance-map",
                "description": "Each canonical MathML artifact remains traceable to the source DOCX object.",
                "evidence_fields": (
                    "formula_id",
                    "doc_part_path",
                    "relationship_id",
                    "embedding_target",
                    "raw_payload_sha256",
                ),
            },
        ],
        "disqualifying_conditions": [
            "preview-only fixture",
            "marker-text payload instead of binary native payload",
            "missing or corrupt OLE payload",
            "conflicting MathType or AxMath vendor signal",
            "no valid canonical MathML output",
            "unclear redistribution or use permission for public fixture promotion",
        ],
        "promotion_gate": [
            "Detector evidence proves native Equation.3 source identity.",
            "Canonical MathML output validates as XML and preserves formula count.",
            "Conversion report keeps provenance for every source formula.",
            "DOCX/Word validation is attempted only after canonical conversion evidence exists.",
        ],
    }


def _dry_run_output_root(context: DryRunContext) -> Path:
    return Path(context.output_dir_hint) / "equation-editor-3-ole"


def _execution_output_root(context: ExecutionContext) -> Path:
    return Path(context.output_dir) / "equation-editor-3-ole"


def _write_json(path: Path, payload: dict[str, object]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _action_summaries(step: ExecutionStep) -> tuple[dict[str, object], ...]:
    summaries: list[dict[str, object]] = []
    for action in step.actions:
        if action.action_id == "probe-header-and-classid":
            summaries.append(
                {
                    "action_id": action.action_id,
                    "description": action.description,
                    "status": "review-gated",
                    "supported": True,
                    "blocking": action.blocking,
                }
            )
            continue

        if action.action_id == "attempt-mtef-conversion":
            summaries.append(
                {
                    "action_id": action.action_id,
                    "description": action.description,
                    "status": "manual-gate",
                    "supported": False,
                    "blocking": action.blocking,
                }
            )
            continue

        if action.action_id == "fallback-manual-triage":
            summaries.append(
                {
                    "action_id": action.action_id,
                    "description": action.description,
                    "status": "manual-gate",
                    "supported": False,
                    "blocking": action.blocking,
                }
            )
            continue

        if action.action_id == "word-roundtrip-validation":
            summaries.append(
                {
                    "action_id": action.action_id,
                    "description": action.description,
                    "status": "review-gated",
                    "supported": False,
                    "blocking": action.blocking,
                }
            )
            continue

        summaries.append(
            {
                "action_id": action.action_id,
                "description": action.description,
                "status": "manual-gate",
                "supported": False,
                "blocking": action.blocking,
            }
        )

    return tuple(summaries)


def _write_blocker_record(step: ExecutionStep, context: ExecutionContext) -> Path:
    output_root = _execution_output_root(context)
    blocker_record_path = output_root / BLOCKER_RECORD_FILENAME
    return _write_json(
        blocker_record_path,
        {
            "artifact_type": "equation3-blocker-record",
            "provider": "equation3",
            "source_family": SOURCE_FAMILY,
            "canonical_target": canonical_mathml_contract_for_source_family(SOURCE_FAMILY).to_dict(),
            "status": "blocked",
            "blocking": True,
            "supported": False,
            "conversion_claim": False,
            "fixture_status": "insufficient",
            "fixture_gap": "Equation Editor 3.0 fixture coverage is not strong enough for a deliverable conversion claim.",
            "next_ready_condition": "Add stronger Equation Editor 3.0 fixtures that exercise probe evidence and conversion output, then rerun execute.",
            "fixture_admissibility": equation3_fixture_admissibility_requirements(),
            "execution_plan_path": context.execution_plan_path,
            "input_path": context.input_path,
            "output_root": str(output_root),
            "formula_count": step.formula_count,
            "manual_review_required": step.requires_manual_review,
            "probe": {
                "runner": RUNNER,
                "action_id": "probe-header-and-classid",
                "signals": ["prog-id", "class-id", "eqnolefilehdr", "mtef-v3-header"],
            },
            "actions": list(_action_summaries(step)),
        },
    )


def _probe_dry_run_report(
    step: ExecutionStep,
    action_id: str,
    description: str,
    context: DryRunContext,
) -> DryRunActionReport:
    output_root = _dry_run_output_root(context)
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=False,
        supported=True,
        status="ready",
        runner=RUNNER,
        argv=(
            "probe-equation3-evidence",
            "--execution-plan",
            context.execution_plan_path,
            "--output-dir",
            str(output_root),
            "--signals",
            "prog-id,class-id,eqnolefilehdr,mtef-v3-header",
        ),
        cwd=context.workspace_root,
        notes=(
            "Supported probe skeleton only: confirm Equation.3, ClassID, and EQNOLEFILEHDR/MTEF v3 evidence.",
            "This action does not claim conversion readiness or deliverable Word output.",
            "Next stage requires a fixture-backed canonical MathML conversion attempt that satisfies the blocker checklist.",
            f"Formula count from execution plan: {step.formula_count}.",
        ),
    )


def _manual_gate_dry_run_report(
    action_id: str,
    description: str,
    blocking: bool,
    context: DryRunContext,
    *,
    note: str,
) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=blocking,
        supported=False,
        status="manual-gate",
        runner="manual",
        cwd=context.workspace_root,
        notes=(
            note,
            "Real Equation Editor 3.0 fixture coverage is not strong enough to advertise an executable conversion binding.",
            "Use the blocker-record fixture admissibility checklist before promoting this source line.",
        ),
    )


def _review_gated_dry_run_report(
    action_id: str,
    description: str,
    blocking: bool,
    context: DryRunContext,
) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=blocking,
        supported=False,
        status="review-gated",
        runner="manual-validation",
        cwd=context.workspace_root,
        notes=(
            "Word roundtrip validation requires fixture-backed conversion output and human review before delivery claims.",
            "No automated DOCX/PDF validation binding is available for Equation Editor 3.0 yet.",
        ),
    )


def build_equation3_dry_run_reports(
    step: ExecutionStep,
    context: DryRunContext,
) -> tuple[DryRunActionReport, ...]:
    if step.source_family != SOURCE_FAMILY:
        raise ValueError(f"Equation Editor 3.0 dry-run binding cannot handle source_family={step.source_family!r}.")

    if not step.actions:
        return (
            _manual_gate_dry_run_report(
                "manual-triage",
                "No Equation Editor 3.0 execution actions are registered for this step.",
                True,
                context,
                note="Manual triage is required before any Equation Editor 3.0 conversion attempt.",
            ),
        )

    reports: list[DryRunActionReport] = []
    for action in step.actions:
        action_id = action.action_id
        if action_id == "probe-header-and-classid":
            reports.append(_probe_dry_run_report(step, action_id, action.description, context))
            continue

        if action_id == "attempt-mtef-conversion":
            reports.append(
                _manual_gate_dry_run_report(
                    action_id,
                    action.description,
                    True,
                    context,
                    note="MTEF v3 conversion remains a manual gate until native payload fixtures are strengthened.",
                )
            )
            continue

        if action_id == "fallback-manual-triage":
            reports.append(
                _manual_gate_dry_run_report(
                    action_id,
                    action.description,
                    True,
                    context,
                    note="Ambiguous or incomplete Equation Editor 3.0 evidence must stay in manual triage.",
                )
            )
            continue

        if action_id == "word-roundtrip-validation":
            reports.append(_review_gated_dry_run_report(action_id, action.description, True, context))
            continue

        reports.append(
            _manual_gate_dry_run_report(
                action_id,
                action.description,
                True,
                context,
                note="Unrecognized Equation Editor 3.0 action; manual triage is required before binding execution.",
            )
        )

    return tuple(reports)


def _probe_execution_report(
    step: ExecutionStep,
    action_id: str,
    description: str,
    context: ExecutionContext,
    *,
    output_path: Path,
) -> ActionExecutionReport:
    output_root = _execution_output_root(context)
    return ActionExecutionReport(
        action_id=action_id,
        description=description,
        blocking=False,
        supported=True,
        status="review-gated",
        runner=RUNNER,
        argv=(
            "probe-equation3-evidence",
            "--input",
            context.input_path,
            "--execution-plan",
            context.execution_plan_path,
            "--output-dir",
            str(output_root),
        ),
        cwd=context.workspace_root,
        output_paths=(str(output_path),),
        notes=(
            "Execution is intentionally limited to an evidence/probe skeleton for Equation Editor 3.0.",
            "No payload conversion, OMML replacement, or deliverable Word artifact is produced by this provider.",
            "The blocker record lists the exact fixture admissibility requirements for the next canonical MathML stage.",
            f"Blocker record written to {output_path}.",
            f"Formula count from execution plan: {step.formula_count}.",
        ),
    )


def _manual_gate_execution_report(
    action_id: str,
    description: str,
    blocking: bool,
    context: ExecutionContext,
    *,
    note: str,
    output_path: Path,
) -> ActionExecutionReport:
    return ActionExecutionReport(
        action_id=action_id,
        description=description,
        blocking=blocking,
        supported=False,
        status="manual-gate",
        runner="manual",
        cwd=context.workspace_root,
        output_paths=(str(output_path),),
        notes=(
            note,
            "Execution report is a gate record only; it is not a successful conversion result.",
            f"Blocker record written to {output_path}.",
        ),
    )


def _review_gated_execution_report(
    action_id: str,
    description: str,
    blocking: bool,
    context: ExecutionContext,
    *,
    output_path: Path,
) -> ActionExecutionReport:
    return ActionExecutionReport(
        action_id=action_id,
        description=description,
        blocking=blocking,
        supported=False,
        status="review-gated",
        runner="manual-validation",
        cwd=context.workspace_root,
        output_paths=(str(output_path),),
        notes=(
            "Word open/PDF export/visual parity checks remain review-gated for Equation Editor 3.0.",
            "This status is not a deliverable conversion proof.",
            f"Blocker record written to {output_path}.",
        ),
    )


def execute_equation3_step(
    step: ExecutionStep,
    context: ExecutionContext,
) -> tuple[ActionExecutionReport, ...]:
    if step.source_family != SOURCE_FAMILY:
        raise ValueError(f"Equation Editor 3.0 executor cannot handle source_family={step.source_family!r}.")

    blocker_record_path = _write_blocker_record(step, context)

    if not step.actions:
        return (
            _manual_gate_execution_report(
                "manual-triage",
                "No Equation Editor 3.0 execution actions are registered for this step.",
                True,
                context,
                note="Manual triage is required before Equation Editor 3.0 execution.",
                output_path=blocker_record_path,
            ),
        )

    reports: list[ActionExecutionReport] = []
    for action in step.actions:
        action_id = action.action_id
        if action_id == "probe-header-and-classid":
            reports.append(
                _probe_execution_report(
                    step,
                    action_id,
                    action.description,
                    context,
                    output_path=blocker_record_path,
                )
            )
            continue

        if action_id == "attempt-mtef-conversion":
            reports.append(
                _manual_gate_execution_report(
                    action_id,
                    action.description,
                    True,
                    context,
                    note="MTEF v3 conversion is intentionally not executable until real Equation.3 fixtures are strengthened.",
                    output_path=blocker_record_path,
                )
            )
            continue

        if action_id == "fallback-manual-triage":
            reports.append(
                _manual_gate_execution_report(
                    action_id,
                    action.description,
                    True,
                    context,
                    note="Manual triage remains required for ambiguous Equation Editor 3.0 evidence.",
                    output_path=blocker_record_path,
                )
            )
            continue

        if action_id == "word-roundtrip-validation":
            reports.append(
                _review_gated_execution_report(
                    action_id,
                    action.description,
                    True,
                    context,
                    output_path=blocker_record_path,
                )
            )
            continue

        reports.append(
            _manual_gate_execution_report(
                action_id,
                action.description,
                True,
                context,
                note="Unrecognized Equation Editor 3.0 action; manual triage is required.",
                output_path=blocker_record_path,
            )
        )

    return tuple(reports)
