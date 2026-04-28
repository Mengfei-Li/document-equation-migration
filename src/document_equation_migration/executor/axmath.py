from __future__ import annotations

import json
from pathlib import Path

from ..canonical_target import canonical_mathml_contract_for_source_family
from ..execution_plan.model import ExecutionAction, ExecutionStep
from .model import ActionExecutionReport, DryRunActionReport, DryRunContext, ExecutionContext


AXMATH_SOURCE_FAMILY = "axmath-ole"
AXMATH_OUTPUT_DIR = "axmath-export-assisted"
_INPUT_DOCX_PLACEHOLDER = "<input-docx-from-plan>"


def axmath_export_admissibility_requirements() -> dict[str, object]:
    return {
        "target_stage": "export-to-canonical-mathml",
        "minimum_export_set": (
            "At least one reviewed AxMath export batch that maps detected axmath-ole formulas to "
            "canonical MathML artifacts, either directly from MathML export or through a separately "
            "validated LaTeX-to-MathML step."
        ),
        "accepted_export_channels": [
            {
                "id": "direct-mathml",
                "description": "AxMath or an approved vendor workflow exports MathML that can be normalized to canonical MathML.",
                "required_evidence": [
                    "raw exported MathML artifact",
                    "canonical MathML artifact",
                    "XML parse success",
                    "formula-count parity",
                ],
            },
            {
                "id": "latex-plus-validated-converter",
                "description": "AxMath exports LaTeX and a separately validated converter produces canonical MathML.",
                "required_evidence": [
                    "raw exported LaTeX artifact",
                    "documented LaTeX-to-MathML converter identity and version",
                    "canonical MathML artifact",
                    "semantic review against source AxMath object",
                ],
            },
        ],
        "required_candidate_properties": [
            {
                "id": "axmath-identity",
                "description": "The source formula is classified as AxMath OLE.",
                "evidence_fields": (
                    "source_family=axmath-ole",
                    "provenance.prog_id_raw=Equation.AxMath or field_code_raw contains EMBED Equation.AxMath",
                    "axmath.word_addin_artifacts",
                ),
            },
            {
                "id": "export-provenance",
                "description": "Each exported artifact is traceable to a detected AxMath source object.",
                "evidence_fields": (
                    "formula_id",
                    "doc_part_path",
                    "relationship_id",
                    "embedding_target",
                    "raw_payload_sha256",
                ),
            },
            {
                "id": "canonical-output",
                "description": "The reviewed export path emits canonical MathML artifacts with formula-count parity.",
                "evidence_fields": (
                    "canonical-mathml/*.xml",
                    "canonicalization-summary.json",
                    "canonical_mathml_count equals reviewed export formula count",
                    "unsupported_fragment_count recorded",
                ),
            },
            {
                "id": "semantic-review",
                "description": "A human or deterministic review confirms exported semantics against source AxMath objects.",
                "evidence_fields": (
                    "reviewed-export-report.json",
                    "semantic_review_status",
                    "review_notes",
                ),
            },
        ],
        "disqualifying_conditions": [
            "native static parser claim without verified parser binding",
            "unreviewed export artifacts",
            "export batch without source formula provenance",
            "LaTeX export without a validated LaTeX-to-MathML step",
            "formula-count mismatch between AxMath detections and accepted canonical artifacts",
            "custom symbols or vendor macros without documented semantic handling",
            "unclear permission for publishing source or exported fixture artifacts",
        ],
        "promotion_gate": [
            "Detector evidence proves AxMath source identity.",
            "Export artifacts are reviewed and linked to source formula provenance.",
            "Canonical MathML output validates as XML and preserves formula count.",
            "Any LaTeX route identifies and validates its LaTeX-to-MathML converter.",
            "Manual semantic review is recorded before delivery claims.",
        ],
    }


def _workspace_root(context: DryRunContext | ExecutionContext) -> Path:
    return Path(context.workspace_root)


def _dry_run_output_root(context: DryRunContext) -> Path:
    output_dir = Path(context.output_dir_hint)
    if not output_dir.is_absolute():
        output_dir = _workspace_root(context) / output_dir
    return output_dir / AXMATH_OUTPUT_DIR


def _execution_output_root(context: ExecutionContext) -> Path:
    return Path(context.output_dir) / AXMATH_OUTPUT_DIR


def _write_json(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_input_docx(context: DryRunContext) -> str:
    plan_path_text = context.execution_plan_path.strip()
    if not plan_path_text:
        return _INPUT_DOCX_PLACEHOLDER

    plan_path = Path(plan_path_text)
    if not plan_path.is_absolute():
        plan_path = _workspace_root(context) / plan_path

    try:
        payload = json.loads(plan_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return _INPUT_DOCX_PLACEHOLDER

    input_path_text = str(payload.get("input_path", "")).strip()
    if not input_path_text:
        return _INPUT_DOCX_PLACEHOLDER

    input_path = Path(input_path_text)
    if input_path.is_absolute():
        return str(input_path)
    return str((_workspace_root(context) / input_path).resolve())


def _validate_step(step: ExecutionStep) -> None:
    if step.source_family != AXMATH_SOURCE_FAMILY:
        raise ValueError(f"AxMath executor cannot handle source_family={step.source_family!r}.")


def _review_status(step: ExecutionStep) -> str:
    return "review-gated" if step.requires_manual_review else "validation-gated"


def _write_export_gate_record(
    step: ExecutionStep,
    context: ExecutionContext,
    output_root: Path,
) -> Path:
    record_path = output_root / "blocker-record.json"
    next_ready_condition = (
        "Re-run with --allow-external-tools after preparing an approved AxMath/vendor export workflow "
        "that can emit reviewed MathML or LaTeX artifacts."
        if not context.allow_external_tools
        else "Use a verified AxMath/vendor export workflow to emit reviewed MathML or LaTeX artifacts, "
        "then complete the import and manual review gate before delivery."
    )
    _write_json(
        record_path,
        {
            "artifact_type": "axmath-export-assisted-blocker-record",
            "provider": step.provider,
            "source_family": step.source_family,
            "canonical_target": canonical_mathml_contract_for_source_family(step.source_family).to_dict(),
            "route_kind": step.route_kind,
            "action_id": "export-assisted-conversion",
            "status": "blocked-external-tool" if not context.allow_external_tools else "validation-gated",
            "gate_state": "blocked" if not context.allow_external_tools else "validation-gated",
            "blocking": True,
            "runner": "external-axmath-export",
            "input_path": context.input_path,
            "execution_plan_path": context.execution_plan_path,
            "output_root": str(output_root),
            "external_export_dependency": {
                "required": True,
                "kind": "vendor-export-workflow",
                "description": (
                    "An approved AxMath/vendor export workflow must emit reviewed canonical MathML artifacts, "
                    "or LaTeX artifacts with a separately validated LaTeX-to-MathML conversion step."
                ),
                "allow_external_tools": context.allow_external_tools,
                "verified_cli_binding": False,
            },
            "export_admissibility": axmath_export_admissibility_requirements(),
            "required_evidence": [
                "reviewed canonical MathML artifact(s), or LaTeX plus validated MathML conversion",
                "manual semantic review of exported formulas",
                "render parity check against the source document",
            ],
            "review_requirements": [
                "compare the exported formulas against the source AxMath OLE content",
                "confirm the export can be consumed by the downstream import/review gate",
            ],
            "next_ready_condition": next_ready_condition,
        },
    )
    return record_path


def _classify_dry_run_report(action: ExecutionAction, context: DryRunContext) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=action.blocking,
        supported=True,
        status="export-gate",
        runner="axmath-route-gate",
        cwd=str(_workspace_root(context)),
        notes=(
            "AxMath intake uses detector/classifier evidence and an export-assisted route.",
            "No native AxMath static parser is assumed or previewed by this provider.",
        ),
    )


def _export_dry_run_report(
    action: ExecutionAction,
    context: DryRunContext,
) -> DryRunActionReport:
    input_docx = _load_input_docx(context)
    output_root = _dry_run_output_root(context)
    return DryRunActionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=action.blocking,
        supported=True,
        status="export-gate",
        runner="external-axmath-export",
        cwd=str(_workspace_root(context)),
        notes=(
            "AxMath conversion is export-assisted: an approved AxMath/vendor export workflow must create MathML or LaTeX artifacts.",
            "No command is registered because this provider does not have a verified native static parser or CLI binding.",
            "Next stage requires reviewed export artifacts that satisfy the blocker-record export admissibility checklist.",
            f"Input document for the export gate: {input_docx}",
            f"Expected reviewed export artifact root: {output_root}",
        ),
    )


def _import_dry_run_report(action: ExecutionAction, context: DryRunContext) -> DryRunActionReport:
    output_root = _dry_run_output_root(context)
    return DryRunActionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=True,
        supported=True,
        status="validation-gated",
        runner="manual-import-gate",
        cwd=str(_workspace_root(context)),
        notes=(
            "Import is blocked until reviewed AxMath export artifacts exist.",
            f"Expected reviewed export artifact root: {output_root}",
            "The provider does not synthesize converted math from AxMath OLE payloads directly.",
            "Use the blocker-record export admissibility checklist before accepting canonical MathML artifacts.",
        ),
    )


def _review_dry_run_report(
    step: ExecutionStep,
    action: ExecutionAction,
    context: DryRunContext,
) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=action.blocking or step.requires_manual_review,
        supported=True,
        status=_review_status(step),
        runner="manual-review",
        cwd=str(_workspace_root(context)),
        notes=(
            "Review gate is mandatory for AxMath export-assisted output before delivery claims.",
            "Check exported MathML/LaTeX semantics and rendered formula parity against the source document.",
        ),
    )


def _unknown_dry_run_report(action: ExecutionAction, context: DryRunContext) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=True,
        supported=False,
        status="manual-gate",
        runner="manual",
        cwd=str(_workspace_root(context)),
        notes=("Unrecognized AxMath action; manual triage is required before execution binding.",),
    )


def build_axmath_dry_run_reports(
    step: ExecutionStep,
    context: DryRunContext,
) -> tuple[DryRunActionReport, ...]:
    _validate_step(step)

    reports: list[DryRunActionReport] = []
    for action in step.actions:
        if action.action_id == "classify-axmath-object":
            reports.append(_classify_dry_run_report(action, context))
            continue
        if action.action_id == "export-assisted-conversion":
            reports.append(_export_dry_run_report(action, context))
            continue
        if action.action_id == "import-converted-math":
            reports.append(_import_dry_run_report(action, context))
            continue
        if action.action_id == "manual-spot-check":
            reports.append(_review_dry_run_report(step, action, context))
            continue
        reports.append(_unknown_dry_run_report(action, context))

    if reports:
        return tuple(reports)

    return (
        DryRunActionReport(
            action_id="axmath-export-triage",
            description="Triage AxMath export-assisted conversion requirements.",
            blocking=True,
            supported=False,
            status="manual-gate",
            runner="manual",
            cwd=str(_workspace_root(context)),
            notes=("No AxMath execution actions were supplied by the execution plan.",),
        ),
    )


def _execution_report(
    action: ExecutionAction,
    *,
    status: str,
    runner: str,
    context: ExecutionContext,
    blocking: bool | None = None,
    supported: bool = True,
    output_paths: tuple[Path, ...] = (),
    notes: tuple[str, ...] = (),
) -> ActionExecutionReport:
    return ActionExecutionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=action.blocking if blocking is None else blocking,
        supported=supported,
        status=status,
        runner=runner,
        cwd=context.workspace_root,
        output_paths=tuple(str(path) for path in output_paths),
        notes=notes,
    )


def execute_axmath_step(
    step: ExecutionStep,
    context: ExecutionContext,
) -> tuple[ActionExecutionReport, ...]:
    _validate_step(step)

    output_root = _execution_output_root(context)
    reports: list[ActionExecutionReport] = []
    for action in step.actions:
        if action.action_id == "classify-axmath-object":
            reports.append(
                _execution_report(
                    action,
                    status="export-gate",
                    runner="axmath-route-gate",
                    context=context,
                    notes=(
                        "Proceeding from execution-plan classifier evidence only.",
                        "No native AxMath static parse was attempted.",
                    ),
                )
            )
            continue

        if action.action_id == "export-assisted-conversion":
            gate_path = _write_export_gate_record(step, context, output_root)
            if context.allow_external_tools:
                reports.append(
                    _execution_report(
                        action,
                        status="validation-gated",
                        runner="external-axmath-export",
                        context=context,
                        output_paths=(gate_path,),
                        notes=(
                            "External tools are allowed by the execution context, but no verified AxMath CLI binding is registered.",
                            "Run an approved AxMath export workflow and record reviewed MathML/LaTeX artifacts before import.",
                            "The blocker record lists export admissibility requirements for canonical MathML promotion.",
                            f"Expected reviewed export artifact root: {output_root}",
                        ),
                    )
                )
                continue

            reports.append(
                _execution_report(
                    action,
                    status="blocked-external-tool",
                    runner="external-axmath-export",
                    context=context,
                    output_paths=(gate_path,),
                    notes=(
                        "AxMath execution requires an external vendor/export workflow and is blocked unless explicitly allowed.",
                        "Use --allow-external-tools only after the AxMath export environment and review workflow are prepared.",
                        "This provider does not claim native static parsing support.",
                        "The blocker record lists export admissibility requirements for canonical MathML promotion.",
                    ),
                )
            )
            continue

        if action.action_id == "import-converted-math":
            reports.append(
                _execution_report(
                    action,
                    status="skipped-until-export-artifacts",
                    runner="manual-import-gate",
                    context=context,
                    blocking=True,
                    notes=(
                        "Import is skipped until reviewed AxMath export artifacts are available.",
                        f"Expected reviewed export artifact root: {output_root}",
                    ),
                )
            )
            continue

        if action.action_id == "manual-spot-check":
            reports.append(
                _execution_report(
                    action,
                    status=_review_status(step),
                    runner="manual-review",
                    context=context,
                    blocking=action.blocking or step.requires_manual_review,
                    notes=(
                        "Manual semantic and render review remains required for AxMath export-assisted output.",
                        "This step remains non-deliverable until the review gate is satisfied.",
                    ),
                )
            )
            continue

        reports.append(
            _execution_report(
                action,
                status="manual-gate",
                runner="manual",
                context=context,
                blocking=True,
                supported=False,
                notes=("Unrecognized AxMath action; manual triage is required.",),
            )
        )

    if reports:
        return tuple(reports)

    return (
        ActionExecutionReport(
            action_id="axmath-export-triage",
            description="Triage AxMath export-assisted conversion requirements.",
            blocking=True,
            supported=False,
            status="manual-gate",
            runner="manual",
            cwd=context.workspace_root,
            notes=("No AxMath execution actions were supplied by the execution plan.",),
        ),
    )
