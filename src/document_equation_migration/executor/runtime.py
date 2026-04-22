from __future__ import annotations

import json
from pathlib import Path

from ..execution_plan.model import ExecutionAction, ExecutionPlan, ExecutionStep
from .model import (
    ActionExecutionReport,
    DryRunActionReport,
    DryRunContext,
    DryRunExecutionReport,
    DryRunStepReport,
    ExecutionContext,
    PlanExecutionReport,
    StepExecutionReport,
)
from .registry import DRY_RUN_BINDING_REGISTRY, EXECUTION_BINDING_REGISTRY


def _workspace_root() -> Path:
    return Path(__file__).resolve().parents[3]


def _default_output_dir(plan: ExecutionPlan) -> str:
    return str(_workspace_root() / "out" / plan.document_id)


def _resolve_input_path(plan: ExecutionPlan, execution_plan_path: str) -> str:
    input_path = Path(plan.input_path)
    if input_path.is_absolute():
        return str(input_path)

    if execution_plan_path:
        plan_path = Path(execution_plan_path)
        if not plan_path.is_absolute():
            plan_path = _workspace_root() / plan_path
        candidate = plan_path.parent / input_path
        if candidate.exists():
            return str(candidate.resolve())

    return str((_workspace_root() / input_path).resolve())


def load_execution_plan(path: str | Path) -> ExecutionPlan:
    plan_path = Path(path)
    payload = json.loads(plan_path.read_text(encoding="utf-8"))
    return ExecutionPlan.from_dict(payload)


def _build_fallback_action_report(action: ExecutionAction) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=action.blocking,
        supported=False,
        status="manual-gate",
        runner="manual",
        notes=("No concrete dry-run binding is registered for this action yet.",),
    )


def _build_fallback_step_reports(step: ExecutionStep) -> tuple[DryRunActionReport, ...]:
    if not step.actions:
        return (
            DryRunActionReport(
                action_id="manual-triage",
                description="No execution actions are registered for this step.",
                blocking=True,
                supported=False,
                status="manual-gate",
                runner="manual",
                notes=("The provider returned no actions; manual classification is required.",),
            ),
        )
    return tuple(_build_fallback_action_report(action) for action in step.actions)


def _step_status(actions: tuple[DryRunActionReport, ...]) -> str:
    if actions and all(action.supported for action in actions):
        return "runnable"
    return "manual-only"


def _build_fallback_execution_reports(step: ExecutionStep) -> tuple[ActionExecutionReport, ...]:
    if not step.actions:
        return (
            ActionExecutionReport(
                action_id="manual-triage",
                description="No execution actions are registered for this step.",
                blocking=True,
                supported=False,
                status="manual-gate",
                runner="manual",
                notes=("The provider returned no actions; manual classification is required.",),
            ),
        )
    return tuple(
        ActionExecutionReport(
            action_id=action.action_id,
            description=action.description,
            blocking=action.blocking,
            supported=False,
            status="manual-gate",
            runner="manual",
            notes=("No concrete execution binding is registered for this action yet.",),
        )
        for action in step.actions
    )


def _execution_step_status(actions: tuple[ActionExecutionReport, ...]) -> str:
    statuses = {action.status for action in actions}
    if not actions:
        return "manual-only"
    if "failed" in statuses:
        return "failed"
    if "blocked-external-tool" in statuses:
        return "blocked"
    if statuses == {"manual-gate"} or any(not action.supported for action in actions):
        return "manual-only"
    if "validation-gated" in statuses:
        return "validation-gated"
    if "review-gated" in statuses:
        return "review-gated"
    if any(status.startswith("skipped") for status in statuses):
        return "completed-with-skips"
    return "completed"


def build_dry_run_execution_report(
    plan: ExecutionPlan,
    *,
    execution_plan_path: str = "",
) -> DryRunExecutionReport:
    context = DryRunContext(
        workspace_root=str(_workspace_root()),
        execution_plan_path=execution_plan_path,
        output_dir_hint=_default_output_dir(plan),
    )

    step_reports: list[DryRunStepReport] = []
    runnable_step_count = 0
    manual_only_step_count = 0

    for step in plan.steps:
        builder = DRY_RUN_BINDING_REGISTRY.get(step.provider)
        if builder is None:
            action_reports = _build_fallback_step_reports(step)
            step_notes = step.notes + ("Dry-run execution falls back to manual gate for this provider.",)
        else:
            try:
                action_reports = builder(step, context)
                step_notes = step.notes
            except NotImplementedError:
                action_reports = _build_fallback_step_reports(step)
                step_notes = step.notes + ("Provider-specific dry-run binding is not implemented yet.",)

        status = _step_status(action_reports)
        if status == "runnable":
            runnable_step_count += 1
        else:
            manual_only_step_count += 1

        step_reports.append(
            DryRunStepReport(
                source_family=step.source_family,
                provider=step.provider,
                route_kind=step.route_kind,
                formula_count=step.formula_count,
                next_action=step.next_action,
                requires_manual_review=step.requires_manual_review,
                status=status,
                actions=action_reports,
                notes=step_notes,
            )
        )

    return DryRunExecutionReport(
        document_id=plan.document_id,
        input_path=plan.input_path,
        detector_version=plan.detector_version,
        execution_plan_path=execution_plan_path,
        mode="dry-run",
        step_count=len(step_reports),
        runnable_step_count=runnable_step_count,
        manual_only_step_count=manual_only_step_count,
        manual_review_required=plan.manual_review_required,
        steps=tuple(step_reports),
    )


def build_execution_report(
    plan: ExecutionPlan,
    *,
    execution_plan_path: str = "",
    output_dir: str | None = None,
    allow_external_tools: bool = False,
) -> PlanExecutionReport:
    resolved_output_dir = str(Path(output_dir).resolve()) if output_dir else _default_output_dir(plan)
    context = ExecutionContext(
        workspace_root=str(_workspace_root()),
        execution_plan_path=execution_plan_path,
        input_path=_resolve_input_path(plan, execution_plan_path),
        output_dir=resolved_output_dir,
        allow_external_tools=allow_external_tools,
    )

    step_reports: list[StepExecutionReport] = []
    completed_step_count = 0
    blocked_step_count = 0
    manual_only_step_count = 0

    for step in plan.steps:
        runner = EXECUTION_BINDING_REGISTRY.get(step.provider)
        if runner is None:
            action_reports = _build_fallback_execution_reports(step)
            step_notes = step.notes + ("Execution falls back to manual gate for this provider.",)
        else:
            try:
                action_reports = runner(step, context)
                step_notes = step.notes
            except Exception as exc:  # noqa: BLE001 - execution reports must capture provider failures.
                action_reports = (
                    ActionExecutionReport(
                        action_id="provider-execution",
                        description=f"Execute provider {step.provider}.",
                        blocking=True,
                        supported=True,
                        status="failed",
                        runner=step.provider,
                        notes=(f"Provider execution failed before producing action reports: {exc}",),
                    ),
                )
                step_notes = step.notes + ("Provider execution raised an exception.",)

        status = _execution_step_status(action_reports)
        if status in {"completed", "completed-with-skips"}:
            completed_step_count += 1
        elif status == "manual-only":
            manual_only_step_count += 1
        else:
            blocked_step_count += 1

        step_reports.append(
            StepExecutionReport(
                source_family=step.source_family,
                provider=step.provider,
                route_kind=step.route_kind,
                formula_count=step.formula_count,
                next_action=step.next_action,
                requires_manual_review=step.requires_manual_review,
                status=status,
                actions=action_reports,
                notes=step_notes,
            )
        )

    return PlanExecutionReport(
        document_id=plan.document_id,
        input_path=plan.input_path,
        detector_version=plan.detector_version,
        execution_plan_path=execution_plan_path,
        mode="execute",
        output_dir=resolved_output_dir,
        step_count=len(step_reports),
        completed_step_count=completed_step_count,
        blocked_step_count=blocked_step_count,
        manual_only_step_count=manual_only_step_count,
        manual_review_required=plan.manual_review_required,
        allow_external_tools=allow_external_tools,
        steps=tuple(step_reports),
    )
