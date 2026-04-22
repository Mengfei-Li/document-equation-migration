from __future__ import annotations

import subprocess
from dataclasses import dataclass, field


def _command_line(argv: tuple[str, ...]) -> str:
    if not argv:
        return ""
    return subprocess.list2cmdline(list(argv))


@dataclass(frozen=True, slots=True)
class DryRunContext:
    workspace_root: str
    execution_plan_path: str
    output_dir_hint: str


@dataclass(frozen=True, slots=True)
class DryRunActionReport:
    action_id: str
    description: str
    blocking: bool
    supported: bool
    status: str
    runner: str
    argv: tuple[str, ...] = ()
    cwd: str = ""
    notes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, object]:
        return {
            "action_id": self.action_id,
            "description": self.description,
            "blocking": self.blocking,
            "supported": self.supported,
            "status": self.status,
            "runner": self.runner,
            "argv": list(self.argv),
            "command_line": _command_line(self.argv),
            "cwd": self.cwd,
            "notes": list(self.notes),
        }


@dataclass(frozen=True, slots=True)
class DryRunStepReport:
    source_family: str
    provider: str
    route_kind: str
    formula_count: int
    next_action: str
    requires_manual_review: bool
    status: str
    actions: tuple[DryRunActionReport, ...] = ()
    notes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, object]:
        return {
            "source_family": self.source_family,
            "provider": self.provider,
            "route_kind": self.route_kind,
            "formula_count": self.formula_count,
            "next_action": self.next_action,
            "requires_manual_review": self.requires_manual_review,
            "status": self.status,
            "actions": [action.to_dict() for action in self.actions],
            "notes": list(self.notes),
        }


@dataclass(frozen=True, slots=True)
class DryRunExecutionReport:
    document_id: str
    input_path: str
    detector_version: str
    execution_plan_path: str
    mode: str
    step_count: int
    runnable_step_count: int
    manual_only_step_count: int
    manual_review_required: bool
    steps: tuple[DryRunStepReport, ...] = field(default_factory=tuple)

    def to_dict(self) -> dict[str, object]:
        return {
            "document_id": self.document_id,
            "input_path": self.input_path,
            "detector_version": self.detector_version,
            "execution_plan_path": self.execution_plan_path,
            "mode": self.mode,
            "step_count": self.step_count,
            "runnable_step_count": self.runnable_step_count,
            "manual_only_step_count": self.manual_only_step_count,
            "manual_review_required": self.manual_review_required,
            "steps": [step.to_dict() for step in self.steps],
        }


@dataclass(frozen=True, slots=True)
class ExecutionContext:
    workspace_root: str
    execution_plan_path: str
    input_path: str
    output_dir: str
    allow_external_tools: bool = False


@dataclass(frozen=True, slots=True)
class ActionExecutionReport:
    action_id: str
    description: str
    blocking: bool
    supported: bool
    status: str
    runner: str
    argv: tuple[str, ...] = ()
    cwd: str = ""
    exit_code: int | None = None
    stdout_path: str = ""
    stderr_path: str = ""
    output_paths: tuple[str, ...] = ()
    notes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, object]:
        return {
            "action_id": self.action_id,
            "description": self.description,
            "blocking": self.blocking,
            "supported": self.supported,
            "status": self.status,
            "runner": self.runner,
            "argv": list(self.argv),
            "command_line": _command_line(self.argv),
            "cwd": self.cwd,
            "exit_code": self.exit_code,
            "stdout_path": self.stdout_path,
            "stderr_path": self.stderr_path,
            "output_paths": list(self.output_paths),
            "notes": list(self.notes),
        }


@dataclass(frozen=True, slots=True)
class StepExecutionReport:
    source_family: str
    provider: str
    route_kind: str
    formula_count: int
    next_action: str
    requires_manual_review: bool
    status: str
    actions: tuple[ActionExecutionReport, ...] = ()
    notes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, object]:
        return {
            "source_family": self.source_family,
            "provider": self.provider,
            "route_kind": self.route_kind,
            "formula_count": self.formula_count,
            "next_action": self.next_action,
            "requires_manual_review": self.requires_manual_review,
            "status": self.status,
            "actions": [action.to_dict() for action in self.actions],
            "notes": list(self.notes),
        }


@dataclass(frozen=True, slots=True)
class PlanExecutionReport:
    document_id: str
    input_path: str
    detector_version: str
    execution_plan_path: str
    mode: str
    output_dir: str
    step_count: int
    completed_step_count: int
    blocked_step_count: int
    manual_only_step_count: int
    manual_review_required: bool
    allow_external_tools: bool
    steps: tuple[StepExecutionReport, ...] = field(default_factory=tuple)

    def to_dict(self) -> dict[str, object]:
        return {
            "document_id": self.document_id,
            "input_path": self.input_path,
            "detector_version": self.detector_version,
            "execution_plan_path": self.execution_plan_path,
            "mode": self.mode,
            "output_dir": self.output_dir,
            "step_count": self.step_count,
            "completed_step_count": self.completed_step_count,
            "blocked_step_count": self.blocked_step_count,
            "manual_only_step_count": self.manual_only_step_count,
            "manual_review_required": self.manual_review_required,
            "allow_external_tools": self.allow_external_tools,
            "steps": [step.to_dict() for step in self.steps],
        }
