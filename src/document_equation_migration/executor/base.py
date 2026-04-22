from __future__ import annotations

from typing import Protocol

from ..execution_plan.model import ExecutionStep
from .model import ActionExecutionReport, DryRunActionReport, DryRunContext, ExecutionContext


class DryRunBindingBuilder(Protocol):
    def __call__(
        self,
        step: ExecutionStep,
        context: DryRunContext,
    ) -> tuple[DryRunActionReport, ...]: ...


class ExecutionBindingRunner(Protocol):
    def __call__(
        self,
        step: ExecutionStep,
        context: ExecutionContext,
    ) -> tuple[ActionExecutionReport, ...]: ...
