from __future__ import annotations

from collections.abc import Mapping
from typing import Protocol

from .model import ExecutionStep

RouteEntry = Mapping[str, object]


class ExecutionStepBuilder(Protocol):
    def __call__(self, route_entry: RouteEntry) -> ExecutionStep: ...
