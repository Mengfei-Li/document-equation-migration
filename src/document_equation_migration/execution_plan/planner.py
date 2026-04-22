from __future__ import annotations

from collections.abc import Mapping
from typing import Any

from ..source_taxonomy import SourceFamily
from .axmath import build_axmath_execution_step
from .base import ExecutionStepBuilder, RouteEntry
from .equation3 import build_equation3_execution_step
from .mathtype import build_mathtype_execution_step
from .model import ExecutionAction, ExecutionPlan, ExecutionStep
from .odf import build_odf_execution_step
from .omml import build_omml_execution_step

PROVIDER_REGISTRY: dict[SourceFamily, ExecutionStepBuilder] = {
    SourceFamily.MATHTYPE_OLE: build_mathtype_execution_step,
    SourceFamily.OMML_NATIVE: build_omml_execution_step,
    SourceFamily.EQUATION_EDITOR_3_OLE: build_equation3_execution_step,
    SourceFamily.AXMATH_OLE: build_axmath_execution_step,
    SourceFamily.ODF_NATIVE: build_odf_execution_step,
    SourceFamily.LIBREOFFICE_TRANSFORMED: build_odf_execution_step,
}


def _string(value: object, *, default: str = "") -> str:
    if value is None:
        return default
    return str(value)


def _integer(value: object, *, default: int = 0) -> int:
    if isinstance(value, bool):
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return default
        try:
            return int(value)
        except ValueError:
            return default
    return default


def _boolean(value: object, *, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes"}:
            return True
        if lowered in {"0", "false", "no"}:
            return False
    return default


def _build_default_step(source_family: str, route_entry: RouteEntry) -> ExecutionStep:
    return ExecutionStep(
        source_family=source_family,
        formula_count=_integer(route_entry.get("formula_count")),
        route_kind=_string(route_entry.get("route_kind")),
        confidence_policy=_string(route_entry.get("confidence_policy")),
        requires_manual_review=_boolean(route_entry.get("requires_manual_review"), default=True),
        provider="default",
        next_action=_string(route_entry.get("next_action")),
        actions=(
            ExecutionAction(
                action_id="manual-triage",
                description="Capture sample and classify conversion route manually.",
                blocking=True,
            ),
        ),
        notes=("No source-line provider is registered for this source family.",),
    )


def _build_step(route_entry: Mapping[str, Any]) -> ExecutionStep:
    source_family_raw = _string(route_entry.get("source_family"))
    try:
        source_family = SourceFamily(source_family_raw)
    except ValueError:
        fallback_family = source_family_raw or SourceFamily.UNKNOWN_OLE.value
        return _build_default_step(fallback_family, route_entry)
    builder = PROVIDER_REGISTRY.get(source_family)
    if builder is None:
        return _build_default_step(source_family.value, route_entry)
    return builder(route_entry)


def build_execution_plan(routing_report: Mapping[str, Any]) -> ExecutionPlan:
    route_plan = routing_report.get("route_plan")
    if not isinstance(route_plan, list):
        raise ValueError("routing_report.route_plan must be a list")

    steps = tuple(_build_step(item) for item in route_plan)
    recommended_sequence = tuple(_string(item) for item in routing_report.get("recommended_sequence", []))
    return ExecutionPlan(
        document_id=_string(routing_report.get("document_id")),
        input_path=_string(routing_report.get("input_path")),
        detector_version=_string(routing_report.get("detector_version")),
        formula_count=_integer(routing_report.get("formula_count")),
        recommended_sequence=recommended_sequence,
        steps=steps,
    )
