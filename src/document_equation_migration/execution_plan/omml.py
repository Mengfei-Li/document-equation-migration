from __future__ import annotations

from .base import RouteEntry
from .model import ExecutionAction, ExecutionStep


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


def build_omml_execution_step(route_entry: RouteEntry) -> ExecutionStep:
    formula_count = _integer(route_entry.get("formula_count"))
    route_kind = _string(route_entry.get("route_kind"), default="primary-source-first")
    confidence_policy = _string(route_entry.get("confidence_policy"), default="high")
    requires_manual_review = _boolean(route_entry.get("requires_manual_review"), default=False)
    next_action = _string(route_entry.get("next_action"), default="run-omml-native-pipeline")

    return ExecutionStep(
        source_family="omml-native",
        formula_count=formula_count,
        route_kind=route_kind,
        confidence_policy=confidence_policy,
        requires_manual_review=requires_manual_review,
        provider="omml",
        next_action=next_action,
        actions=(
            ExecutionAction(
                action_id="extract-omml",
                description="Extract OMML equations from OOXML document parts.",
                metadata={
                    "source_family": "omml-native",
                    "formula_count": formula_count,
                    "route_kind": route_kind,
                },
            ),
            ExecutionAction(
                action_id="normalize-omml",
                description="Normalize OMML structure for deterministic downstream conversion.",
                metadata={
                    "confidence_policy": confidence_policy,
                    "preserve_semantics": True,
                },
            ),
            ExecutionAction(
                action_id="render-check",
                description="Run render parity checks against source equations.",
                blocking=requires_manual_review,
                metadata={"requires_manual_review": requires_manual_review},
            ),
            ExecutionAction(
                action_id="package-omml-output",
                description="Package normalized OMML output and execution metadata.",
                metadata={"next_action": next_action},
            ),
        ),
        notes=(
            "OMML-native route stays on a primary-source-first pipeline with no OLE dependency.",
            "Native WordprocessingML math objects usually keep conversion risk low and predictable.",
            "Manual review remains policy-driven and only escalates when routing marks the sample.",
        ),
    )
