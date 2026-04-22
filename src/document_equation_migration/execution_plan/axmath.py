from __future__ import annotations

from .base import RouteEntry
from .model import ExecutionAction, ExecutionStep

AXMATH_SOURCE_FAMILY = "axmath-ole"
DEFAULT_ROUTE_KIND = "export-assisted"
DEFAULT_CONFIDENCE_POLICY = "medium"
DEFAULT_NEXT_ACTION = "run-axmath-export-assisted-pipeline"


def _to_str(value: object, *, default: str = "") -> str:
    if value is None:
        return default
    return str(value)


def _to_int(value: object, *, default: int = 0) -> int:
    if isinstance(value, bool):
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return default
        try:
            return int(text)
        except ValueError:
            return default
    return default


def _to_bool(value: object, *, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes"}:
            return True
        if lowered in {"0", "false", "no"}:
            return False
    return default


def build_axmath_execution_step(route_entry: RouteEntry) -> ExecutionStep:
    source_family = _to_str(route_entry.get("source_family"), default=AXMATH_SOURCE_FAMILY)
    if source_family != AXMATH_SOURCE_FAMILY:
        raise ValueError(
            f"AxMath execution provider only accepts source_family={AXMATH_SOURCE_FAMILY!r}, "
            f"got {source_family!r}."
        )

    formula_count = _to_int(route_entry.get("formula_count"))
    route_kind = _to_str(route_entry.get("route_kind"), default=DEFAULT_ROUTE_KIND)
    confidence_policy = _to_str(route_entry.get("confidence_policy"), default=DEFAULT_CONFIDENCE_POLICY)
    requires_manual_review = _to_bool(route_entry.get("requires_manual_review"), default=True)
    next_action = _to_str(route_entry.get("next_action"), default=DEFAULT_NEXT_ACTION)

    return ExecutionStep(
        source_family=AXMATH_SOURCE_FAMILY,
        formula_count=formula_count,
        route_kind=route_kind,
        confidence_policy=confidence_policy,
        requires_manual_review=requires_manual_review,
        provider="axmath",
        next_action=next_action,
        actions=(
            ExecutionAction(
                action_id="classify-axmath-object",
                description="Classify AxMath OLE objects and confirm export-assisted compatibility.",
                blocking=True,
                metadata={
                    "source_family": AXMATH_SOURCE_FAMILY,
                    "route_kind": route_kind,
                },
            ),
            ExecutionAction(
                action_id="export-assisted-conversion",
                description="Run export-assisted conversion to produce normalized math payloads.",
                blocking=True,
                metadata={
                    "next_action": next_action,
                    "confidence_policy": confidence_policy,
                },
            ),
            ExecutionAction(
                action_id="import-converted-math",
                description="Import converted math back into the target document structure.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="manual-spot-check",
                description="Perform manual spot checks on converted formulas before delivery.",
                blocking=requires_manual_review,
                metadata={
                    "requires_manual_review": requires_manual_review,
                    "manual_review_signal_raw": route_entry.get("requires_manual_review"),
                },
            ),
        ),
        notes=(
            "AxMath follows an export-assisted route and is not a primary-source-first pipeline.",
            "Retain manual review signal from route_entry.requires_manual_review for downstream gating.",
        ),
    )
