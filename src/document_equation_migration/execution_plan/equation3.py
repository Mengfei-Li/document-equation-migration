from __future__ import annotations

from .base import RouteEntry
from .model import ExecutionAction, ExecutionStep

SOURCE_FAMILY = "equation-editor-3-ole"
PROVIDER_NAME = "equation3"


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
        raw = value.strip()
        if not raw:
            return default
        try:
            return int(raw)
        except ValueError:
            return default
    return default


def _boolean(value: object, *, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value != 0
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes"}:
            return True
        if lowered in {"0", "false", "no"}:
            return False
    return default


def _manual_review_required(route_entry: RouteEntry) -> bool:
    signal_keys = (
        "requires_manual_review",
        "manual_review_required",
        "manual_review_signal",
        "manual_review",
    )
    has_explicit_signal = False
    for key in signal_keys:
        if key in route_entry:
            has_explicit_signal = True
            if _boolean(route_entry.get(key), default=False):
                return True
    if has_explicit_signal:
        return False
    return True


def build_equation3_execution_step(route_entry: RouteEntry) -> ExecutionStep:
    requires_manual_review = _manual_review_required(route_entry)
    notes = [
        "Equation Editor 3.0 has an internal limited MTEF v3 to canonical MathML path for supported script, root, fraction, slash-fraction, bar, fence, limit, matrix, pile, BigOp (sum/integral/product/coproduct/integral-op), character structures, and narrow legacy post-END footers."
    ]
    if requires_manual_review:
        notes.append("Preserve manual-review gate for unsupported MTEF records, legacy .doc ingestion, and deliverability claims.")

    return ExecutionStep(
        source_family=SOURCE_FAMILY,
        formula_count=_integer(route_entry.get("formula_count"), default=0),
        route_kind=_string(route_entry.get("route_kind"), default="primary-candidate"),
        confidence_policy=_string(route_entry.get("confidence_policy"), default="medium"),
        requires_manual_review=requires_manual_review,
        provider=PROVIDER_NAME,
        next_action=_string(route_entry.get("next_action"), default="run-equation3-probe-and-conversion"),
        actions=(
            ExecutionAction(
                action_id="probe-header-and-classid",
                description="Probe OLE header and ClassID to confirm Equation Editor 3.0 payload.",
            ),
            ExecutionAction(
                action_id="attempt-mtef-conversion",
                description="Attempt limited MTEF v3 to canonical MathML conversion as the primary Equation Editor 3.0 path.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="fallback-manual-triage",
                description="Fallback to manual triage when probe or conversion results are ambiguous.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="word-roundtrip-validation",
                description="Validate downstream Word roundtrip only after canonical MathML evidence is accepted.",
                blocking=True,
            ),
        ),
        notes=tuple(notes),
    )
