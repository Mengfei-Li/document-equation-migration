from __future__ import annotations

from .base import RouteEntry
from .model import ExecutionAction, ExecutionStep


def build_odf_execution_step(route_entry: RouteEntry) -> ExecutionStep:
    source_family = _string(route_entry.get("source_family"))
    normalized_source_family = source_family.strip().lower()

    if normalized_source_family == "libreoffice-transformed":
        return ExecutionStep(
            source_family=source_family,
            formula_count=_integer(route_entry.get("formula_count")),
            route_kind=_string(route_entry.get("route_kind"), default="bridge-source"),
            confidence_policy=_string(route_entry.get("confidence_policy"), default="low"),
            requires_manual_review=_manual_review_signal(route_entry, default=True),
            provider="odf",
            next_action=_string(route_entry.get("next_action"), default="run-libreoffice-bridge-review-pipeline"),
            actions=(
                ExecutionAction(
                    action_id="inspect-transform-chain",
                    description="Inspect LibreOffice transformation chain and identify source-of-truth risks.",
                    blocking=True,
                ),
                ExecutionAction(
                    action_id="bridge-review",
                    description="Review bridge output fidelity against expected math structure before acceptance.",
                    blocking=True,
                ),
                ExecutionAction(
                    action_id="decide-reconvert-or-accept",
                    description="Decide whether to reconvert from native source or accept bridge output.",
                    blocking=True,
                ),
            ),
            notes=(
                "ODF-native is the primary line; LibreOffice-transformed output is treated as bridge evidence.",
                "Bridge decisions should prefer reconversion when transform-chain confidence is insufficient.",
            ),
        )

    return ExecutionStep(
        source_family=source_family,
        formula_count=_integer(route_entry.get("formula_count")),
        route_kind=_string(route_entry.get("route_kind"), default="primary-source-first"),
        confidence_policy=_string(route_entry.get("confidence_policy"), default="medium"),
        requires_manual_review=_manual_review_signal(route_entry, default=False),
        provider="odf",
        next_action=_string(route_entry.get("next_action"), default="run-odf-native-pipeline"),
        actions=(
            ExecutionAction(
                action_id="extract-odf-formula",
                description="Extract native ODF formula payloads from source document.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="convert-odf-mathml",
                description="Convert ODF formula payloads to canonical MathML for downstream pipelines.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="emit-target-format",
                description="Emit target format from canonical representation.",
                metadata={"target_hint": _string(route_entry.get("next_action"))},
            ),
        ),
        notes=(
            "ODF-native is the primary conversion line and should be preferred when native payloads are available.",
            "LibreOffice bridge remains a secondary path for transformed inputs and reconciliation.",
        ),
    )


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
        text = value.strip()
        if not text:
            return default
        try:
            return int(text)
        except ValueError:
            return default
    return default


def _boolean(value: object, *, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes", "y"}:
            return True
        if lowered in {"0", "false", "no", "n"}:
            return False
    return default


def _manual_review_signal(route_entry: RouteEntry, *, default: bool) -> bool:
    if "requires_manual_review" in route_entry:
        return _boolean(route_entry.get("requires_manual_review"), default=default)
    if "manual_review_required" in route_entry:
        return _boolean(route_entry.get("manual_review_required"), default=default)

    signals = route_entry.get("signals")
    if isinstance(signals, (list, tuple, set)):
        lowered = {str(item).strip().lower() for item in signals}
        if "manual-review" in lowered or "manual_review" in lowered:
            return True

    return default
