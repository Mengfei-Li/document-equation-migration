from __future__ import annotations

from collections.abc import Mapping, Sequence

from ..axmath_controlled_acceptance import (
    AXMATH_CONTROLLED_ACCEPTANCE_KEY,
    normalize_axmath_controlled_acceptance_contract,
)
from ..axmath_structural_metadata import classify_axmath_structural_metadata_rows
from .base import RouteEntry
from .model import ExecutionAction, ExecutionStep

AXMATH_SOURCE_FAMILY = "axmath-ole"
DEFAULT_ROUTE_KIND = "export-assisted"
DEFAULT_CONFIDENCE_POLICY = "medium"
DEFAULT_NEXT_ACTION = "run-axmath-export-assisted-pipeline"
STRUCTURAL_METADATA_KEY = "axmath_structural_metadata"


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


def _structural_metadata_rows(value: object) -> tuple[Mapping[str, object], ...]:
    if isinstance(value, Mapping):
        for key in ("rows", "records"):
            rows = value.get(key)
            if isinstance(rows, Sequence) and not isinstance(rows, str | bytes | bytearray):
                return tuple(row for row in rows if isinstance(row, Mapping))
        return (value,)
    if isinstance(value, Sequence) and not isinstance(value, str | bytes | bytearray):
        return tuple(row for row in value if isinstance(row, Mapping))
    raise TypeError("AxMath structural metadata must be a mapping or sequence of mappings.")


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
    structural_metadata = route_entry.get(STRUCTURAL_METADATA_KEY)
    controlled_acceptance = route_entry.get(AXMATH_CONTROLLED_ACCEPTANCE_KEY)

    actions: list[ExecutionAction] = [
        ExecutionAction(
            action_id="classify-axmath-object",
            description="Classify AxMath OLE objects and confirm export-assisted compatibility.",
            blocking=True,
            metadata={
                "source_family": AXMATH_SOURCE_FAMILY,
                "route_kind": route_kind,
            },
        ),
    ]
    notes = [
        "AxMath conversion follows an export-assisted route; public BYO Contents inspection is structural metadata-only.",
        "Retain manual review signal from route_entry.requires_manual_review for downstream gating.",
    ]

    if structural_metadata is not None:
        structural_classification = classify_axmath_structural_metadata_rows(
            _structural_metadata_rows(structural_metadata)
        )
        actions.append(
            ExecutionAction(
                action_id="classify-axmath-structural-metadata",
                description=(
                    "Classify explicit AxMath structural metadata rows without inspecting payload bodies."
                ),
                blocking=False,
                metadata=structural_classification.to_dict(),
            )
        )
        notes.append(
            "Explicit AxMath structural metadata is metadata-only and does not establish semantic native parser or conversion support."
        )

    if controlled_acceptance is not None:
        acceptance_contract = normalize_axmath_controlled_acceptance_contract(controlled_acceptance)
        actions.append(
            ExecutionAction(
                action_id="record-axmath-controlled-acceptance-contract",
                description=(
                    "Record public-safe AxMath controlled acceptance metadata without publishing private artifacts."
                ),
                blocking=False,
                metadata=acceptance_contract.to_dict(),
            )
        )
        notes.append(
            "AxMath controlled acceptance metadata is contract-only and does not establish a public native parser binding."
        )

    actions.extend(
        [
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
        ]
    )

    return ExecutionStep(
        source_family=AXMATH_SOURCE_FAMILY,
        formula_count=formula_count,
        route_kind=route_kind,
        confidence_policy=confidence_policy,
        requires_manual_review=requires_manual_review,
        provider="axmath",
        next_action=next_action,
        actions=tuple(actions),
        notes=tuple(notes),
    )
