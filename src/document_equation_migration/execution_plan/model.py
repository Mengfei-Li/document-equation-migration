from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


def _as_dict(value: object) -> dict[str, Any]:
    if isinstance(value, dict):
        return dict(value)
    return {}


@dataclass(frozen=True, slots=True)
class ExecutionAction:
    action_id: str
    description: str
    blocking: bool = False
    metadata: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExecutionAction":
        return cls(
            action_id=str(payload.get("action_id", "")),
            description=str(payload.get("description", "")),
            blocking=bool(payload.get("blocking", False)),
            metadata=_as_dict(payload.get("metadata", {})),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "action_id": self.action_id,
            "description": self.description,
            "blocking": self.blocking,
            "metadata": dict(self.metadata),
        }


@dataclass(frozen=True, slots=True)
class ExecutionStep:
    source_family: str
    formula_count: int
    route_kind: str
    confidence_policy: str
    requires_manual_review: bool
    provider: str
    next_action: str
    metadata: dict[str, Any] = field(default_factory=dict)
    actions: tuple[ExecutionAction, ...] = ()
    notes: tuple[str, ...] = ()

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExecutionStep":
        actions = tuple(
            ExecutionAction.from_dict(item)
            for item in payload.get("actions", [])
            if isinstance(item, dict)
        )
        notes = tuple(str(item) for item in payload.get("notes", []))
        return cls(
            source_family=str(payload.get("source_family", "")),
            formula_count=int(payload.get("formula_count", 0)),
            route_kind=str(payload.get("route_kind", "")),
            confidence_policy=str(payload.get("confidence_policy", "")),
            requires_manual_review=bool(payload.get("requires_manual_review", False)),
            provider=str(payload.get("provider", "")),
            next_action=str(payload.get("next_action", "")),
            metadata=_as_dict(payload.get("metadata", {})),
            actions=actions,
            notes=notes,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_family": self.source_family,
            "formula_count": self.formula_count,
            "route_kind": self.route_kind,
            "confidence_policy": self.confidence_policy,
            "requires_manual_review": self.requires_manual_review,
            "provider": self.provider,
            "next_action": self.next_action,
            "metadata": dict(self.metadata),
            "actions": [action.to_dict() for action in self.actions],
            "notes": list(self.notes),
        }


@dataclass(frozen=True, slots=True)
class ExecutionPlan:
    document_id: str
    input_path: str
    detector_version: str
    formula_count: int
    recommended_sequence: tuple[str, ...]
    steps: tuple[ExecutionStep, ...]

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExecutionPlan":
        steps = tuple(
            ExecutionStep.from_dict(item)
            for item in payload.get("steps", [])
            if isinstance(item, dict)
        )
        return cls(
            document_id=str(payload.get("document_id", "")),
            input_path=str(payload.get("input_path", "")),
            detector_version=str(payload.get("detector_version", "")),
            formula_count=int(payload.get("formula_count", 0)),
            recommended_sequence=tuple(str(item) for item in payload.get("recommended_sequence", [])),
            steps=steps,
        )

    @property
    def manual_review_required(self) -> bool:
        return any(step.requires_manual_review for step in self.steps)

    def to_dict(self) -> dict[str, Any]:
        return {
            "document_id": self.document_id,
            "input_path": self.input_path,
            "detector_version": self.detector_version,
            "formula_count": self.formula_count,
            "recommended_sequence": list(self.recommended_sequence),
            "manual_review_required": self.manual_review_required,
            "steps": [step.to_dict() for step in self.steps],
        }
