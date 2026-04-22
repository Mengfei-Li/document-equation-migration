from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from ..container_scan import ContainerScanResult
from ..manifest import FormulaRecord, ProvenanceRecord, ValidationRecord
from ..source_taxonomy import SourceFamily, normalize_source_family


@dataclass(slots=True)
class DetectorContext:
    scan_result: ContainerScanResult
    detector_version: str


class FormulaDetector(ABC):
    source_family: SourceFamily
    name: str
    priority: int = 100

    @abstractmethod
    def detect(self, context: DetectorContext) -> list[FormulaRecord]:
        raise NotImplementedError

    def supports(self, scan_result: ContainerScanResult) -> bool:
        return bool(scan_result.story_parts or scan_result.embedding_targets or scan_result.object_parts)


FORMULA_RECORD_KEYS = {
    "formula_id",
    "source_family",
    "source_role",
    "doc_part_path",
    "story_type",
    "storage_kind",
    "relationship_id",
    "embedding_target",
    "preview_target",
    "paragraph_index",
    "run_index",
    "object_sequence",
    "canonical_mathml_status",
    "omml_status",
    "latex_status",
    "risk_level",
    "risk_flags",
    "failure_mode",
    "confidence",
}


def formula_record_from_mapping(
    mapping: dict[str, Any],
    default_source_family: SourceFamily,
) -> FormulaRecord:
    data = dict(mapping)
    source_specific = dict(data.pop("source_specific", {}))
    provenance = ProvenanceRecord(**dict(data.pop("provenance", {})))
    validation = ValidationRecord(**dict(data.pop("validation", {})))
    for key in list(data):
        if key not in FORMULA_RECORD_KEYS:
            source_specific[key] = data.pop(key)
    return FormulaRecord(
        formula_id=data.pop("formula_id"),
        source_family=data.pop("source_family", default_source_family),
        source_role=data.pop("source_role", "native-source"),
        doc_part_path=data.pop("doc_part_path", ""),
        story_type=data.pop("story_type", "other"),
        storage_kind=data.pop("storage_kind", "unknown"),
        relationship_id=data.pop("relationship_id", "") or "",
        embedding_target=data.pop("embedding_target", "") or "",
        preview_target=data.pop("preview_target", "") or "",
        paragraph_index=data.pop("paragraph_index", None),
        run_index=data.pop("run_index", None),
        object_sequence=data.pop("object_sequence", None),
        canonical_mathml_status=data.pop("canonical_mathml_status", "unverified"),
        omml_status=data.pop("omml_status", "not-applicable"),
        latex_status=data.pop("latex_status", "not-applicable"),
        risk_level=data.pop("risk_level", "manual-review"),
        risk_flags=data.pop("risk_flags", []),
        failure_mode=data.pop("failure_mode", "") or "",
        confidence=float(data.pop("confidence", 0.0)),
        provenance=provenance,
        validation=validation,
        source_specific=source_specific,
    )


class FunctionDetector(FormulaDetector):
    def __init__(
        self,
        *,
        source_family: SourceFamily | str,
        name: str,
        handler: Callable[[str | Path], dict[str, Any]],
        priority: int = 100,
    ) -> None:
        self.source_family = normalize_source_family(source_family)
        self.name = name
        self._handler = handler
        self.priority = priority

    def detect(self, context: DetectorContext) -> list[FormulaRecord]:
        payload = self._handler(context.scan_result.input_path)
        if isinstance(payload, list):
            formulas = payload
        elif isinstance(payload, dict):
            formulas = payload.get("formulas", [])
        else:
            raise TypeError(f"Unsupported detector payload type: {type(payload)!r}")
        return [
            item
            if isinstance(item, FormulaRecord)
            else formula_record_from_mapping(item, self.source_family)
            for item in formulas
        ]
