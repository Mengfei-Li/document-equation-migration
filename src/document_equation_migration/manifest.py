from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import UTC, datetime
from enum import Enum
from pathlib import Path
from typing import Any

from .container_scan import ContainerScanResult
from .source_taxonomy import SourceFamily, SourceRole, normalize_source_family, normalize_source_role


def _serialize(value: Any) -> Any:
    if isinstance(value, Enum):
        return value.value
    if dataclass_isinstance(value):
        data = asdict(value)
        return {key: _serialize(item) for key, item in data.items()}
    if isinstance(value, dict):
        return {key: _serialize(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_serialize(item) for item in value]
    return value

def dataclass_isinstance(value: Any) -> bool:
    return hasattr(value, "__dataclass_fields__")


@dataclass(slots=True)
class ProvenanceRecord:
    prog_id_raw: str = ""
    field_code_raw: str = ""
    ole_stream_names: list[str] = field(default_factory=list)
    raw_payload_status: str = "unknown"
    raw_payload_sha256: str = ""
    transform_chain: list[str] = field(default_factory=list)
    generator_raw: str = ""
    generator_id: str = ""
    evidence_sources: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ValidationRecord:
    word_validation_status: str = "unvalidated"
    pdf_export_status: str = "unvalidated"
    visual_compare_status: str = "unvalidated"
    validation_report_ref: str = ""


@dataclass(slots=True)
class FormulaRecord:
    formula_id: str
    source_family: SourceFamily
    source_role: SourceRole
    doc_part_path: str
    story_type: str
    storage_kind: str
    relationship_id: str = ""
    embedding_target: str = ""
    preview_target: str = ""
    paragraph_index: int | None = None
    run_index: int | None = None
    object_sequence: int | None = None
    canonical_mathml_status: str = "unverified"
    omml_status: str = "not-applicable"
    latex_status: str = "not-applicable"
    risk_level: str = "manual-review"
    risk_flags: list[str] = field(default_factory=list)
    failure_mode: str = ""
    confidence: float = 0.0
    provenance: ProvenanceRecord = field(default_factory=ProvenanceRecord)
    source_specific: dict[str, Any] = field(default_factory=dict)
    validation: ValidationRecord = field(default_factory=ValidationRecord)

    def __post_init__(self) -> None:
        self.source_family = normalize_source_family(self.source_family)
        self.source_role = normalize_source_role(self.source_role)

    def to_dict(self) -> dict[str, Any]:
        return {
            "formula_id": self.formula_id,
            "source_family": self.source_family.value,
            "source_role": self.source_role.value,
            "doc_part_path": self.doc_part_path,
            "story_type": self.story_type,
            "storage_kind": self.storage_kind,
            "relationship_id": self.relationship_id,
            "embedding_target": self.embedding_target,
            "preview_target": self.preview_target,
            "paragraph_index": self.paragraph_index,
            "run_index": self.run_index,
            "object_sequence": self.object_sequence,
            "canonical_mathml_status": self.canonical_mathml_status,
            "omml_status": self.omml_status,
            "latex_status": self.latex_status,
            "risk_level": self.risk_level,
            "risk_flags": self.risk_flags,
            "failure_mode": self.failure_mode,
            "confidence": self.confidence,
            "provenance": _serialize(self.provenance),
            "source_specific": _serialize(self.source_specific),
            "validation": _serialize(self.validation),
        }


@dataclass(slots=True)
class DocumentRecord:
    document_id: str
    input_path: str
    input_sha256: str
    container_format: str
    detector_version: str
    source_counts: dict[str, int] = field(default_factory=dict)
    generated_at: str = field(default_factory=lambda: datetime.now(UTC).isoformat())
    notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "document_id": self.document_id,
            "input_path": self.input_path,
            "input_sha256": self.input_sha256,
            "container_format": self.container_format,
            "detector_version": self.detector_version,
            "source_counts": self.source_counts,
            "generated_at": self.generated_at,
            "notes": self.notes,
        }


@dataclass(slots=True)
class Manifest:
    document: DocumentRecord
    formulas: list[FormulaRecord] = field(default_factory=list)

    @classmethod
    def from_scan(
        cls,
        scan_result: ContainerScanResult,
        detector_version: str,
        formulas: list[FormulaRecord] | None = None,
        notes: list[str] | None = None,
    ) -> "Manifest":
        scan_notes = list(notes or [])
        if not formulas:
            scan_notes.append("No registered detectors produced formula records yet.")
        scan_notes.append(
            f"container-scan: entries={scan_result.entry_count}, "
            f"story-parts={len(scan_result.story_parts)}, "
            f"embeddings={len(scan_result.embedding_targets)}"
        )
        manifest = cls(
            document=DocumentRecord(
                document_id=Path(scan_result.input_path).stem,
                input_path=scan_result.input_path,
                input_sha256=scan_result.input_sha256,
                container_format=scan_result.container_format,
                detector_version=detector_version,
                notes=scan_notes,
            ),
            formulas=list(formulas or []),
        )
        manifest.update_source_counts()
        return manifest

    def update_source_counts(self) -> None:
        counts: dict[str, int] = {}
        for formula in self.formulas:
            key = formula.source_family.value
            counts[key] = counts.get(key, 0) + 1
        self.document.source_counts = counts

    def to_dict(self) -> dict[str, Any]:
        return {
            "document": self.document.to_dict(),
            "formulas": [formula.to_dict() for formula in self.formulas],
        }

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)
