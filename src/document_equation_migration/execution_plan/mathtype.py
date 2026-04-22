from __future__ import annotations

from .base import RouteEntry
from .model import ExecutionAction, ExecutionStep

_MATHTYPE_SOURCE_FAMILY = "mathtype-ole"
_DEFAULT_ROUTE_KIND = "primary-source-first"
_DEFAULT_NEXT_ACTION = "run-mathtype-source-first-pipeline"
_DEFAULT_CONFIDENCE_POLICY = "high"
_DEFAULT_LAYOUT_FACTOR = 1.01375


def _as_string(value: object, *, default: str = "") -> str:
    if value is None:
        return default
    return str(value)


def _as_int(value: object, *, default: int = 0) -> int:
    if isinstance(value, bool):
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return default
        try:
            return int(stripped)
        except ValueError:
            return default
    return default


def _as_bool(value: object, *, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes"}:
            return True
        if lowered in {"0", "false", "no"}:
            return False
    return default


def _as_float(value: object, *, default: float = 0.0) -> float:
    if isinstance(value, bool):
        return default
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return default
        try:
            return float(stripped)
        except ValueError:
            return default
    return default


def _as_mapping(value: object) -> dict[str, object]:
    if isinstance(value, dict):
        return {str(key): item for key, item in value.items()}
    return {}


def _normalize_experimental_options(options: dict[str, object]) -> dict[str, object]:
    normalized = dict(options)
    preserve_layout = _as_bool(normalized.get("preserve_mathtype_layout"), default=False)
    if "preserve_mathtype_layout" in normalized:
        normalized["preserve_mathtype_layout"] = preserve_layout
    if preserve_layout or "mathtype_layout_factor" in normalized:
        normalized["mathtype_layout_factor"] = _as_float(
            normalized.get("mathtype_layout_factor"),
            default=_DEFAULT_LAYOUT_FACTOR,
        )
    if "resume_mathtype_pipeline" in normalized:
        normalized["resume_mathtype_pipeline"] = _as_bool(
            normalized.get("resume_mathtype_pipeline"),
            default=False,
        )
    if "mathtype_start_index" in normalized:
        normalized["mathtype_start_index"] = max(
            0,
            _as_int(normalized.get("mathtype_start_index"), default=0),
        )
    if "mathtype_end_index" in normalized:
        normalized["mathtype_end_index"] = max(
            0,
            _as_int(normalized.get("mathtype_end_index"), default=0),
        )
    return normalized


def _build_step_metadata(route_entry: RouteEntry) -> dict[str, object]:
    metadata = _as_mapping(route_entry.get("metadata"))
    existing_experimental_options = _as_mapping(metadata.get("experimental_options"))
    if existing_experimental_options:
        metadata["experimental_options"] = _normalize_experimental_options(existing_experimental_options)

    route_experimental_options = _as_mapping(route_entry.get("experimental_options"))
    if route_experimental_options:
        metadata["experimental_options"] = _normalize_experimental_options(route_experimental_options)

    return metadata


def build_mathtype_execution_step(route_entry: RouteEntry) -> ExecutionStep:
    source_family = _as_string(route_entry.get("source_family"), default=_MATHTYPE_SOURCE_FAMILY)
    if source_family != _MATHTYPE_SOURCE_FAMILY:
        raise ValueError(
            f"MathType execution provider expects source_family={_MATHTYPE_SOURCE_FAMILY!r}, "
            f"got {source_family!r}."
        )

    formula_count = _as_int(route_entry.get("formula_count"))
    route_kind = _as_string(route_entry.get("route_kind"), default=_DEFAULT_ROUTE_KIND)
    confidence_policy = _as_string(route_entry.get("confidence_policy"), default=_DEFAULT_CONFIDENCE_POLICY)
    requires_manual_review = _as_bool(route_entry.get("requires_manual_review"), default=False)
    next_action = _as_string(route_entry.get("next_action"), default=_DEFAULT_NEXT_ACTION)
    metadata = _build_step_metadata(route_entry)

    return ExecutionStep(
        source_family=_MATHTYPE_SOURCE_FAMILY,
        formula_count=formula_count,
        route_kind=route_kind,
        confidence_policy=confidence_policy,
        requires_manual_review=requires_manual_review,
        provider="mathtype",
        next_action=next_action,
        metadata=metadata,
        actions=(
            ExecutionAction(
                action_id="extract-equation-native",
                description="Extract Equation Native payload from MathType OLE containers.",
                blocking=True,
                metadata={"input": "ole-object", "output": "equation-native"},
            ),
            ExecutionAction(
                action_id="mtef-to-mathml",
                description="Decode MTEF from Equation Native and convert it to MathML.",
                blocking=True,
                metadata={"input": "equation-native+mtef", "output": "mathml"},
            ),
            ExecutionAction(
                action_id="normalize-mathml",
                description="Normalize MathML structure for deterministic downstream conversion.",
                blocking=True,
                metadata={"input": "mathml", "output": "normalized-mathml"},
            ),
            ExecutionAction(
                action_id="mathml-to-omml",
                description="Convert normalized MathML to OMML for Word-native equations.",
                blocking=True,
                metadata={"input": "normalized-mathml", "output": "omml"},
            ),
            ExecutionAction(
                action_id="replace-ole-with-omml",
                description="Replace original MathType OLE instances with OMML formulas in DOCX.",
                blocking=True,
                metadata={"input": "docx+omml", "output": "docx-with-omml"},
            ),
            ExecutionAction(
                action_id="validate-word-output",
                description="Open in Word-compatible flow, verify rendering, and confirm export readiness.",
                blocking=False,
                metadata={"input": "docx-with-omml", "checks": ["open", "render", "export-pdf"]},
            ),
        ),
        notes=(
            "Primary-source-first: preserve native MathType payload semantics before any fallback handling.",
            "The route prioritizes Equation Native/MTEF extraction and canonical MathML normalization prior to OMML replacement.",
        ),
    )
