from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .execution_plan import build_execution_plan
from .manifest import Manifest
from .source_taxonomy import SourceFamily


@dataclass(frozen=True, slots=True)
class RouteSpec:
    route_kind: str
    priority: int
    next_action: str
    confidence_policy: str
    requires_manual_review: bool = False


ROUTE_SPECS: dict[SourceFamily, RouteSpec] = {
    SourceFamily.MATHTYPE_OLE: RouteSpec(
        route_kind="primary-source-first",
        priority=10,
        next_action="run-mathtype-source-first-pipeline",
        confidence_policy="high",
    ),
    SourceFamily.OMML_NATIVE: RouteSpec(
        route_kind="primary-source-first",
        priority=20,
        next_action="run-omml-native-pipeline",
        confidence_policy="high",
    ),
    SourceFamily.EQUATION_EDITOR_3_OLE: RouteSpec(
        route_kind="primary-candidate",
        priority=30,
        next_action="run-equation3-probe-and-conversion",
        confidence_policy="medium",
        requires_manual_review=True,
    ),
    SourceFamily.ODF_NATIVE: RouteSpec(
        route_kind="primary-source-first",
        priority=40,
        next_action="run-odf-native-pipeline",
        confidence_policy="medium",
    ),
    SourceFamily.AXMATH_OLE: RouteSpec(
        route_kind="export-assisted",
        priority=50,
        next_action="run-axmath-export-assisted-pipeline",
        confidence_policy="medium",
        requires_manual_review=True,
    ),
    SourceFamily.LIBREOFFICE_TRANSFORMED: RouteSpec(
        route_kind="bridge-source",
        priority=60,
        next_action="run-libreoffice-bridge-review-pipeline",
        confidence_policy="low",
        requires_manual_review=True,
    ),
    SourceFamily.GRAPHIC_FALLBACK: RouteSpec(
        route_kind="fallback",
        priority=70,
        next_action="manual-review-or-ocr-fallback",
        confidence_policy="low",
        requires_manual_review=True,
    ),
    SourceFamily.UNKNOWN_OLE: RouteSpec(
        route_kind="manual-classification",
        priority=80,
        next_action="manual-classification-required",
        confidence_policy="low",
        requires_manual_review=True,
    ),
}


def build_routing_report(manifest: Manifest) -> dict[str, Any]:
    source_counts = dict(manifest.document.source_counts)
    route_plan: list[dict[str, Any]] = []
    manual_review_reasons: list[str] = []

    for source_family_text, formula_count in source_counts.items():
        source_family = SourceFamily(source_family_text)
        spec = ROUTE_SPECS[source_family]
        route_plan.append(
            {
                "source_family": source_family.value,
                "formula_count": formula_count,
                "route_kind": spec.route_kind,
                "priority": spec.priority,
                "next_action": spec.next_action,
                "confidence_policy": spec.confidence_policy,
                "requires_manual_review": spec.requires_manual_review,
            }
        )
        if spec.requires_manual_review:
            manual_review_reasons.append(source_family.value)

    route_plan.sort(key=lambda item: (item["priority"], item["source_family"]))
    total_formula_count = sum(source_counts.values())

    return {
        "document_id": manifest.document.document_id,
        "input_path": manifest.document.input_path,
        "detector_version": manifest.document.detector_version,
        "formula_count": total_formula_count,
        "source_counts": source_counts,
        "recommended_sequence": [item["source_family"] for item in route_plan],
        "route_plan": route_plan,
        "manual_review_required": bool(manual_review_reasons),
        "manual_review_reasons": manual_review_reasons,
    }


def build_execution_plan_report(manifest: Manifest) -> dict[str, Any]:
    routing_report = build_routing_report(manifest)
    return build_execution_plan(routing_report).to_dict()
