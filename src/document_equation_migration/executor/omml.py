from __future__ import annotations

import hashlib
import json
import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from ..canonical_target import canonical_mathml_contract_for_source_family
from ..execution_plan.model import ExecutionStep
from ..omml_to_mathml import omml_fragment_to_mathml
from .model import ActionExecutionReport, DryRunActionReport, DryRunContext, ExecutionContext


OMML_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"

ET.register_namespace("m", OMML_NAMESPACE)


def _output_root(context: DryRunContext) -> Path:
    return Path(context.output_dir_hint) / "omml-native"


def _execution_output_root(context: ExecutionContext) -> Path:
    return Path(context.output_dir) / "omml-native"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _safe_stem(part_path: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9_.-]+", "-", part_path).strip("-")
    return stem or "part"


def _candidate_story_parts(names: list[str]) -> list[str]:
    parts = []
    for name in sorted(names):
        if not name.startswith("word/") or not name.endswith(".xml") or "/_rels/" in name:
            continue
        if (
            name == "word/document.xml"
            or name in {"word/comments.xml", "word/endnotes.xml", "word/footnotes.xml"}
            or name.startswith("word/header")
            or name.startswith("word/footer")
        ):
            parts.append(name)
    return parts


def _write_json(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _read_json(path: Path) -> dict[str, object] | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _mathml_property_signals(root: ET.Element) -> dict[str, object]:
    nodes = list(root.iter())
    return {
        "root_attributes": dict(root.attrib),
        "root_display": root.attrib.get("display", ""),
        "mathml_attribute_count": sum(len(node.attrib) for node in nodes),
        "has_semantics": any(_local_name(node.tag) == "semantics" for node in nodes),
        "has_annotation": any(_local_name(node.tag) == "annotation" for node in nodes),
        "has_mfrac_linethickness": any(
            "linethickness" in node.attrib
            for node in nodes
            if _local_name(node.tag) == "mfrac"
        ),
        "has_mfrac_bevelled": any(
            node.attrib.get("bevelled") == "true"
            for node in nodes
            if _local_name(node.tag) == "mfrac"
        ),
        "has_mfenced_separators": any(
            "separators" in node.attrib
            for node in nodes
            if _local_name(node.tag) == "mfenced"
        ),
        "has_movablelimits": any("movablelimits" in node.attrib for node in nodes),
        "has_mathvariant": any("mathvariant" in node.attrib for node in nodes),
        "has_accent": any(node.attrib.get("accent") == "true" for node in nodes),
        "has_accentunder": any(node.attrib.get("accentunder") == "true" for node in nodes),
    }


def _property_summary(items: list[dict[str, object]]) -> dict[str, object]:
    property_keys = (
        "has_semantics",
        "has_annotation",
        "has_mfrac_linethickness",
        "has_mfrac_bevelled",
        "has_mfenced_separators",
        "has_movablelimits",
        "has_mathvariant",
        "has_accent",
        "has_accentunder",
    )
    signals = [item.get("property_signals", {}) for item in items]
    root_display_values = sorted(
        {
            str(signal.get("root_display"))
            for signal in signals
            if isinstance(signal, dict) and signal.get("root_display")
        }
    )
    return {
        "mathml_attribute_count": sum(
            int(signal.get("mathml_attribute_count", 0))
            for signal in signals
            if isinstance(signal, dict)
        ),
        "root_display_values": root_display_values,
        "signal_counts": {
            key: sum(1 for signal in signals if isinstance(signal, dict) and signal.get(key))
            for key in property_keys
        },
    }


def _read_manifest_summary(output_root: Path) -> dict[str, object]:
    manifest_path = output_root / "manifest.json"
    if not manifest_path.exists():
        return {
            "manifest_path": str(manifest_path),
            "manifest_present": False,
            "formula_count": None,
        }

    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {
            "manifest_path": str(manifest_path),
            "manifest_present": False,
            "formula_count": None,
        }

    return {
        "manifest_path": str(manifest_path),
        "manifest_present": True,
        "formula_count": manifest.get("formula_count"),
    }


def _read_normalization_summary(output_root: Path) -> dict[str, object]:
    summary_path = output_root / "normalization-summary.json"
    if not summary_path.exists():
        return {
            "normalization_summary_path": str(summary_path),
            "normalization_summary_present": False,
            "normalized_count": None,
        }

    try:
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {
            "normalization_summary_path": str(summary_path),
            "normalization_summary_present": False,
            "normalized_count": None,
        }

    return {
        "normalization_summary_path": str(summary_path),
        "normalization_summary_present": True,
        "normalized_count": summary.get("normalized_count"),
        "strategy": summary.get("strategy"),
    }


def _read_canonicalization_summary(output_root: Path) -> dict[str, object]:
    summary_path = output_root / "canonicalization-summary.json"
    if not summary_path.exists():
        return {
            "canonicalization_summary_path": str(summary_path),
            "canonicalization_summary_present": False,
            "canonical_mathml_count": None,
        }

    try:
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {
            "canonicalization_summary_path": str(summary_path),
            "canonicalization_summary_present": False,
            "canonical_mathml_count": None,
        }

    return {
        "canonicalization_summary_path": str(summary_path),
        "canonicalization_summary_present": True,
        "expected_formula_count": summary.get("expected_formula_count"),
        "canonical_mathml_count": summary.get("canonical_mathml_count"),
        "unsupported_fragment_count": summary.get("unsupported_fragment_count"),
        "formula_count_parity": summary.get("formula_count_parity"),
        "strategy": summary.get("strategy"),
        "canonical_mathml_dir": summary.get("canonical_mathml_dir"),
        "property_summary": summary.get("property_summary"),
        "source_to_canonical_provenance_count": len(
            summary.get("source_to_canonical_provenance", [])
            if isinstance(summary.get("source_to_canonical_provenance"), list)
            else []
        ),
    }


def _read_package_metadata_summary(output_root: Path) -> dict[str, object]:
    metadata_path = output_root / "package" / "execution-metadata.json"
    if not metadata_path.exists():
        return {
            "package_metadata_path": str(metadata_path),
            "package_metadata_present": False,
            "source_family": None,
            "provider": None,
            "next_action": None,
            "validation_target_docx": None,
        }

    try:
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {
            "package_metadata_path": str(metadata_path),
            "package_metadata_present": False,
            "source_family": None,
            "provider": None,
            "next_action": None,
            "validation_target_docx": None,
        }

    return {
        "package_metadata_path": str(metadata_path),
        "package_metadata_present": True,
        "source_family": metadata.get("source_family"),
        "provider": metadata.get("provider"),
        "next_action": metadata.get("next_action"),
        "validation_target_docx": metadata.get("validation_target_docx"),
    }


def _read_validation_target_summary(output_root: Path) -> dict[str, object]:
    validation_target_path = output_root / "package" / "validation-target.docx"
    return {
        "validation_target_docx": str(validation_target_path),
        "validation_target_present": validation_target_path.exists(),
    }


def _read_validation_plan_summary(output_root: Path) -> dict[str, object]:
    validation_plan_path = output_root / "validation" / "validation-plan.json"
    if not validation_plan_path.exists():
        return {
            "validation_plan_path": str(validation_plan_path),
            "validation_plan_present": False,
            "status": None,
            "review_mode": None,
            "planned_check_count": None,
        }

    try:
        validation_plan = json.loads(validation_plan_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {
            "validation_plan_path": str(validation_plan_path),
            "validation_plan_present": False,
            "status": None,
            "review_mode": None,
            "planned_check_count": None,
        }

    planned_checks = validation_plan.get("planned_checks")
    planned_check_count = len(planned_checks) if isinstance(planned_checks, list) else None
    return {
        "validation_plan_path": str(validation_plan_path),
        "validation_plan_present": True,
        "status": validation_plan.get("status"),
        "review_mode": validation_plan.get("review_mode"),
        "planned_check_count": planned_check_count,
    }


def _write_validation_plan(step: ExecutionStep, context: ExecutionContext, output_root: Path) -> Path:
    validation_dir = output_root / "validation"
    validation_plan_path = validation_dir / "validation-plan.json"
    review_mode = "required" if step.requires_manual_review else "spot-check"
    gate_status = "manual-review-required" if step.requires_manual_review else "pending-external-validation"
    _write_json(
        validation_plan_path,
        {
            "artifact_type": "omml-validation-plan",
            "provider": step.provider,
            "source_family": step.source_family,
            "route_kind": step.route_kind,
            "status": gate_status,
            "review_mode": review_mode,
            "runner": "manual-validation",
            "input_path": context.input_path,
            "execution_plan_path": context.execution_plan_path,
            "output_root": str(output_root),
            "evidence": {
                "manifest": _read_manifest_summary(output_root),
                "normalization_summary_path": str(output_root / "normalization-summary.json"),
                "canonicalization_summary_path": str(output_root / "canonicalization-summary.json"),
                "package_metadata_path": str(output_root / "package" / "execution-metadata.json"),
                "validation_target": _read_validation_target_summary(output_root),
            },
            "planned_checks": [
                {
                    "id": "word-open",
                    "description": "Open the source or packaged DOCX in Microsoft Word.",
                    "status": "not-run",
                    "requires_external_application": True,
                },
                {
                    "id": "word-export-pdf",
                    "description": "Export the Word document to PDF for render evidence.",
                    "status": "not-run",
                    "requires_external_application": True,
                },
                {
                    "id": "render-parity",
                    "description": "Compare formula render positions and visible math content against source evidence.",
                    "status": "not-run",
                    "requires_external_application": True,
                },
                {
                    "id": "formula-count",
                    "description": "Compare detected OMML count with extracted and normalized artifact counts.",
                    "status": "planned",
                    "requires_external_application": False,
                },
            ],
            "non_goals": [
                "This executor does not invoke Microsoft Word.",
                "This executor does not export PDF evidence.",
                "This executor does not claim deliverable render parity.",
            ],
        },
    )
    return validation_plan_path


def _write_validation_evidence(step: ExecutionStep, context: ExecutionContext, output_root: Path) -> Path:
    validation_dir = output_root / "validation"
    validation_evidence_path = validation_dir / "validation-evidence.json"
    review_mode = "required" if step.requires_manual_review else "spot-check"
    gate_status = "manual-review-required" if step.requires_manual_review else "pending-external-validation"
    manifest_summary = _read_manifest_summary(output_root)
    normalization_summary = _read_normalization_summary(output_root)
    canonicalization_summary = _read_canonicalization_summary(output_root)
    canonicalization_payload = _read_json(output_root / "canonicalization-summary.json") or {}
    package_metadata = _read_package_metadata_summary(output_root)
    validation_plan = _read_validation_plan_summary(output_root)
    validation_target = _read_validation_target_summary(output_root)
    formula_count = manifest_summary.get("formula_count")
    normalized_count = normalization_summary.get("normalized_count")
    canonical_count = canonicalization_summary.get("canonical_mathml_count")

    evidence_checks: list[dict[str, object]] = [
        {
            "id": "manifest-present",
            "status": "passed" if manifest_summary["manifest_present"] else "missing",
        },
        {
            "id": "normalization-summary-present",
            "status": "passed" if normalization_summary["normalization_summary_present"] else "missing",
        },
        {
            "id": "canonicalization-summary-present",
            "status": "passed" if canonicalization_summary["canonicalization_summary_present"] else "missing",
        },
        {
            "id": "package-metadata-present",
            "status": "passed" if package_metadata["package_metadata_present"] else "missing",
        },
        {
            "id": "validation-plan-present",
            "status": "passed" if validation_plan["validation_plan_present"] else "missing",
        },
        {
            "id": "validation-target-present",
            "status": "passed" if validation_target["validation_target_present"] else "missing",
            "validation_target_docx": validation_target["validation_target_docx"],
        },
        {
            "id": "word-open",
            "status": "not-run",
            "requires_external_application": True,
        },
        {
            "id": "word-export-pdf",
            "status": "not-run",
            "requires_external_application": True,
        },
        {
            "id": "render-parity",
            "status": "not-run",
            "requires_external_application": True,
        },
    ]
    if isinstance(formula_count, int) and isinstance(normalized_count, int):
        evidence_checks.insert(
            4,
            {
                "id": "formula-count-vs-normalized-count",
                "status": "passed" if formula_count == normalized_count else "mismatch",
                "manifest_formula_count": formula_count,
                "normalized_count": normalized_count,
            },
        )
    if isinstance(formula_count, int) and isinstance(canonical_count, int):
        evidence_checks.insert(
            5,
            {
                "id": "formula-count-vs-canonical-mathml-count",
                "status": "passed" if formula_count == canonical_count else "mismatch",
                "manifest_formula_count": formula_count,
                "canonical_mathml_count": canonical_count,
            },
        )

    _write_json(
        validation_evidence_path,
        {
            "artifact_type": "omml-validation-evidence",
            "provider": step.provider,
            "source_family": step.source_family,
            "route_kind": step.route_kind,
            "status": "evidence-collected",
            "gate_status": gate_status,
            "review_mode": review_mode,
            "input_path": context.input_path,
            "execution_plan_path": context.execution_plan_path,
            "output_root": str(output_root),
            "canonical_target": canonical_mathml_contract_for_source_family(step.source_family).to_dict(),
            "artifacts": {
                "manifest": manifest_summary,
                "normalization_summary": normalization_summary,
                "canonicalization_summary": canonicalization_summary,
                "package_metadata": package_metadata,
                "validation_target": validation_target,
                "validation_plan": validation_plan,
            },
            "canonical_artifact_gate": {
                "formula_count_parity": canonicalization_summary.get("formula_count_parity"),
                "property_summary": canonicalization_summary.get("property_summary"),
                "source_to_canonical_provenance_count": canonicalization_summary.get(
                    "source_to_canonical_provenance_count"
                ),
            },
            "source_to_canonical_provenance": canonicalization_payload.get(
                "source_to_canonical_provenance", []
            ),
            "evidence_checks": evidence_checks,
            "observations": [
                "Manifest, normalization, package and validation-plan artifacts were collected locally.",
                "Canonical MathML artifacts are the structured target for downstream pipelines.",
                "Validation target DOCX is packaged for the shared DOCX deliverability workflow.",
                "This evidence does not claim Microsoft Word or PDF validation completed.",
                "Word and PDF checks remain external/manual validation gates.",
            ],
        },
    )
    return validation_evidence_path


def _package_validation_target(input_path: Path, output_root: Path) -> Path:
    package_root = output_root / "package"
    package_root.mkdir(parents=True, exist_ok=True)
    validation_target_path = package_root / "validation-target.docx"
    shutil.copyfile(input_path, validation_target_path)
    return validation_target_path


def _extract_omml_fragments(input_path: Path, output_root: Path) -> tuple[Path, tuple[Path, ...]]:
    extracted_dir = output_root / "extracted"
    extracted_dir.mkdir(parents=True, exist_ok=True)
    items: list[dict[str, object]] = []
    output_paths: list[Path] = []

    with zipfile.ZipFile(input_path) as zf:
        for part_path in _candidate_story_parts(zf.namelist()):
            try:
                root = ET.fromstring(zf.read(part_path))
            except ET.ParseError:
                continue
            part_index = 0
            for element in root.iter():
                kind = _local_name(element.tag)
                if kind not in {"oMath", "oMathPara"}:
                    continue
                part_index += 1
                item_id = f"omml-{len(items) + 1:04d}"
                output_path = extracted_dir / f"{item_id}-{_safe_stem(part_path)}-{kind}.xml"
                xml_text = ET.tostring(element, encoding="unicode")
                output_path.write_text(f'<?xml version="1.0" encoding="UTF-8"?>\n{xml_text}\n', encoding="utf-8")
                output_paths.append(output_path)
                items.append(
                    {
                        "id": item_id,
                        "part_path": part_path,
                        "part_index": part_index,
                        "kind": kind,
                        "extracted_path": str(output_path),
                    }
                )

    manifest_path = output_root / "manifest.json"
    _write_json(
        manifest_path,
        {
            "input_path": str(input_path),
            "formula_count": len(items),
            "items": items,
        },
    )
    return manifest_path, tuple(output_paths)


def _normalize_omml_fragments(output_root: Path) -> tuple[Path, tuple[Path, ...]]:
    extracted_dir = output_root / "extracted"
    normalized_dir = output_root / "normalized"
    normalized_dir.mkdir(parents=True, exist_ok=True)
    normalized_paths: list[Path] = []

    for source_path in sorted(extracted_dir.glob("*.xml")):
        target_path = normalized_dir / source_path.name
        shutil.copyfile(source_path, target_path)
        normalized_paths.append(target_path)

    summary_path = output_root / "normalization-summary.json"
    _write_json(
        summary_path,
        {
            "strategy": "preserve-native-omml",
            "normalized_count": len(normalized_paths),
            "items": [str(path) for path in normalized_paths],
        },
    )
    return summary_path, tuple(normalized_paths)


def _canonicalize_omml_fragments(output_root: Path) -> tuple[Path, tuple[Path, ...]]:
    normalized_dir = output_root / "normalized"
    canonical_dir = output_root / "canonical-mathml"
    canonical_dir.mkdir(parents=True, exist_ok=True)
    canonical_paths: list[Path] = []
    unsupported_items: list[dict[str, object]] = []
    provenance_items: list[dict[str, object]] = []
    manifest = _read_json(output_root / "manifest.json") or {}
    manifest_items = manifest.get("items", [])
    manifest_by_filename = {
        Path(str(item.get("extracted_path"))).name: item
        for item in manifest_items
        if isinstance(item, dict) and item.get("extracted_path")
    }

    for source_path in sorted(normalized_dir.glob("*.xml")):
        target_path = canonical_dir / source_path.name
        source_text = source_path.read_text(encoding="utf-8")
        mathml_text = omml_fragment_to_mathml(source_text)
        target_text = f'<?xml version="1.0" encoding="UTF-8"?>\n{mathml_text}\n'
        item = manifest_by_filename.get(source_path.name, {})
        formula_id = str(item.get("id") or source_path.stem)
        property_signals: dict[str, object] = {}
        root_tag = ""
        try:
            root = ET.fromstring(mathml_text)
            root_tag = root.tag
            property_signals = _mathml_property_signals(root)
            if _local_name(root.tag) != "math":
                unsupported_items.append(
                    {
                        "formula_id": formula_id,
                        "source_omml_path": str(source_path),
                        "status": "not-mathml-root",
                        "root_tag": root.tag,
                    }
                )
        except ET.ParseError as exc:
            unsupported_items.append(
                {
                    "formula_id": formula_id,
                    "source_omml_path": str(source_path),
                    "status": "xml-parse-error",
                    "error": str(exc),
                }
            )
        if "data-omml-unsupported=" in mathml_text:
            unsupported_items.append(
                {
                    "formula_id": formula_id,
                    "source_omml_path": str(source_path),
                    "status": "unsupported-omml-structure",
                }
            )
        target_path.write_text(
            target_text,
            encoding="utf-8",
        )
        canonical_paths.append(target_path)
        provenance_items.append(
            {
                "formula_id": formula_id,
                "source_omml_path": str(source_path),
                "source_part_path": item.get("part_path"),
                "source_part_index": item.get("part_index"),
                "source_kind": item.get("kind"),
                "canonical_artifact_path": str(target_path),
                "source_sha256": _sha256_text(source_text),
                "canonical_sha256": _sha256_text(target_text),
                "preservation_status": "converted-omml-to-canonical-mathml",
                "root_tag": root_tag,
                "property_signals": property_signals,
            }
        )

    summary_path = output_root / "canonicalization-summary.json"
    expected_count = int(manifest.get("formula_count", len(canonical_paths)) or 0)
    unsupported_count = len(unsupported_items)
    _write_json(
        summary_path,
        {
            "strategy": "internal-basic-omml-to-presentation-mathml",
            "expected_formula_count": expected_count,
            "canonical_mathml_count": len(canonical_paths),
            "unsupported_fragment_count": unsupported_count,
            "formula_count_parity": (
                "passed"
                if expected_count == len(canonical_paths) and unsupported_count == 0
                else "mismatch"
            ),
            "canonical_mathml_dir": str(canonical_dir),
            "source_to_canonical_provenance": provenance_items,
            "property_summary": _property_summary(provenance_items),
            "unsupported_fragments": unsupported_items,
            "items": [str(path) for path in canonical_paths],
        },
    )
    return summary_path, tuple(canonical_paths)


def _package_omml_output(step: ExecutionStep, context: ExecutionContext, output_root: Path) -> Path:
    package_root = output_root / "package"
    package_root.mkdir(parents=True, exist_ok=True)
    metadata_path = package_root / "execution-metadata.json"
    validation_target_path = _package_validation_target(Path(context.input_path), output_root)
    _write_json(
        metadata_path,
        {
            "source_family": step.source_family,
            "provider": step.provider,
            "next_action": step.next_action,
            "input_path": context.input_path,
            "execution_plan_path": context.execution_plan_path,
            "output_root": str(output_root),
            "manifest_path": str(output_root / "manifest.json"),
            "normalization_summary_path": str(output_root / "normalization-summary.json"),
            "canonicalization_summary_path": str(output_root / "canonicalization-summary.json"),
            "canonical_mathml_dir": str(output_root / "canonical-mathml"),
            "validation_target_docx": str(validation_target_path),
            "validation_plan_path": str(output_root / "validation" / "validation-plan.json"),
            "validation_evidence_path": str(output_root / "validation" / "validation-evidence.json"),
        },
    )
    return metadata_path


def _build_extract_report(
    step: ExecutionStep,
    action_id: str,
    description: str,
    context: DryRunContext,
) -> DryRunActionReport:
    output_root = _output_root(context)
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=False,
        supported=True,
        status="ready",
        runner="internal-omml-native",
        argv=(
            "extract",
            "--execution-plan",
            context.execution_plan_path,
            "--output-dir",
            str(output_root),
            "--parts",
            "word/document.xml,word/comments.xml,word/endnotes.xml,word/footnotes.xml",
            "--preserve-native-omml",
        ),
        cwd=context.workspace_root,
        notes=(
            "Native-preserve intake: dry-run assumes OMML is read directly from OOXML parts without conversion.",
            f"Planned extraction output root: {output_root}",
        ),
    )


def _build_normalize_report(
    step: ExecutionStep,
    action_id: str,
    description: str,
    context: DryRunContext,
) -> DryRunActionReport:
    output_root = _output_root(context)
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=False,
        supported=True,
        status="ready",
        runner="internal-omml-native",
        argv=(
            "normalize",
            "--input-dir",
            str(output_root),
            "--profile",
            "word-native-safe",
            "--preserve-semantics",
            "--keep-run-layout",
        ),
        cwd=context.workspace_root,
        notes=(
            "Low-risk path: normalization is scoped to deterministic OMML cleanup, not format translation.",
            "No external script is assumed; this is intended for an internal Python runner phase.",
        ),
    )


def _build_render_check_report(
    step: ExecutionStep,
    action_id: str,
    description: str,
    context: DryRunContext,
) -> DryRunActionReport:
    output_root = _output_root(context)
    review_mode = "required" if step.requires_manual_review else "spot-check"
    status = "review-gated" if step.requires_manual_review else "ready"
    notes = [
        "Render check writes a validation-plan artifact instead of invoking Word or PDF export.",
        f"Review mode for this step: {review_mode}.",
        f"Planned validation plan: {output_root / 'validation' / 'validation-plan.json'}",
        "Validation evidence is synthesized later from manifest, normalization, package and plan artifacts.",
    ]
    if step.requires_manual_review:
        notes.append("Routing requested manual review; treat this action as the verification gate.")
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=step.requires_manual_review,
        supported=True,
        status=status,
        runner="internal-omml-native",
        argv=(
            "render-check",
            "--input-dir",
            str(output_root),
            "--mode",
            "word-parity",
            "--review",
            review_mode,
        ),
        cwd=context.workspace_root,
        notes=tuple(notes),
    )


def _build_canonical_mathml_report(
    step: ExecutionStep,
    action_id: str,
    description: str,
    context: DryRunContext,
) -> DryRunActionReport:
    output_root = _output_root(context)
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=False,
        supported=True,
        status="ready",
        runner="internal-omml-to-mathml",
        argv=(
            "canonicalize",
            "--input-dir",
            str(output_root / "normalized"),
            "--output-dir",
            str(output_root / "canonical-mathml"),
            "--profile",
            "basic-presentation-mathml",
        ),
        cwd=context.workspace_root,
        notes=(
            "Converts normalized OMML fragments into canonical MathML artifacts for downstream pipelines.",
            "Unsupported OMML structures are preserved as review-marked MathML rows rather than silently dropped.",
        ),
    )


def _build_package_report(
    step: ExecutionStep,
    action_id: str,
    description: str,
    context: DryRunContext,
) -> DryRunActionReport:
    output_root = _output_root(context)
    package_root = output_root / "package"
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=False,
        supported=True,
        status="ready",
        runner="internal-omml-native",
        argv=(
            "package-output",
            "--input-dir",
            str(output_root),
            "--package-dir",
            str(package_root),
            "--execution-plan",
            context.execution_plan_path,
            "--label",
            step.next_action or "run-omml-native-pipeline",
        ),
        cwd=context.workspace_root,
        notes=(
            "Package step keeps OMML as the primary deliverable and avoids inventing a converter dependency.",
            f"Planned package output root: {package_root}",
        ),
    )


def _build_unknown_action_report(
    action_id: str,
    description: str,
    context: DryRunContext,
) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=True,
        supported=False,
        status="manual-gate",
        runner="manual",
        cwd=context.workspace_root,
        notes=("Unrecognized OMML dry-run action; manual triage is required before execution binding.",),
    )


def build_omml_dry_run_reports(
    step: ExecutionStep,
    context: DryRunContext,
) -> tuple[DryRunActionReport, ...]:
    if step.source_family != "omml-native":
        raise ValueError(f"OMML dry-run binding cannot handle source_family={step.source_family!r}.")

    builders = {
        "extract-omml": _build_extract_report,
        "normalize-omml": _build_normalize_report,
        "omml-to-canonical-mathml": _build_canonical_mathml_report,
        "render-check": _build_render_check_report,
        "package-omml-output": _build_package_report,
    }

    reports: list[DryRunActionReport] = []
    for action in step.actions:
        builder = builders.get(action.action_id)
        if builder is None:
            reports.append(_build_unknown_action_report(action.action_id, action.description, context))
            continue
        reports.append(builder(step, action.action_id, action.description, context))

    return tuple(reports)


def _execution_report(
    *,
    action_id: str,
    description: str,
    blocking: bool,
    status: str,
    runner: str,
    context: ExecutionContext,
    supported: bool = True,
    output_paths: tuple[Path, ...] = (),
    notes: tuple[str, ...] = (),
) -> ActionExecutionReport:
    return ActionExecutionReport(
        action_id=action_id,
        description=description,
        blocking=blocking,
        supported=supported,
        status=status,
        runner=runner,
        cwd=context.workspace_root,
        output_paths=tuple(str(path) for path in output_paths),
        notes=notes,
    )


def execute_omml_step(
    step: ExecutionStep,
    context: ExecutionContext,
) -> tuple[ActionExecutionReport, ...]:
    if step.source_family != "omml-native":
        raise ValueError(f"OMML executor cannot handle source_family={step.source_family!r}.")

    output_root = _execution_output_root(context)
    input_path = Path(context.input_path)
    reports: list[ActionExecutionReport] = []
    extract_succeeded = False
    normalization_succeeded = False

    for action in step.actions:
        action_id = action.action_id
        if action_id == "extract-omml":
            if not input_path.exists():
                reports.append(
                    _execution_report(
                        action_id=action_id,
                        description=action.description,
                        blocking=True,
                        status="failed",
                        runner="internal-omml-native",
                        context=context,
                        notes=(f"Input document does not exist: {input_path}",),
                    )
                )
                continue
            try:
                manifest_path, extracted_paths = _extract_omml_fragments(input_path, output_root)
            except (OSError, zipfile.BadZipFile) as exc:
                reports.append(
                    _execution_report(
                        action_id=action_id,
                        description=action.description,
                        blocking=True,
                        status="failed",
                        runner="internal-omml-native",
                        context=context,
                        notes=(f"Failed to extract OMML from DOCX: {exc}",),
                    )
                )
                continue
            extract_succeeded = True
            reports.append(
                _execution_report(
                    action_id=action_id,
                    description=action.description,
                    blocking=action.blocking,
                    status="completed",
                    runner="internal-omml-native",
                    context=context,
                    output_paths=(manifest_path, *extracted_paths),
                    notes=(
                        f"Extracted {len(extracted_paths)} OMML fragment(s).",
                        "Native OMML was preserved as XML fragments; no format translation was performed.",
                    ),
                )
            )
            continue

        if action_id == "normalize-omml":
            if not extract_succeeded:
                reports.append(
                    _execution_report(
                        action_id=action_id,
                        description=action.description,
                        blocking=action.blocking,
                        status="skipped-after-failure",
                        runner="internal-omml-native",
                        context=context,
                        notes=("Extraction did not complete, so normalization was skipped.",),
                    )
                )
                continue
            summary_path, normalized_paths = _normalize_omml_fragments(output_root)
            normalization_succeeded = True
            reports.append(
                _execution_report(
                    action_id=action_id,
                    description=action.description,
                    blocking=action.blocking,
                    status="completed",
                    runner="internal-omml-native",
                    context=context,
                    output_paths=(summary_path, *normalized_paths),
                    notes=(
                        "Current normalization is a native-preserving packaging pass.",
                        f"Normalized artifact count: {len(normalized_paths)}.",
                    ),
                )
            )
            continue

        if action_id == "omml-to-canonical-mathml":
            if not normalization_succeeded:
                reports.append(
                    _execution_report(
                        action_id=action_id,
                        description=action.description,
                        blocking=action.blocking,
                        status="skipped-after-failure",
                        runner="internal-omml-to-mathml",
                        context=context,
                        notes=("OMML normalization did not complete, so canonical MathML conversion was skipped.",),
                    )
                )
                continue
            summary_path, canonical_paths = _canonicalize_omml_fragments(output_root)
            reports.append(
                _execution_report(
                    action_id=action_id,
                    description=action.description,
                    blocking=action.blocking,
                    status="completed",
                    runner="internal-omml-to-mathml",
                    context=context,
                    output_paths=(summary_path, *canonical_paths),
                    notes=(
                        "Converted normalized OMML fragments to canonical MathML artifacts.",
                        f"Canonical MathML artifact count: {len(canonical_paths)}.",
                    ),
                )
            )
            continue

        if action_id == "render-check":
            validation_plan_path = _write_validation_plan(step, context, output_root)
            status = "review-gated" if step.requires_manual_review else "skipped"
            reports.append(
                _execution_report(
                    action_id=action_id,
                    description=action.description,
                    blocking=step.requires_manual_review,
                    status=status,
                    runner="manual-validation",
                    context=context,
                    output_paths=(validation_plan_path,),
                    notes=(
                        "Wrote OMML validation plan; automated Word render parity was not executed.",
                        "Use the DOCX/PDF validation workflow before claiming deliverable Word output.",
                    ),
                )
            )
            continue

        if action_id == "package-omml-output":
            if not extract_succeeded:
                reports.append(
                    _execution_report(
                        action_id=action_id,
                        description=action.description,
                        blocking=action.blocking,
                        status="skipped-after-failure",
                        runner="internal-omml-native",
                        context=context,
                        notes=("Extraction did not complete, so packaging was skipped.",),
                    )
                )
                continue
            metadata_path = _package_omml_output(step, context, output_root)
            validation_evidence_path = _write_validation_evidence(step, context, output_root)
            validation_target_path = output_root / "package" / "validation-target.docx"
            reports.append(
                _execution_report(
                    action_id=action_id,
                    description=action.description,
                    blocking=action.blocking,
                    status="completed",
                    runner="internal-omml-native",
                    context=context,
                    output_paths=(metadata_path, validation_target_path, validation_evidence_path),
                    notes=(
                        "Packaged OMML execution metadata, validation target DOCX, and validation evidence for downstream review.",
                    ),
                )
            )
            continue

        reports.append(
            _execution_report(
                action_id=action_id,
                description=action.description,
                blocking=True,
                status="manual-gate",
                runner="manual",
                context=context,
                supported=False,
                notes=("Unrecognized OMML action; manual triage is required.",),
            )
        )

    return tuple(reports)
