from __future__ import annotations

import hashlib
import json
import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from ..canonical_target import canonical_mathml_contract_for_source_family
from ..execution_plan.model import ExecutionAction, ExecutionStep
from .model import ActionExecutionReport, DryRunActionReport, DryRunContext, ExecutionContext


MATHML_NAMESPACE = "http://www.w3.org/1998/Math/MathML"

ET.register_namespace("math", MATHML_NAMESPACE)

_INPUT_PLACEHOLDER = "<input-odf-from-plan>"
_NATIVE_SOURCE = "odf-native"
_BRIDGE_SOURCE = "libreoffice-transformed"


def _qname(namespace: str, local_name: str) -> str:
    return f"{{{namespace}}}{local_name}"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _resolve_under(base: str, path_text: str) -> Path:
    path = Path(path_text)
    if path.is_absolute():
        return path
    return Path(base) / path


def _dry_output_root(context: DryRunContext, source_family: str) -> Path:
    return _resolve_under(context.workspace_root, context.output_dir_hint) / source_family


def _execution_output_root(context: ExecutionContext, source_family: str) -> Path:
    return Path(context.output_dir) / source_family


def _load_input_path(context: DryRunContext) -> str:
    plan_path_text = context.execution_plan_path.strip()
    if not plan_path_text:
        return _INPUT_PLACEHOLDER

    plan_path = _resolve_under(context.workspace_root, plan_path_text)
    try:
        payload = json.loads(plan_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return _INPUT_PLACEHOLDER

    input_path_text = str(payload.get("input_path", "")).strip()
    if not input_path_text:
        return _INPUT_PLACEHOLDER

    input_path = Path(input_path_text)
    if input_path.is_absolute():
        return str(input_path)
    return str((plan_path.parent / input_path).resolve())


def _safe_stem(part_path: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9_.-]+", "-", part_path).strip("-")
    return stem or "content"


def _write_json(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _read_json(path: Path) -> dict[str, object] | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _candidate_content_members(names: list[str]) -> list[str]:
    return sorted(name for name in names if name == "content.xml" or name.endswith("/content.xml"))


def _iter_content_roots(input_path: Path) -> tuple[tuple[str, ET.Element], ...]:
    if zipfile.is_zipfile(input_path):
        roots: list[tuple[str, ET.Element]] = []
        with zipfile.ZipFile(input_path) as zf:
            for member_name in _candidate_content_members(zf.namelist()):
                roots.append((member_name, ET.fromstring(zf.read(member_name))))
        return tuple(roots)

    return (("content.xml", ET.fromstring(input_path.read_bytes())),)


def _iter_math_nodes(root: ET.Element) -> tuple[ET.Element, ...]:
    math_tag = _qname(MATHML_NAMESPACE, "math")
    return tuple(element for element in root.iter() if element.tag == math_tag)


def _mathml_property_signals(root: ET.Element) -> dict[str, object]:
    nodes = list(root.iter())
    root_display = root.attrib.get("display", "")
    return {
        "root_attributes": dict(root.attrib),
        "root_display": root_display,
        "mathml_attribute_count": sum(len(node.attrib) for node in nodes),
        "has_semantics": any(_local_name(node.tag) == "semantics" for node in nodes),
        "has_annotation": any(_local_name(node.tag) == "annotation" for node in nodes),
        "has_mfrac_linethickness": any("linethickness" in node.attrib for node in nodes if _local_name(node.tag) == "mfrac"),
        "has_mfrac_bevelled": any(node.attrib.get("bevelled") == "true" for node in nodes if _local_name(node.tag) == "mfrac"),
        "has_mfenced_separators": any("separators" in node.attrib for node in nodes if _local_name(node.tag) == "mfenced"),
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
            int(signal.get("mathml_attribute_count", 0)) for signal in signals if isinstance(signal, dict)
        ),
        "root_display_values": root_display_values,
        "signal_counts": {
            key: sum(1 for signal in signals if isinstance(signal, dict) and signal.get(key))
            for key in property_keys
        },
    }


def _storage_kind(member_path: str, root: ET.Element, math_node: ET.Element) -> str:
    if member_path != "content.xml":
        return "odf-draw-object-subdocument"
    if root is math_node:
        return "odf-formula-root"
    return "odf-content-inline-mathml"


def _extract_odf_mathml(input_path: Path, output_root: Path) -> tuple[Path, tuple[Path, ...]]:
    extracted_dir = output_root / "extracted"
    extracted_dir.mkdir(parents=True, exist_ok=True)
    formulas: list[dict[str, object]] = []
    output_paths: list[Path] = []

    for member_path, root in _iter_content_roots(input_path):
        member_index = 0
        for math_node in _iter_math_nodes(root):
            member_index += 1
            formula_id = f"odf-native-{len(formulas) + 1:04d}"
            output_path = extracted_dir / f"{formula_id}-{_safe_stem(member_path)}.xml"
            xml_text = ET.tostring(math_node, encoding="unicode")
            output_text = f'<?xml version="1.0" encoding="UTF-8"?>\n{xml_text}\n'
            output_path.write_text(output_text, encoding="utf-8")
            output_paths.append(output_path)
            formulas.append(
                {
                    "formula_id": formula_id,
                    "source_family": _NATIVE_SOURCE,
                    "source_role": "native-source",
                    "doc_part_path": member_path,
                    "story_type": "main",
                    "storage_kind": _storage_kind(member_path, root, math_node),
                    "object_sequence": len(formulas) + 1,
                    "part_sequence": member_index,
                    "canonical_mathml_status": "available",
                    "omml_status": "not-applicable",
                    "latex_status": "not-applicable",
                    "risk_level": "low",
                    "risk_flags": [],
                    "failure_mode": "",
                    "confidence": 0.99,
                    "artifact_path": str(output_path),
                    "provenance": {
                        "raw_payload_sha256": _sha256_text(xml_text),
                        "transform_chain": [],
                        "evidence_sources": [str(input_path), member_path],
                    },
                    "source_specific": {
                        "mathml": {
                            "root_attributes": dict(math_node.attrib),
                            "property_signals": _mathml_property_signals(math_node),
                        },
                    },
                }
            )

    manifest_path = output_root / "manifest.json"
    _write_json(
        manifest_path,
        {
            "input_path": str(input_path),
            "source_family": _NATIVE_SOURCE,
            "source_role": "native-source",
            "formula_count": len(formulas),
            "source_counts": {_NATIVE_SOURCE: len(formulas)},
            "formulas": formulas,
        },
    )
    return manifest_path, tuple(output_paths)


def _canonicalize_mathml(output_root: Path) -> tuple[Path, tuple[Path, ...]]:
    extracted_dir = output_root / "extracted"
    canonical_dir = output_root / "canonical-mathml"
    canonical_dir.mkdir(parents=True, exist_ok=True)
    canonical_paths: list[Path] = []
    unsupported_items: list[dict[str, object]] = []
    provenance_items: list[dict[str, object]] = []
    manifest = _read_json(output_root / "manifest.json") or {}
    formulas = manifest.get("formulas", [])
    formula_by_artifact = {
        str(item.get("artifact_path")): item
        for item in formulas
        if isinstance(item, dict) and item.get("artifact_path")
    }

    for source_path in sorted(extracted_dir.glob("*.xml")):
        target_path = canonical_dir / source_path.name
        shutil.copyfile(source_path, target_path)
        canonical_paths.append(target_path)
        formula = formula_by_artifact.get(str(source_path), {})
        source_text = source_path.read_text(encoding="utf-8")
        canonical_text = target_path.read_text(encoding="utf-8")
        property_signals: dict[str, object] = {}
        status = "accepted"
        try:
            root = ET.fromstring(source_text)
            property_signals = _mathml_property_signals(root)
            if root.tag != _qname(MATHML_NAMESPACE, "math"):
                status = "not-mathml-root"
        except ET.ParseError as exc:
            status = "xml-parse-error"
            unsupported_items.append(
                {
                    "source_mathml_path": str(source_path),
                    "status": status,
                    "error": str(exc),
                }
            )
        if status != "accepted" and status != "xml-parse-error":
            unsupported_items.append(
                {
                    "source_mathml_path": str(source_path),
                    "status": status,
                }
            )
        provenance_items.append(
            {
                "formula_id": formula.get("formula_id", source_path.stem),
                "source_mathml_path": str(source_path),
                "canonical_artifact_path": str(target_path),
                "source_sha256": _sha256_text(source_text),
                "canonical_sha256": _sha256_text(canonical_text),
                "preservation_status": "byte-identical-after-extraction" if source_text == canonical_text else "changed",
                "property_signals": property_signals,
            }
        )

    summary_path = output_root / "canonicalization-summary.json"
    expected_count = int(manifest.get("formula_count", len(canonical_paths)) or 0)
    unsupported_count = len(unsupported_items)
    _write_json(
        summary_path,
        {
            "strategy": "preserve-existing-odf-mathml",
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


def _write_validation_evidence(step: ExecutionStep, context: ExecutionContext, output_root: Path) -> Path:
    evidence_path = output_root / "validation-evidence.json"
    manifest_path = output_root / "manifest.json"
    canonical_summary_path = output_root / "canonicalization-summary.json"
    manifest = _read_json(manifest_path) or {}
    canonical_summary = _read_json(canonical_summary_path) or {}
    canonical_dir = output_root / "canonical-mathml"
    _write_json(
        evidence_path,
        {
            "artifact_type": "odf-validation-evidence",
            "source_family": step.source_family,
            "canonical_target": canonical_mathml_contract_for_source_family(step.source_family).to_dict(),
            "provider": step.provider,
            "route_kind": step.route_kind,
            "next_action": step.next_action,
            "input_path": context.input_path,
            "execution_plan_path": context.execution_plan_path,
            "output_root": str(output_root),
            "status": "evidence-recorded",
            "runner": "internal-odf-native",
            "manifest": {
                "manifest_path": str(manifest_path),
                "present": manifest_path.exists(),
                "formula_count": manifest.get("formula_count"),
                "source_family": manifest.get("source_family"),
            },
            "canonicalization": {
                "summary_path": str(canonical_summary_path),
                "present": canonical_summary_path.exists(),
                "canonical_mathml_count": canonical_summary.get("canonical_mathml_count"),
                "unsupported_fragment_count": canonical_summary.get("unsupported_fragment_count"),
                "formula_count_parity": canonical_summary.get("formula_count_parity"),
                "property_summary": canonical_summary.get("property_summary"),
                "canonical_mathml_dir": str(canonical_dir),
            },
            "source_to_canonical_provenance": canonical_summary.get("source_to_canonical_provenance", []),
            "evidence_paths": [
                str(manifest_path),
                str(canonical_summary_path),
                str(canonical_dir),
            ],
            "next_ready_condition": "Downstream target-format delivery still requires a later conversion binding.",
            "reason": "This provider slice extracts native ODF MathML and records validation evidence, but does not claim target-format delivery.",
        },
    )
    return evidence_path


def _write_blocker_record(step: ExecutionStep, context: ExecutionContext) -> Path:
    output_root = _execution_output_root(context, _BRIDGE_SOURCE)
    blocker_path = output_root / "blocker-record.json"
    _write_json(
        blocker_path,
        {
            "artifact_type": "odf-blocker-record",
            "source_family": step.source_family,
            "canonical_target": canonical_mathml_contract_for_source_family(step.source_family).to_dict(),
            "provider": step.provider,
            "route_kind": step.route_kind,
            "next_action": step.next_action,
            "input_path": context.input_path,
            "output_root": str(output_root),
            "status": "blocked",
            "review_status": "review-gated",
            "runner": "manual-review",
            "blocker_kind": "bridge-review",
            "blocker_reason": "LibreOffice transformed output is bridge provenance evidence, not a native source.",
            "required_evidence": [
                "native ODF source extraction",
                "transform-chain provenance review",
            ],
            "next_ready_condition": "Reconvert from native source or attach provenance review evidence before accepting the bridge output.",
        },
    )
    return blocker_path


def _dry_report(
    action: ExecutionAction,
    *,
    status: str,
    runner: str,
    context: DryRunContext,
    argv: tuple[str, ...] = (),
    blocking: bool | None = None,
    notes: tuple[str, ...] = (),
) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=action.blocking if blocking is None else blocking,
        supported=True,
        status=status,
        runner=runner,
        argv=argv,
        cwd=context.workspace_root,
        notes=notes,
    )


def _manual_dry_report(action: ExecutionAction, context: DryRunContext) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=True,
        supported=False,
        status="manual-gate",
        runner="manual",
        cwd=context.workspace_root,
        notes=("No ODF dry-run binding is registered for this action.",),
    )


def _build_native_dry_run_reports(
    step: ExecutionStep,
    context: DryRunContext,
) -> tuple[DryRunActionReport, ...]:
    input_path = _load_input_path(context)
    output_root = _dry_output_root(context, _NATIVE_SOURCE)
    reports: list[DryRunActionReport] = []

    for action in step.actions:
        if action.action_id == "extract-odf-formula":
            reports.append(
                _dry_report(
                    action,
                    status="ready",
                    runner="internal-odf-native",
                    context=context,
                    argv=(
                        "extract",
                        "--input",
                        input_path,
                        "--output-dir",
                        str(output_root / "extracted"),
                        "--manifest",
                        str(output_root / "manifest.json"),
                    ),
                    notes=("Extract MathML <math> payloads from FODT or ODF package content.xml members.",),
                )
            )
            continue

        if action.action_id == "convert-odf-mathml":
            reports.append(
                _dry_report(
                    action,
                    status="ready",
                    runner="internal-odf-native",
                    context=context,
                    argv=(
                        "canonicalize",
                        "--input-dir",
                        str(output_root / "extracted"),
                        "--output-dir",
                        str(output_root / "canonical-mathml"),
                        "--preserve-existing-mathml",
                    ),
                    notes=("ODF formula payloads are already MathML; this slice preserves them as canonical artifacts.",),
                )
            )
            continue

        if action.action_id == "emit-target-format":
            reports.append(
                _dry_report(
                    action,
                    status="validation-gated",
                    runner="manual-validation",
                    context=context,
                    argv=("review-target-format-gate", "--manifest", str(output_root / "manifest.json")),
                    blocking=True,
                    notes=("Target DOCX/OMML/LaTeX emission is not claimed by this extraction slice.",),
                )
            )
            continue

        reports.append(_manual_dry_report(action, context))

    return tuple(reports)


def _build_bridge_dry_run_reports(
    step: ExecutionStep,
    context: DryRunContext,
) -> tuple[DryRunActionReport, ...]:
    output_root = _dry_output_root(context, _BRIDGE_SOURCE)
    blocker_path = output_root / "blocker-record.json"
    reports: list[DryRunActionReport] = []
    for action in step.actions:
        if action.action_id in {"inspect-transform-chain", "bridge-review", "decide-reconvert-or-accept"}:
            reports.append(
                _dry_report(
                    action,
                    status="review-gated",
                    runner="manual-review",
                    context=context,
                    argv=("review-libreoffice-bridge", "--blocker", str(blocker_path)),
                    blocking=True,
                    notes=(
                        "LibreOffice transformed output is treated as bridge evidence.",
                        f"Review writes a blocker record at {blocker_path}.",
                        "Accept only after provenance review or reconvert from a native source.",
                    ),
                )
            )
            continue
        reports.append(_manual_dry_report(action, context))
    return tuple(reports)


def build_odf_dry_run_reports(
    step: ExecutionStep,
    context: DryRunContext,
) -> tuple[DryRunActionReport, ...]:
    if step.source_family == _NATIVE_SOURCE:
        return _build_native_dry_run_reports(step, context)
    if step.source_family == _BRIDGE_SOURCE:
        return _build_bridge_dry_run_reports(step, context)
    raise ValueError(f"ODF dry-run binding cannot handle source_family={step.source_family!r}.")


def _execution_report(
    *,
    action: ExecutionAction,
    status: str,
    runner: str,
    context: ExecutionContext,
    supported: bool = True,
    blocking: bool | None = None,
    output_paths: tuple[Path, ...] = (),
    notes: tuple[str, ...] = (),
) -> ActionExecutionReport:
    return ActionExecutionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=action.blocking if blocking is None else blocking,
        supported=supported,
        status=status,
        runner=runner,
        cwd=context.workspace_root,
        output_paths=tuple(str(path) for path in output_paths),
        notes=notes,
    )


def _manual_execution_report(action: ExecutionAction, context: ExecutionContext) -> ActionExecutionReport:
    return _execution_report(
        action=action,
        status="manual-gate",
        runner="manual",
        context=context,
        supported=False,
        blocking=True,
        notes=("No ODF execution binding is registered for this action.",),
    )


def _execute_bridge_step(
    step: ExecutionStep,
    context: ExecutionContext,
) -> tuple[ActionExecutionReport, ...]:
    blocker_path = _write_blocker_record(step, context)
    reports: list[ActionExecutionReport] = []
    for action in step.actions:
        if action.action_id in {"inspect-transform-chain", "bridge-review", "decide-reconvert-or-accept"}:
            reports.append(
                _execution_report(
                    action=action,
                    status="review-gated",
                    runner="manual-review",
                    context=context,
                    blocking=True,
                    output_paths=(blocker_path,),
                    notes=(
                        "LibreOffice transformed output stays behind a bridge provenance review gate.",
                        f"Bridge review writes a blocker record at {blocker_path}.",
                        "This executor deliberately does not extract it as a native source.",
                    ),
                )
            )
            continue
        reports.append(_manual_execution_report(action, context))
    return tuple(reports)


def execute_odf_step(
    step: ExecutionStep,
    context: ExecutionContext,
) -> tuple[ActionExecutionReport, ...]:
    if step.source_family == _BRIDGE_SOURCE:
        return _execute_bridge_step(step, context)
    if step.source_family != _NATIVE_SOURCE:
        raise ValueError(f"ODF executor cannot handle source_family={step.source_family!r}.")

    output_root = _execution_output_root(context, _NATIVE_SOURCE)
    input_path = Path(context.input_path)
    reports: list[ActionExecutionReport] = []
    extract_succeeded = False
    extracted_count = 0

    for action in step.actions:
        if action.action_id == "extract-odf-formula":
            if not input_path.exists():
                reports.append(
                    _execution_report(
                        action=action,
                        status="failed",
                        runner="internal-odf-native",
                        context=context,
                        blocking=True,
                        notes=(f"Input ODF document does not exist: {input_path}",),
                    )
                )
                continue
            try:
                manifest_path, extracted_paths = _extract_odf_mathml(input_path, output_root)
            except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
                reports.append(
                    _execution_report(
                        action=action,
                        status="failed",
                        runner="internal-odf-native",
                        context=context,
                        blocking=True,
                        notes=(f"Failed to extract ODF MathML payloads: {exc}",),
                    )
                )
                continue
            extract_succeeded = True
            extracted_count = len(extracted_paths)
            reports.append(
                _execution_report(
                    action=action,
                    status="completed",
                    runner="internal-odf-native",
                    context=context,
                    output_paths=(manifest_path, *extracted_paths),
                    notes=(
                        f"Extracted {extracted_count} ODF MathML fragment(s).",
                        "Native MathML was preserved as XML artifacts.",
                    ),
                )
            )
            continue

        if action.action_id == "convert-odf-mathml":
            if not extract_succeeded:
                reports.append(
                    _execution_report(
                        action=action,
                        status="skipped-after-failure",
                        runner="internal-odf-native",
                        context=context,
                        notes=("Extraction did not complete, so canonicalization was skipped.",),
                    )
                )
                continue
            if extracted_count == 0:
                reports.append(
                    _execution_report(
                        action=action,
                        status="skipped-empty",
                        runner="internal-odf-native",
                        context=context,
                        notes=("No ODF MathML fragments were found to canonicalize.",),
                    )
                )
                continue
            summary_path, canonical_paths = _canonicalize_mathml(output_root)
            reports.append(
                _execution_report(
                    action=action,
                    status="completed",
                    runner="internal-odf-native",
                    context=context,
                    output_paths=(summary_path, *canonical_paths),
                    notes=(
                        "Canonicalization currently preserves existing ODF MathML.",
                        f"Canonical artifact count: {len(canonical_paths)}.",
                    ),
                )
            )
            continue

        if action.action_id == "emit-target-format":
            if not extract_succeeded:
                reports.append(
                    _execution_report(
                        action=action,
                        status="skipped-after-failure",
                        runner="manual-validation",
                        context=context,
                        notes=("Extraction did not complete, so target-format review was skipped.",),
                    )
                )
                continue
            evidence_path = _write_validation_evidence(step, context, output_root)
            reports.append(
                _execution_report(
                    action=action,
                    status="validation-gated",
                    runner="manual-validation",
                    context=context,
                    blocking=True,
                    output_paths=(evidence_path,),
                    notes=(
                        "No final target format was emitted by this ODF-native extraction slice.",
                        f"Validation evidence was recorded at {evidence_path}.",
                        "Use a later validated conversion binding before claiming deliverable output.",
                    ),
                )
            )
            continue

        reports.append(_manual_execution_report(action, context))

    return tuple(reports)
