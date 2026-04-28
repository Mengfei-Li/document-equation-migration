from __future__ import annotations

import hashlib
import json
import subprocess
from pathlib import Path
from xml.etree import ElementTree as ET

from ..canonical_target import canonical_mathml_contract_for_source_family
from ..execution_plan.model import ExecutionAction, ExecutionStep
from .model import ActionExecutionReport, DryRunActionReport, DryRunContext, ExecutionContext


_INPUT_DOCX_PLACEHOLDER = "<input-docx-from-plan>"
_PROVIDER_NAME = "mathtype"
_SOURCE_FAMILY = "mathtype-ole"
_DEFAULT_LAYOUT_FACTOR = 1.01375


def mathtype_canonical_artifact_requirements() -> dict[str, object]:
    return {
        "target_stage": "equation-native-mtef-to-canonical-mathml",
        "minimum_artifact_set": (
            "At least one MathType OLE source object with extracted Equation Native / MTEF evidence, "
            "converter-emitted MathML, accepted canonical MathML, formula-count parity, and source-to-artifact provenance."
        ),
        "accepted_artifact_sets": [
            {
                "id": "normalized-mathml-promoted-to-canonical",
                "description": (
                    "The MathType converter emits MathML that is normalized and accepted as canonical MathML "
                    "with explicit summary and provenance records."
                ),
                "required_evidence": [
                    "raw Equation Native / MTEF payload artifact or digest",
                    "converted/*.mml or converted/*.mathml",
                    "canonical-mathml/*.xml",
                    "canonicalization-summary.json",
                    "validation-evidence.json",
                    "source-to-canonical provenance map",
                ],
            },
            {
                "id": "converter-native-control",
                "description": (
                    "A controlled converter-native fixture demonstrates the external MathType toolchain and "
                    "produces canonical MathML evidence before any downstream OMML / Word route is considered."
                ),
                "required_evidence": [
                    "fixture source attribution or permission tier",
                    "raw payload SHA-256",
                    "converter stdout/stderr logs",
                    "canonical MathML XML parse success",
                    "formula-count parity",
                ],
            },
        ],
        "required_candidate_properties": [
            {
                "id": "mathtype-identity",
                "description": "The source formula is classified as MathType OLE.",
                "evidence_fields": (
                    "source_family=mathtype-ole",
                    "provenance.prog_id_raw=Equation.DSMT* or field_code_raw contains EMBED Equation.DSMT",
                    "source_specific.mathtype.class_id_raw",
                ),
            },
            {
                "id": "equation-native-payload",
                "description": "Each source formula exposes real Equation Native / MTEF payload evidence.",
                "evidence_fields": (
                    "source_role=native-source",
                    "provenance.raw_payload_status=present",
                    "provenance.raw_payload_sha256",
                    "provenance.payload_stream_name",
                    "mtef_version",
                ),
            },
            {
                "id": "canonical-output",
                "description": "Accepted output is canonical MathML, not only OMML or Word-rendered output.",
                "evidence_fields": (
                    "canonical-mathml/*.xml",
                    "canonicalization-summary.json",
                    "canonical_mathml_count equals accepted source formula count",
                    "unsupported_fragment_count recorded",
                ),
            },
            {
                "id": "provenance-map",
                "description": "Every canonical MathML artifact remains traceable to the original DOCX object.",
                "evidence_fields": (
                    "formula_id",
                    "doc_part_path",
                    "relationship_id",
                    "embedding_target",
                    "raw_payload_sha256",
                    "canonical_artifact_path",
                ),
            },
        ],
        "disqualifying_conditions": [
            "marker-text payload instead of binary Equation Native / MTEF payload",
            "preview-only image without native payload",
            "empty, BOM-only, or XML-invalid converter output",
            "Word-only OMML output without accepted canonical MathML artifacts",
            "formula-count mismatch between detected MathType sources and accepted canonical artifacts",
            "missing source-to-artifact provenance",
            "unclear redistribution or use permission for public fixture promotion",
        ],
        "promotion_gate": [
            "Detector evidence proves MathType OLE source identity.",
            "External tool prerequisites are recorded before live conversion.",
            "Canonical MathML artifacts validate as XML and preserve formula count.",
            "Each accepted artifact has source provenance and raw-payload digest evidence.",
            "Downstream OMML / Word validation is attempted only after the canonical artifact gate passes.",
        ],
    }


def _workspace_root(context: DryRunContext) -> Path:
    return Path(context.workspace_root)


def _output_dir(context: DryRunContext) -> Path:
    output_dir = Path(context.output_dir_hint)
    if output_dir.is_absolute():
        return output_dir
    return _workspace_root(context) / output_dir


def _load_input_docx(context: DryRunContext) -> str:
    plan_path_text = context.execution_plan_path.strip()
    if not plan_path_text:
        return _INPUT_DOCX_PLACEHOLDER

    plan_path = Path(plan_path_text)
    if not plan_path.is_absolute():
        plan_path = _workspace_root(context) / plan_path

    try:
        payload = json.loads(plan_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return _INPUT_DOCX_PLACEHOLDER

    input_path_text = str(payload.get("input_path", "")).strip()
    if not input_path_text:
        return _INPUT_DOCX_PLACEHOLDER

    input_path = Path(input_path_text)
    if input_path.is_absolute():
        return str(input_path)
    return str((_workspace_root(context) / input_path).resolve())


def _input_stem(input_docx: str) -> str:
    if input_docx == _INPUT_DOCX_PLACEHOLDER:
        return "document"
    return Path(input_docx).stem or "document"


def _execution_output_root(context: ExecutionContext) -> Path:
    return Path(context.output_dir) / _PROVIDER_NAME


def _as_mapping(value: object) -> dict[str, object]:
    if isinstance(value, dict):
        return {str(key): item for key, item in value.items()}
    return {}


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


def _as_float(value: object, *, default: float) -> float:
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


def _mathtype_experimental_options(step: ExecutionStep) -> dict[str, object]:
    metadata = _as_mapping(step.metadata)
    experimental_options = _as_mapping(metadata.get("experimental_options"))
    if experimental_options:
        return experimental_options

    fallback_options: dict[str, object] = {}
    for key in (
        "preserve_mathtype_layout",
        "mathtype_layout_factor",
        "resume_mathtype_pipeline",
        "mathtype_start_index",
        "mathtype_end_index",
    ):
        if key in metadata:
            fallback_options[key] = metadata[key]
    return fallback_options


def _mathtype_layout_args(step: ExecutionStep, *, powershell: bool) -> tuple[str, ...]:
    experimental_options = _mathtype_experimental_options(step)
    if not _as_bool(experimental_options.get("preserve_mathtype_layout"), default=False):
        return ()

    factor = _as_float(
        experimental_options.get("mathtype_layout_factor"),
        default=_DEFAULT_LAYOUT_FACTOR,
    )
    factor_text = f"{factor:g}"
    if powershell:
        return ("-PreserveMathTypeLayout", "-MathTypeLayoutFactor", factor_text)
    return ("--preserve-mathtype-layout", "--mathtype-layout-factor", factor_text)


def _layout_option_notes(step: ExecutionStep) -> tuple[str, ...]:
    if not _mathtype_layout_args(step, powershell=True):
        return ()
    return ("Guarded MathType layout preservation is enabled via execution-step metadata.",)


def _mathtype_resume_args(step: ExecutionStep) -> tuple[str, ...]:
    experimental_options = _mathtype_experimental_options(step)
    args: list[str] = []
    if _as_bool(experimental_options.get("resume_mathtype_pipeline"), default=False):
        args.append("-Resume")

    start_index = _as_int(experimental_options.get("mathtype_start_index"), default=0)
    end_index = _as_int(experimental_options.get("mathtype_end_index"), default=0)
    if start_index > 0:
        args.extend(("-StartIndex", str(start_index)))
    if end_index > 0:
        args.extend(("-EndIndex", str(end_index)))
    return tuple(args)


def _resume_option_notes(step: ExecutionStep) -> tuple[str, ...]:
    if not _mathtype_resume_args(step):
        return ()
    return (
        "MathType wrapper resume/chunk options are enabled via execution-step metadata.",
        "Use -Resume with optional -StartIndex/-EndIndex to continue long runs without deleting converted artifacts.",
    )


def _canonical_artifact_notes() -> tuple[str, ...]:
    return (
        "Structured-core acceptance requires canonical MathML artifacts, formula-count parity, and source-to-artifact provenance.",
        "OMML replacement and Word export are downstream validation surfaces, not the canonical target itself.",
    )


def _write_json(path: Path, payload: dict[str, object]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _blocker_record_path(output_root: Path) -> Path:
    return output_root / "blocker-record.json"


def _validation_evidence_path(output_root: Path) -> Path:
    return output_root / "validation-evidence.json"


def _canonicalization_summary_path(output_root: Path) -> Path:
    return output_root / "canonicalization-summary.json"


def _canonical_mathml_dir(output_root: Path) -> Path:
    return output_root / "canonical-mathml"


def _converted_mathml_paths(output_root: Path) -> tuple[Path, ...]:
    converted_dir = output_root / "converted"
    if not converted_dir.exists():
        return ()
    paths: list[Path] = []
    for pattern in ("*.mml", "*.mathml"):
        paths.extend(converted_dir.glob(pattern))
    return tuple(sorted(paths))


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _is_mathml_root(root: ET.Element) -> bool:
    return _local_name(root.tag) == "math"


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


def _write_canonical_mathml_artifacts(
    *,
    output_root: Path,
    step: ExecutionStep,
) -> dict[str, object]:
    canonical_dir = _canonical_mathml_dir(output_root)
    canonical_dir.mkdir(parents=True, exist_ok=True)
    canonical_items: list[dict[str, object]] = []
    unsupported_items: list[dict[str, object]] = []

    for index, source_path in enumerate(_converted_mathml_paths(output_root), start=1):
        try:
            text = source_path.read_text(encoding="utf-8-sig").strip()
        except OSError as exc:
            unsupported_items.append(
                {
                    "source_path": str(source_path),
                    "status": "read-error",
                    "error": str(exc),
                }
            )
            continue

        if not text:
            unsupported_items.append(
                {
                    "source_path": str(source_path),
                    "status": "empty-or-bom-only",
                }
            )
            continue

        try:
            root = ET.fromstring(text)
        except ET.ParseError as exc:
            unsupported_items.append(
                {
                    "source_path": str(source_path),
                    "status": "xml-parse-error",
                    "error": str(exc),
                }
            )
            continue

        if not _is_mathml_root(root):
            unsupported_items.append(
                {
                    "source_path": str(source_path),
                    "status": "not-mathml-root",
                    "root_tag": root.tag,
                }
            )
            continue

        formula_id = f"mathtype-canonical-{len(canonical_items) + 1:04d}"
        target_path = canonical_dir / f"{formula_id}.xml"
        target_text = f'<?xml version="1.0" encoding="UTF-8"?>\n{text}\n'
        target_path.write_text(target_text, encoding="utf-8")
        canonical_items.append(
            {
                "formula_id": formula_id,
                "source_mathml_path": str(source_path),
                "canonical_artifact_path": str(target_path),
                "source_family": _SOURCE_FAMILY,
                "provider": _PROVIDER_NAME,
                "source_index": index,
                "source_sha256": _sha256_text(text),
                "canonical_sha256": _sha256_text(target_text),
                "preservation_status": "mathml-content-preserved-with-xml-declaration",
                "root_tag": root.tag,
                "property_signals": _mathml_property_signals(root),
            }
        )

    expected_count = step.formula_count
    canonical_count = len(canonical_items)
    unsupported_count = len(unsupported_items)
    parity_status = (
        "passed"
        if expected_count == canonical_count and unsupported_count == 0
        else "mismatch"
        if expected_count != canonical_count
        else "unsupported-fragments"
    )
    gate_status = (
        "passed"
        if canonical_count > 0 and parity_status == "passed"
        else "blocked-canonical-artifact"
    )
    summary: dict[str, object] = {
        "artifact_type": "mathtype-canonicalization-summary",
        "provider": _PROVIDER_NAME,
        "source_family": _SOURCE_FAMILY,
        "strategy": "materialize-normalized-mathml",
        "expected_formula_count": expected_count,
        "canonical_mathml_count": canonical_count,
        "unsupported_fragment_count": unsupported_count,
        "formula_count_parity": parity_status,
        "gate_status": gate_status,
        "canonical_mathml_dir": str(canonical_dir),
        "source_to_canonical_provenance": canonical_items,
        "property_summary": _property_summary(canonical_items),
        "unsupported_fragments": unsupported_items,
    }
    _write_json(_canonicalization_summary_path(output_root), summary)
    return summary


def _step_summaries(step: ExecutionStep, status: str) -> list[dict[str, object]]:
    return [
        {
            "action_id": action.action_id,
            "description": action.description,
            "blocking": action.blocking,
            "status": status,
        }
        for action in step.actions
    ]


def _write_blocker_record(
    *,
    output_root: Path,
    step: ExecutionStep,
    context: ExecutionContext,
    status: str,
    reason: str,
    required_evidence: tuple[str, ...],
    next_ready_condition: str,
    pipeline_artifact_path: str = "",
) -> Path:
    blocker_path = _blocker_record_path(output_root)
    payload: dict[str, object] = {
        "artifact_type": "mathtype-blocker-record",
        "provider": _PROVIDER_NAME,
        "source_family": _SOURCE_FAMILY,
        "canonical_target": canonical_mathml_contract_for_source_family(_SOURCE_FAMILY).to_dict(),
        "status": status,
        "gate_state": status,
        "conversion_claim": False,
        "reason": reason,
        "required_evidence": list(required_evidence),
        "next_ready_condition": next_ready_condition,
        "canonical_artifact_admissibility": mathtype_canonical_artifact_requirements(),
        "input_path": context.input_path,
        "execution_plan_path": context.execution_plan_path,
        "output_root": str(output_root),
        "pipeline_artifact_path": pipeline_artifact_path,
        "actions": _step_summaries(step, status),
    }
    return _write_json(blocker_path, payload)


def _write_validation_evidence(
    *,
    output_root: Path,
    step: ExecutionStep,
    context: ExecutionContext,
    executed_report: ActionExecutionReport,
    blocker_path: Path,
    covered_status: str,
) -> Path:
    evidence_path = _validation_evidence_path(output_root)
    covered_actions = [
        {
            "action_id": action.action_id,
            "description": action.description,
            "status": covered_status,
        }
        for action in step.actions
        if action.action_id != executed_report.action_id and action.action_id != "validate-word-output"
    ]
    canonicalization_summary = _write_canonical_mathml_artifacts(output_root=output_root, step=step)
    payload: dict[str, object] = {
        "artifact_type": "mathtype-validation-evidence",
        "provider": _PROVIDER_NAME,
        "source_family": _SOURCE_FAMILY,
        "canonical_target": canonical_mathml_contract_for_source_family(_SOURCE_FAMILY).to_dict(),
        "status": executed_report.status,
        "input_path": context.input_path,
        "execution_plan_path": context.execution_plan_path,
        "output_root": str(output_root),
        "pipeline": {
            "action_id": executed_report.action_id,
            "runner": executed_report.runner,
            "argv": list(executed_report.argv),
            "cwd": executed_report.cwd,
            "exit_code": executed_report.exit_code,
            "stdout_path": executed_report.stdout_path,
            "stderr_path": executed_report.stderr_path,
            "output_paths": list(executed_report.output_paths),
        },
        "covered_actions": covered_actions,
        "validation_gate": {
            "action_id": "validate-word-output",
            "status": "validation-gated",
            "artifact_path": str(blocker_path),
        },
        "canonical_artifact_gate": {
            "status": canonicalization_summary["gate_status"],
            "conversion_claim": False,
            "admissibility": mathtype_canonical_artifact_requirements(),
            "canonicalization_summary_path": str(_canonicalization_summary_path(output_root)),
            "canonical_mathml_count": canonicalization_summary["canonical_mathml_count"],
            "unsupported_fragment_count": canonicalization_summary["unsupported_fragment_count"],
            "formula_count_parity": canonicalization_summary["formula_count_parity"],
            "property_summary": canonicalization_summary["property_summary"],
        },
        "source_to_canonical_provenance": canonicalization_summary["source_to_canonical_provenance"],
        "next_ready_condition": (
            "Accept canonical MathML artifacts with provenance before treating downstream OMML / Word output as deliverable."
        ),
    }
    return _write_json(evidence_path, payload)


def _powershell_report(
    action: ExecutionAction,
    *,
    context: DryRunContext,
    script_name: str,
    script_args: tuple[str, ...],
    notes: tuple[str, ...],
) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=action.blocking,
        supported=True,
        status="ready",
        runner="powershell",
        argv=(
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(_workspace_root(context) / script_name),
            *script_args,
        ),
        cwd=str(_workspace_root(context)),
        notes=notes,
    )


def _python_report(
    action: ExecutionAction,
    *,
    context: DryRunContext,
    script_name: str,
    script_args: tuple[str, ...],
    notes: tuple[str, ...],
) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action.action_id,
        description=action.description,
        blocking=action.blocking,
        supported=True,
        status="ready",
        runner="python",
        argv=(str(_workspace_root(context) / script_name), *script_args),
        cwd=str(_workspace_root(context)),
        notes=notes,
    )


def _manual_review_note(step: ExecutionStep) -> tuple[str, ...]:
    if not step.requires_manual_review:
        return ()
    return ("Execution plan flags this step for manual review after the dry-run command path.",)


def build_mathtype_dry_run_reports(
    step: ExecutionStep,
    context: DryRunContext,
) -> tuple[DryRunActionReport, ...]:
    if step.source_family != _SOURCE_FAMILY:
        raise ValueError(f"MathType dry-run binding cannot handle source_family={step.source_family!r}.")

    input_docx = _load_input_docx(context)
    output_dir = _output_dir(context)
    document_stem = _input_stem(input_docx)
    ole_dir = output_dir / "ole_bins"
    converted_dir = output_dir / "converted"
    output_docx = output_dir / f"{document_stem}.omml.docx"
    output_pdf = output_dir / f"{document_stem}.omml.pdf"
    pipeline_layout_args = _mathtype_layout_args(step, powershell=True)
    pipeline_resume_args = _mathtype_resume_args(step)
    replace_layout_args = _mathtype_layout_args(step, powershell=False)
    layout_option_notes = _layout_option_notes(step)
    resume_option_notes = _resume_option_notes(step)

    reports: list[DryRunActionReport] = []
    shared_probe_note = (
        "This preview reuses probe_formula_pipeline.ps1, which chains the Java bridge, "
        "normalize_mathml.py, and MML2OMML.XSL inside one staged conversion script."
    )
    for action in step.actions:
        action_id = action.action_id
        if action_id == "extract-equation-native":
            reports.append(
                _powershell_report(
                    action,
                    context=context,
                    script_name="run_docx_open_source_pipeline.ps1",
                    script_args=(
                        "-InputDocx",
                        input_docx,
                        "-OutputDir",
                        str(output_dir),
                        "-OutputDocx",
                        str(output_docx),
                        "-SkipLatexPreview",
                        *pipeline_resume_args,
                        *pipeline_layout_args,
                    ),
                    notes=(
                        "The repository does not expose a standalone DOCX-to-OLE extraction CLI, "
                        "so dry-run preview binds extraction to the document-level pipeline wrapper.",
                        f"Expected extracted OLE payloads will be staged under {ole_dir}.",
                        "The same wrapper also owns later replacement/summary stages in the end-to-end pipeline.",
                        *layout_option_notes,
                        *resume_option_notes,
                        *_manual_review_note(step),
                    ),
                )
            )
            continue

        if action_id == "mtef-to-mathml":
            reports.append(
                _powershell_report(
                    action,
                    context=context,
                    script_name="probe_formula_pipeline.ps1",
                    script_args=(
                        "-InputDir",
                        str(ole_dir),
                        "-OutputDir",
                        str(converted_dir),
                        "-Limit",
                        "0",
                        "-SkipLatexPreview",
                    ),
                    notes=(
                        "This action drives the MathType Java bridge through probe_formula_pipeline.ps1.",
                        f"MathML outputs are expected under {converted_dir}.",
                        shared_probe_note,
                        *_canonical_artifact_notes(),
                        *_manual_review_note(step),
                    ),
                )
            )
            continue

        if action_id == "normalize-mathml":
            reports.append(
                _powershell_report(
                    action,
                    context=context,
                    script_name="probe_formula_pipeline.ps1",
                    script_args=(
                        "-InputDir",
                        str(ole_dir),
                        "-OutputDir",
                        str(converted_dir),
                        "-Limit",
                        "0",
                        "-SkipLatexPreview",
                    ),
                    notes=(
                        "normalize_mathml.py is invoked inside probe_formula_pipeline.ps1 immediately after MathML emission.",
                        "This dry-run reuses the same staged probe command rather than inventing a second wrapper.",
                        shared_probe_note,
                        *_canonical_artifact_notes(),
                        *_manual_review_note(step),
                    ),
                )
            )
            continue

        if action_id == "mathml-to-omml":
            reports.append(
                _powershell_report(
                    action,
                    context=context,
                    script_name="probe_formula_pipeline.ps1",
                    script_args=(
                        "-InputDir",
                        str(ole_dir),
                        "-OutputDir",
                        str(converted_dir),
                        "-Limit",
                        "0",
                        "-SkipLatexPreview",
                    ),
                    notes=(
                        "MML2OMML.XSL is applied inside probe_formula_pipeline.ps1 after MathML normalization.",
                        "This step shares the same command preview as the preceding probe stages by design.",
                        "This downstream OMML step does not by itself satisfy the canonical MathML artifact gate.",
                        shared_probe_note,
                        *_manual_review_note(step),
                    ),
                )
            )
            continue

        if action_id == "replace-ole-with-omml":
            reports.append(
                _python_report(
                    action,
                    context=context,
                    script_name="replace_docx_ole_with_omml.py",
                    script_args=(
                        input_docx,
                        str(converted_dir),
                        str(output_docx),
                        *replace_layout_args,
                    ),
                    notes=(
                        "This action binds directly to replace_docx_ole_with_omml.py using the converted OMML artifacts.",
                        f"The converted OMML directory is expected at {converted_dir}.",
                        "run_docx_open_source_pipeline.ps1 invokes this same script during the full document pipeline.",
                        *layout_option_notes,
                        *_manual_review_note(step),
                    ),
                )
            )
            continue

        if action_id == "validate-word-output":
            reports.append(
                _powershell_report(
                    action,
                    context=context,
                    script_name="export_word_pdf.ps1",
                    script_args=(
                        "-InputDocx",
                        str(output_docx),
                        "-OutputPdf",
                        str(output_pdf),
                    ),
                    notes=(
                        "Word validation is previewed via export_word_pdf.ps1 so the converted DOCX is opened by Word and exported to PDF.",
                        f"Expected validation artifact: {output_pdf}.",
                        "For structural follow-up, the repository also provides docx_math_object_map.py and XML summary files in the full pipeline.",
                        *_manual_review_note(step),
                    ),
                )
            )
            continue

        reports.append(
            DryRunActionReport(
                action_id=action.action_id,
                description=action.description,
                blocking=action.blocking,
                supported=False,
                status="manual-gate",
                runner="manual",
                cwd=str(_workspace_root(context)),
                notes=("No MathType dry-run command binding is registered for this action.",),
            )
        )

    return tuple(reports)


def _to_blocked_execution_report(
    dry_run_report: DryRunActionReport,
    *,
    context: ExecutionContext,
    blocker_path: Path,
) -> ActionExecutionReport:
    return ActionExecutionReport(
        action_id=dry_run_report.action_id,
        description=dry_run_report.description,
        blocking=dry_run_report.blocking,
        supported=dry_run_report.supported,
        status="blocked-external-tool",
        runner=dry_run_report.runner,
        argv=dry_run_report.argv,
        cwd=dry_run_report.cwd,
        notes=(
            "MathType execution uses existing PowerShell/Python pipeline scripts and is gated by --allow-external-tools.",
            "Run first with --dry-run, then pass --execute --allow-external-tools only when Java, Office XSL, and local script dependencies are ready.",
            *dry_run_report.notes,
        ),
        output_paths=(str(blocker_path),),
    )


def _run_external_pipeline(
    dry_run_report: DryRunActionReport,
    *,
    context: ExecutionContext,
    output_root: Path,
) -> ActionExecutionReport:
    logs_dir = output_root / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    stdout_path = logs_dir / f"{dry_run_report.action_id}.stdout.txt"
    stderr_path = logs_dir / f"{dry_run_report.action_id}.stderr.txt"

    executable = "powershell" if dry_run_report.runner == "powershell" else "python"
    argv = (executable, *dry_run_report.argv)
    completed = subprocess.run(
        argv,
        cwd=dry_run_report.cwd or context.workspace_root,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    stdout_path.write_text(completed.stdout, encoding="utf-8")
    stderr_path.write_text(completed.stderr, encoding="utf-8")
    status = "completed" if completed.returncode == 0 else "failed"
    return ActionExecutionReport(
        action_id=dry_run_report.action_id,
        description=dry_run_report.description,
        blocking=dry_run_report.blocking,
        supported=dry_run_report.supported,
        status=status,
        runner=dry_run_report.runner,
        argv=argv,
        cwd=dry_run_report.cwd,
        exit_code=completed.returncode,
        stdout_path=str(stdout_path),
        stderr_path=str(stderr_path),
        output_paths=(str(output_root),),
        notes=(
            "Executed the guarded MathType document-level pipeline entry point.",
            "Inspect stdout/stderr logs and downstream pipeline summary before treating output as deliverable.",
        ),
    )


def execute_mathtype_step(
    step: ExecutionStep,
    context: ExecutionContext,
) -> tuple[ActionExecutionReport, ...]:
    if step.source_family != _SOURCE_FAMILY:
        raise ValueError(f"MathType executor cannot handle source_family={step.source_family!r}.")

    output_root = _execution_output_root(context)
    dry_run_context = DryRunContext(
        workspace_root=context.workspace_root,
        execution_plan_path=context.execution_plan_path,
        output_dir_hint=str(output_root),
    )
    dry_run_reports = build_mathtype_dry_run_reports(step, dry_run_context)
    if not context.allow_external_tools:
        blocker_path = _write_blocker_record(
            output_root=output_root,
            step=step,
            context=context,
            status="blocked-external-tool",
            reason="MathType execution requires --allow-external-tools before the guarded pipeline can run.",
            required_evidence=(
                "Execute the provider with --allow-external-tools enabled.",
                "Record external tool availability for the document-level pipeline scripts.",
                "Capture a validation artifact from the Word open/export gate after pipeline completion.",
            ),
            next_ready_condition="Re-run with --allow-external-tools and provide Word open/export validation evidence.",
        )
        return tuple(
            _to_blocked_execution_report(report, context=context, blocker_path=blocker_path)
            for report in dry_run_reports
        )

    reports: list[ActionExecutionReport] = []
    pipeline_has_run = False
    pipeline_succeeded = False
    pipeline_stdout_path = ""
    pipeline_stderr_path = ""
    blocker_path: Path | None = None
    validation_evidence_path: Path | None = None
    for dry_run_report in dry_run_reports:
        if dry_run_report.action_id == "extract-equation-native" and not pipeline_has_run:
            executed = _run_external_pipeline(dry_run_report, context=context, output_root=output_root)
            pipeline_has_run = True
            pipeline_succeeded = executed.status == "completed"
            pipeline_stdout_path = executed.stdout_path
            pipeline_stderr_path = executed.stderr_path
            blocker_path = _write_blocker_record(
                output_root=output_root,
                step=step,
                context=context,
                status="validation-gated" if pipeline_succeeded else "blocked-external-tool",
                reason=(
                    "The guarded MathType document-level pipeline completed, but Word open/export parity remains a validation gate."
                    if pipeline_succeeded
                    else "The guarded MathType document-level pipeline did not complete successfully."
                ),
                required_evidence=(
                    "Capture a Word open/export parity report.",
                    "Attach the PDF export artifact produced from the converted DOCX.",
                    "Confirm the validation gate before treating output as deliverable.",
                ),
                next_ready_condition=(
                    "Provide the Word open/export parity evidence and update the blocker record."
                    if pipeline_succeeded
                    else "Fix the guarded pipeline failure and rerun the execution path."
                ),
                pipeline_artifact_path=(
                    str(_validation_evidence_path(output_root)) if pipeline_succeeded else ""
                ),
            )
            if pipeline_succeeded:
                validation_evidence_path = _write_validation_evidence(
                    output_root=output_root,
                    step=step,
                    context=context,
                    executed_report=executed,
                    blocker_path=blocker_path,
                    covered_status="covered-by-document-pipeline",
                )
                reports.append(
                    ActionExecutionReport(
                        action_id=executed.action_id,
                        description=executed.description,
                        blocking=executed.blocking,
                        supported=executed.supported,
                        status=executed.status,
                        runner=executed.runner,
                        argv=executed.argv,
                        cwd=executed.cwd,
                        exit_code=executed.exit_code,
                        stdout_path=executed.stdout_path,
                        stderr_path=executed.stderr_path,
                        output_paths=(str(validation_evidence_path),),
                        notes=(
                            "Executed the guarded MathType document-level pipeline entry point.",
                            "Inspect stdout/stderr logs and downstream pipeline summary before treating output as deliverable.",
                        ),
                    )
                )
            else:
                reports.append(
                    ActionExecutionReport(
                        action_id=executed.action_id,
                        description=executed.description,
                        blocking=executed.blocking,
                        supported=executed.supported,
                        status=executed.status,
                        runner=executed.runner,
                        argv=executed.argv,
                        cwd=executed.cwd,
                        exit_code=executed.exit_code,
                        stdout_path=executed.stdout_path,
                        stderr_path=executed.stderr_path,
                        output_paths=(str(blocker_path),),
                        notes=(
                            "Executed the guarded MathType document-level pipeline entry point.",
                            "Inspect stdout/stderr logs and downstream pipeline summary before treating output as deliverable.",
                        ),
                    )
                )
            continue

        covered_output_paths = (
            (str(validation_evidence_path),)
            if pipeline_succeeded and validation_evidence_path is not None
            else (str(blocker_path),)
            if blocker_path is not None
            else ()
        )

        if dry_run_report.action_id == "validate-word-output" and pipeline_succeeded:
            reports.append(
                ActionExecutionReport(
                    action_id=dry_run_report.action_id,
                    description=dry_run_report.description,
                    blocking=dry_run_report.blocking,
                    supported=dry_run_report.supported,
                    status="validation-gated",
                    runner="manual-validation",
                    argv=dry_run_report.argv,
                    cwd=dry_run_report.cwd,
                    stdout_path=pipeline_stdout_path,
                    stderr_path=pipeline_stderr_path,
                    output_paths=(str(blocker_path),) if blocker_path is not None else (),
                    notes=(
                        "The guarded MathType document-level pipeline completed, but Word open/export parity remains a validation gate.",
                        "Inspect the blocker record, pipeline logs, and generated artifacts before treating output as deliverable.",
                        *dry_run_report.notes,
                    ),
                )
            )
            continue

        reports.append(
            ActionExecutionReport(
                action_id=dry_run_report.action_id,
                description=dry_run_report.description,
                blocking=dry_run_report.blocking,
                supported=dry_run_report.supported,
                status="covered-by-document-pipeline" if pipeline_succeeded else "skipped-after-failure",
                runner=dry_run_report.runner,
                argv=dry_run_report.argv,
                cwd=dry_run_report.cwd,
                stdout_path=pipeline_stdout_path if pipeline_succeeded else "",
                stderr_path=pipeline_stderr_path if pipeline_succeeded else "",
                output_paths=covered_output_paths,
                notes=(
                    "The MathType source-first script is a document-level pipeline, so this action is covered by the guarded pipeline run.",
                    "Review the validation evidence file alongside the blocker record before claiming deliverable output.",
                    *dry_run_report.notes,
                ),
            )
        )
    return tuple(reports)
