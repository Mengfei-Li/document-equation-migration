from __future__ import annotations

import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    import olefile
except ImportError:  # pragma: no cover - dependency is declared, this is defensive.
    olefile = None

from ..canonical_mathml_evidence import mathml_property_signals, property_summary, sha256_text
from ..canonical_target import canonical_mathml_contract_for_source_family
from ..detectors.equation_editor_3_ole import detect_equation_editor_3_ole
from ..equation3_mtef import Equation3MtefError, convert_equation3_payload_to_mathml, local_name
from ..execution_plan.model import ExecutionStep
from .model import ActionExecutionReport, DryRunActionReport, DryRunContext, ExecutionContext


SOURCE_FAMILY = "equation-editor-3-ole"
RUNNER = "internal-equation3-probe"
BLOCKER_RECORD_FILENAME = "blocker-record.json"
CANONICALIZATION_SUMMARY_FILENAME = "canonicalization-summary.json"
VALIDATION_EVIDENCE_FILENAME = "validation-evidence.json"


def equation3_fixture_admissibility_requirements() -> dict[str, object]:
    return {
        "target_stage": "fixture-backed-canonical-mathml-conversion",
        "minimum_fixture_set": (
            "At least one redistributable or explicitly authorized DOCX or legacy DOC containing a real "
            "Equation.3 native OLE payload, not only a preview image or marker payload."
        ),
        "required_candidate_properties": [
            {
                "id": "equation3-identity",
                "description": "The detector record identifies the source as Equation Editor 3.0.",
                "evidence_fields": (
                    "source_family=equation-editor-3-ole",
                    "provenance.prog_id_raw=Equation.3 or field_code_raw contains EMBED Equation",
                    "source_specific.equation_editor_3.class_id_raw",
                ),
            },
            {
                "id": "native-payload",
                "description": "The fixture exposes a readable native OLE payload, not just a rendered preview.",
                "evidence_fields": (
                    "source_role=native-source",
                    "provenance.raw_payload_status=present",
                    "provenance.raw_payload_sha256",
                    "provenance.payload_stream_name",
                ),
            },
            {
                "id": "mtef-v3-header",
                "description": "The native payload has Equation Editor 3.0 / MTEF v3 header evidence.",
                "evidence_fields": (
                    "source_specific.equation_editor_3.native_header_size_bytes",
                    "source_specific.equation_editor_3.mtef_version=3",
                    "source_specific.equation_editor_3.selected_route=mtef-v3-mainline",
                ),
            },
            {
                "id": "canonical-output",
                "description": (
                    "A conversion attempt emits valid canonical MathML artifacts with formula-count parity. "
                    "The current internal converter is limited to the observed and synthetic-covered MTEF v3 script, "
                    "root, fraction, slash-fraction, bar, fence, limit, matrix, "
                    "BigOp (sum/integral/product/coproduct/integral-op), and character structures."
                ),
                "evidence_fields": (
                    "canonical-mathml/*.xml",
                    "canonicalization-summary.json",
                    "canonical_mathml_count equals fixture formula count",
                    "unsupported_fragment_count recorded",
                ),
            },
            {
                "id": "provenance-map",
                "description": "Each canonical MathML artifact remains traceable to the source DOCX object.",
                "evidence_fields": (
                    "formula_id",
                    "doc_part_path",
                    "relationship_id",
                    "embedding_target",
                    "raw_payload_sha256",
                ),
            },
        ],
        "disqualifying_conditions": [
            "preview-only fixture",
            "marker-text payload instead of binary native payload",
            "missing or corrupt OLE payload",
            "conflicting MathType or AxMath vendor signal",
            "no valid canonical MathML output",
            "unclear redistribution or use permission for public fixture promotion",
        ],
        "promotion_gate": [
            "Detector evidence proves native Equation.3 source identity.",
            "Canonical MathML output validates as XML and preserves formula count.",
            "Conversion report keeps provenance for every source formula.",
            "DOCX/Word validation is attempted only after canonical conversion evidence exists.",
        ],
        "current_productized_slice": {
            "status": "implemented-limited",
            "supported_mtef_version": 3,
            "supported_containers": [
                "DOCX OLE embeddings exposing Equation Native payloads",
                "legacy .doc OLE compound files exposing ObjectPool/*/Equation Native streams",
            ],
            "supported_records": [
                "slot",
                "char",
                "tmpl script templates: tmSUP, tmSUB, tmSUBSUP",
                "tmpl root templates: tmROOT, tmNTHROOT",
                "tmpl fraction templates: tmFRACT, tmFRACT_SMALL",
                "tmpl slash-fraction templates: tmSLFRACT, tmSLFRACT_BASELINE, tmSLFRACT_SMALL",
                "tmpl bar templates: tmUBAR, tmUBAR_DOUBLE, tmOBAR, tmOBAR_DOUBLE",
                "tmpl fence templates: tmANGLE, tmPAREN, tmBRACE, tmBRACK, tmBAR, tmDBAR, tmFLOOR, tmCEILING",
                "tmpl limit templates: tmLIM_UPPER, tmLIM_LOWER, tmLIM_BOTH",
                "tmpl BigOp templates: tmSINT_NO_LIMITS, tmSINT_LOWER, tmSINT_BOTH, tmSUM_NO_LIMITS, tmSUM_LOWER, tmSUM_BOTH, tmISUM_LOWER, tmISUM_BOTH, tmPROD_NO_LIMITS, tmPROD_LOWER, tmPROD_BOTH, tmIPROD_LOWER, tmIPROD_BOTH, tmCOPROD_NO_LIMITS, tmCOPROD_LOWER, tmCOPROD_BOTH, tmINTOP_UPPER, tmINTOP_LOWER, tmINTOP_BOTH",
                "matrix records with supported line-based cells",
                "full/sub/sub2 placeholder markers",
                "font/size/ruler records as ignored formatting metadata",
                "embellishment records parsed; prime mapped to msup (others currently ignored)",
                "legacy Equation Native trailer: optional 16-bit word after END (ignored)",
            ],
            "unsupported_records": [
                "unsupported matrix cell object records",
                "unknown template selectors",
                "unexpected trailing bytes beyond the allowed trailer",
            ],
            "general_converter_claim": False,
        },
    }


def _dry_run_output_root(context: DryRunContext) -> Path:
    return Path(context.output_dir_hint) / "equation-editor-3-ole"


def _execution_output_root(context: ExecutionContext) -> Path:
    return Path(context.output_dir) / "equation-editor-3-ole"


def _write_json(path: Path, payload: dict[str, object]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _canonicalization_summary_path(output_root: Path) -> Path:
    return output_root / CANONICALIZATION_SUMMARY_FILENAME


def _validation_evidence_path(output_root: Path) -> Path:
    return output_root / VALIDATION_EVIDENCE_FILENAME


def _canonical_mathml_dir(output_root: Path) -> Path:
    return output_root / "canonical-mathml"


def _is_legacy_doc_input(input_path: Path) -> bool:
    return input_path.suffix.lower() == ".doc"


def _read_legacy_doc_stream(input_path: Path, stream_name: str) -> bytes:
    if olefile is None:
        raise Equation3MtefError("olefile dependency is required for legacy .doc Equation3 ingestion.")
    with olefile.OleFileIO(input_path) as ole:
        return ole.openstream(stream_name.split("/")).read()


def _read_record_payload(
    input_path: Path,
    record: dict[str, object],
    zip_file: zipfile.ZipFile | None,
) -> bytes:
    embedding_target = str(record.get("embedding_target") or "")
    if _is_legacy_doc_input(input_path):
        if not embedding_target:
            raise Equation3MtefError("Legacy .doc Equation3 record is missing an OLE stream target.")
        return _read_legacy_doc_stream(input_path, embedding_target)

    if zip_file is None:
        raise Equation3MtefError("DOCX Equation3 record requires an open ZIP container.")
    if not embedding_target or embedding_target not in zip_file.namelist():
        raise Equation3MtefError(f"Missing DOCX embedding target: {embedding_target}")
    return zip_file.read(embedding_target)


def _canonicalize_detected_equation3(step: ExecutionStep, context: ExecutionContext) -> dict[str, object]:
    input_path = Path(context.input_path)
    output_root = _execution_output_root(context)
    canonical_dir = _canonical_mathml_dir(output_root)
    canonical_dir.mkdir(parents=True, exist_ok=True)

    records = detect_equation_editor_3_ole(input_path)
    canonical_items: list[dict[str, object]] = []
    unsupported_items: list[dict[str, object]] = []

    zip_file: zipfile.ZipFile | None = None
    if not _is_legacy_doc_input(input_path):
        zip_file = zipfile.ZipFile(input_path)
    try:
        for index, record in enumerate(records, start=1):
            formula_id = f"equation3-canonical-{index:04d}"
            embedding_target = str(record.get("embedding_target") or "")
            provenance = record.get("provenance", {})
            source_specific = record.get("source_specific", {}).get("equation_editor_3", {})

            if record.get("source_role") != "native-source":
                unsupported_items.append(
                    {
                        "formula_id": record.get("formula_id", formula_id),
                        "reason": "not-native-source",
                        "source_role": record.get("source_role"),
                    }
                )
                continue
            if source_specific.get("mtef_version") != 3:
                unsupported_items.append(
                    {
                        "formula_id": record.get("formula_id", formula_id),
                        "reason": "unsupported-mtef-version",
                        "mtef_version": source_specific.get("mtef_version"),
                    }
                )
                continue
            try:
                payload_data = _read_record_payload(input_path, record, zip_file)
                native_payload, result = convert_equation3_payload_to_mathml(
                    payload_data,
                    preferred_stream_name=str(provenance.get("payload_stream_name") or "") or None,
                )
                target_path = canonical_dir / f"{formula_id}.xml"
                target_path.write_text(result.mathml_text, encoding="utf-8")
                root = ET.fromstring(result.mathml_text)
                if local_name(root.tag) != "math":
                    raise Equation3MtefError("Canonical MathML root is not math.")
            except Exception as exc:  # noqa: BLE001 - artifact gate must capture exact unsupported formula.
                unsupported_items.append(
                    {
                        "formula_id": record.get("formula_id", formula_id),
                        "reason": f"{type(exc).__name__}: {exc}",
                        "embedding_target": embedding_target,
                    }
                )
                continue

            mathml_text = target_path.read_text(encoding="utf-8")
            property_signals = mathml_property_signals(root)
            canonical_items.append(
                {
                    "formula_id": formula_id,
                    "source_formula_id": record.get("formula_id", ""),
                    "doc_part_path": record.get("doc_part_path", ""),
                    "relationship_id": record.get("relationship_id", ""),
                    "embedding_target": embedding_target,
                    "payload_stream_name": native_payload.stream_name,
                    "raw_payload_sha256": provenance.get("raw_payload_sha256") or native_payload.source_stream_sha256,
                    "equation_native_sha256": native_payload.equation_native_sha256,
                    "mtef_payload_sha256": result.mtef_payload_sha256,
                    "canonical_artifact_path": str(target_path),
                    "canonical_sha256": sha256_text(mathml_text),
                    "preservation_status": "converted-equation3-mtef-v3-to-canonical-mathml-limited",
                    "mtef_version": result.mtef_version,
                    "mtef_product": result.product,
                    "mtef_product_version": result.product_version,
                    "mtef_product_subversion": result.product_subversion,
                    "record_counts": result.record_counts,
                    "template_selector_counts": result.template_selector_counts,
                    "parsed_bytes": result.parsed_bytes,
                    "mtef_payload_bytes": result.mtef_payload_bytes,
                    "property_signals": property_signals,
                }
            )
    finally:
        if zip_file is not None:
            zip_file.close()

    canonical_count = len(canonical_items)
    expected_count = int(step.formula_count or len(records))
    unsupported_count = len(unsupported_items)
    formula_count_parity = "passed" if expected_count == canonical_count and unsupported_count == 0 else "failed"
    gate_status = "passed-limited" if canonical_count > 0 and formula_count_parity == "passed" else "blocked-canonical-artifact"

    summary = {
        "artifact_type": "equation3-canonicalization-summary",
        "source_family": SOURCE_FAMILY,
        "target_format": "canonical-mathml",
        "target_stage": "equation3-mtef-v3-to-canonical-mathml-limited",
        "gate_status": gate_status,
        "limited_conversion_claim": gate_status == "passed-limited",
        "general_converter_claim": False,
        "deliverability_claim": False,
        "word_visual_fill_back_claim": False,
        "input_path": context.input_path,
        "detected_formula_count": len(records),
        "expected_formula_count": expected_count,
        "canonical_mathml_count": canonical_count,
        "unsupported_fragment_count": unsupported_count,
        "formula_count_parity": formula_count_parity,
        "canonical_mathml_dir": str(canonical_dir),
        "source_to_canonical_provenance": canonical_items,
        "unsupported_fragments": unsupported_items,
        "property_summary": property_summary(canonical_items),
        "supported_slice": equation3_fixture_admissibility_requirements()["current_productized_slice"],
        "claim_boundary": {
            "accepted": (
                "MTEF v3 Equation Native payloads using the implemented script, root, fraction, slash-fraction, bar, fence, limit, matrix, BigOp, and character slice can be converted to canonical MathML.",
            ),
            "not_accepted": (
                "Universal Equation Editor 3.0 support.",
                "Layout/order fidelity for every legacy .doc object arrangement.",
                "Public redistribution eligibility for research-control samples.",
                "Word visual fill-back or DOCX-route deliverability.",
            ),
        },
    }
    _write_json(_canonicalization_summary_path(output_root), summary)
    _write_json(
        _validation_evidence_path(output_root),
        {
            "artifact_type": "equation3-validation-evidence",
            "source_family": SOURCE_FAMILY,
            "status": gate_status,
            "canonicalization_summary_path": str(_canonicalization_summary_path(output_root)),
            "canonical_mathml_count": canonical_count,
            "unsupported_fragment_count": unsupported_count,
            "formula_count_parity": formula_count_parity,
            "property_summary": summary["property_summary"],
            "claim_boundary": summary["claim_boundary"],
        },
    )
    return summary


def _action_summaries(step: ExecutionStep) -> tuple[dict[str, object], ...]:
    summaries: list[dict[str, object]] = []
    for action in step.actions:
        if action.action_id == "probe-header-and-classid":
            summaries.append(
                {
                    "action_id": action.action_id,
                    "description": action.description,
                    "status": "review-gated",
                    "supported": True,
                    "blocking": action.blocking,
                }
            )
            continue

        if action.action_id == "attempt-mtef-conversion":
            summaries.append(
                {
                    "action_id": action.action_id,
                    "description": action.description,
                    "status": "manual-gate",
                    "supported": False,
                    "blocking": action.blocking,
                }
            )
            continue

        if action.action_id == "fallback-manual-triage":
            summaries.append(
                {
                    "action_id": action.action_id,
                    "description": action.description,
                    "status": "manual-gate",
                    "supported": False,
                    "blocking": action.blocking,
                }
            )
            continue

        if action.action_id == "word-roundtrip-validation":
            summaries.append(
                {
                    "action_id": action.action_id,
                    "description": action.description,
                    "status": "review-gated",
                    "supported": False,
                    "blocking": action.blocking,
                }
            )
            continue

        summaries.append(
            {
                "action_id": action.action_id,
                "description": action.description,
                "status": "manual-gate",
                "supported": False,
                "blocking": action.blocking,
            }
        )

    return tuple(summaries)


def _write_blocker_record(
    step: ExecutionStep,
    context: ExecutionContext,
    *,
    conversion_attempt: dict[str, object] | None = None,
) -> Path:
    output_root = _execution_output_root(context)
    blocker_record_path = output_root / BLOCKER_RECORD_FILENAME
    return _write_json(
        blocker_record_path,
        {
            "artifact_type": "equation3-blocker-record",
            "provider": "equation3",
            "source_family": SOURCE_FAMILY,
            "canonical_target": canonical_mathml_contract_for_source_family(SOURCE_FAMILY).to_dict(),
            "status": "blocked",
            "blocking": True,
            "supported": False,
            "conversion_claim": False,
            "fixture_status": "insufficient",
            "fixture_gap": "Equation Editor 3.0 fixture coverage is not strong enough for a deliverable conversion claim.",
            "next_ready_condition": "Add stronger Equation Editor 3.0 fixtures that exercise probe evidence and conversion output, then rerun execute.",
            "fixture_admissibility": equation3_fixture_admissibility_requirements(),
            "execution_plan_path": context.execution_plan_path,
            "input_path": context.input_path,
            "output_root": str(output_root),
            "formula_count": step.formula_count,
            "manual_review_required": step.requires_manual_review,
            "probe": {
                "runner": RUNNER,
                "action_id": "probe-header-and-classid",
                "signals": ["prog-id", "class-id", "eqnolefilehdr", "mtef-v3-header"],
            },
            "actions": list(_action_summaries(step)),
            "conversion_attempt": conversion_attempt or {},
        },
    )


def _probe_dry_run_report(
    step: ExecutionStep,
    action_id: str,
    description: str,
    context: DryRunContext,
) -> DryRunActionReport:
    output_root = _dry_run_output_root(context)
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=False,
        supported=True,
        status="ready",
        runner=RUNNER,
        argv=(
            "probe-equation3-evidence",
            "--execution-plan",
            context.execution_plan_path,
            "--output-dir",
            str(output_root),
            "--signals",
            "prog-id,class-id,eqnolefilehdr,mtef-v3-header",
        ),
        cwd=context.workspace_root,
        notes=(
            "Confirm Equation.3, ClassID, and EQNOLEFILEHDR/MTEF v3 evidence.",
            "The internal converter can now attempt the limited MTEF v3 script, root, fraction, slash-fraction, bar, fence, limit, matrix, BigOp, and character slice.",
            "This action does not claim universal Equation Editor 3.0 support or deliverable Word output.",
            f"Formula count from execution plan: {step.formula_count}.",
        ),
    )


def _manual_gate_dry_run_report(
    action_id: str,
    description: str,
    blocking: bool,
    context: DryRunContext,
    *,
    note: str,
) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=blocking,
        supported=False,
        status="manual-gate",
        runner="manual",
        cwd=context.workspace_root,
        notes=(
            note,
            "Real Equation Editor 3.0 fixture coverage is not strong enough to advertise an executable conversion binding.",
            "Use the blocker-record fixture admissibility checklist before promoting this source line.",
        ),
    )


def _review_gated_dry_run_report(
    action_id: str,
    description: str,
    blocking: bool,
    context: DryRunContext,
) -> DryRunActionReport:
    return DryRunActionReport(
        action_id=action_id,
        description=description,
        blocking=blocking,
        supported=False,
        status="review-gated",
        runner="manual-validation",
        cwd=context.workspace_root,
        notes=(
            "Word roundtrip validation requires fixture-backed conversion output and human review before delivery claims.",
            "No automated DOCX/PDF validation binding is available for Equation Editor 3.0 yet.",
        ),
    )


def build_equation3_dry_run_reports(
    step: ExecutionStep,
    context: DryRunContext,
) -> tuple[DryRunActionReport, ...]:
    if step.source_family != SOURCE_FAMILY:
        raise ValueError(f"Equation Editor 3.0 dry-run binding cannot handle source_family={step.source_family!r}.")

    if not step.actions:
        return (
            _manual_gate_dry_run_report(
                "manual-triage",
                "No Equation Editor 3.0 execution actions are registered for this step.",
                True,
                context,
                note="Manual triage is required before any Equation Editor 3.0 conversion attempt.",
            ),
        )

    reports: list[DryRunActionReport] = []
    for action in step.actions:
        action_id = action.action_id
        if action_id == "probe-header-and-classid":
            reports.append(_probe_dry_run_report(step, action_id, action.description, context))
            continue

        if action_id == "attempt-mtef-conversion":
            output_root = _dry_run_output_root(context)
            reports.append(
                DryRunActionReport(
                    action_id=action_id,
                    description=action.description,
                    blocking=action.blocking,
                    supported=True,
                    status="ready",
                    runner="internal-equation3-mtef-v3-limited",
                    argv=(
                        "convert-equation3-mtef-v3",
                        "--input",
                        "<execution-plan-input>",
                        "--output-dir",
                        str(output_root),
                    ),
                    cwd=context.workspace_root,
                    notes=(
                        "Attempts canonical MathML conversion for the implemented MTEF v3 script, root, fraction, slash-fraction, bar, fence, limit, matrix, BigOp, and character slice.",
                        "Unsupported records are reported as canonical artifact blockers instead of guessed.",
                        "This is not a universal Equation Editor 3.0 converter claim.",
                    ),
                )
            )
            continue

        if action_id == "fallback-manual-triage":
            reports.append(
                DryRunActionReport(
                    action_id=action_id,
                    description=action.description,
                    blocking=False,
                    supported=True,
                    status="skipped-until-needed",
                    runner="manual",
                    cwd=context.workspace_root,
                    notes=("Manual triage is used only if MTEF parsing or canonical validation is blocked.",),
                )
            )
            continue

        if action_id == "word-roundtrip-validation":
            reports.append(
                DryRunActionReport(
                    action_id=action_id,
                    description=action.description,
                    blocking=False,
                    supported=True,
                    status="skipped-downstream",
                    runner="manual-validation",
                    cwd=context.workspace_root,
                    notes=(
                        "Word roundtrip validation is downstream of the canonical MathML target and is not run in this source-core step.",
                    ),
                )
            )
            continue

        reports.append(
            _manual_gate_dry_run_report(
                action_id,
                action.description,
                True,
                context,
                note="Unrecognized Equation Editor 3.0 action; manual triage is required before binding execution.",
            )
        )

    return tuple(reports)


def _probe_execution_report(
    step: ExecutionStep,
    action_id: str,
    description: str,
    context: ExecutionContext,
    *,
    output_path: Path,
) -> ActionExecutionReport:
    output_root = _execution_output_root(context)
    return ActionExecutionReport(
        action_id=action_id,
        description=description,
        blocking=False,
        supported=True,
        status="review-gated",
        runner=RUNNER,
        argv=(
            "probe-equation3-evidence",
            "--input",
            context.input_path,
            "--execution-plan",
            context.execution_plan_path,
            "--output-dir",
            str(output_root),
        ),
        cwd=context.workspace_root,
        output_paths=(str(output_path),),
        notes=(
            "Execution probes Equation.3 source identity and MTEF v3 header evidence.",
            "Payload conversion is handled by the limited canonical MathML action.",
            "No OMML replacement or deliverable Word artifact is produced by this provider.",
            f"Blocker record written to {output_path}.",
            f"Formula count from execution plan: {step.formula_count}.",
        ),
    )


def _manual_gate_execution_report(
    action_id: str,
    description: str,
    blocking: bool,
    context: ExecutionContext,
    *,
    note: str,
    output_path: Path,
) -> ActionExecutionReport:
    return ActionExecutionReport(
        action_id=action_id,
        description=description,
        blocking=blocking,
        supported=False,
        status="manual-gate",
        runner="manual",
        cwd=context.workspace_root,
        output_paths=(str(output_path),),
        notes=(
            note,
            "Execution report is a gate record only; it is not a successful conversion result.",
            f"Blocker record written to {output_path}.",
        ),
    )


def _review_gated_execution_report(
    action_id: str,
    description: str,
    blocking: bool,
    context: ExecutionContext,
    *,
    output_path: Path,
) -> ActionExecutionReport:
    return ActionExecutionReport(
        action_id=action_id,
        description=description,
        blocking=blocking,
        supported=False,
        status="review-gated",
        runner="manual-validation",
        cwd=context.workspace_root,
        output_paths=(str(output_path),),
        notes=(
            "Word open/PDF export/visual parity checks remain review-gated for Equation Editor 3.0.",
            "This status is not a deliverable conversion proof.",
            f"Blocker record written to {output_path}.",
        ),
    )


def execute_equation3_step(
    step: ExecutionStep,
    context: ExecutionContext,
) -> tuple[ActionExecutionReport, ...]:
    if step.source_family != SOURCE_FAMILY:
        raise ValueError(f"Equation Editor 3.0 executor cannot handle source_family={step.source_family!r}.")

    conversion_summary: dict[str, object] | None = None
    conversion_error = ""
    try:
        conversion_summary = _canonicalize_detected_equation3(step, context)
    except Exception as exc:  # noqa: BLE001 - convert attempt is reported as a gate, not raised.
        conversion_error = f"{type(exc).__name__}: {exc}"

    conversion_passed = bool(conversion_summary and conversion_summary.get("gate_status") == "passed-limited")
    blocker_record_path: Path | None = None
    if not conversion_passed:
        blocker_record_path = _write_blocker_record(
            step,
            context,
            conversion_attempt=conversion_summary or {"status": "failed-before-summary", "error": conversion_error},
        )

    if not step.actions:
        return (
            _manual_gate_execution_report(
                "manual-triage",
                "No Equation Editor 3.0 execution actions are registered for this step.",
                True,
                context,
                note="Manual triage is required before Equation Editor 3.0 execution.",
                output_path=blocker_record_path or _canonicalization_summary_path(_execution_output_root(context)),
            ),
        )

    reports: list[ActionExecutionReport] = []
    for action in step.actions:
        action_id = action.action_id
        if action_id == "probe-header-and-classid":
            reports.append(
                ActionExecutionReport(
                    action_id=action_id,
                    description=action.description,
                    blocking=False,
                    supported=True,
                    status="completed" if conversion_passed else "review-gated",
                    runner=RUNNER,
                    argv=(
                        "probe-equation3-evidence",
                        "--input",
                        context.input_path,
                        "--execution-plan",
                        context.execution_plan_path,
                        "--output-dir",
                        str(_execution_output_root(context)),
                    ),
                    cwd=context.workspace_root,
                    output_paths=(
                        str(_canonicalization_summary_path(_execution_output_root(context)))
                        if conversion_passed
                        else str(blocker_record_path),
                    ),
                    notes=(
                        "Equation3 source identity and MTEF v3 evidence were evaluated.",
                        f"Formula count from execution plan: {step.formula_count}.",
                    ),
                )
            )
            continue

        if action_id == "attempt-mtef-conversion":
            if conversion_passed:
                output_root = _execution_output_root(context)
                reports.append(
                    ActionExecutionReport(
                        action_id=action_id,
                        description=action.description,
                        blocking=False,
                        supported=True,
                        status="completed",
                        runner="internal-equation3-mtef-v3-limited",
                        argv=(
                            "convert-equation3-mtef-v3",
                            "--input",
                            context.input_path,
                            "--output-dir",
                            str(output_root),
                        ),
                        cwd=context.workspace_root,
                        output_paths=(
                            str(_canonicalization_summary_path(output_root)),
                            str(_validation_evidence_path(output_root)),
                            str(_canonical_mathml_dir(output_root)),
                        ),
                        notes=(
                            "Converted supported Equation3 MTEF v3 payloads to canonical MathML artifacts.",
                            "Formula-count parity passed for this input.",
                            "This remains a limited observed-structure converter, not universal Equation Editor 3.0 support.",
                        ),
                    )
                )
            else:
                reports.append(
                    _manual_gate_execution_report(
                        action_id,
                        action.description,
                        True,
                        context,
                        note="MTEF v3 conversion did not satisfy the limited canonical MathML gate.",
                        output_path=blocker_record_path or _canonicalization_summary_path(_execution_output_root(context)),
                    )
                )
            continue

        if action_id == "fallback-manual-triage":
            if conversion_passed:
                reports.append(
                    ActionExecutionReport(
                        action_id=action_id,
                        description=action.description,
                        blocking=False,
                        supported=True,
                        status="skipped-not-needed",
                        runner="manual",
                        cwd=context.workspace_root,
                        output_paths=(str(_canonicalization_summary_path(_execution_output_root(context))),),
                        notes=("Manual triage was not needed because the limited canonical MathML gate passed.",),
                    )
                )
            else:
                reports.append(
                    _manual_gate_execution_report(
                        action_id,
                        action.description,
                        True,
                        context,
                        note="Manual triage remains required for ambiguous or unsupported Equation Editor 3.0 evidence.",
                        output_path=blocker_record_path or _canonicalization_summary_path(_execution_output_root(context)),
                    )
                )
            continue

        if action_id == "word-roundtrip-validation":
            if conversion_passed:
                reports.append(
                    ActionExecutionReport(
                        action_id=action_id,
                        description=action.description,
                        blocking=False,
                        supported=True,
                        status="skipped-downstream",
                        runner="manual-validation",
                        cwd=context.workspace_root,
                        output_paths=(str(_validation_evidence_path(_execution_output_root(context))),),
                        notes=(
                            "Word roundtrip validation is downstream of canonical MathML and was intentionally skipped.",
                            "This is not a DOCX deliverability claim.",
                        ),
                    )
                )
            else:
                reports.append(
                    _review_gated_execution_report(
                        action_id,
                        action.description,
                        True,
                        context,
                        output_path=blocker_record_path or _canonicalization_summary_path(_execution_output_root(context)),
                    )
                )
            continue

        reports.append(
            _manual_gate_execution_report(
                action_id,
                action.description,
                True,
                context,
                note="Unrecognized Equation Editor 3.0 action; manual triage is required.",
                output_path=blocker_record_path or _canonicalization_summary_path(_execution_output_root(context)),
            )
        )

    return tuple(reports)
