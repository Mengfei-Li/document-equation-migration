import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

import document_equation_migration.detectors.equation_editor_3_ole as equation3_detector_module
import document_equation_migration.executor.equation3 as equation3_executor_module
from document_equation_migration.execution_plan.model import ExecutionAction, ExecutionStep
from document_equation_migration.executor.equation3 import (
    build_equation3_dry_run_reports,
    equation3_fixture_admissibility_requirements,
    execute_equation3_step,
)
from document_equation_migration.executor.model import DryRunContext, ExecutionContext
from document_equation_migration.equation3_mtef import (
    EQNOLEFILEHDR_SIZE,
    Equation3MathMLResult,
    NativePayload,
    local_name,
)


FIXTURE_ROOT = Path(__file__).resolve().parent / "fixtures" / "equation_editor_3_ole"


def _equation3_step(*, source_family: str = "equation-editor-3-ole", formula_count: int = 2) -> ExecutionStep:
    return ExecutionStep(
        source_family=source_family,
        formula_count=formula_count,
        route_kind="primary-candidate",
        confidence_policy="medium",
        requires_manual_review=True,
        provider="equation3",
        next_action="run-equation3-probe-and-conversion",
        actions=(
            ExecutionAction(
                action_id="probe-header-and-classid",
                description="Probe OLE header and ClassID to confirm Equation Editor 3.0 payload.",
            ),
            ExecutionAction(
                action_id="attempt-mtef-conversion",
                description="Attempt MTEF-oriented conversion as primary Equation Editor 3.0 path.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="fallback-manual-triage",
                description="Fallback to manual triage when probe or conversion results are ambiguous.",
                blocking=True,
            ),
            ExecutionAction(
                action_id="word-roundtrip-validation",
                description="Validate converted output through Word roundtrip before delivery.",
                blocking=True,
            ),
        ),
        notes=("Equation3 sample",),
    )


def _dry_run_context(tmp_path: Path) -> DryRunContext:
    return DryRunContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(tmp_path / "execution-plan.json"),
        output_dir_hint=str(tmp_path / "out"),
    )


def _execution_context(tmp_path: Path) -> ExecutionContext:
    return ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(tmp_path / "execution-plan.json"),
        input_path=str(tmp_path / "input.docx"),
        output_dir=str(tmp_path / "out"),
    )


def _typeface_byte(typeface: int) -> int:
    return (typeface - 128) % 256


def _char(codepoint: int, *, typeface: int = 3, options: int = 1) -> bytes:
    return bytes([(options << 4) | 2, _typeface_byte(typeface), codepoint & 0xFF, codepoint >> 8])


def _subscript(slot: bytes) -> bytes:
    return b"\x03\x0f\x01\x00" + b"\x0b" + b"\x01" + slot + b"\x00" + b"\x11" + b"\x00"


def _supported_equation_native_stream() -> bytes:
    expression = (
        b"\x0a"
        + b"\x01"
        + _char(ord("b"))
        + _subscript(_char(ord("k")))
        + b"\x0a"
        + _char(ord("="), typeface=6, options=0)
        + _char(ord("a"))
        + _subscript(_char(ord("k")))
        + b"\x00"
        + b"\x00"
    )
    return bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression


def _write_supported_equation3_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", (FIXTURE_ROOT / "document_equation3.xml").read_text(encoding="utf-8"))
        zf.writestr(
            "word/_rels/document.xml.rels",
            (FIXTURE_ROOT / "document_equation3.rels.xml").read_text(encoding="utf-8"),
        )
        zf.writestr("word/embeddings/oleObject1.bin", _supported_equation_native_stream())
        zf.writestr("word/media/image1.wmf", b"WMF-SYNTHETIC")


def test_equation3_dry_run_is_provider_binding_not_generic_fallback(tmp_path: Path) -> None:
    reports = build_equation3_dry_run_reports(_equation3_step(), _dry_run_context(tmp_path))

    assert [report.action_id for report in reports] == [
        "probe-header-and-classid",
        "attempt-mtef-conversion",
        "fallback-manual-triage",
        "word-roundtrip-validation",
    ]
    assert reports[0].supported is True
    assert reports[0].status == "ready"
    assert reports[0].runner == "internal-equation3-probe"
    assert reports[0].argv[0] == "probe-equation3-evidence"
    assert "limited MTEF v2/v3 script, root, fraction, slash-fraction, bar, fence, limit, matrix, pile" in (
        "\n".join(reports[0].notes)
    )
    assert reports[1].supported is True
    assert reports[1].status == "ready"
    assert reports[1].runner == "internal-equation3-mtef-v2v3-limited"
    assert "limit, matrix, pile, BigOp, standalone sum operator, character, and narrow legacy footer slice" in "\n".join(
        reports[1].notes
    )
    assert "not a universal Equation Editor 3.0 converter claim" in "\n".join(reports[1].notes)
    assert reports[2].status == "skipped-until-needed"
    assert reports[3].status == "skipped-downstream"

    combined_notes = "\n".join("\n".join(report.notes) for report in reports)
    assert "No concrete dry-run binding is registered" not in combined_notes
    assert "generic fallback" not in combined_notes.lower()


def test_equation3_execute_writes_blocker_record_and_keeps_gate_status(tmp_path: Path) -> None:
    reports = execute_equation3_step(_equation3_step(), _execution_context(tmp_path))

    assert [report.status for report in reports] == [
        "review-gated",
        "manual-gate",
        "manual-gate",
        "review-gated",
    ]
    assert reports[0].supported is True
    assert reports[0].runner == "internal-equation3-probe"
    assert reports[1].supported is False
    assert reports[3].runner == "manual-validation"
    assert all(report.status != "completed" for report in reports)
    assert len({report.output_paths for report in reports}) == 1

    blocker_record_path = Path(reports[0].output_paths[0])
    assert blocker_record_path.name == "blocker-record.json"
    assert blocker_record_path.exists()

    blocker_record = json.loads(blocker_record_path.read_text(encoding="utf-8"))
    assert blocker_record["artifact_type"] == "equation3-blocker-record"
    assert blocker_record["provider"] == "equation3"
    assert blocker_record["source_family"] == "equation-editor-3-ole"
    assert blocker_record["status"] == "blocked"
    assert blocker_record["blocking"] is True
    assert blocker_record["conversion_claim"] is False
    assert blocker_record["fixture_status"] == "insufficient"
    assert "deliverable conversion claim" in blocker_record["fixture_gap"]
    assert "stronger Equation Editor 3.0 fixtures" in blocker_record["next_ready_condition"]
    assert blocker_record["fixture_admissibility"]["target_stage"] == "fixture-backed-canonical-mathml-conversion"
    required_property_ids = {
        item["id"] for item in blocker_record["fixture_admissibility"]["required_candidate_properties"]
    }
    assert required_property_ids == {
        "equation3-identity",
        "native-payload",
        "mtef-header",
        "canonical-output",
        "provenance-map",
    }
    assert "preview-only fixture" in blocker_record["fixture_admissibility"]["disqualifying_conditions"]
    assert any(
        "Canonical MathML output validates" in gate
        for gate in blocker_record["fixture_admissibility"]["promotion_gate"]
    )
    assert blocker_record["probe"]["runner"] == "internal-equation3-probe"
    assert blocker_record["probe"]["signals"] == ["prog-id", "class-id", "eqnolefilehdr", "mtef-header"]
    assert [action["action_id"] for action in blocker_record["actions"]] == [
        "probe-header-and-classid",
        "attempt-mtef-conversion",
        "fallback-manual-triage",
        "word-roundtrip-validation",
    ]
    assert blocker_record["actions"][0]["status"] == "review-gated"
    assert blocker_record["actions"][1]["status"] == "manual-gate"
    assert blocker_record["actions"][3]["supported"] is False

    combined_notes = "\n".join("\n".join(report.notes) for report in reports)
    assert "MTEF v2/v3 conversion did not satisfy" in combined_notes
    assert "gate record only" in combined_notes
    assert "Blocker record written to" in combined_notes
    assert "deliverable conversion proof" in combined_notes


def test_equation3_execute_writes_limited_canonical_mathml_for_supported_payload(tmp_path: Path) -> None:
    input_path = tmp_path / "supported-equation3.docx"
    _write_supported_equation3_docx(input_path)
    context = ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(tmp_path / "execution-plan.json"),
        input_path=str(input_path),
        output_dir=str(tmp_path / "out"),
    )

    reports = execute_equation3_step(_equation3_step(formula_count=1), context)

    assert [report.status for report in reports] == [
        "completed",
        "completed",
        "skipped-not-needed",
        "skipped-downstream",
    ]
    assert all(report.supported for report in reports)
    output_root = tmp_path / "out" / "equation-editor-3-ole"
    assert not (output_root / "blocker-record.json").exists()
    summary = json.loads((output_root / "canonicalization-summary.json").read_text(encoding="utf-8"))
    assert summary["gate_status"] == "passed-limited"
    assert summary["limited_conversion_claim"] is True
    assert summary["general_converter_claim"] is False
    assert summary["deliverability_claim"] is False
    assert summary["canonical_mathml_count"] == 1
    assert summary["formula_count_parity"] == "passed"
    assert summary["source_to_canonical_provenance"][0]["preservation_status"] == (
        "converted-equation3-mtef-v2v3-to-canonical-mathml-limited"
    )
    assert summary["source_to_canonical_provenance"][0]["typeface_counts"] == {
        "3:fnVARIABLE": 4,
        "6:fnSYMBOL": 1,
    }
    canonical_path = output_root / "canonical-mathml" / "equation3-canonical-0001.xml"
    root = ET.parse(canonical_path).getroot()
    assert local_name(root.tag) == "math"
    assert "".join(root.itertext()) == "bk=ak"


def test_equation3_execute_does_not_write_invalid_mathml_artifact(monkeypatch, tmp_path: Path) -> None:
    input_path = tmp_path / "invalid-equation3.docx"
    _write_supported_equation3_docx(input_path)
    context = ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(tmp_path / "execution-plan.json"),
        input_path=str(input_path),
        output_dir=str(tmp_path / "out"),
    )

    def fake_convert(payload_data: bytes, *, preferred_stream_name: str | None = None):
        return NativePayload(
            raw_payload=payload_data,
            equation_native_stream=payload_data,
            stream_name=preferred_stream_name or "",
            source_stream_sha256="raw",
            equation_native_sha256="native",
        ), Equation3MathMLResult(
            mathml_text=(
                '<?xml version="1.0" encoding="UTF-8"?>\n'
                f'<math xmlns="http://www.w3.org/1998/Math/MathML">{chr(2)}</math>\n'
            ),
            mtef_version=3,
            platform=1,
            product=1,
            product_version=3,
            product_subversion=0,
            record_counts={},
            template_selector_counts={},
            typeface_counts={},
            parsed_bytes=0,
            mtef_payload_bytes=0,
            mtef_payload_sha256="mtef",
        )

    monkeypatch.setattr(equation3_executor_module, "convert_equation3_payload_to_mathml", fake_convert)

    execute_equation3_step(_equation3_step(formula_count=1), context)

    output_root = tmp_path / "out" / "equation-editor-3-ole"
    summary = json.loads((output_root / "canonicalization-summary.json").read_text(encoding="utf-8"))
    assert summary["canonical_mathml_count"] == 0
    assert summary["unsupported_fragment_count"] == 1
    assert list((output_root / "canonical-mathml").glob("*.xml")) == []


def test_equation3_execute_reads_legacy_doc_equation_native_stream(tmp_path: Path) -> None:
    native_stream = _supported_equation_native_stream()

    class FakeStream:
        def __init__(self, data: bytes) -> None:
            self.data = data

        def read(self) -> bytes:
            return self.data

    class FakeOle:
        streams = {
            "WordDocument": b"EMBED Equation.3",
            "ObjectPool/_1/\x01CompObj": b"Microsoft Equation 3.0\x00Equation.3\x00",
            "ObjectPool/_1/Equation Native": native_stream,
        }

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb) -> None:
            return None

        def listdir(self):
            return [name.split("/") for name in self.streams]

        def openstream(self, name):
            key = "/".join(name) if isinstance(name, list) else name
            return FakeStream(self.streams[key])

    class FakeOlefile:
        @staticmethod
        def isOleFile(path) -> bool:
            return True

        @staticmethod
        def OleFileIO(path) -> FakeOle:
            return FakeOle()

    input_path = tmp_path / "legacy.doc"
    input_path.write_bytes(b"OLE-CFB")
    context = ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(tmp_path / "execution-plan.json"),
        input_path=str(input_path),
        output_dir=str(tmp_path / "out"),
    )
    original_detector_olefile = equation3_detector_module.olefile
    original_executor_olefile = equation3_executor_module.olefile
    equation3_detector_module.olefile = FakeOlefile
    equation3_executor_module.olefile = FakeOlefile
    try:
        reports = execute_equation3_step(_equation3_step(formula_count=1), context)
    finally:
        equation3_detector_module.olefile = original_detector_olefile
        equation3_executor_module.olefile = original_executor_olefile

    assert [report.status for report in reports] == [
        "completed",
        "completed",
        "skipped-not-needed",
        "skipped-downstream",
    ]
    output_root = tmp_path / "out" / "equation-editor-3-ole"
    summary = json.loads((output_root / "canonicalization-summary.json").read_text(encoding="utf-8"))
    assert summary["input_path"] == str(input_path)
    assert summary["canonical_mathml_count"] == 1
    provenance = summary["source_to_canonical_provenance"][0]
    assert provenance["embedding_target"] == "ObjectPool/_1/Equation Native"
    assert provenance["payload_stream_name"] == "ObjectPool/_1/Equation Native"
    assert summary["claim_boundary"]["not_accepted"][0] == "Universal Equation Editor 3.0 support."


def test_equation3_provider_rejects_wrong_source_family(tmp_path: Path) -> None:
    step = _equation3_step(source_family="mathtype-ole")

    with pytest.raises(ValueError):
        build_equation3_dry_run_reports(step, _dry_run_context(tmp_path))

    with pytest.raises(ValueError):
        execute_equation3_step(step, _execution_context(tmp_path))


def test_equation3_fixture_admissibility_keeps_public_promotion_gated() -> None:
    requirements = equation3_fixture_admissibility_requirements()

    assert requirements["target_stage"] == "fixture-backed-canonical-mathml-conversion"
    assert "real Equation.3" in requirements["minimum_fixture_set"]
    assert "preview-only fixture" in requirements["disqualifying_conditions"]
    assert "unclear redistribution or use permission for public fixture promotion" in requirements[
        "disqualifying_conditions"
    ]
    assert any("Canonical MathML output validates" in item for item in requirements["promotion_gate"])
    supported_records = "\n".join(requirements["current_productized_slice"]["supported_records"])
    assert "tmLIM_UPPER" in supported_records
    assert "tmISUM_LOWER" in supported_records
    assert "tmIPROD_BOTH" in supported_records
    assert "tmCOPROD_NO_LIMITS" in supported_records
    assert "tmINTOP_BOTH" in supported_records
    assert "tmSUMOP" in supported_records
    assert "pile records with supported line-based rows" in supported_records
    assert "observed short footer envelopes" in supported_records
