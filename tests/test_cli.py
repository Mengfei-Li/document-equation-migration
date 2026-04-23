import json
import zipfile
from pathlib import Path

import document_equation_migration.cli as cli_module
from document_equation_migration.cli import main
from document_equation_migration.execution_plan import build_execution_plan


DOCX_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>
    </w:p>
  </w:body>
</w:document>
"""

LIBREOFFICE_FIXTURE_ROOT = Path(__file__).resolve().parent / "fixtures" / "libreoffice_transformed"


def make_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", DOCX_XML)


def build_odf_archive(fixture_dir: Path, suffix: str, output_dir: Path) -> Path:
    archive_path = output_dir / f"{fixture_dir.name}{suffix}"
    with zipfile.ZipFile(archive_path, "w") as zf:
        mimetype_path = fixture_dir / "mimetype"
        if mimetype_path.exists():
            zf.writestr("mimetype", mimetype_path.read_text(encoding="utf-8"), compress_type=zipfile.ZIP_STORED)
        for path in sorted(fixture_dir.rglob("*")):
            if not path.is_file():
                continue
            relative_path = path.relative_to(fixture_dir).as_posix()
            if relative_path == "mimetype":
                continue
            zf.write(path, arcname=relative_path, compress_type=zipfile.ZIP_DEFLATED)
    return archive_path


def test_scan_writes_manifest_and_summary_files(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.docx"
    manifest_path = tmp_path / "out" / "manifest.json"
    summary_path = tmp_path / "out" / "summary.txt"
    routing_path = tmp_path / "out" / "routing.json"
    execution_plan_path = tmp_path / "out" / "execution-plan.json"
    make_docx(input_path)

    exit_code = main(
        [
            "scan",
            str(input_path),
            "--output",
            str(manifest_path),
            "--routing",
            str(routing_path),
            "--summary",
            str(summary_path),
            "--execution-plan",
            str(execution_plan_path),
        ]
    )

    assert exit_code == 0
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert payload["document"]["container_format"] == "docx"
    assert len(payload["formulas"]) == 1
    assert payload["formulas"][0]["source_family"] == "omml-native"
    routing = json.loads(routing_path.read_text(encoding="utf-8"))
    assert routing["formula_count"] == 1
    assert routing["recommended_sequence"] == ["omml-native"]
    assert routing["route_plan"][0]["next_action"] == "run-omml-native-pipeline"
    execution_plan = json.loads(execution_plan_path.read_text(encoding="utf-8"))
    assert execution_plan["formula_count"] == 1
    assert execution_plan["steps"][0]["source_family"] == "omml-native"
    assert execution_plan["steps"][0]["provider"] == "omml"
    assert execution_plan["steps"][0]["actions"]
    summary = summary_path.read_text(encoding="utf-8")
    assert "Document Equation Migration scan summary" in summary
    assert "format: docx" in summary
    assert "formula_count: 1" in summary


def test_run_plan_writes_dry_run_execution_report(tmp_path: Path) -> None:
    plan_path = tmp_path / "execution-plan.json"
    report_path = tmp_path / "out" / "execution-report.json"
    plan_payload = build_execution_plan(
        {
            "document_id": "sample",
            "input_path": "sample.docx",
            "detector_version": "0.1.0",
            "formula_count": 2,
            "recommended_sequence": ["mathtype-ole", "equation-editor-3-ole"],
            "route_plan": [
                {
                    "source_family": "mathtype-ole",
                    "formula_count": 1,
                    "route_kind": "primary-source-first",
                    "priority": 10,
                    "next_action": "run-mathtype-source-first-pipeline",
                    "confidence_policy": "high",
                    "requires_manual_review": False,
                },
                {
                    "source_family": "equation-editor-3-ole",
                    "formula_count": 1,
                    "route_kind": "primary-candidate",
                    "priority": 30,
                    "next_action": "run-equation3-probe-and-conversion",
                    "confidence_policy": "medium",
                    "requires_manual_review": True,
                },
            ],
        }
    ).to_dict()
    plan_path.write_text(json.dumps(plan_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    exit_code = main(
        [
            "run-plan",
            str(plan_path),
            "--dry-run",
            "--output",
            str(report_path),
        ]
    )

    assert exit_code == 0
    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert payload["mode"] == "dry-run"
    assert payload["step_count"] == 2
    assert payload["runnable_step_count"] == 1
    assert payload["steps"][0]["provider"] == "mathtype"
    assert payload["steps"][0]["status"] == "runnable"
    assert payload["steps"][1]["status"] == "manual-only"


def test_scan_routes_libreoffice_bridge_as_transformed_source(tmp_path: Path) -> None:
    input_path = build_odf_archive(LIBREOFFICE_FIXTURE_ROOT / "libreoffice_bridge", ".odt", tmp_path)
    manifest_path = tmp_path / "out" / "manifest.json"
    summary_path = tmp_path / "out" / "summary.txt"
    routing_path = tmp_path / "out" / "routing.json"
    execution_plan_path = tmp_path / "out" / "execution-plan.json"

    exit_code = main(
        [
            "scan",
            str(input_path),
            "--output",
            str(manifest_path),
            "--routing",
            str(routing_path),
            "--summary",
            str(summary_path),
            "--execution-plan",
            str(execution_plan_path),
        ]
    )

    assert exit_code == 0
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert payload["document"]["container_format"] == "odt"
    assert payload["document"]["source_counts"] == {"libreoffice-transformed": 1}
    assert not any("Detector runtime warning" in note for note in payload["document"]["notes"])
    assert len(payload["formulas"]) == 1
    assert payload["formulas"][0]["source_family"] == "libreoffice-transformed"
    assert payload["formulas"][0]["provenance"]["generator_id"] == "libreoffice"
    assert payload["formulas"][0]["provenance"]["generator_raw"].startswith("LibreOffice/")
    routing = json.loads(routing_path.read_text(encoding="utf-8"))
    assert routing["formula_count"] == 1
    assert routing["recommended_sequence"] == ["libreoffice-transformed"]
    assert routing["route_plan"][0]["route_kind"] == "bridge-source"
    execution_plan = json.loads(execution_plan_path.read_text(encoding="utf-8"))
    assert execution_plan["formula_count"] == 1
    assert execution_plan["steps"][0]["source_family"] == "libreoffice-transformed"
    summary = summary_path.read_text(encoding="utf-8")
    assert "libreoffice-transformed: 1" in summary


def test_run_plan_dry_run_surfaces_mathtype_guarded_layout_args(tmp_path: Path) -> None:
    plan_path = tmp_path / "execution-plan.json"
    report_path = tmp_path / "out" / "execution-report.json"
    plan_payload = build_execution_plan(
        {
            "document_id": "sample",
            "input_path": "sample.docx",
            "detector_version": "0.1.0",
            "formula_count": 1,
            "recommended_sequence": ["mathtype-ole"],
            "route_plan": [
                {
                    "source_family": "mathtype-ole",
                    "formula_count": 1,
                    "route_kind": "primary-source-first",
                    "priority": 10,
                    "next_action": "run-mathtype-source-first-pipeline",
                    "confidence_policy": "high",
                    "requires_manual_review": False,
                    "experimental_options": {
                        "preserve_mathtype_layout": True,
                        "mathtype_layout_factor": 1.02,
                        "resume_mathtype_pipeline": True,
                        "mathtype_start_index": 216,
                        "mathtype_end_index": 238,
                    },
                }
            ],
        }
    ).to_dict()
    plan_path.write_text(json.dumps(plan_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    exit_code = main(
        [
            "run-plan",
            str(plan_path),
            "--dry-run",
            "--output",
            str(report_path),
        ]
    )

    assert exit_code == 0
    payload = json.loads(report_path.read_text(encoding="utf-8"))
    extract_action = payload["steps"][0]["actions"][0]
    replace_action = payload["steps"][0]["actions"][4]
    assert "-PreserveMathTypeLayout" in extract_action["argv"]
    assert "-MathTypeLayoutFactor" in extract_action["argv"]
    assert "1.02" in extract_action["argv"]
    assert "-Resume" in extract_action["argv"]
    assert "-StartIndex" in extract_action["argv"]
    assert "216" in extract_action["argv"]
    assert "-EndIndex" in extract_action["argv"]
    assert "238" in extract_action["argv"]
    assert "--preserve-mathtype-layout" in replace_action["argv"]
    assert "--mathtype-layout-factor" in replace_action["argv"]
    assert "1.02" in replace_action["argv"]


def test_run_plan_execute_writes_omml_execution_report_and_artifacts(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.docx"
    plan_path = tmp_path / "execution-plan.json"
    report_path = tmp_path / "out" / "execution-report.json"
    artifact_dir = tmp_path / "artifacts"
    make_docx(input_path)
    plan_payload = build_execution_plan(
        {
            "document_id": "sample",
            "input_path": str(input_path),
            "detector_version": "0.1.0",
            "formula_count": 1,
            "recommended_sequence": ["omml-native"],
            "route_plan": [
                {
                    "source_family": "omml-native",
                    "formula_count": 1,
                    "route_kind": "primary-source-first",
                    "priority": 10,
                    "next_action": "run-omml-native-pipeline",
                    "confidence_policy": "high",
                    "requires_manual_review": False,
                }
            ],
        }
    ).to_dict()
    plan_path.write_text(json.dumps(plan_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    exit_code = main(
        [
            "run-plan",
            str(plan_path),
            "--execute",
            "--output-dir",
            str(artifact_dir),
            "--output",
            str(report_path),
        ]
    )

    assert exit_code == 0
    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert payload["mode"] == "execute"
    assert payload["completed_step_count"] == 1
    assert payload["steps"][0]["provider"] == "omml"
    manifest_path = artifact_dir / "omml-native" / "manifest.json"
    assert manifest_path.exists()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["formula_count"] == 1


def test_validate_docx_writes_validation_report(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.docx"
    output_dir = tmp_path / "validation"
    make_docx(input_path)

    exit_code = main(
        [
            "validate-docx",
            str(input_path),
            "--output-dir",
            str(output_dir),
            "--provider",
            "omml",
            "--source-family",
            "omml-native",
        ]
    )

    assert exit_code == 0
    report_path = output_dir / "validation-report.json"
    text_path = output_dir / "validation-report.txt"
    assert report_path.exists()
    assert text_path.exists()
    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert payload["artifact_type"] == "docx-validation-report"
    assert payload["conclusion"] == "research-only"
    assert payload["provider"] == "omml"
    assert payload["source_family"] == "omml-native"


def test_validate_docx_accepts_target_from_metadata(tmp_path: Path, monkeypatch) -> None:
    output_dir = tmp_path / "validation"
    metadata_path = tmp_path / "execution-metadata.json"
    metadata_path.write_text(
        json.dumps({"validation_target_docx": str(tmp_path / "target.docx")}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    captured: dict[str, object] = {}

    def fake_validate_docx_artifact(**kwargs):
        captured.update(kwargs)
        return {
            "artifact_type": "docx-validation-report",
            "provider": kwargs["provider"],
            "source_family": kwargs["source_family"],
            "conclusion": "research-only",
            "checks": {
                "target_docx": {"status": "passed"},
                "word_export": {"status": "skipped"},
                "visual_compare": {"status": "skipped"},
            },
            "residual_risks": [],
        }

    monkeypatch.setattr(cli_module, "validate_docx_artifact", fake_validate_docx_artifact)

    exit_code = main(
        [
            "validate-docx",
            "--target-from-metadata",
            str(metadata_path),
            "--output-dir",
            str(output_dir),
            "--provider",
            "omml",
            "--source-family",
            "omml-native",
        ]
    )

    assert exit_code == 0
    assert captured["target_docx"] == ""
    assert captured["target_from_metadata"] == str(metadata_path)


def test_validate_docx_passes_visual_gate_arguments(tmp_path: Path, monkeypatch) -> None:
    input_path = tmp_path / "sample.docx"
    output_dir = tmp_path / "validation"
    reference_pdf = tmp_path / "reference.pdf"
    make_docx(input_path)
    reference_pdf.write_bytes(b"%PDF-1.4\n")
    captured: dict[str, object] = {}

    def fake_validate_docx_artifact(**kwargs):
        captured.update(kwargs)
        return {
            "artifact_type": "docx-validation-report",
            "provider": kwargs["provider"],
            "source_family": kwargs["source_family"],
            "conclusion": "review-gated",
            "checks": {
                "target_docx": {"status": "passed"},
                "word_export": {"status": "passed"},
                "visual_compare": {"status": "review-gated"},
            },
            "residual_risks": [],
        }

    monkeypatch.setattr(cli_module, "validate_docx_artifact", fake_validate_docx_artifact)

    exit_code = main(
        [
            "validate-docx",
            str(input_path),
            "--output-dir",
            str(output_dir),
            "--provider",
            "mathtype",
            "--source-family",
            "mathtype-ole",
            "--reference-pdf",
            str(reference_pdf),
            "--allow-word-export",
            "--visual-compare",
            "--visual-max-changed-ratio-per-page",
            "0.03",
            "--visual-max-unmatched-pages",
            "1",
        ]
    )

    assert exit_code == 0
    assert captured["visual_max_changed_ratio_per_page"] == 0.03
    assert captured["visual_max_unmatched_pages"] == 1
