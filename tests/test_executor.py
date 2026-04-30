import json
import zipfile
from pathlib import Path

from document_equation_migration.executor import (
    build_dry_run_execution_report,
    build_execution_report,
    load_execution_plan,
)
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


def make_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", DOCX_XML)


def contract_artifact_names(step: dict[str, object]) -> set[str]:
    names: set[str] = set()
    for action in step["actions"]:
        for output_path in action["output_paths"]:
            names.add(Path(output_path).name)
    return names


def make_execution_plan_dict() -> dict[str, object]:
    routing_report = {
        "document_id": "sample",
        "input_path": "sample.docx",
        "detector_version": "0.1.0",
        "formula_count": 3,
        "recommended_sequence": [
            "mathtype-ole",
            "omml-native",
            "equation-editor-3-ole",
        ],
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
                "source_family": "omml-native",
                "formula_count": 1,
                "route_kind": "primary-source-first",
                "priority": 20,
                "next_action": "run-omml-native-pipeline",
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
    return build_execution_plan(routing_report).to_dict()


def test_load_execution_plan_round_trip(tmp_path: Path) -> None:
    plan_path = tmp_path / "execution-plan.json"
    payload = make_execution_plan_dict()
    plan_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    plan = load_execution_plan(plan_path)

    assert plan.document_id == "sample"
    assert len(plan.steps) == 3
    assert plan.steps[0].provider == "mathtype"


def test_build_dry_run_execution_report_uses_fallback_for_unbound_providers() -> None:
    plan = load_execution_plan(
        Path(__file__).parent / "fixtures" / "executor" / "sample-execution-plan.json"
    )

    report = build_dry_run_execution_report(plan, execution_plan_path="sample-execution-plan.json").to_dict()

    assert report["mode"] == "dry-run"
    assert report["step_count"] == 3
    assert report["runnable_step_count"] == 3
    assert report["manual_only_step_count"] == 0
    assert report["manual_review_required"] is True
    assert report["steps"][0]["provider"] == "mathtype"
    assert report["steps"][0]["status"] == "runnable"
    assert report["steps"][0]["actions"][0]["runner"] == "powershell"
    assert report["steps"][1]["provider"] == "omml"
    assert report["steps"][1]["status"] == "runnable"
    assert report["steps"][1]["actions"][0]["runner"] == "internal-omml-native"
    assert report["steps"][2]["provider"] == "equation3"
    assert report["steps"][2]["status"] == "runnable"
    assert report["steps"][2]["actions"][0]["runner"] == "internal-equation3-probe"


def test_build_execution_report_executes_omml_native_slice(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.docx"
    plan_path = tmp_path / "execution-plan.json"
    output_dir = tmp_path / "out"
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

    report = build_execution_report(
        load_execution_plan(plan_path),
        execution_plan_path=str(plan_path),
        output_dir=str(output_dir),
    ).to_dict()

    assert report["mode"] == "execute"
    assert report["step_count"] == 1
    assert report["completed_step_count"] == 1
    assert report["blocked_step_count"] == 0
    assert report["steps"][0]["provider"] == "omml"
    assert report["steps"][0]["status"] == "completed-with-skips"
    action_statuses = [item["status"] for item in report["steps"][0]["actions"]]
    assert action_statuses == ["completed", "completed", "completed", "skipped", "completed"]
    assert report["steps"][0]["canonical_target"]["target_format"] == "canonical-mathml"
    assert report["steps"][0]["canonical_target"]["contract_status"] == "implemented-basic"
    manifest_path = Path(report["steps"][0]["actions"][0]["output_paths"][0])
    assert manifest_path.exists()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["formula_count"] == 1
    canonical_action = report["steps"][0]["actions"][2]
    canonical_summary_path = Path(canonical_action["output_paths"][0])
    canonical_summary = json.loads(canonical_summary_path.read_text(encoding="utf-8"))
    assert canonical_summary["canonical_mathml_count"] == 1
    package_action = report["steps"][0]["actions"][4]
    validation_target_path = Path(package_action["output_paths"][1])
    assert validation_target_path.exists()
    validation_evidence_path = Path(package_action["output_paths"][2])
    validation_evidence = json.loads(validation_evidence_path.read_text(encoding="utf-8"))
    assert validation_evidence["canonical_target"]["target_format"] == "canonical-mathml"
    assert validation_evidence["artifacts"]["canonicalization_summary"]["canonical_mathml_count"] == 1
    assert validation_evidence["artifacts"]["validation_target"]["validation_target_present"] is True
    assert validation_evidence["artifacts"]["validation_target"]["validation_target_docx"] == str(validation_target_path)


def test_build_execution_report_blocks_mathtype_external_tools_by_default() -> None:
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
                }
            ],
        }
    )

    report = build_execution_report(plan_payload).to_dict()

    assert report["mode"] == "execute"
    assert report["completed_step_count"] == 0
    assert report["blocked_step_count"] == 1
    assert report["steps"][0]["provider"] == "mathtype"
    assert report["steps"][0]["status"] == "blocked"
    assert report["steps"][0]["actions"][0]["status"] == "blocked-external-tool"


def test_dry_run_registry_covers_all_source_line_providers() -> None:
    plan_payload = build_execution_plan(
        {
            "document_id": "sample",
            "input_path": "sample.docx",
            "detector_version": "0.1.0",
            "formula_count": 6,
            "recommended_sequence": [
                "mathtype-ole",
                "omml-native",
                "equation-editor-3-ole",
                "axmath-ole",
                "odf-native",
                "libreoffice-transformed",
            ],
            "route_plan": [
                {
                    "source_family": "mathtype-ole",
                    "formula_count": 1,
                    "route_kind": "primary-source-first",
                    "next_action": "run-mathtype-source-first-pipeline",
                    "confidence_policy": "high",
                    "requires_manual_review": False,
                },
                {
                    "source_family": "omml-native",
                    "formula_count": 1,
                    "route_kind": "primary-source-first",
                    "next_action": "run-omml-native-pipeline",
                    "confidence_policy": "high",
                    "requires_manual_review": False,
                },
                {
                    "source_family": "equation-editor-3-ole",
                    "formula_count": 1,
                    "route_kind": "primary-candidate",
                    "next_action": "run-equation3-probe-and-conversion",
                    "confidence_policy": "medium",
                    "requires_manual_review": True,
                },
                {
                    "source_family": "axmath-ole",
                    "formula_count": 1,
                    "route_kind": "export-assisted",
                    "next_action": "run-axmath-export-assisted-pipeline",
                    "confidence_policy": "medium",
                    "requires_manual_review": True,
                },
                {
                    "source_family": "odf-native",
                    "formula_count": 1,
                    "route_kind": "primary-source-first",
                    "next_action": "run-odf-native-pipeline",
                    "confidence_policy": "medium",
                    "requires_manual_review": False,
                },
                {
                    "source_family": "libreoffice-transformed",
                    "formula_count": 1,
                    "route_kind": "bridge-source",
                    "next_action": "run-libreoffice-bridge-review-pipeline",
                    "confidence_policy": "low",
                    "requires_manual_review": True,
                },
            ],
        }
    )

    report = build_dry_run_execution_report(plan_payload).to_dict()

    providers = {step["provider"] for step in report["steps"]}
    assert providers == {"mathtype", "omml", "equation3", "axmath", "odf"}
    by_provider = {step["provider"]: step for step in report["steps"]}
    assert by_provider["mathtype"]["canonical_target"]["target_format"] == "canonical-mathml"
    assert by_provider["equation3"]["canonical_target"]["contract_status"] == "implemented-limited"
    assert by_provider["axmath"]["canonical_target"]["contract_status"] == "export-gated"
    assert by_provider["equation3"]["actions"][0]["runner"] == "internal-equation3-probe"
    assert by_provider["axmath"]["actions"][0]["runner"] == "axmath-route-gate"
    odf_steps = [step for step in report["steps"] if step["provider"] == "odf"]
    assert {step["source_family"] for step in odf_steps} == {"odf-native", "libreoffice-transformed"}


def test_build_execution_report_mixed_plan_emits_evidence_or_blocker_contract(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.docx"
    plan_path = tmp_path / "execution-plan.json"
    output_dir = tmp_path / "execution"
    make_docx(input_path)

    plan_payload = build_execution_plan(
        {
            "document_id": "sample",
            "input_path": str(input_path),
            "detector_version": "0.1.0",
            "formula_count": 6,
            "recommended_sequence": [
                "mathtype-ole",
                "omml-native",
                "equation-editor-3-ole",
                "axmath-ole",
                "odf-native",
                "libreoffice-transformed",
            ],
            "route_plan": [
                {
                    "source_family": "mathtype-ole",
                    "formula_count": 1,
                    "route_kind": "primary-source-first",
                    "next_action": "run-mathtype-source-first-pipeline",
                    "confidence_policy": "high",
                    "requires_manual_review": False,
                },
                {
                    "source_family": "omml-native",
                    "formula_count": 1,
                    "route_kind": "primary-source-first",
                    "next_action": "run-omml-native-pipeline",
                    "confidence_policy": "high",
                    "requires_manual_review": False,
                },
                {
                    "source_family": "equation-editor-3-ole",
                    "formula_count": 1,
                    "route_kind": "primary-candidate",
                    "next_action": "run-equation3-probe-and-conversion",
                    "confidence_policy": "medium",
                    "requires_manual_review": True,
                },
                {
                    "source_family": "axmath-ole",
                    "formula_count": 1,
                    "route_kind": "export-assisted",
                    "next_action": "run-axmath-export-assisted-pipeline",
                    "confidence_policy": "medium",
                    "requires_manual_review": True,
                },
                {
                    "source_family": "odf-native",
                    "formula_count": 1,
                    "route_kind": "primary-source-first",
                    "next_action": "run-odf-native-pipeline",
                    "confidence_policy": "medium",
                    "requires_manual_review": False,
                },
                {
                    "source_family": "libreoffice-transformed",
                    "formula_count": 1,
                    "route_kind": "bridge-source",
                    "next_action": "run-libreoffice-bridge-review-pipeline",
                    "confidence_policy": "low",
                    "requires_manual_review": True,
                },
            ],
        }
    ).to_dict()
    plan_path.write_text(json.dumps(plan_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    report = build_execution_report(
        load_execution_plan(plan_path),
        execution_plan_path=str(plan_path),
        output_dir=str(output_dir),
    ).to_dict()

    assert report["mode"] == "execute"
    assert report["step_count"] == 6

    contract_artifacts = {"validation-evidence.json", "blocker-record.json"}
    by_source = {step["source_family"]: step for step in report["steps"]}
    for source_family in [
        "mathtype-ole",
        "omml-native",
        "equation-editor-3-ole",
        "axmath-ole",
        "odf-native",
        "libreoffice-transformed",
    ]:
        assert contract_artifacts & contract_artifact_names(by_source[source_family])
