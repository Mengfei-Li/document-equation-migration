import json
import zipfile
from pathlib import Path

from document_equation_migration.execution_plan.model import ExecutionAction, ExecutionStep
from document_equation_migration.executor.model import ExecutionContext
from document_equation_migration.executor.omml import execute_omml_step


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

COMPLEX_DOCX_XML = """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <m:oMath>
        <m:f>
          <m:num><m:r><m:t>1</m:t></m:r></m:num>
          <m:den><m:r><m:t>2</m:t></m:r></m:den>
        </m:f>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:sSup>
          <m:e><m:r><m:t>x</m:t></m:r></m:e>
          <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
        </m:sSup>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:sSub>
          <m:e><m:r><m:t>a</m:t></m:r></m:e>
          <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
        </m:sSub>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:rad><m:e><m:r><m:t>y</m:t></m:r></m:e></m:rad>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:d>
          <m:dPr>
            <m:begChr m:val="(" />
            <m:endChr m:val=")" />
          </m:dPr>
          <m:e><m:r><m:t>z</m:t></m:r></m:e>
        </m:d>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:nary>
          <m:naryPr><m:chr m:val="&#x2211;" /></m:naryPr>
          <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
          <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
          <m:e><m:r><m:t>a</m:t></m:r></m:e>
        </m:nary>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:eqArr>
          <m:e><m:r><m:t>x</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>1</m:t></m:r></m:e>
          <m:e><m:r><m:t>y</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>2</m:t></m:r></m:e>
        </m:eqArr>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:m>
          <m:mr>
            <m:e><m:r><m:t>a</m:t></m:r></m:e>
            <m:e><m:r><m:t>b</m:t></m:r></m:e>
          </m:mr>
          <m:mr>
            <m:e><m:r><m:t>c</m:t></m:r></m:e>
            <m:e><m:r><m:t>d</m:t></m:r></m:e>
          </m:mr>
        </m:m>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:acc>
          <m:accPr><m:chr m:val="~" /></m:accPr>
          <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:acc>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:bar>
          <m:barPr><m:pos m:val="top" /></m:barPr>
          <m:e><m:r><m:t>y</m:t></m:r></m:e>
        </m:bar>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:func>
          <m:fName><m:r><m:t>sin</m:t></m:r></m:fName>
          <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:func>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:limLow>
          <m:e><m:r><m:t>lim</m:t></m:r></m:e>
          <m:lim><m:r><m:t>0</m:t></m:r></m:lim>
        </m:limLow>
      </m:oMath>
    </w:p>
    <w:p>
      <m:oMath>
        <m:limUpp>
          <m:e><m:r><m:t>x</m:t></m:r></m:e>
          <m:lim><m:r><m:t>^</m:t></m:r></m:lim>
        </m:limUpp>
      </m:oMath>
    </w:p>
  </w:body>
</w:document>
""".encode("utf-8")

def make_docx(path: Path, document_xml: bytes = DOCX_XML) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", document_xml)


def make_omml_step(*, requires_manual_review: bool = False) -> ExecutionStep:
    return ExecutionStep(
        source_family="omml-native",
        formula_count=1,
        route_kind="primary-source-first",
        confidence_policy="high",
        requires_manual_review=requires_manual_review,
        provider="omml",
        next_action="run-omml-native-pipeline",
        actions=(
            ExecutionAction(
                action_id="extract-omml",
                description="Extract OMML equations from OOXML document parts.",
            ),
            ExecutionAction(
                action_id="normalize-omml",
                description="Normalize OMML structure for deterministic downstream conversion.",
            ),
            ExecutionAction(
                action_id="omml-to-canonical-mathml",
                description="Convert normalized OMML fragments to canonical MathML.",
            ),
            ExecutionAction(
                action_id="render-check",
                description="Create validation plan for render parity checks.",
                blocking=requires_manual_review,
            ),
            ExecutionAction(
                action_id="package-omml-output",
                description="Package normalized OMML output and execution metadata.",
            ),
        ),
    )


def make_context(tmp_path: Path, input_path: Path) -> ExecutionContext:
    return ExecutionContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(tmp_path / "execution-plan.json"),
        input_path=str(input_path),
        output_dir=str(tmp_path / "out"),
    )


def test_execute_omml_step_writes_manifest_and_validation_plan(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.docx"
    make_docx(input_path)

    reports = execute_omml_step(make_omml_step(), make_context(tmp_path, input_path))

    assert [report.action_id for report in reports] == [
        "extract-omml",
        "normalize-omml",
        "omml-to-canonical-mathml",
        "render-check",
        "package-omml-output",
    ]
    assert [report.status for report in reports] == ["completed", "completed", "completed", "skipped", "completed"]

    manifest_path = Path(reports[0].output_paths[0])
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["formula_count"] == 1
    assert manifest["items"][0]["kind"] == "oMath"
    assert Path(manifest["items"][0]["extracted_path"]).exists()

    canonical_summary_path = Path(reports[2].output_paths[0])
    canonical_summary = json.loads(canonical_summary_path.read_text(encoding="utf-8"))
    assert canonical_summary["strategy"] == "internal-basic-omml-to-presentation-mathml"
    assert canonical_summary["canonical_mathml_count"] == 1
    canonical_mathml_path = Path(reports[2].output_paths[1])
    canonical_mathml = canonical_mathml_path.read_text(encoding="utf-8")
    assert "<math:math" in canonical_mathml
    assert "<math:mi>x</math:mi>" in canonical_mathml

    package_report = reports[4]
    validation_target_path = Path(package_report.output_paths[1])
    validation_evidence_path = Path(package_report.output_paths[2])
    assert validation_target_path.exists()
    assert validation_evidence_path.exists()

    validation_evidence = json.loads(validation_evidence_path.read_text(encoding="utf-8"))
    assert validation_evidence["artifact_type"] == "omml-validation-evidence"
    assert validation_evidence["status"] == "evidence-collected"
    assert validation_evidence["gate_status"] == "pending-external-validation"
    assert validation_evidence["canonical_target"]["target_format"] == "canonical-mathml"
    assert validation_evidence["canonical_target"]["contract_status"] == "implemented-basic"
    assert validation_evidence["artifacts"]["manifest"]["formula_count"] == 1
    assert validation_evidence["artifacts"]["normalization_summary"]["normalized_count"] == 1
    assert validation_evidence["artifacts"]["canonicalization_summary"]["canonical_mathml_count"] == 1
    assert validation_evidence["artifacts"]["package_metadata"]["provider"] == "omml"
    assert validation_evidence["artifacts"]["validation_target"]["validation_target_present"] is True
    assert validation_evidence["artifacts"]["validation_target"]["validation_target_docx"] == str(validation_target_path)
    assert validation_evidence["artifacts"]["validation_plan"]["status"] == "pending-external-validation"
    assert validation_evidence["artifacts"]["validation_plan"]["review_mode"] == "spot-check"
    assert {item["id"] for item in validation_evidence["evidence_checks"]} >= {
        "manifest-present",
        "normalization-summary-present",
        "canonicalization-summary-present",
        "package-metadata-present",
        "validation-target-present",
        "validation-plan-present",
        "word-open",
        "word-export-pdf",
        "render-parity",
    }
    assert any(
        item["id"] == "word-open" and item["status"] == "not-run"
        for item in validation_evidence["evidence_checks"]
    )

    validation_plan_path = Path(reports[3].output_paths[0])
    validation_plan = json.loads(validation_plan_path.read_text(encoding="utf-8"))
    assert validation_plan["artifact_type"] == "omml-validation-plan"
    assert validation_plan["status"] == "pending-external-validation"
    assert validation_plan["review_mode"] == "spot-check"
    assert validation_plan["evidence"]["manifest"]["formula_count"] == 1
    assert {item["id"] for item in validation_plan["planned_checks"]} == {
        "word-open",
        "word-export-pdf",
        "render-parity",
        "formula-count",
    }
    assert all(item["status"] != "completed" for item in validation_plan["planned_checks"])


def test_execute_omml_step_writes_canonical_mathml_for_common_structures(tmp_path: Path) -> None:
    input_path = tmp_path / "complex.docx"
    make_docx(input_path, COMPLEX_DOCX_XML)

    reports = execute_omml_step(make_omml_step(), make_context(tmp_path, input_path))

    manifest = json.loads(Path(reports[0].output_paths[0]).read_text(encoding="utf-8"))
    canonical_summary = json.loads(Path(reports[2].output_paths[0]).read_text(encoding="utf-8"))
    assert manifest["formula_count"] == 13
    assert canonical_summary["canonical_mathml_count"] == 13
    assert canonical_summary["unsupported_fragment_count"] == 0

    canonical_text = "\n".join(
        Path(path).read_text(encoding="utf-8")
        for path in reports[2].output_paths[1:]
    )
    assert "<math:mfrac>" in canonical_text
    assert "<math:msup>" in canonical_text
    assert "<math:msub>" in canonical_text
    assert "<math:msqrt>" in canonical_text
    assert "<math:mfenced" in canonical_text
    assert "<math:munderover>" in canonical_text
    assert "<math:mtable>" in canonical_text
    assert '<math:mover accent="true">' in canonical_text
    assert f"<math:mo>{chr(0x00AF)}</math:mo>" in canonical_text
    assert f"<math:mo>{chr(0x2061)}</math:mo>" in canonical_text
    assert "<math:munder>" in canonical_text


def test_execute_omml_step_marks_manual_review_validation_gate(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.docx"
    make_docx(input_path)

    reports = execute_omml_step(
        make_omml_step(requires_manual_review=True),
        make_context(tmp_path, input_path),
    )

    render_report = reports[3]
    validation_plan = json.loads(Path(render_report.output_paths[0]).read_text(encoding="utf-8"))
    validation_target_path = Path(reports[4].output_paths[1])
    validation_evidence = json.loads(Path(reports[4].output_paths[2]).read_text(encoding="utf-8"))

    assert render_report.status == "review-gated"
    assert render_report.blocking is True
    assert validation_target_path.exists()
    assert validation_plan["status"] == "manual-review-required"
    assert validation_plan["review_mode"] == "required"
    assert validation_evidence["artifact_type"] == "omml-validation-evidence"
    assert validation_evidence["status"] == "evidence-collected"
    assert validation_evidence["gate_status"] == "manual-review-required"
    assert validation_evidence["artifacts"]["validation_plan"]["status"] == "manual-review-required"
    assert validation_evidence["artifacts"]["validation_plan"]["review_mode"] == "required"
