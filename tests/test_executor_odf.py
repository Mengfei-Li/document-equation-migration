from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from document_equation_migration.execution_plan.odf import build_odf_execution_step
from document_equation_migration.executor.model import DryRunContext, ExecutionContext
from document_equation_migration.executor.odf import build_odf_dry_run_reports, execute_odf_step


FODT_WITH_INLINE_MATH = b"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:math="http://www.w3.org/1998/Math/MathML">
  <office:body>
    <office:text>
      <text:p>
        <math:math>
          <math:mrow>
            <math:mi>x</math:mi>
            <math:mo>+</math:mo>
            <math:mn>1</math:mn>
          </math:mrow>
        </math:math>
      </text:p>
    </office:text>
  </office:body>
</office:document-content>
"""

FODT_WITH_COMMON_MATHML_STRUCTURES = b"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:math="http://www.w3.org/1998/Math/MathML">
  <office:body>
    <office:text>
      <text:p>
        <math:math><math:mfrac><math:mn>1</math:mn><math:mn>2</math:mn></math:mfrac></math:math>
      </text:p>
      <text:p>
        <math:math><math:msup><math:mi>x</math:mi><math:mn>2</math:mn></math:msup></math:math>
      </text:p>
      <text:p>
        <math:math><math:msub><math:mi>a</math:mi><math:mi>i</math:mi></math:msub></math:math>
      </text:p>
      <text:p>
        <math:math><math:msqrt><math:mi>y</math:mi></math:msqrt></math:math>
      </text:p>
      <text:p>
        <math:math><math:mfenced open="(" close=")"><math:mi>z</math:mi></math:mfenced></math:math>
      </text:p>
      <text:p>
        <math:math>
          <math:mrow>
            <math:munderover><math:mo>&#x2211;</math:mo><math:mi>i</math:mi><math:mi>n</math:mi></math:munderover>
            <math:mi>a</math:mi>
          </math:mrow>
        </math:math>
      </text:p>
      <text:p>
        <math:math>
          <math:mtable>
            <math:mtr><math:mtd><math:mrow><math:mi>x</math:mi><math:mo>=</math:mo><math:mn>1</math:mn></math:mrow></math:mtd></math:mtr>
            <math:mtr><math:mtd><math:mrow><math:mi>y</math:mi><math:mo>=</math:mo><math:mn>2</math:mn></math:mrow></math:mtd></math:mtr>
          </math:mtable>
        </math:math>
      </text:p>
      <text:p>
        <math:math>
          <math:mtable>
            <math:mtr><math:mtd><math:mn>1</math:mn></math:mtd><math:mtd><math:mn>0</math:mn></math:mtd></math:mtr>
            <math:mtr><math:mtd><math:mn>0</math:mn></math:mtd><math:mtd><math:mn>1</math:mn></math:mtd></math:mtr>
          </math:mtable>
        </math:math>
      </text:p>
      <text:p>
        <math:math><math:mover accent="true"><math:mi>x</math:mi><math:mo>~</math:mo></math:mover></math:math>
      </text:p>
      <text:p>
        <math:math><math:mover><math:mi>x</math:mi><math:mo>&#x00AF;</math:mo></math:mover></math:math>
      </text:p>
      <text:p>
        <math:math><math:mrow><math:mi>sin</math:mi><math:mo>&#x2061;</math:mo><math:mi>x</math:mi></math:mrow></math:math>
      </text:p>
      <text:p>
        <math:math><math:munder><math:mi>lim</math:mi><math:mn>0</math:mn></math:munder></math:math>
      </text:p>
      <text:p>
        <math:math><math:mover><math:mi>x</math:mi><math:mo>^</math:mo></math:mover></math:math>
      </text:p>
    </office:text>
  </office:body>
</office:document-content>
"""

FODT_WITH_PROPERTY_RICH_MATHML = b"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:math="http://www.w3.org/1998/Math/MathML">
  <office:body>
    <office:text>
      <text:p>
        <math:math display="block">
          <math:mfrac linethickness="0"><math:mi>a</math:mi><math:mi>b</math:mi></math:mfrac>
        </math:math>
      </text:p>
      <text:p>
        <math:math>
          <math:mfrac bevelled="true"><math:mn>1</math:mn><math:mn>2</math:mn></math:mfrac>
        </math:math>
      </text:p>
      <text:p>
        <math:math>
          <math:mfenced open="[" close="]" separators=";">
            <math:mi>x</math:mi><math:mi>y</math:mi>
          </math:mfenced>
        </math:math>
      </text:p>
      <text:p>
        <math:math>
          <math:mrow>
            <math:munderover>
              <math:mo movablelimits="true">&#x2211;</math:mo>
              <math:mi>i</math:mi>
              <math:mi>n</math:mi>
            </math:munderover>
            <math:mi>a</math:mi>
          </math:mrow>
        </math:math>
      </text:p>
      <text:p>
        <math:math><math:mi mathvariant="bold-italic">v</math:mi></math:math>
      </text:p>
      <text:p>
        <math:math>
          <math:munder accentunder="true"><math:mi>x</math:mi><math:mo>&#x00AF;</math:mo></math:munder>
        </math:math>
      </text:p>
      <text:p>
        <math:math>
          <math:semantics>
            <math:mrow><math:mi>x</math:mi><math:mo>+</math:mo><math:mn>1</math:mn></math:mrow>
            <math:annotation encoding="application/x-tex">x+1</math:annotation>
          </math:semantics>
        </math:math>
      </text:p>
    </office:text>
  </office:body>
</office:document-content>
"""

ODT_CONTENT_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:body>
    <office:text>
      <text:p>
        <draw:frame>
          <draw:object xlink:href="./Object 1" xlink:type="simple" />
        </draw:frame>
      </text:p>
    </office:text>
  </office:body>
</office:document-content>
"""

ODF_FORMULA_CONTENT_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<math:math xmlns:math="http://www.w3.org/1998/Math/MathML">
  <math:mrow>
    <math:mi>a</math:mi>
    <math:mo>=</math:mo>
    <math:mi>b</math:mi>
  </math:mrow>
</math:math>
"""


def _native_step():
    return build_odf_execution_step(
        {
            "source_family": "odf-native",
            "formula_count": 1,
            "route_kind": "primary-source-first",
            "next_action": "run-odf-native-pipeline",
            "confidence_policy": "medium",
            "requires_manual_review": False,
        }
    )


def _bridge_step():
    return build_odf_execution_step(
        {
            "source_family": "libreoffice-transformed",
            "formula_count": 1,
            "route_kind": "bridge-source",
            "next_action": "run-libreoffice-bridge-review-pipeline",
            "confidence_policy": "low",
            "requires_manual_review": True,
        }
    )


def _execution_context(tmp_path: Path, input_path: Path) -> ExecutionContext:
    return ExecutionContext(
        workspace_root=str(PROJECT_ROOT),
        execution_plan_path=str(tmp_path / "execution-plan.json"),
        input_path=str(input_path),
        output_dir=str(tmp_path / "out"),
    )


def _dry_run_context(tmp_path: Path) -> DryRunContext:
    plan_path = tmp_path / "execution-plan.json"
    plan_path.write_text(
        json.dumps({"input_path": "sample.fodt"}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return DryRunContext(
        workspace_root=str(tmp_path),
        execution_plan_path=str(plan_path),
        output_dir_hint=str(tmp_path / "out"),
    )


def _make_odt(path: Path) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "mimetype",
            "application/vnd.oasis.opendocument.text",
            compress_type=zipfile.ZIP_STORED,
        )
        zf.writestr("content.xml", ODT_CONTENT_XML)
        zf.writestr("Object 1/content.xml", ODF_FORMULA_CONTENT_XML)


def test_build_odf_native_dry_run_reports_marks_extraction_ready_and_delivery_gated(tmp_path: Path) -> None:
    reports = build_odf_dry_run_reports(_native_step(), _dry_run_context(tmp_path))

    assert [item.action_id for item in reports] == [
        "extract-odf-formula",
        "convert-odf-mathml",
        "emit-target-format",
    ]
    assert [item.status for item in reports] == ["ready", "ready", "validation-gated"]
    assert reports[0].runner == "internal-odf-native"
    assert reports[2].runner == "manual-validation"
    assert reports[2].blocking is True


def test_execute_odf_native_extracts_mathml_from_fodt(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.fodt"
    input_path.write_bytes(FODT_WITH_INLINE_MATH)

    reports = execute_odf_step(_native_step(), _execution_context(tmp_path, input_path))

    assert [item.status for item in reports] == ["completed", "completed", "validation-gated"]
    manifest_path = Path(reports[0].output_paths[0])
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["formula_count"] == 1
    formula = manifest["formulas"][0]
    assert formula["source_family"] == "odf-native"
    assert formula["source_role"] == "native-source"
    assert formula["doc_part_path"] == "content.xml"
    assert formula["storage_kind"] == "odf-content-inline-mathml"

    extracted_path = Path(reports[0].output_paths[1])
    canonical_path = Path(reports[1].output_paths[1])
    evidence_path = Path(reports[2].output_paths[0])
    assert "<math:math" in extracted_path.read_text(encoding="utf-8")
    assert canonical_path.exists()
    assert evidence_path.name == "validation-evidence.json"
    evidence = json.loads(evidence_path.read_text(encoding="utf-8"))
    assert evidence["artifact_type"] == "odf-validation-evidence"
    assert evidence["source_family"] == "odf-native"
    assert evidence["status"] == "evidence-recorded"
    assert evidence["manifest"]["formula_count"] == 1
    assert evidence["canonicalization"]["canonical_mathml_count"] == 1


def test_execute_odf_native_preserves_common_canonical_mathml_structures(tmp_path: Path) -> None:
    input_path = tmp_path / "complex.fodt"
    input_path.write_bytes(FODT_WITH_COMMON_MATHML_STRUCTURES)

    reports = execute_odf_step(_native_step(), _execution_context(tmp_path, input_path))

    manifest = json.loads(Path(reports[0].output_paths[0]).read_text(encoding="utf-8"))
    canonical_summary = json.loads(Path(reports[1].output_paths[0]).read_text(encoding="utf-8"))
    assert manifest["formula_count"] == 13
    assert canonical_summary["canonical_mathml_count"] == 13

    canonical_text = "\n".join(
        Path(path).read_text(encoding="utf-8")
        for path in reports[1].output_paths[1:]
    )
    assert "<math:mfrac>" in canonical_text
    assert "<math:msup>" in canonical_text
    assert "<math:msub>" in canonical_text
    assert "<math:msqrt>" in canonical_text
    assert "<math:mfenced" in canonical_text
    assert "<math:munderover>" in canonical_text
    assert f"<math:mo>{chr(0x2211)}</math:mo>" in canonical_text
    assert canonical_text.count("<math:mtable>") == 2
    assert "<math:mtr>" in canonical_text
    assert "<math:mtd>" in canonical_text
    assert "<math:mover accent=\"true\">" in canonical_text
    assert f"<math:mo>{chr(0x00AF)}</math:mo>" in canonical_text
    assert "<math:mi>sin</math:mi>" in canonical_text
    assert f"<math:mo>{chr(0x2061)}</math:mo>" in canonical_text
    assert "<math:munder>" in canonical_text
    assert "<math:mo>^</math:mo>" in canonical_text


def test_execute_odf_native_preserves_mathml_properties_and_metadata(tmp_path: Path) -> None:
    input_path = tmp_path / "property-rich.fodt"
    input_path.write_bytes(FODT_WITH_PROPERTY_RICH_MATHML)

    reports = execute_odf_step(_native_step(), _execution_context(tmp_path, input_path))

    manifest = json.loads(Path(reports[0].output_paths[0]).read_text(encoding="utf-8"))
    canonical_summary = json.loads(Path(reports[1].output_paths[0]).read_text(encoding="utf-8"))
    evidence = json.loads(Path(reports[2].output_paths[0]).read_text(encoding="utf-8"))
    assert manifest["formula_count"] == 7
    assert canonical_summary["expected_formula_count"] == 7
    assert canonical_summary["canonical_mathml_count"] == 7
    assert canonical_summary["unsupported_fragment_count"] == 0
    assert canonical_summary["formula_count_parity"] == "passed"
    assert len(canonical_summary["source_to_canonical_provenance"]) == 7
    assert {item["preservation_status"] for item in canonical_summary["source_to_canonical_provenance"]} == {
        "byte-identical-after-extraction"
    }

    property_summary = canonical_summary["property_summary"]
    assert property_summary["mathml_attribute_count"] >= 8
    assert property_summary["root_display_values"] == ["block"]
    assert property_summary["signal_counts"]["has_mfrac_linethickness"] == 1
    assert property_summary["signal_counts"]["has_mfrac_bevelled"] == 1
    assert property_summary["signal_counts"]["has_mfenced_separators"] == 1
    assert property_summary["signal_counts"]["has_movablelimits"] == 1
    assert property_summary["signal_counts"]["has_mathvariant"] == 1
    assert property_summary["signal_counts"]["has_accentunder"] == 1
    assert property_summary["signal_counts"]["has_semantics"] == 1
    assert property_summary["signal_counts"]["has_annotation"] == 1
    assert evidence["canonicalization"]["formula_count_parity"] == "passed"
    assert evidence["canonicalization"]["property_summary"] == property_summary
    assert len(evidence["source_to_canonical_provenance"]) == 7

    canonical_text = "\n".join(
        Path(path).read_text(encoding="utf-8")
        for path in reports[1].output_paths[1:]
    )
    assert 'display="block"' in canonical_text
    assert 'linethickness="0"' in canonical_text
    assert 'bevelled="true"' in canonical_text
    assert 'separators=";"' in canonical_text
    assert 'movablelimits="true"' in canonical_text
    assert 'mathvariant="bold-italic"' in canonical_text
    assert 'accentunder="true"' in canonical_text
    assert 'encoding="application/x-tex"' in canonical_text


def test_execute_odf_native_extracts_mathml_from_odt_subdocument(tmp_path: Path) -> None:
    input_path = tmp_path / "sample.odt"
    _make_odt(input_path)

    reports = execute_odf_step(_native_step(), _execution_context(tmp_path, input_path))

    manifest = json.loads(Path(reports[0].output_paths[0]).read_text(encoding="utf-8"))
    assert manifest["formula_count"] == 1
    formula = manifest["formulas"][0]
    assert formula["doc_part_path"] == "Object 1/content.xml"
    assert formula["storage_kind"] == "odf-draw-object-subdocument"
    assert Path(formula["artifact_path"]).exists()


def test_libreoffice_transformed_is_bridge_review_gate_not_native_extract(tmp_path: Path) -> None:
    input_path = tmp_path / "bridge.odt"
    _make_odt(input_path)
    dry_reports = build_odf_dry_run_reports(_bridge_step(), _dry_run_context(tmp_path))

    assert [item.action_id for item in dry_reports] == [
        "inspect-transform-chain",
        "bridge-review",
        "decide-reconvert-or-accept",
    ]
    assert {item.status for item in dry_reports} == {"review-gated"}
    assert {item.runner for item in dry_reports} == {"manual-review"}
    assert all(item.blocking for item in dry_reports)

    execution_reports = execute_odf_step(_bridge_step(), _execution_context(tmp_path, input_path))

    assert {item.status for item in execution_reports} == {"review-gated"}
    assert {item.runner for item in execution_reports} == {"manual-review"}
    assert not (tmp_path / "out" / "odf-native" / "manifest.json").exists()
    blocker_path = Path(execution_reports[0].output_paths[0])
    assert blocker_path.name == "blocker-record.json"
    blocker = json.loads(blocker_path.read_text(encoding="utf-8"))
    assert blocker["artifact_type"] == "odf-blocker-record"
    assert blocker["source_family"] == "libreoffice-transformed"
    assert blocker["status"] == "blocked"
    assert blocker["review_status"] == "review-gated"
    assert blocker["blocker_kind"] == "bridge-review"
