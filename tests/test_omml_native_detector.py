import importlib.util
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "src" / "document_equation_migration" / "detectors" / "omml_native.py"
FIXTURE_ROOT = Path(__file__).resolve().parent / "fixtures" / "omml_native"


def load_detector_module():
    spec = importlib.util.spec_from_file_location("omml_native_detector", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def build_docx_from_fixture(tmp_path: Path, fixture_name: str) -> Path:
    fixture_dir = FIXTURE_ROOT / fixture_name
    output_path = tmp_path / f"{fixture_name}.docx"
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(fixture_dir.rglob("*")):
            if file_path.is_file():
                zf.write(file_path, file_path.relative_to(fixture_dir).as_posix())
    return output_path


def test_detects_inline_omml_in_main_story(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "main_inline")

    result = detector.detect_omml_native(docx_path)

    assert result["source_counts"] == {"omml-native": 1}
    assert len(result["formulas"]) == 1
    formula = result["formulas"][0]
    assert formula["source_family"] == "omml-native"
    assert formula["source_role"] == "native-source"
    assert formula["doc_part_path"] == "word/document.xml"
    assert formula["story_type"] == "main"
    assert formula["storage_kind"] == "omml-inline"
    assert formula["paragraph_index"] == 1
    assert formula["run_index"] == 2
    assert formula["omml"]["container_element"] == "m:oMath"
    assert formula["omml"]["display_mode"] == "inline"
    assert formula["omml"]["has_mathPr"] is False
    assert formula["risk_level"] == "low"
    assert formula["xpath"].endswith("/m:oMath[1]")


def test_detects_display_omml_and_math_settings(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "main_display")

    result = detector.detect_omml_native(docx_path)

    assert result["source_counts"] == {"omml-native": 1}
    formula = result["formulas"][0]
    assert formula["doc_part_path"] == "word/document.xml"
    assert formula["story_type"] == "main"
    assert formula["storage_kind"] == "omml-display"
    assert formula["paragraph_index"] == 1
    assert formula["run_index"] is None
    assert formula["omml"]["container_element"] == "m:oMathPara"
    assert formula["omml"]["display_mode"] == "display"
    assert formula["omml"]["has_mathPr"] is True
    assert formula["omml"]["math_child_count"] == 2
    assert "word/settings.xml" in formula["provenance"]["evidence_sources"]


def test_detects_non_main_story_formula(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "comment_story")

    result = detector.detect_omml_native(docx_path)

    assert result["source_counts"] == {"omml-native": 1}
    formula = result["formulas"][0]
    assert formula["doc_part_path"] == "word/comments.xml"
    assert formula["story_type"] == "comment"
    assert formula["risk_level"] == "medium"
    assert "story-part-nonmain" in formula["risk_flags"]
    assert formula["omml"]["container_element"] == "m:oMath"
    assert formula["omml"]["display_mode"] == "inline"


def test_returns_empty_for_docx_without_omml(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "no_omml")

    result = detector.detect_omml_native(docx_path)

    assert result["source_counts"] == {}
    assert result["formulas"] == []
