import importlib.util
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "src" / "document_equation_migration" / "detectors" / "axmath_ole.py"
FIXTURE_ROOT = Path(__file__).resolve().parent / "fixtures" / "axmath_ole"


def load_detector_module():
    spec = importlib.util.spec_from_file_location("axmath_ole_detector", MODULE_PATH)
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


def test_detects_axmath_from_prog_id_and_relationship_payload(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "main_prog_id")

    result = detector.detect_axmath_ole(docx_path)

    assert result["source_counts"] == {"axmath-ole": 1}
    assert len(result["formulas"]) == 1
    formula = result["formulas"][0]
    assert formula["source_family"] == "axmath-ole"
    assert formula["source_role"] == "native-source"
    assert formula["doc_part_path"] == "word/document.xml"
    assert formula["story_type"] == "main"
    assert formula["storage_kind"] == "ole-embedded"
    assert formula["relationship_id"] == "rIdOle1"
    assert formula["embedding_target"] == "word/embeddings/oleObject1.bin"
    assert formula["preview_target"] == "word/media/image1.wmf"
    assert formula["paragraph_index"] == 1
    assert formula["run_index"] == 2
    assert formula["risk_level"] == "medium"
    assert formula["provenance"]["prog_id_raw"] == "Equation.AxMath"
    assert formula["provenance"]["field_code_raw"] is None
    assert formula["provenance"]["raw_payload_status"] == "present"
    assert "word/styles.xml" in formula["provenance"]["evidence_sources"]
    assert formula["axmath"]["word_addin_artifacts"] == ["AMDisplayEquation"]
    assert formula["axmath"]["export_channels"] == ["latex", "mathml"]
    assert formula["axmath"]["export_route_verified"] is False
    assert "native-static-parse-unverified" in formula["risk_flags"]
    assert formula["xpath"].endswith("/w:object[1]")


def test_detects_axmath_from_field_code_when_prog_id_is_missing(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "main_field_code")

    result = detector.detect_axmath_ole(docx_path)

    assert result["source_counts"] == {"axmath-ole": 1}
    formula = result["formulas"][0]
    assert formula["provenance"]["prog_id_raw"] is None
    assert formula["provenance"]["field_code_raw"] == "EMBED Equation.AxMath"
    assert formula["relationship_id"] == "rIdOle2"
    assert formula["embedding_target"] == "word/embeddings/oleObject2.bin"
    assert formula["risk_level"] == "high"
    assert "prog-id-missing" in formula["risk_flags"]
    assert formula["axmath"]["custom_symbol_present"] is True
    assert "custom-symbols-present" in formula["risk_flags"]
    assert formula["confidence"] == 0.97


def test_marks_non_main_story_as_higher_risk(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "comment_story")

    result = detector.detect_axmath_ole(docx_path)

    assert result["source_counts"] == {"axmath-ole": 1}
    formula = result["formulas"][0]
    assert formula["doc_part_path"] == "word/comments.xml"
    assert formula["story_type"] == "comment"
    assert formula["relationship_id"] == "rIdCommentOle1"
    assert formula["risk_level"] == "high"
    assert "story-part-nonmain" in formula["risk_flags"]
    assert formula["axmath"]["word_addin_artifacts"] == []


def test_returns_empty_for_non_axmath_ole_candidates(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "no_axmath")

    result = detector.detect_axmath_ole(docx_path)

    assert result["source_counts"] == {}
    assert result["formulas"] == []
