import base64
import importlib.util
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = (
    ROOT
    / "src"
    / "document_equation_migration"
    / "detectors"
    / "mathtype_ole.py"
)
FIXTURE_ROOT = Path(__file__).resolve().parent / "fixtures" / "mathtype_ole"


def load_detector_module():
    spec = importlib.util.spec_from_file_location("mathtype_ole_detector", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def build_docx_from_fixture(tmp_path: Path, fixture_name: str) -> Path:
    fixture_dir = FIXTURE_ROOT / fixture_name
    output_path = tmp_path / f"{fixture_name}.docx"
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(fixture_dir.rglob("*")):
            if not file_path.is_file():
                continue

            relative_path = file_path.relative_to(fixture_dir).as_posix()
            if file_path.suffix == ".b64":
                arcname = relative_path[: -len(".b64")]
                data = base64.b64decode(file_path.read_text(encoding="ascii"))
                zf.writestr(arcname, data)
                continue

            zf.write(file_path, relative_path)
    return output_path


def test_detects_main_story_mathtype_and_extracts_probe_signals(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "main_story")

    result = detector.detect_mathtype_ole(docx_path)

    assert result["source_counts"] == {"mathtype-ole": 1}
    assert len(result["formulas"]) == 1
    formula = result["formulas"][0]
    assert formula["source_family"] == "mathtype-ole"
    assert formula["source_role"] == "native-source"
    assert formula["doc_part_path"] == "word/document.xml"
    assert formula["story_type"] == "main"
    assert formula["storage_kind"] == "ole-embedded"
    assert formula["relationship_id"] == "rId1"
    assert formula["embedding_target"] == "word/embeddings/oleObject1.bin"
    assert formula["preview_target"] == "word/media/image1.wmf"
    assert formula["paragraph_index"] == 1
    assert formula["run_index"] == 3
    assert formula["risk_level"] == "low"
    assert formula["failure_mode"] is None
    assert formula["confidence"] >= 0.95
    assert formula["paragraph_text"] == "Before [OLE]After"
    assert formula["text_before"] == "Before "
    assert formula["text_after"] == "After"
    assert "o:OLEObject" in formula["xpath"]

    provenance = formula["provenance"]
    assert provenance["prog_id_raw"] == "Equation.DSMT4"
    assert provenance["field_code_raw"] == "EMBED Equation.DSMT4"
    assert provenance["ole_stream_names"] == ["Equation Native"]
    assert provenance["raw_payload_status"] == "present"
    assert len(provenance["raw_payload_sha256"]) == 64
    assert "word/_rels/document.xml.rels" in provenance["evidence_sources"]

    mathtype = formula["mathtype"]
    assert mathtype["equation_native_stream_exists"] is True
    assert mathtype["equation_native_size_bytes"] == 4096
    assert mathtype["mtef_version"] == 5
    assert mathtype["application_key"] == "DSMT7"
    assert mathtype["product_version"] == 7
    assert mathtype["product_subversion"] == 0


def test_detects_comment_story_mathtype_as_non_main_risk(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "comment_story")

    result = detector.detect_mathtype_ole(docx_path)

    assert result["source_counts"] == {"mathtype-ole": 1}
    formula = result["formulas"][0]
    assert formula["doc_part_path"] == "word/comments.xml"
    assert formula["story_type"] == "comment"
    assert formula["risk_level"] == "medium"
    assert "story-part-nonmain" in formula["risk_flags"]
    assert formula["preview_target"] is None
    assert formula["mathtype"]["equation_native_stream_exists"] is True


def test_flags_missing_equation_native_stream(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "missing_equation_native")

    result = detector.detect_mathtype_ole(docx_path)

    assert result["source_counts"] == {"mathtype-ole": 1}
    formula = result["formulas"][0]
    assert formula["risk_level"] == "high"
    assert "missing-equation-native-stream" in formula["risk_flags"]
    assert formula["failure_mode"] == "missing-equation-native-stream"
    assert formula["canonical_mathml_status"] == "missing"
    assert formula["provenance"]["raw_payload_status"] == "present"
    assert formula["provenance"]["ole_stream_names"] == ["NotEquation"]
    assert formula["mathtype"]["equation_native_stream_exists"] is False
    assert formula["mathtype"]["application_key"] is None


def test_ignores_non_mathtype_ole_objects(tmp_path: Path):
    detector = load_detector_module()
    docx_path = build_docx_from_fixture(tmp_path, "non_mathtype_ole")

    result = detector.detect_mathtype_ole(docx_path)

    assert result["source_counts"] == {}
    assert result["formulas"] == []
