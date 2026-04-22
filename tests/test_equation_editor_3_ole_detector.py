import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path


TESTS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TESTS_DIR.parent
MODULE_PATH = REPO_ROOT / "src" / "document_equation_migration" / "detectors" / "equation_editor_3_ole.py"
FIXTURE_DIR = TESTS_DIR / "fixtures" / "equation_editor_3_ole"


def _load_module():
    spec = importlib.util.spec_from_file_location("equation_editor_3_ole", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def _fixture_text(name: str) -> str:
    return (FIXTURE_DIR / name).read_text(encoding="utf-8")


def _fixture_bytes_from_hex(name: str) -> bytes:
    raw = _fixture_text(name)
    compact = "".join(raw.split())
    return bytes.fromhex(compact)


def _build_docx(
    target: Path,
    *,
    document_xml_name: str,
    document_rels_name: str,
    embedding_hex_name: str | None = None,
    preview_bytes: bytes | None = None,
):
    with zipfile.ZipFile(target, "w") as zf:
        zf.writestr("word/document.xml", _fixture_text(document_xml_name))
        zf.writestr("word/_rels/document.xml.rels", _fixture_text(document_rels_name))
        if embedding_hex_name is not None:
            zf.writestr("word/embeddings/oleObject1.bin", _fixture_bytes_from_hex(embedding_hex_name))
        if preview_bytes is not None:
            zf.writestr("word/media/image1.wmf", preview_bytes)


class EquationEditor3OleDetectorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.module = _load_module()

    def test_probe_eqnolefilehdr_detects_equation_editor_mtef_v3(self):
        payload = _fixture_bytes_from_hex("equation3_native_payload.hex")

        probe = self.module.probe_eqnolefilehdr(payload)

        self.assertTrue(probe["header_detected"])
        self.assertEqual(probe["native_header_size_bytes"], 28)
        self.assertEqual(probe["mtef_version"], 3)
        self.assertEqual(probe["mtef_generating_product"], 1)

    def test_detects_equation_editor_3_docx_from_progid_and_header(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            docx_path = Path(tmp_dir) / "synthetic-eq3.docx"
            _build_docx(
                docx_path,
                document_xml_name="document_equation3.xml",
                document_rels_name="document_equation3.rels.xml",
                embedding_hex_name="equation3_native_payload.hex",
                preview_bytes=b"WMF-SYNTHETIC",
            )

            records = self.module.detect_equation_editor_3_ole(docx_path)

        self.assertEqual(len(records), 1)
        record = records[0]
        self.assertEqual(record["source_family"], "equation-editor-3-ole")
        self.assertEqual(record["source_role"], "native-source")
        self.assertEqual(record["story_type"], "main")
        self.assertEqual(record["embedding_target"], "word/embeddings/oleObject1.bin")
        self.assertEqual(record["preview_target"], "word/media/image1.wmf")
        self.assertEqual(record["provenance"]["prog_id_raw"], "Equation.3")
        self.assertEqual(record["source_specific"]["equation_editor_3"]["selected_route"], "mtef-v3-mainline")
        self.assertEqual(record["risk_level"], "low")
        self.assertGreaterEqual(record["confidence"], 0.75)

    def test_marks_preview_only_when_payload_is_missing(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            docx_path = Path(tmp_dir) / "synthetic-eq3-preview-only.docx"
            _build_docx(
                docx_path,
                document_xml_name="document_equation3.xml",
                document_rels_name="document_preview_only.rels.xml",
                embedding_hex_name=None,
                preview_bytes=b"WMF-SYNTHETIC",
            )

            records = self.module.detect_equation_editor_3_ole(docx_path)

        self.assertEqual(len(records), 1)
        record = records[0]
        self.assertEqual(record["source_role"], "preview-only")
        self.assertEqual(record["provenance"]["raw_payload_status"], "missing")
        self.assertEqual(record["source_specific"]["equation_editor_3"]["selected_route"], "preview-only")
        self.assertEqual(record["risk_level"], "manual-review")

    def test_field_code_fallback_detects_equation3_without_progid(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            docx_path = Path(tmp_dir) / "synthetic-eq3-field-code.docx"
            _build_docx(
                docx_path,
                document_xml_name="document_field_code_only.xml",
                document_rels_name="document_equation3.rels.xml",
                embedding_hex_name="equation3_native_payload.hex",
                preview_bytes=b"WMF-SYNTHETIC",
            )

            records = self.module.detect_equation_editor_3_ole(docx_path)

        self.assertEqual(len(records), 1)
        record = records[0]
        self.assertIsNone(record["provenance"]["prog_id_raw"])
        self.assertIn("EMBED Equation", record["provenance"]["field_code_raw"])
        self.assertEqual(record["source_family"], "equation-editor-3-ole")
        self.assertGreaterEqual(record["confidence"], 0.4)

    def test_does_not_misclassify_mathtype_as_equation_editor_3(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            docx_path = Path(tmp_dir) / "synthetic-mathtype.docx"
            _build_docx(
                docx_path,
                document_xml_name="document_mathtype.xml",
                document_rels_name="document_equation3.rels.xml",
                embedding_hex_name="mathtype_like_payload.hex",
                preview_bytes=b"WMF-SYNTHETIC",
            )

            records = self.module.detect_equation_editor_3_ole(docx_path)

        self.assertEqual(records, [])


if __name__ == "__main__":
    unittest.main()
