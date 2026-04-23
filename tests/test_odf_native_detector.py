import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from document_equation_migration.detectors.odf_native import detect_odf_native


FIXTURE_ROOT = PROJECT_ROOT / "tests" / "fixtures" / "odf_native"
LIBREOFFICE_FIXTURE_ROOT = PROJECT_ROOT / "tests" / "fixtures" / "libreoffice_transformed"


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


class OdfNativeDetectorTests(unittest.TestCase):
    def test_detects_standalone_formula_package(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = build_odf_archive(FIXTURE_ROOT / "standalone_formula", ".odf", Path(temp_dir))

            result = detect_odf_native(archive_path)

            self.assertEqual(result["container_format"], "odf")
            self.assertEqual(result["formula_count"], 1)
            formula = result["formulas"][0]
            self.assertEqual(formula["source_family"], "odf-native")
            self.assertEqual(formula["source_role"], "native-source")
            self.assertEqual(formula["storage_kind"], "odf-formula-root")
            self.assertEqual(formula["doc_part_path"], "content.xml")
            self.assertEqual(formula["canonical_mathml_status"], "available")

    def test_detects_embedded_subdocument_formula(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = build_odf_archive(FIXTURE_ROOT / "embedded_native", ".odt", Path(temp_dir))

            result = detect_odf_native(archive_path)

            self.assertEqual(result["container_format"], "odt")
            self.assertEqual(result["formula_count"], 1)
            formula = result["formulas"][0]
            self.assertEqual(formula["storage_kind"], "odf-draw-object-subdocument")
            self.assertEqual(formula["embedding_target"], "./Object 1")
            self.assertEqual(formula["doc_part_path"], "Object 1/content.xml")
            self.assertEqual(formula["risk_level"], "low")

    def test_returns_empty_when_no_mathml_payload_exists(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = build_odf_archive(FIXTURE_ROOT / "no_math", ".odt", Path(temp_dir))

            result = detect_odf_native(archive_path)

            self.assertEqual(result["formula_count"], 0)
            self.assertEqual(result["formulas"], [])

    def test_suppresses_native_classification_for_libreoffice_bridge_fixture(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = build_odf_archive(LIBREOFFICE_FIXTURE_ROOT / "libreoffice_bridge", ".odt", Path(temp_dir))

            result = detect_odf_native(archive_path)

            self.assertEqual(result["container_format"], "odt")
            self.assertEqual(result["formula_count"], 0)
            self.assertEqual(result["source_counts"], {"odf-native": 0})
            self.assertEqual(result["formulas"], [])


if __name__ == "__main__":
    unittest.main()
