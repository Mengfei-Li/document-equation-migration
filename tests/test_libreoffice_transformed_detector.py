import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from document_equation_migration.detectors.libreoffice_transformed import detect_libreoffice_transformed


FIXTURE_ROOT = PROJECT_ROOT / "tests" / "fixtures" / "libreoffice_transformed"


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


class LibreOfficeTransformedDetectorTests(unittest.TestCase):
    def test_flags_bridge_result_only_when_generator_and_provenance_exist(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = build_odf_archive(FIXTURE_ROOT / "libreoffice_bridge", ".odt", Path(temp_dir))

            result = detect_libreoffice_transformed(archive_path)

            self.assertEqual(result["formula_count"], 1)
            formula = result["formulas"][0]
            self.assertEqual(formula["source_family"], "libreoffice-transformed")
            self.assertEqual(formula["source_role"], "transformed-source")
            self.assertEqual(formula["storage_kind"], "odf-draw-object-subdocument")
            self.assertEqual(formula["risk_level"], "high")
            self.assertIn("transformed-source", formula["risk_flags"])
            self.assertEqual(formula["libreoffice"]["original_origin"], "omml")
            self.assertEqual(formula["libreoffice"]["conversion_mode"], "headless-convert-to")
            self.assertEqual(formula["libreoffice"]["input_filter"], "Office Open XML Text")
            self.assertTrue(formula["libreoffice"]["profile_isolated"])
            self.assertEqual(formula["libreoffice"]["producer_version"], "25.2.0.3")

    def test_requires_bridge_provenance_not_just_libreoffice_generator(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = build_odf_archive(FIXTURE_ROOT / "libreoffice_no_bridge", ".odt", Path(temp_dir))

            result = detect_libreoffice_transformed(archive_path)

            self.assertEqual(result["formula_count"], 0)
            self.assertEqual(result["formulas"], [])


if __name__ == "__main__":
    unittest.main()
