import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from document_equation_migration.manifest import DocumentRecord, FormulaRecord, Manifest
from document_equation_migration.source_taxonomy import SourceFamily, SourceRole


class ManifestTests(unittest.TestCase):
    def test_manifest_updates_source_counts(self) -> None:
        manifest = Manifest(
            document=DocumentRecord(
                document_id="sample",
                input_path="sample.docx",
                input_sha256="abc",
                container_format="docx",
                detector_version="0.1.0",
            ),
            formulas=[
                FormulaRecord(
                    formula_id="f-001",
                    source_family=SourceFamily.MATHTYPE_OLE,
                    source_role=SourceRole.NATIVE_SOURCE,
                    doc_part_path="word/document.xml",
                    story_type="main",
                    storage_kind="ole-embedded",
                ),
                FormulaRecord(
                    formula_id="f-002",
                    source_family=SourceFamily.OMML_NATIVE,
                    source_role=SourceRole.NATIVE_SOURCE,
                    doc_part_path="word/document.xml",
                    story_type="main",
                    storage_kind="omml-inline",
                ),
                FormulaRecord(
                    formula_id="f-003",
                    source_family=SourceFamily.OMML_NATIVE,
                    source_role=SourceRole.NATIVE_SOURCE,
                    doc_part_path="word/header1.xml",
                    story_type="header",
                    storage_kind="omml-inline",
                ),
            ],
        )

        manifest.update_source_counts()

        self.assertEqual(manifest.document.source_counts["mathtype-ole"], 1)
        self.assertEqual(manifest.document.source_counts["omml-native"], 2)

    def test_manifest_json_serializes_enum_values(self) -> None:
        manifest = Manifest(
            document=DocumentRecord(
                document_id="sample",
                input_path="sample.docx",
                input_sha256="abc",
                container_format="docx",
                detector_version="0.1.0",
            ),
            formulas=[
                FormulaRecord(
                    formula_id="f-001",
                    source_family=SourceFamily.UNKNOWN_OLE,
                    source_role=SourceRole.PREVIEW_ONLY,
                    doc_part_path="word/document.xml",
                    story_type="main",
                    storage_kind="graphic-only",
                )
            ],
        )
        payload = manifest.to_dict()

        self.assertEqual(payload["formulas"][0]["source_family"], "unknown-ole")
        self.assertEqual(payload["formulas"][0]["source_role"], "preview-only")


if __name__ == "__main__":
    unittest.main()
