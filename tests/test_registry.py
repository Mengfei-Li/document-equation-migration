import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from document_equation_migration.container_scan import ContainerScanResult
from document_equation_migration.detectors.base import DetectorContext, FormulaDetector
from document_equation_migration.detectors.registry import DetectorRegistry
from document_equation_migration.source_taxonomy import SourceFamily


class DummyDetector(FormulaDetector):
    source_family = SourceFamily.UNKNOWN_OLE
    name = "dummy"

    def detect(self, context: DetectorContext) -> list:
        return []


class RegistryTests(unittest.TestCase):
    def test_registry_registers_and_lists_detector(self) -> None:
        registry = DetectorRegistry()
        registry.register(DummyDetector())
        self.assertEqual(registry.available_source_families(), ["unknown-ole"])

    def test_registry_rejects_duplicate_source_family(self) -> None:
        registry = DetectorRegistry()
        registry.register(DummyDetector())
        with self.assertRaises(ValueError):
            registry.register(DummyDetector())

    def test_registry_run_executes_registered_detectors(self) -> None:
        registry = DetectorRegistry()
        registry.register(DummyDetector())
        context = DetectorContext(
            scan_result=ContainerScanResult(
                input_path="sample.docx",
                input_sha256="abc",
                container_format="docx",
                package_kind="ooxml-zip",
                entry_count=1,
                story_parts=[],
            ),
            detector_version="0.1.0",
        )
        self.assertEqual(registry.run(context), [])


if __name__ == "__main__":
    unittest.main()
