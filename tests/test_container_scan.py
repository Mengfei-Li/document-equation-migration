import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

import document_equation_migration.container_scan as container_scan_module
from document_equation_migration.container_scan import scan_container


DOCX_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <w:r>
        <w:object>
          <v:shape id="_x0000_i1025">
            <v:imagedata r:id="rId1" />
          </v:shape>
          <o:OLEObject ProgID="Equation.DSMT4" r:id="rId2" />
        </w:object>
      </w:r>
      <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>
    </w:p>
  </w:body>
</w:document>
"""

ODT_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                         xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                         xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
                         xmlns:math="http://www.w3.org/1998/Math/MathML">
  <office:body>
    <office:text>
      <text:p>
        <draw:object xlink:href="./Object 1" xmlns:xlink="http://www.w3.org/1999/xlink" />
        <math:math><math:mi>x</math:mi></math:math>
      </text:p>
    </office:text>
  </office:body>
</office:document-content>
"""

FODT_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                 xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
                 xmlns:math="http://www.w3.org/1998/Math/MathML">
  <office:body>
    <math:math><math:mi>y</math:mi></math:math>
    <draw:object />
  </office:body>
</office:document>
"""


class ContainerScanTests(unittest.TestCase):
    def test_scan_legacy_doc_detects_ole_equation_native_streams(self) -> None:
        class FakeStream:
            def __init__(self, data: bytes) -> None:
                self.data = data

            def read(self) -> bytes:
                return self.data

        class FakeOle:
            streams = {
                "WordDocument": b"EMBED Equation.3 EMBED Equation.3",
                "ObjectPool/_1/\x01CompObj": b"Microsoft Equation 3.0\x00Equation.3\x00",
                "ObjectPool/_1/Equation Native": b"native-1",
                "ObjectPool/_2/Equation Native": b"native-2",
            }

            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb) -> None:
                return None

            def listdir(self):
                return [name.split("/") for name in self.streams]

            def exists(self, name: str) -> bool:
                return name in self.streams

            def openstream(self, name):
                key = "/".join(name) if isinstance(name, list) else name
                return FakeStream(self.streams[key])

        class FakeOlefile:
            @staticmethod
            def isOleFile(path) -> bool:
                return True

            @staticmethod
            def OleFileIO(path) -> FakeOle:
                return FakeOle()

        with tempfile.TemporaryDirectory() as tmp_dir:
            path = Path(tmp_dir) / "legacy.doc"
            path.write_bytes(b"OLE-CFB")
            original_olefile = container_scan_module.olefile
            container_scan_module.olefile = FakeOlefile
            try:
                result = scan_container(path)
            finally:
                container_scan_module.olefile = original_olefile

        self.assertEqual(result.container_format, "doc")
        self.assertEqual(result.package_kind, "ole-cfb")
        self.assertEqual(
            result.embedding_targets,
            ["ObjectPool/_1/Equation Native", "ObjectPool/_2/Equation Native"],
        )
        self.assertEqual(result.object_parts, ["ObjectPool/_1", "ObjectPool/_2"])
        self.assertEqual(result.story_parts[0].ole_object_count, 2)
        self.assertEqual(result.story_parts[0].field_code_count, 2)

    def test_scan_docx_detects_story_parts_embeddings_and_media(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            path = Path(tmp_dir) / "sample.docx"
            with zipfile.ZipFile(path, "w") as zf:
                zf.writestr("word/document.xml", DOCX_XML)
                zf.writestr("word/embeddings/oleObject1.bin", b"Equation Native")
                zf.writestr("word/media/image1.wmf", b"preview")
                zf.writestr("word/_rels/document.xml.rels", b"<Relationships/>")

            result = scan_container(path)

            self.assertEqual(result.container_format, "docx")
            self.assertEqual(result.embedding_targets, ["word/embeddings/oleObject1.bin"])
            self.assertEqual(result.media_targets, ["word/media/image1.wmf"])
            self.assertEqual(result.story_parts[0].omml_count, 1)
            self.assertEqual(result.story_parts[0].ole_object_count, 1)

    def test_scan_odt_detects_math_and_object_parts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            path = Path(tmp_dir) / "sample.odt"
            with zipfile.ZipFile(path, "w") as zf:
                zf.writestr("content.xml", ODT_XML)
                zf.writestr("Object 1/content.xml", ODT_XML)

            result = scan_container(path)

            self.assertEqual(result.container_format, "odt")
            self.assertEqual(result.object_parts, ["Object 1/content.xml"])
            self.assertGreaterEqual(result.story_parts[0].odf_math_count, 1)

    def test_scan_fodt_detects_flat_math(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            path = Path(tmp_dir) / "sample.fodt"
            path.write_bytes(FODT_XML)

            result = scan_container(path)

            self.assertEqual(result.container_format, "fodt")
            self.assertEqual(result.entry_count, 1)
            self.assertEqual(result.story_parts[0].odf_math_count, 1)


if __name__ == "__main__":
    unittest.main()
