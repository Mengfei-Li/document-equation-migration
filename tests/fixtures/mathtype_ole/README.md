# MathType Detector Fixtures

This directory contains minimal synthetic test fixtures for MathType OLE detection. It does not contain real private documents.

The fixtures model DOCX container fragments rather than complete user documents:

- `main_story/`
  - Positive MathType OLE case in the main document story.
  - `word/embeddings/oleObject1.bin.b64` is a base64-encoded synthetic CFB marker payload with artificial MathType/MTEF marker streams.
- `comment_story/`
  - Positive MathType OLE case in a comments story.
  - It reuses the same synthetic marker payload style as `main_story/`.
- `missing_equation_native/`
  - Negative/blocked control case for an object without an `Equation Native` stream.
- `non_mathtype_ole/`
  - Negative control case for a non-MathType OLE object.
- `live_control/`
  - Positive MathType OLE case with a real binary MTEF OLE payload from an MIT-licensed upstream test fixture.
  - `word/embeddings/oleObject1.bin.b64` is base64-encoded text and is decoded only into temporary test files.

The marker fixtures are detector/routing fixtures only. `live_control/` provides source material for live-conversion tooling tests when the documented external prerequisites are available, but the default public test suite only verifies fixture integrity, source detection, and temporary DOCX packaging. These fixtures are not intended to represent production MathType output or visual parity.
