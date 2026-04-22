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

These fixtures are detector/routing fixtures only. They are not live-conversion proof, not redistributable evidence for real MathType formulas, and not intended to represent production MathType output.
