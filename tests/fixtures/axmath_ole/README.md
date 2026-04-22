# AxMath Detector Fixtures

This directory contains minimal synthetic test fixtures for AxMath OLE detection. It does not contain real private documents.

The fixtures model a few DOCX container fragments:

- `main_prog_id/`
  - Positive AxMath case using an `Equation.AxMath` program identifier.
  - `word/embeddings/oleObject1.bin` is a short ASCII marker payload, not a real user equation.
  - `word/media/image1.wmf` is an 8-byte preview marker used only to exercise preview-target detection.
- `main_field_code/`
  - Positive AxMath case using field-code evidence.
  - `word/embeddings/oleObject2.bin` is a short ASCII marker payload.
- `comment_story/`
  - Positive AxMath case in a comments story.
  - `word/embeddings/oleObject3.bin` is a short ASCII marker payload.
- `no_axmath/`
  - Negative control case.
  - `word/embeddings/oleObject4.bin` is a short ASCII marker payload for a non-AxMath object.

These fixtures are detector/routing fixtures only. They are not live-conversion proof and are not intended to represent production AxMath output.
