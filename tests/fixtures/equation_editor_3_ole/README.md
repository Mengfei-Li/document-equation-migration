# Equation Editor 3.0 Detector Fixtures

This directory contains only minimal synthetic test fixtures. It does not contain real private documents.

- `document_equation3.xml` / `document_equation3.rels.xml`
  - Positive `Equation.3` case in the main document story.
- `document_preview_only.rels.xml`
  - Fallback case with only a preview image and no native payload.
- `document_field_code_only.xml`
  - Detection case without `ProgID`, using only the `EMBED Equation` field code and header probe.
- `document_mathtype.xml`
  - Negative `Equation.DSMT4` case.
- `equation3_native_payload.hex`
  - Synthetic `EQNOLEFILEHDR + MTEF v3` payload.
- `mathtype_like_payload.hex`
  - Synthetic non-Equation Editor 3.0 payload.
