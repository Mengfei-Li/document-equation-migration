# Equation Editor 3.0 Fixtures

This directory contains minimal Equation Editor 3.0 fixtures. It does not contain private documents or full source Word files.

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

## Apache POI native-stream controls

- `apache_poi/`
  - Minimal Apache-derived `Equation Native` stream controls with Apache-2.0 attribution and fixed upstream source metadata.
  - The full upstream `.doc` files are not vendored in the public fixture tree.
