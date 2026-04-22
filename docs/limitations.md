# Limitations

This project is a research-grade converter. Use it with review, especially for high-stakes documents.

## Conversion Quality

- Complex formulas may need spot checks.
- Some MathType constructs may not map cleanly to MathML or OMML.
- Rule-based normalization only covers known defects.
- A successful output file does not prove formula-level semantic equivalence.
- The bundled patch improves the tested conversion route but is not a full upstream replacement for a maintained MathType parser.

## Layout

MathType OLE objects often include preview images with fixed dimensions. Replacing them with native Word math changes rendering metrics. The result can be editable and mathematically useful while still having different pagination or line breaks.

Editable OMML output and print-identical layout are separate goals. A document can be useful for editing, search, and downstream processing while still requiring human visual review before publication.

In this research-preview release, `review-gated` MathType output is not an automated deliverable claim. It means Word export succeeded and traceable visual evidence exists, but a human reviewer must inspect layout drift and formula risk before production use.

For the current release-facing evidence boundary and manual-review interpretation, see [MathType evidence pack](mathtype-evidence.md).

The guarded MathType layout-preservation option is opt-in. It is based on current sample evidence and should not be treated as a universal correction for all documents.

Current MathType output is not lossless, pixel-identical, or production-ready without review. The evidence supports a research-preview manual-review workflow: complete conversion and replacement counts, Word export, page-count checks, unmatched-page checks, visual changed-ratio metrics, and human inspection.

Public marker fixtures are detector/routing fixtures only. They are useful for source discovery and dry-run planning, but they are not proof that the live MTEF-to-OMML conversion path works. Real binary MTEF OLE payloads used for local control runs are not included in the public repository unless a fixture record can document the upstream source, SHA-256, license, NOTICE attribution, and artifact hygiene.

## Document Structure

The replacement logic is conservative. If a Word run contains unexpected mixed content around an OLE object, the script should stop rather than silently corrupt the document.

## Platform

The OMML route is Windows-first because it relies on Office's `MML2OMML.XSL`. Optional PDF validation requires Word COM automation and therefore does not run on GitHub-hosted Linux runners.

## Data Safety

Private or copyrighted `.docx` files are not used as public test fixtures. Small synthetic fixtures are preferred when a test needs to demonstrate one formula pattern at a time.
