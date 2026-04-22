# Research-Preview Release Notes

This repository is a research-preview release, not a production-ready or lossless converter release.

## Release Positioning

This release should be understood as follows:

- `document-equation-migration` is a research-preview toolkit for detecting formula sources and experimenting with source-first migration paths.
- The current strongest live conversion path is MathType OLE to MathML to OMML to editable Word equations.
- Current MathType outputs can be useful for editing, search, review, and downstream processing.
- Current MathType outputs still require human review before production use.

This release does not claim:

- lossless MathType conversion
- pixel-identical Word layout
- production-ready converter
- fully automated deliverable output for all documents
- visual parity guaranteed

## MathType Evidence Summary

The current evidence supports the following limited claims:

- two real MathType documents have been converted and exported through Word in local validation
- the guarded layout-preservation option can recover page count on the current validated samples
- conversion/replacement counts, Word export, and visual comparison artifacts are available for the current evidence set
- the guarded layout-preservation option is reachable from the execution surface
- wrapper resume/chunk behavior exists for long MathType runs
- external MathType live-conversion prerequisites are documented

The current evidence does not support these claims:

- universal MathType support
- semantic equivalence for every formula
- pixel-identical layout
- visual gate pass for current guarded real-sample outputs
- public live-convertible fixture coverage

## Release-Facing Status Terms

Use these terms consistently:

- `automated deliverable candidate`: Word export passes, conversion/replacement counts are complete, and the configured visual gate passes.
- `manual-review candidate`: Word export passes and conversion/replacement counts are complete, but visual comparison remains review-gated or requires human inspection.
- `review-gated`: evidence exists, but the result must not be treated as automatically ready for production.
- `blocked`: required output, conversion, Word export, or validation evidence is missing or failed.

For the current research preview, guarded MathType outputs are manual-review candidates, not automated deliverable candidates.

## Fixture Policy

Public marker fixtures are for detector, routing, and dry-run tests. They are not live-conversion proof.

Real binary MTEF OLE payloads may be used as local research controls under `research-artifacts/`, but they are not automatically approved as public fixtures.

A public live-convertible fixture requires a separate task that records:

- upstream source path
- SHA-256
- license and attribution
- NOTICE update
- minimal fixture scope
- confirmation that generated DOCX/PDF/WMF/BIN artifacts are not committed

## Dependency Notes

MathType live conversion requires external tools:

- full JDK 17 or newer with `java.exe`, `javac.exe`, and `jdk.charsets`
- Microsoft Office `MML2OMML.XSL`
- bootstrapped `transpect/mathtype-extension` and `jure/mathtype_to_mathml` sources
- optional Word desktop for PDF export validation
- optional visual comparison dependencies for PDF visual gates

The public release does not vendor JDK/JRE archives, Office files, complete third-party repositories, generated DOCX/PDF outputs, or private sample documents.
