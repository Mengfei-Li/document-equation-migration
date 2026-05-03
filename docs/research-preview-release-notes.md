# Research-Preview Release Notes

This repository is a research-preview release, not a production-ready or lossless converter release.

For a consolidated explanation of the current MathType claim boundary, see [MathType evidence pack](mathtype-evidence.md).

For the Equation Editor 3.0 source-core claim boundary, see [Equation Editor 3.0 evidence pack](equation3-evidence.md).

## Release Positioning

This release should be understood as follows:

- `document-equation-migration` is a research-preview toolkit for detecting formula sources and experimenting with source-first migration paths.
- The current strongest live conversion path is MathType OLE to MathML to OMML to editable Word equations.
- Equation Editor 3.0 has an implemented limited source-core path to canonical MathML, not a Word replacement or universal-conversion claim.
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
- the current evidence boundary is summarized in a dedicated public evidence-pack document

The current evidence does not support these claims:

- universal MathType support
- semantic equivalence for every formula
- pixel-identical layout
- visual gate pass for current guarded real-sample outputs
- public live-convertible fixture coverage

## Equation Editor 3.0 Evidence Summary

The current evidence supports the following limited claims:

- supported Equation Editor 3.0 native payloads can be converted to canonical MathML through the internal MTEF v2/v3 path
- legacy binary Word `.doc` ObjectPool `Equation Native` streams are a supported source surface for that limited path
- public Apache-derived native-stream controls exist for conversion-positive and unsupported-regression coverage
- unsupported Equation3 structures produce blocker evidence instead of guessed MathML

The current evidence does not support these claims:

- universal Equation Editor 3.0 support
- a statistically valid global historical coverage percentage
- Word visual fill-back, DOCX/PDF deliverability, or pixel parity for Equation3 output
- treating MathType, `Equation.DSMT*`, or AxMath material as Equation Editor 3.0 evidence

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

Public live-convertible native payload controls require:

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

Related public documentation:

- [MathType evidence pack](mathtype-evidence.md)
- [Limitations](limitations.md)
- [Dependencies](dependencies.md)
