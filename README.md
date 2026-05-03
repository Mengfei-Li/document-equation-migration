# Document Equation Migration

Research-grade tools for detecting and migrating structured equation sources in Word and related document formats, including MathType OLE, native OMML, Equation Editor 3.0 OLE, ODF MathML, and related bridge sources.

This project is Windows-first because the OMML conversion path uses Microsoft Office's `MML2OMML.XSL`, and the optional PDF validation path uses Word COM automation.

## Status

This repository is a research preview, not a guaranteed lossless converter.

Current published release: `v0.2.0-research-preview`.

The current strongest deliverable-oriented route is still MathType OLE to MathML / OMML / editable Word equations. The detector-first executor also has source-core canonical MathML slices for native OMML, ODF MathML, and an implemented limited Equation Editor 3.0 path.

## What It Does

- Extracts `oleObject*.bin` files from a Word `.docx`.
- Scans supported DOCX, legacy `.doc` OLE, and ODF/FODT containers for formula source evidence.
- Writes formula-source manifests, routing reports, execution plans, canonical-target contracts, and evidence/blocker records.
- Converts MathType OLE / MTEF content to intermediate XML.
- Converts the intermediate XML to MathML.
- Normalizes common MathML defects found in MathType-to-MathML conversion output.
- Converts MathML to OMML with Office's `MML2OMML.XSL`.
- Replaces OLE formula objects in a copy of the original `.docx`.
- Produces a LaTeX validation preview and risk classification output.
- Converts the implemented limited Equation Editor 3.0 MTEF v2/v3 slice to canonical MathML with provenance for supported payloads.

## What It Does Not Promise

- It does not guarantee pixel-identical layout after conversion.
- It does not guarantee semantic equivalence for every possible MathType equation.
- It does not claim universal support for every historical Equation Editor 3.0 document.
- It does not claim a statistically valid global Equation Editor 3.0 coverage percentage.
- It does not claim Word/DOCX/PDF deliverability for Equation Editor 3.0 output.
- It does not include proprietary or third-party sample documents.
- It does not vendor a JDK or large third-party runtime binaries.
- It does not replace legal review for documents that you do not own or cannot redistribute.

## Pipeline

```text
DOCX
  -> word/embeddings/oleObject*.bin
  -> MathType / MTEF XML
  -> MathML
  -> normalized MathML
  -> OMML
  -> new DOCX with editable Word math
```

## Requirements

- Windows.
- Python 3.11 or newer.
- Full Java JDK 17 or newer with `javac.exe` and `jdk.charsets`. A JRE or stripped runtime image is not enough for first-run compilation or JRuby/Nokogiri extraction.
- Microsoft Office with `MML2OMML.XSL`.
- Optional: `pandoc` for LaTeX validation previews.
- Optional: Microsoft Word desktop for PDF export validation.

Python packages:

```powershell
python -m pip install -r requirements.txt
```

Optional visual PDF comparison packages:

```powershell
python -m pip install -r requirements-visual.txt
```

Prepare third-party converter sources:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\bootstrap_third_party.ps1
```

The bootstrap script clones:

- `transpect/mathtype-extension`
- `jure/mathtype_to_mathml`

It also applies the local quality patch in `patches/mathtype_to_mathml-quality-fixes.patch`. You must comply with the licenses of those projects and their dependencies.

For the full external-tool requirements, known Java charset failure mode, and troubleshooting guidance, see [Dependencies](docs/dependencies.md).

## Quick Start

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
powershell -ExecutionPolicy Bypass -File .\scripts\bootstrap_third_party.ps1

powershell -ExecutionPolicy Bypass -File .\run_docx_open_source_pipeline.ps1 `
  -InputDocx .\input.docx `
  -OutputDir .\out `
  -MathtypeExtensionDir .\third_party\mathtype-extension `
  -MathTypeToMathMlDir .\third_party\mathtype_to_mathml `
  -Mml2OmmlXsl "C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"
```

If you do not have `pandoc` installed and only need the converted `.docx`, add `-SkipLatexPreview` to the pipeline command.

Main output:

- `out\<input-name>.omml.docx`: converted Word document.
- `out\pipeline_summary.txt`: conversion counts.
- `out\converted\summary.csv`: per-equation conversion summary.
- `out\<input-name>.omml.validation.tex`: LaTeX validation preview, unless `-SkipLatexPreview` is used.
- `out\<input-name>.omml.ole_map.json`: mapping between formulas and document context.

## Detector-First MVP Core

The repository now also contains an experimental shared core package for source detection and manifest generation:

- `src/document_equation_migration/source_taxonomy.py`
- `src/document_equation_migration/manifest.py`
- `src/document_equation_migration/container_scan.py`
- `src/document_equation_migration/detectors/base.py`
- `src/document_equation_migration/detectors/registry.py`
- `src/document_equation_migration/cli.py`

This shared core does not replace the existing MathType conversion scripts. It establishes a detector-first entry point that inventories formula sources before routing them to source-specific conversion paths.

Install the package locally for development:

```powershell
python -m pip install -e ".[test]"
```

Scan a document and write a manifest, routing report, execution plan, plus a human-readable summary:

```powershell
dem scan .\input.docx --output .\out\manifest.json --routing .\out\routing.json --execution-plan .\out\execution-plan.json --summary .\out\summary.txt
```

Equivalent module invocation:

```powershell
python -m document_equation_migration.cli scan .\input.docx --output .\out\manifest.json --routing .\out\routing.json --execution-plan .\out\execution-plan.json
```

Supported detector-first source families currently include:

- `mathtype-ole`
- `omml-native`
- `equation-editor-3-ole`
- `axmath-ole`
- `odf-native`
- `libreoffice-transformed`

The detector-first CLI identifies formula sources and writes a manifest; it does not yet perform full MathML / OMML / LaTeX conversion for every source family.

`routing.json` is a document-level route decision artifact. It includes:

- `recommended_sequence`: source families ordered by route priority
- `route_plan`: next action per source family
- `manual_review_required` and `manual_review_reasons`

`execution-plan.json` is a converter-oriented plan generated from `routing.json`. It includes:

- `steps`: source-family execution steps with provider name and ordered actions
- `manual_review_required`: aggregated gate for downstream validation

Preview the current execution plan without executing any converter commands:

```powershell
dem run-plan .\out\execution-plan.json --dry-run --output .\out\execution-report.json
```

`execution-report.json` is a dry-run executor report. In the current milestone:

- source-specific providers expose concrete or explicitly gated dry-run bindings
- every source-family step includes a `canonical_target` block that names `canonical-mathml` as the shared structured target and records the current gate status for that source line

Run the currently supported execution bindings:

```powershell
dem run-plan .\out\execution-plan.json --execute --output-dir .\out\execution --output .\out\execution-report.json
```

In the current milestone:

- `omml` can execute a native-preserving execution slice that extracts OMML XML fragments, writes a manifest, converts common presentation OMML structures into canonical MathML artifacts, performs a deterministic packaging pass, and records execution metadata
- `mathtype` is wired to the existing PowerShell/Python document pipeline, but external tools are blocked unless you explicitly pass `--allow-external-tools`; Word validation remains a separate gate
- `equation3` provides an internal limited Equation Editor 3.0 MTEF v2/v3 to canonical MathML path for supported DOCX OLE embeddings and legacy `.doc` ObjectPool `Equation Native` streams; Word roundtrip remains downstream and is not claimed
- `axmath` is export-assisted and stays behind external export / validation gates; the project does not claim a native static AxMath parser, and canonical MathML evidence must come from reviewed export artifacts or a validated conversion step
- `odf-native` can execute a native MathML extraction slice from ODF/FODT content, while `libreoffice-transformed` remains a bridge provenance review gate
- render parity, Word opening, and PDF export are still validation gates; an execution report alone is not proof of deliverable Word output

Execute-mode provider outputs are evidence-oriented:

- each provider output root should contain either `validation-evidence.json` or `blocker-record.json`
- `validation-evidence.json` and `blocker-record.json` should carry the same `canonical_target` contract used by the execution report
- `validation-plan.json` can exist as a supporting artifact, but it does not replace the evidence/blocker contract on its own
- `validation-gated` and `review-gated` statuses mean the slice produced traceable evidence or a review gate, not that deliverable conversion is complete

Only allow external MathType tools after the dry-run report has been inspected and Java / Office XSL / local script dependencies are ready:

```powershell
dem run-plan .\out\execution-plan.json --execute --allow-external-tools --output-dir .\out\execution --output .\out\execution-report.json
```

For MathType live conversion, verify that `JAVA_EXE` / `JAVAC_EXE` point to a full JDK and that `MML2OMML_XSL` points to an Office-provided `MML2OMML.XSL`. A runtime missing `jdk.charsets` can fail during extraction with `UnsupportedCharsetException: ISO-2022-JP`.

Validate a target DOCX and write a reusable validation report artifact:

```powershell
dem validate-docx .\out\target.docx --output-dir .\out\validation --provider omml --source-family omml-native
```

If an execute output already wrote execution metadata or validation evidence with a packaged validation target, resolve the DOCX directly from that JSON instead of reconstructing the path manually:

```powershell
dem validate-docx --target-from-metadata .\out\execution\omml-native\package\execution-metadata.json --output-dir .\out\validation --provider omml --source-family omml-native
```

For deliverable-oriented Word validation, allow Word PDF export:

```powershell
dem validate-docx .\out\target.docx --output-dir .\out\validation --provider omml --source-family omml-native --allow-word-export
```

If you also have a reference PDF and the optional visual dependencies installed, run visual comparison:

```powershell
dem validate-docx .\out\target.docx --output-dir .\out\validation --provider omml --source-family omml-native --allow-word-export --reference-pdf .\out\reference.pdf --visual-compare
```

You can tighten or relax the shared visual gate explicitly:

```powershell
dem validate-docx .\out\target.docx --output-dir .\out\validation --provider omml --source-family omml-native --allow-word-export --reference-pdf .\out\reference.pdf --visual-compare --visual-max-changed-ratio-per-page 0.02 --visual-max-unmatched-pages 0
```

`validation-report.json` distinguishes:

- `deliverable-ready`: target DOCX exists and Word PDF export passed
- `review-gated`: Word PDF export passed and visual compare ran, but the current visual gate threshold was exceeded
- `research-only`: structural evidence exists, but Word deliverability was not yet validated
- `blocked`: target file is missing, Word export failed, or requested visual comparison failed

Important: `visual_compare = passed` now means both "the compare pipeline ran" and "the current visual gate thresholds were met". If the compare pipeline runs but page-count mismatch or changed-ratio exceeds threshold, the visual check becomes `review-gated` instead of `passed`.

## First-Release Review Gate

For the research-preview release, MathType conversion results should be interpreted conservatively:

- `deliverable-ready` is an automated candidate only when Word export passes, conversion/replacement counts are complete, and the configured visual gate passes.
- `review-gated` can be a manual-review candidate when Word export passes, conversion/replacement counts are complete, source and converted page counts match, unmatched pages are zero, and the visual drift is documented for human review.
- `blocked` means the output should not be presented as a usable converted document until the failed or missing gate is resolved.

Current real MathType evidence supports the guarded layout-preservation path as a manual-review candidate, not as a pixel-identical or lossless converter. The guarded layout option remains opt-in because its current factor is sample-derived and requires broader validation.

For a structured statement of the current claim boundary, evidence classes, and manual-review gate, see [MathType evidence pack](docs/mathtype-evidence.md).

For the current Equation Editor 3.0 source-core claim boundary and public native-stream fixtures, see [Equation Editor 3.0 evidence pack](docs/equation3-evidence.md).

Before using a `review-gated` output in production, review the generated PDF, inspect changed pages, spot-check high-risk formulas, and keep the source document available for comparison.

Run the current test gate:

```powershell
python -m pytest tests -q
```

## Risk Analysis

After generating `summary.csv` and an OLE map, classify equations with:

```powershell
python .\analyze_formula_risks.py `
  .\out\converted\summary.csv `
  .\out\input.omml.ole_map.json `
  .\out\risk_analysis.json `
  .\out\risk_analysis.txt
```

The categories are:

- `auto_replace`: simple formulas that did not trigger known risk rules.
- `spot_check`: complex formulas that deserve sampling.
- `manual_review`: formulas that match patterns associated with likely conversion defects.

Risk analysis is most useful when LaTeX previews are available, so QA runs should keep LaTeX previews enabled when possible.

## Validation

Optional PDF validation requires Microsoft Word desktop:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_word_pdf.ps1 `
  -InputDocx .\out\input.omml.docx `
  -OutputPdf .\out\converted.pdf
```

Visual PDF comparison uses `PyMuPDF` and `Pillow`:

```powershell
python -m pip install -r requirements-visual.txt
python .\compare_pdf_visual.py .\original.pdf .\converted.pdf .\out\visual_compare
```

## Documentation

- [Architecture](docs/architecture.md)
- [Dependencies](docs/dependencies.md)
- [Limitations](docs/limitations.md)
- [MathType evidence pack](docs/mathtype-evidence.md)
- [Equation Editor 3.0 evidence pack](docs/equation3-evidence.md)
- [Research-preview release notes](docs/research-preview-release-notes.md)

## License

This repository's original code is licensed under the MIT License. Third-party tools referenced by this project keep their own licenses.
