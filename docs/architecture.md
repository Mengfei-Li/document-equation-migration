# Architecture

`document-equation-migration` is a source-first equation migration pipeline. It avoids OCR when the original document still contains MathType OLE / MTEF, native OMML, ODF MathML, or other structured formula data.

The shared intermediate target is canonical MathML. Word-native OMML output is a downstream editing format, not the only project target.

## Main Stages

1. Extract the `.docx` archive into a temporary directory.
2. Copy `word/embeddings/oleObject*.bin` files into the output workspace.
3. Use `Ole2XmlCli.java` and `transpect/mathtype-extension` to convert MathType OLE data into MTEF XML.
4. Use patched `mathtype_to_mathml` XSLT to transform MTEF XML into MathML.
5. Run `normalize_mathml.py` to fix known structural defects.
6. Use Office's `MML2OMML.XSL` to convert MathML into OMML.
7. Use `replace_docx_ole_with_omml.py` to replace OLE object runs with OMML nodes in a copy of the original `.docx`.
8. Generate mapping and validation artifacts for review.

## Canonical Target Contract

The detector-first executor records a `canonical_target` block for every routed source family. The block makes the target representation explicit:

- `target_format`: currently `canonical-mathml`
- `contract_status`: whether the source line is implemented, externally gated, fixture gated, export gated, or bridge-review gated
- `expected_artifacts`: where canonical MathML evidence or blocker evidence should appear
- `required_evidence`: what must exist before a conversion claim is made

Current source-line contracts:

- MathType OLE: Equation Native / MTEF to normalized MathML, gated by external converter prerequisites.
- OMML native: internal basic OMML-to-canonical-MathML conversion for common presentation structures.
- Equation Editor 3.0: fixture-gated MTEF v3 candidate path, with no conversion claim until stronger fixtures exist.
- AxMath: export-assisted path requiring reviewed canonical MathML artifacts, or LaTeX plus a separately validated LaTeX-to-MathML step.
- ODF native: preserves existing ODF MathML as canonical MathML artifacts.
- LibreOffice-transformed: bridge provenance review gate, not a native source claim.

## Key Files

- `run_docx_open_source_pipeline.ps1`: document-level pipeline entry point.
- `probe_formula_pipeline.ps1`: formula-level OLE to MathML / OMML converter.
- `java_bridge/Ole2XmlCli.java`: minimal Java wrapper around the external OLE converter.
- `normalize_mathml.py`: post-processing rules for common MathML defects.
- `src/document_equation_migration/canonical_target.py`: source-line contract for canonical MathML convergence.
- `src/document_equation_migration/omml_to_mathml.py`: basic internal OMML-to-presentation-MathML conversion used by the OMML-native executor slice.
- `replace_docx_ole_with_omml.py`: OpenXML replacement logic.
- `docx_math_object_map.py`: mapping between embedded objects and document context.
- `analyze_formula_risks.py`: rule-based risk classification.
- `scripts/bootstrap_third_party.ps1`: clones external converter projects and applies local quality patches.
- `patches/mathtype_to_mathml-quality-fixes.patch`: XSLT fixes used by this pipeline.

## Design Tradeoffs

The pipeline favors repeatable batch processing over perfect layout preservation. After OLE objects are replaced by native OMML, Word may render formulas with different metrics than MathType OLE preview images. This can change line breaks and pagination even when the formula semantics are acceptable.
