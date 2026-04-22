# Architecture

`document-equation-migration` is a file-based conversion pipeline. It avoids OCR when the original document still contains MathType OLE / MTEF data.

## Main Stages

1. Extract the `.docx` archive into a temporary directory.
2. Copy `word/embeddings/oleObject*.bin` files into the output workspace.
3. Use `Ole2XmlCli.java` and `transpect/mathtype-extension` to convert MathType OLE data into MTEF XML.
4. Use patched `mathtype_to_mathml` XSLT to transform MTEF XML into MathML.
5. Run `normalize_mathml.py` to fix known structural defects.
6. Use Office's `MML2OMML.XSL` to convert MathML into OMML.
7. Use `replace_docx_ole_with_omml.py` to replace OLE object runs with OMML nodes in a copy of the original `.docx`.
8. Generate mapping and validation artifacts for review.

## Key Files

- `run_docx_open_source_pipeline.ps1`: document-level pipeline entry point.
- `probe_formula_pipeline.ps1`: formula-level OLE to MathML / OMML converter.
- `java_bridge/Ole2XmlCli.java`: minimal Java wrapper around the external OLE converter.
- `normalize_mathml.py`: post-processing rules for common MathML defects.
- `replace_docx_ole_with_omml.py`: OpenXML replacement logic.
- `docx_math_object_map.py`: mapping between embedded objects and document context.
- `analyze_formula_risks.py`: rule-based risk classification.
- `scripts/bootstrap_third_party.ps1`: clones external converter projects and applies local quality patches.
- `patches/mathtype_to_mathml-quality-fixes.patch`: XSLT fixes used by this pipeline.

## Design Tradeoffs

The pipeline favors repeatable batch processing over perfect layout preservation. After OLE objects are replaced by native OMML, Word may render formulas with different metrics than MathType OLE preview images. This can change line breaks and pagination even when the formula semantics are acceptable.
