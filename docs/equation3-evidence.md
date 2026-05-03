# Equation Editor 3.0 Evidence Pack

This project includes an implemented limited Equation Editor 3.0 path for source-first canonical MathML evidence.

The path is intentionally bounded. It is not a universal Equation Editor 3.0 converter and it does not produce Word replacement output.

## Supported Input Scope

The Equation3 executor can read:

- DOCX OLE embeddings that expose Equation Editor `Equation Native` payloads.
- Legacy binary Word `.doc` OLE compound files that expose `ObjectPool/*/Equation Native` streams.

The source must remain identifiable as Equation Editor 3.0. `Equation.DSMT*`, MathType-marked MTEF3 material, AxMath material, and preview-only objects are not accepted as Equation Editor 3.0 conversion evidence.

## Canonical Output

For supported payloads, the executor writes:

- `canonical-mathml/*.xml`
- `canonicalization-summary.json`
- `validation-evidence.json`

Unsupported payloads are reported as blocker evidence instead of guessed MathML.

The current parser covers the implemented limited MTEF v2/v3 slice named in the executor's `supported_slice`, including common script, fraction, slash-fraction, root, bar, fence, matrix, pile, limit/operator, BigOp, standalone sum operator, selected embellishment, typeface, spacing, and narrow observed legacy post-`END` handling.

## Public Fixtures

The public repository includes minimal Apache-derived native-stream controls under:

- `tests/fixtures/equation_editor_3_ole/apache_poi/`

These are hex-encoded `Equation Native` streams extracted from Apache POI test data at commit `e6a04b49211e23c704fcdbe524d99d2f4486b083`.

They provide:

- two conversion-positive MTEF v3 controls from `Bug61268.doc`
- one unsupported-regression control from `Bug50936_1.doc` for `selector=43 variation=2`

The full upstream `.doc` files are not vendored. `SOURCES.json`, `NOTICE.md`, and `THIRD_PARTY_LICENSES/Apache-POI-LICENSE.txt` record the attribution, fixed upstream URLs, source document hashes, stream hashes, expected outcomes, and upstream license copy.

## Research-Control Evidence

Local research-control sweeps also exercised larger real-document corpora, including official 3GPP standards and a small set of independent web samples. Those files are not public fixtures unless their redistribution status is separately cleared.

The strongest local research-control sweep reported `100 / 100` assembled Equation3-positive documents and `2791 / 2791` formulas converted under the implemented limited source-core gate. That result is not a random global historical coverage estimate.

## Claim Boundary

This project can claim:

- source-first Equation Editor 3.0 detection for the supported DOCX and legacy `.doc` surfaces
- limited MTEF v2/v3 to canonical MathML conversion with formula-count parity when the supported slice is satisfied
- source-to-canonical provenance and blocker records for unsupported fragments
- minimal public native-stream fixtures with Apache-2.0 attribution

This project does not claim:

- universal support for all historical Equation Editor 3.0 documents
- a statistically valid global coverage percentage
- Word visual fill-back, OMML replacement, DOCX/PDF deliverability, or pixel parity for Equation3 output
- that MathType, `Equation.DSMT*`, or AxMath objects are Equation Editor 3.0 evidence
