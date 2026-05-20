# Canonical Evidence Schema

This project uses canonical MathML as the shared structured target for equation
sources. The evidence schema below defines the minimum fields that make a
canonical artifact, a blocked gate, or an export gate auditable across the
current source families.

The schema is a contract for evidence records. It does not upgrade any source
line capability, remove external-tool gates, or turn research-preview output
into deliverable output.

## Source Families

| Source family | Current canonical target status | Evidence mode |
|---|---|---|
| `mathtype-ole` | `external-tool-gated` | Canonical MathML may be accepted only after an approved external conversion run emits valid MathML artifacts and provenance. |
| `omml-native` | `implemented-basic` | Native OMML can emit canonical MathML for the implemented structure subset. |
| `odf-native` | `implemented` | Native ODF MathML can be preserved as canonical MathML with provenance. |
| `libreoffice-transformed` | `bridge-review-gated` | Bridge provenance is review-gated and must not be treated as native-source evidence. |
| `equation-editor-3-ole` | `implemented-limited` | The implemented limited Equation Editor 3.0 slice can emit canonical MathML with explicit claim boundaries. |
| `axmath-ole` | `export-gated` | AxMath requires reviewed export artifacts before canonical MathML can be accepted. |

## Canonicalization Summary

A successful or attempted canonicalization writes `canonicalization-summary.json`.
The common required fields are:

| Field | Meaning |
|---|---|
| `expected_formula_count` | Number of source formulas the step expected to account for. |
| `canonical_mathml_count` | Number of canonical MathML artifacts accepted. |
| `unsupported_fragment_count` | Number of source fragments rejected or blocked during canonicalization. |
| `formula_count_parity` | Count/parity result. Current values are `passed`, `mismatch`, `unsupported-fragments`, and `failed`. |
| `canonical_mathml_dir` | Directory containing accepted canonical MathML XML files. |
| `source_to_canonical_provenance` | Array of provenance records linking source formulas to canonical artifacts. |
| `property_summary` | Aggregated MathML property signals from accepted artifacts. |
| `unsupported_fragments` | Array of rejected fragment records. Empty when no fragments were rejected. |

Identity fields such as `artifact_type`, `source_family`, `provider`,
`target_format`, `target_stage`, `gate_status`, or `strategy` are source-line
specific today. New producers should include them when practical, but the
minimum cross-line testable contract is the field set above.

## Provenance Records

Each accepted canonical artifact must have one provenance record in
`source_to_canonical_provenance`.

| Field | Meaning |
|---|---|
| `formula_id` | Stable formula identifier within the run. |
| `canonical_artifact_path` | Path to the accepted canonical MathML XML artifact. |
| `canonical_sha256` | SHA-256 hash of the accepted canonical MathML text. |
| `preservation_status` | Short status explaining how the artifact was preserved or converted. |
| `property_signals` | Per-artifact MathML property signals used by `property_summary`. |

Source locator fields vary by family. Examples include `source_mathml_path`,
`source_omml_path`, `source_part_path`, `payload_stream_name`,
`embedding_target`, `source_sha256`, `raw_payload_sha256`, and
`equation_native_sha256`. A producer should include the strongest available
source locator and source hash for the source format it controls.

## Property Signals

`property_signals` and `property_summary` are shared utility output. Current
signals include:

- `root_attributes`
- `root_display`
- `mathml_attribute_count`
- `has_semantics`
- `has_annotation`
- `has_mfrac_linethickness`
- `has_mfrac_bevelled`
- `has_mfenced_separators`
- `has_movablelimits`
- `has_mathvariant`
- `has_accent`
- `has_accentunder`

`property_summary` must include aggregate `mathml_attribute_count`,
`root_display_values`, and `signal_counts`.

## Blocker And Export-Gate Records

When canonical MathML is not accepted, the source line writes a blocker or gate
record instead of silently emitting a partial success. The common required
fields are:

| Field | Meaning |
|---|---|
| `artifact_type` | Record type, for example `mathtype-blocker-record` or `axmath-export-assisted-blocker-record`. |
| `source_family` | Source family governed by the record. |
| `canonical_target` | The canonical MathML target contract for that source family. |
| `status` | Current gate status. |
| `required_evidence` | Evidence needed before the line can be promoted. |
| `next_ready_condition` | Concrete condition for the next admissible run. |

Blocker records may also include family-specific sections such as
`canonical_artifact_admissibility`, `fixture_admissibility`,
`export_admissibility`, `review_status`, `blocker_kind`,
`external_export_dependency`, or `actions`.

## Validation Evidence

When a source line writes a validation wrapper, the minimum common fields are:

| Field | Meaning |
|---|---|
| `artifact_type` | Validation evidence record type. |
| `source_family` | Source family governed by the evidence. |
| `status` | Evidence or gate status. |

Validation wrappers should link back to the canonicalization summary, blocker
record, or canonical target contract whenever those artifacts exist. The
current wrappers use different names for this link; the canonicalization summary
remains the primary source for accepted artifact counts and provenance.

## Claim Boundaries

Canonical evidence must preserve source-line claim boundaries. Current records
use one or more of these fields:

- `conversion_claim`
- `limited_conversion_claim`
- `general_converter_claim`
- `deliverability_claim`
- `word_visual_fill_back_claim`
- `claim_boundary`

Blocked lines such as `mathtype-ole`, `axmath-ole`, and
`libreoffice-transformed` must not be upgraded by schema presence alone.
Equation Editor 3.0 may state `limited_conversion_claim=true` only for the
implemented limited slice while keeping `general_converter_claim=false`.

## Test Surface

The executable schema constants live in
`src/document_equation_migration/canonical_mathml_evidence.py`. Focused tests
assert that this document lists the required fields and that minimal successful
summary, provenance, blocker, and validation examples satisfy the schema.
