from __future__ import annotations

from dataclasses import dataclass

from .source_taxonomy import SourceFamily


@dataclass(frozen=True, slots=True)
class CanonicalMathMLContract:
    source_family: str
    source_line: str
    target_format: str
    contract_status: str
    conversion_claim: bool
    binding: str
    expected_artifacts: tuple[str, ...]
    required_evidence: tuple[str, ...]
    notes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, object]:
        return {
            "source_family": self.source_family,
            "source_line": self.source_line,
            "target_format": self.target_format,
            "contract_status": self.contract_status,
            "conversion_claim": self.conversion_claim,
            "binding": self.binding,
            "expected_artifacts": list(self.expected_artifacts),
            "required_evidence": list(self.required_evidence),
            "notes": list(self.notes),
        }


def _contract(
    *,
    source_family: str,
    source_line: str,
    contract_status: str,
    conversion_claim: bool,
    binding: str,
    expected_artifacts: tuple[str, ...],
    required_evidence: tuple[str, ...],
    notes: tuple[str, ...] = (),
) -> CanonicalMathMLContract:
    return CanonicalMathMLContract(
        source_family=source_family,
        source_line=source_line,
        target_format="canonical-mathml",
        contract_status=contract_status,
        conversion_claim=conversion_claim,
        binding=binding,
        expected_artifacts=expected_artifacts,
        required_evidence=required_evidence,
        notes=notes,
    )


_CONTRACTS: dict[SourceFamily, CanonicalMathMLContract] = {
    SourceFamily.MATHTYPE_OLE: _contract(
        source_family=SourceFamily.MATHTYPE_OLE.value,
        source_line="mathtype",
        contract_status="external-tool-gated",
        conversion_claim=False,
        binding="equation-native-mtef-to-normalized-mathml",
        expected_artifacts=(
            "converted/*.mml",
            "converted/*.mathml",
            "canonical-mathml/*.xml",
            "canonicalization-summary.json",
            "validation-evidence.json",
        ),
        required_evidence=(
            "Equation Native / MTEF payload extraction evidence",
            "MathType-to-MathML converter output",
            "normalize_mathml.py post-processing result",
            "formula-count parity between source objects and accepted canonical MathML artifacts",
            "provenance from every source MathType object to each accepted canonical MathML artifact",
        ),
        notes=(
            "The route is real, but live execution remains guarded by external Java and converter prerequisites.",
            "Word OMML replacement is downstream of canonical MathML and is not the target contract itself.",
        ),
    ),
    SourceFamily.OMML_NATIVE: _contract(
        source_family=SourceFamily.OMML_NATIVE.value,
        source_line="omml",
        contract_status="implemented-basic",
        conversion_claim=True,
        binding="internal-omml-to-canonical-mathml",
        expected_artifacts=(
            "canonical-mathml/*.xml",
            "canonicalization-summary.json",
            "validation-evidence.json",
        ),
        required_evidence=(
            "extracted OMML fragments",
            "canonical MathML fragments",
            "formula count parity between extracted and canonicalized fragments",
        ),
        notes=(
            "The internal binding covers common presentation OMML structures and records unsupported structures for review.",
        ),
    ),
    SourceFamily.EQUATION_EDITOR_3_OLE: _contract(
        source_family=SourceFamily.EQUATION_EDITOR_3_OLE.value,
        source_line="equation-editor-3",
        contract_status="implemented-limited",
        conversion_claim=True,
        binding="internal-equation3-mtef-v3-to-canonical-mathml-limited",
        expected_artifacts=(
            "canonical-mathml/*.xml",
            "canonicalization-summary.json",
            "validation-evidence.json",
        ),
        required_evidence=(
            "Equation.3 source identity and non-MathType provenance",
            "MTEF v3 payload extraction evidence",
            "canonical MathML conversion output for the supported observed-structure MTEF v3 slice",
            "formula-count parity between detected Equation3 objects and accepted canonical MathML artifacts",
            "source-to-canonical provenance for every accepted artifact",
        ),
        notes=(
            "The internal binding is limited to the currently implemented MTEF v3 script, root, fraction, bar, fence, matrix, and character structures.",
            "Do not claim universal Equation Editor 3.0 support or legacy .doc direct ingestion from this limited binding.",
            "Do not count Equation.DSMT* / MathType-marked MTEF3 material as Equation Editor 3.0 evidence.",
        ),
    ),
    SourceFamily.AXMATH_OLE: _contract(
        source_family=SourceFamily.AXMATH_OLE.value,
        source_line="axmath",
        contract_status="export-gated",
        conversion_claim=False,
        binding="approved-axmath-export-to-canonical-mathml",
        expected_artifacts=("blocker-record.json",),
        required_evidence=(
            "approved AxMath/vendor export workflow",
            "reviewed MathML export artifacts, or LaTeX plus a validated LaTeX-to-MathML step",
            "semantic review against source AxMath objects",
        ),
        notes=(
            "No native AxMath static parser is claimed.",
        ),
    ),
    SourceFamily.ODF_NATIVE: _contract(
        source_family=SourceFamily.ODF_NATIVE.value,
        source_line="libreoffice-odf",
        contract_status="implemented",
        conversion_claim=True,
        binding="preserve-existing-odf-mathml",
        expected_artifacts=(
            "canonical-mathml/*.xml",
            "canonicalization-summary.json",
            "validation-evidence.json",
        ),
        required_evidence=(
            "native ODF MathML payload",
            "canonical MathML artifact preservation",
            "formula count parity between native payloads and canonical artifacts",
        ),
    ),
    SourceFamily.LIBREOFFICE_TRANSFORMED: _contract(
        source_family=SourceFamily.LIBREOFFICE_TRANSFORMED.value,
        source_line="libreoffice-odf",
        contract_status="bridge-review-gated",
        conversion_claim=False,
        binding="bridge-provenance-review-before-canonical-mathml",
        expected_artifacts=("blocker-record.json",),
        required_evidence=(
            "native source provenance or transform-chain review",
            "canonical MathML artifact accepted only after bridge review",
        ),
        notes=(
            "LibreOffice-transformed output is bridge evidence, not a native source.",
        ),
    ),
}


def canonical_mathml_contract_for_source_family(source_family: str | SourceFamily) -> CanonicalMathMLContract:
    try:
        family = source_family if isinstance(source_family, SourceFamily) else SourceFamily(str(source_family))
    except ValueError:
        return _contract(
            source_family=str(source_family),
            source_line="unknown",
            contract_status="manual-triage",
            conversion_claim=False,
            binding="unregistered-source-to-canonical-mathml",
            expected_artifacts=("blocker-record.json",),
            required_evidence=("manual source classification", "canonical MathML route decision"),
            notes=("The source family is not registered in the canonical target contract.",),
        )
    return _CONTRACTS.get(
        family,
        _contract(
            source_family=family.value,
            source_line="unknown",
            contract_status="manual-triage",
            conversion_claim=False,
            binding="unregistered-source-to-canonical-mathml",
            expected_artifacts=("blocker-record.json",),
            required_evidence=("manual source classification", "canonical MathML route decision"),
        ),
    )
