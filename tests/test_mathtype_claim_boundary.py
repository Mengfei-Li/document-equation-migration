import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MATHTYPE_EVIDENCE = ROOT / "docs" / "mathtype-evidence.md"
FIXTURE_ROOT = ROOT / "tests" / "fixtures" / "mathtype_ole"
LIVE_CONTROL_ROOT = FIXTURE_ROOT / "live_control"


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _section(text: str, heading: str) -> str:
    start = text.index(heading)
    next_heading = text.find("\n## ", start + len(heading))
    if next_heading == -1:
        return text[start:]
    return text[start:next_heading]


def test_mathtype_evidence_keeps_release_claims_bounded() -> None:
    text = _read_text(MATHTYPE_EVIDENCE)
    release_conclusion = _section(text, "## Release-Facing Conclusion")
    should_not_claim = _section(text, "## What Users Should Not Claim Today")

    assert "manual-review candidate" in release_conclusion
    assert "not as an automated deliverable claim" in release_conclusion
    assert "does **not** claim lossless conversion" in release_conclusion
    assert "pixel-identical layout" in release_conclusion
    assert "universal MathType coverage" in release_conclusion

    forbidden_examples = [
        '"MathType conversion is lossless"',
        '"The output is pixel-identical to the source document"',
        '"All MathType documents are supported"',
        '"Public fixtures already prove the full live MTEF conversion path"',
        '"A successful run means the output is production-ready without review"',
    ]
    for phrase in forbidden_examples:
        assert phrase in should_not_claim


def test_nonexternal_live_control_boundary_is_explicit() -> None:
    text = _read_text(MATHTYPE_EVIDENCE)
    section = _section(text, "## Nonexternal Public Evidence Boundary")

    required_positive_scope = [
        "exact source payload integrity",
        "OLE CFB readability",
        "presence of the `Equation Native` stream",
        "detector classification as one `mathtype-ole` source",
        "temporary DOCX packaging from public text fixture files",
    ]
    for phrase in required_positive_scope:
        assert phrase in section

    blocked_claims = [
        "canonical MathML production",
        "OMML replacement",
        "Word export",
        "production readiness",
        "lossless conversion",
        "pixel-identical layout",
        "universal MathType support",
        "general live-conversion coverage",
    ]
    for phrase in blocked_claims:
        assert phrase in section

    assert "Full live conversion remains an opt-in external-tool check" in section
    assert "manual-review boundary" in section


def test_live_control_fixture_metadata_rejects_overclaiming() -> None:
    manifest = json.loads((LIVE_CONTROL_ROOT / "SOURCES.json").read_text(encoding="utf-8"))
    live_readme = _read_text(LIVE_CONTROL_ROOT / "README.md")
    fixture_readme = _read_text(FIXTURE_ROOT / "README.md")
    boundary = manifest["claim_boundary"]

    assert boundary["default_ci_runs_external_converter"] is False
    assert boundary["requires_external_tools_for_conversion"] is True
    assert boundary["canonical_mathml_claim"] is False
    assert boundary["production_ready_claim"] is False
    assert boundary["lossless_claim"] is False
    assert boundary["pixel_identity_claim"] is False
    assert boundary["universal_mathtype_support_claim"] is False
    assert boundary["general_live_conversion_claim"] is False

    assert "not a real user document" in live_readme
    assert "production output proof" in live_readme
    assert "lossless conversion proof" in live_readme
    assert "pixel-identical layout proof" in live_readme
    assert "general live-conversion proof" in live_readme
    assert "do not run the external MathType converter by default" in live_readme

    assert "not intended to represent production MathType output or visual parity" in fixture_readme
