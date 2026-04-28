from document_equation_migration.omml_to_mathml import omml_fragment_to_mathml


def test_omml_run_converts_to_mathml_identifier() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:r><m:t>x</m:t></m:r>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert "<math:math" in mathml
    assert "<math:mi>x</math:mi>" in mathml


def test_omml_fraction_converts_to_mathml_fraction() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:f>
        <m:num><m:r><m:t>1</m:t></m:r></m:num>
        <m:den><m:r><m:t>2</m:t></m:r></m:den>
      </m:f>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert "<math:mfrac>" in mathml
    assert "<math:mn>1</math:mn>" in mathml
    assert "<math:mn>2</math:mn>" in mathml
