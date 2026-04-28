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


def test_omml_scripts_convert_to_mathml_scripts() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:sSubSup>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
        <m:sub><m:r><m:t>1</m:t></m:r></m:sub>
        <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
      </m:sSubSup>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert "<math:msubsup>" in mathml
    assert "<math:mi>x</math:mi>" in mathml
    assert "<math:mn>1</math:mn>" in mathml
    assert "<math:mn>2</math:mn>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_subscript_converts_to_mathml_subscript() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:sSub>
        <m:e><m:r><m:t>a</m:t></m:r></m:e>
        <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
      </m:sSub>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert "<math:msub>" in mathml
    assert "<math:mi>a</math:mi>" in mathml
    assert "<math:mi>i</math:mi>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_radical_converts_to_mathml_sqrt() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:rad>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
      </m:rad>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert "<math:msqrt>" in mathml
    assert "<math:mi>x</math:mi>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_delimiter_converts_to_mathml_fenced() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:d>
        <m:dPr>
          <m:begChr m:val="[" />
          <m:endChr m:val="]" />
        </m:dPr>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
      </m:d>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert '<math:mfenced open="[" close="]">' in mathml
    assert "<math:mi>x</math:mi>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_nary_converts_to_mathml_under_over_operator() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:nary>
        <m:naryPr><m:chr m:val="&#x2211;" /></m:naryPr>
        <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
        <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
      </m:nary>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert "<math:munderover>" in mathml
    assert f"<math:mo>{chr(0x2211)}</math:mo>" in mathml
    assert "<math:mi>i</math:mi>" in mathml
    assert "<math:mi>n</math:mi>" in mathml
    assert "<math:mi>x</math:mi>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_matrix_converts_to_mathml_table() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:m>
        <m:mr>
          <m:e><m:r><m:t>a</m:t></m:r></m:e>
          <m:e><m:r><m:t>b</m:t></m:r></m:e>
        </m:mr>
        <m:mr>
          <m:e><m:r><m:t>c</m:t></m:r></m:e>
          <m:e><m:r><m:t>d</m:t></m:r></m:e>
        </m:mr>
      </m:m>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert "<math:mtable>" in mathml
    assert mathml.count("<math:mtr>") == 2
    assert mathml.count("<math:mtd>") == 4
    assert "<math:mi>a</math:mi>" in mathml
    assert "<math:mi>d</math:mi>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_accent_and_bar_convert_to_mathml_overscripts() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:acc>
        <m:accPr><m:chr m:val="~" /></m:accPr>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
      </m:acc>
      <m:bar>
        <m:barPr><m:pos m:val="top" /></m:barPr>
        <m:e><m:r><m:t>y</m:t></m:r></m:e>
      </m:bar>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert '<math:mover accent="true">' in mathml
    assert "<math:mo>~</math:mo>" in mathml
    assert "<math:mover>" in mathml
    assert f"<math:mo>{chr(0x00AF)}</math:mo>" in mathml
    assert "<math:mi>x</math:mi>" in mathml
    assert "<math:mi>y</math:mi>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_function_converts_to_mathml_function_application() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:func>
        <m:fName><m:r><m:t>sin</m:t></m:r></m:fName>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
      </m:func>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert "<math:mi>sin</math:mi>" in mathml
    assert f"<math:mo>{chr(0x2061)}</math:mo>" in mathml
    assert "<math:mi>x</math:mi>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_limits_convert_to_mathml_under_and_over() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:limLow>
        <m:e><m:r><m:t>lim</m:t></m:r></m:e>
        <m:lim><m:r><m:t>0</m:t></m:r></m:lim>
      </m:limLow>
      <m:limUpp>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
        <m:lim><m:r><m:t>^</m:t></m:r></m:lim>
      </m:limUpp>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert "<math:munder>" in mathml
    assert "<math:mover>" in mathml
    assert "<math:mi>lim</math:mi>" in mathml
    assert "<math:mn>0</math:mn>" in mathml
    assert "<math:mo>^</math:mo>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_group_character_converts_to_mathml_accented_under_or_over() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:groupChr>
        <m:groupChrPr><m:chr m:val="&#x23DE;" /><m:pos m:val="top" /></m:groupChrPr>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
      </m:groupChr>
      <m:groupChr>
        <m:groupChrPr><m:chr m:val="&#x23DF;" /><m:pos m:val="bot" /></m:groupChrPr>
        <m:e><m:r><m:t>y</m:t></m:r></m:e>
      </m:groupChr>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert '<math:mover accent="true">' in mathml
    assert '<math:munder accentunder="true">' in mathml
    assert f"<math:mo>{chr(0x23DE)}</math:mo>" in mathml
    assert f"<math:mo>{chr(0x23DF)}</math:mo>" in mathml
    assert "<math:mi>x</math:mi>" in mathml
    assert "<math:mi>y</math:mi>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_phantom_and_boxes_convert_to_structured_mathml() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:phant>
        <m:phantPr><m:show m:val="0" /></m:phantPr>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
      </m:phant>
      <m:box>
        <m:e><m:r><m:t>y</m:t></m:r></m:e>
      </m:box>
      <m:borderBox>
        <m:e><m:r><m:t>z</m:t></m:r></m:e>
      </m:borderBox>
      <m:borderBox>
        <m:borderBoxPr><m:hideTop m:val="1" /></m:borderBoxPr>
        <m:e><m:r><m:t>w</m:t></m:r></m:e>
      </m:borderBox>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert "<math:mphantom>" in mathml
    assert '<math:mrow data-omml-box="true">' in mathml
    assert '<math:menclose notation="box">' in mathml
    assert '<math:menclose notation="bottom left right">' in mathml
    assert "<math:mi>x</math:mi>" in mathml
    assert "<math:mi>w</math:mi>" in mathml
    assert "data-omml-unsupported" not in mathml


def test_omml_run_style_preserves_mathvariant() -> None:
    payload = """<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
      <m:r><m:rPr><m:sty m:val="b" /></m:rPr><m:t>x</m:t></m:r>
      <m:r><m:rPr><m:sty m:val="bi" /></m:rPr><m:t>y</m:t></m:r>
    </m:oMath>"""

    mathml = omml_fragment_to_mathml(payload)

    assert '<math:mi mathvariant="bold">x</math:mi>' in mathml
    assert '<math:mi mathvariant="bold-italic">y</math:mi>' in mathml
    assert "data-omml-unsupported" not in mathml
