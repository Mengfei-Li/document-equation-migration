from xml.etree import ElementTree as ET

import pytest

from document_equation_migration.equation3_mtef import (
    EQNOLEFILEHDR_SIZE,
    Equation3MtefError,
    convert_equation_native_stream_to_mathml,
    local_name,
)


def _typeface_byte(typeface: int) -> int:
    return (typeface - 128) % 256


def _char(codepoint: int, *, typeface: int = 3, options: int = 1) -> bytes:
    return bytes([(options << 4) | 2, _typeface_byte(typeface), codepoint & 0xFF, codepoint >> 8])


def _char_v2(codepoint: int, *, typeface: int = 3, options: int = 1) -> bytes:
    return bytes([(options << 4) | 2, _typeface_byte(typeface), codepoint & 0xFF])


def _char_with_prime(codepoint: int) -> bytes:
    # CHAR record with xfEMBELL (0x2), followed by an EMBELL list containing embPRIME (5) and END (0).
    return _char(codepoint, options=2) + b"\x06\x05\x00"


def _char_with_hat(codepoint: int) -> bytes:
    # CHAR record with xfEMBELL (0x2), followed by an EMBELL list containing embHAT (9) and END (0).
    return _char(codepoint, options=2) + b"\x06\x09\x00"


def _char_with_tilde(codepoint: int) -> bytes:
    # CHAR record with xfEMBELL (0x2), followed by an EMBELL list containing embTILDE (8) and END (0).
    return _char(codepoint, options=2) + b"\x06\x08\x00"


def _char_with_overbar(codepoint: int) -> bytes:
    # CHAR record with xfEMBELL (0x2), followed by an EMBELL list containing embOBAR (17) and END (0).
    return _char(codepoint, options=2) + b"\x06\x11\x00"

def _char_with_embellishment(codepoint: int, embell_id: int) -> bytes:
    return _char(codepoint, options=2) + b"\x06" + bytes([embell_id]) + b"\x00"


def _subscript(slot: bytes) -> bytes:
    return b"\x03\x0f\x01\x00" + b"\x0b" + b"\x01" + slot + b"\x00" + b"\x11" + b"\x00"


def _leading_subscript(slot: bytes) -> bytes:
    return b"\x03\x2c\x01\x00" + b"\x0b" + b"\x01" + slot + b"\x00" + b"\x11" + b"\x00"


def _superscript(slot: bytes) -> bytes:
    return b"\x03\x0f\x00\x00" + b"\x11" + b"\x0b" + b"\x01" + slot + b"\x00" + b"\x00"


def _subsup(sub: bytes, sup: bytes) -> bytes:
    return b"\x03\x0f\x02\x00" + b"\x0b" + b"\x01" + sub + b"\x00" + b"\x0b" + b"\x01" + sup + b"\x00" + b"\x00"


def _leading_superscript(slot: bytes) -> bytes:
    return b"\x03\x2c\x00\x00" + b"\x11" + b"\x0b" + b"\x01" + slot + b"\x00" + b"\x00"


def _leading_subsup(sub: bytes, sup: bytes) -> bytes:
    return b"\x03\x2c\x02\x00" + b"\x0b" + b"\x01" + sub + b"\x00" + b"\x0b" + b"\x01" + sup + b"\x00" + b"\x00"


def _fraction(numerator: bytes, denominator: bytes, *, variation: int = 0) -> bytes:
    return (
        b"\x03\x0e"
        + bytes([variation])
        + b"\x00"
        + b"\x01"
        + numerator
        + b"\x00"
        + b"\x01"
        + denominator
        + b"\x00"
        + b"\x00"
    )


def _slash_fraction(numerator: bytes, denominator: bytes, *, variation: int = 0) -> bytes:
    return (
        b"\x03\x29"
        + bytes([variation])
        + b"\x00"
        + b"\x01"
        + numerator
        + b"\x00"
        + b"\x01"
        + denominator
        + b"\x00"
        + b"\x00"
    )


def _square_root(radicand: bytes) -> bytes:
    return b"\x03\x0d\x00\x00" + b"\x01" + radicand + b"\x00" + b"\x00"


def _nth_root(index: bytes, radicand: bytes) -> bytes:
    return b"\x03\x0d\x01\x00" + b"\x01" + index + b"\x00" + b"\x01" + radicand + b"\x00" + b"\x00"


def _underbar(main: bytes, *, variation: int = 0) -> bytes:
    return b"\x03\x10" + bytes([variation]) + b"\x00" + b"\x01" + main + b"\x00" + b"\x00"


def _overbar(main: bytes, *, variation: int = 0) -> bytes:
    return b"\x03\x11" + bytes([variation]) + b"\x00" + b"\x01" + main + b"\x00" + b"\x00"


def _under_arrow(main: bytes, *, variation: int = 0) -> bytes:
    return b"\x03" + bytes([46, variation]) + b"\x00" + b"\x01" + main + b"\x00" + b"\x00"


def _over_arrow(main: bytes, *, variation: int = 0) -> bytes:
    return b"\x03" + bytes([47, variation]) + b"\x00" + b"\x01" + main + b"\x00" + b"\x00"


def _over_arc(main: bytes) -> bytes:
    return b"\x03" + bytes([48, 0]) + b"\x00" + b"\x01" + main + b"\x00" + b"\x00"


def _font_record(*, tface: int = 3, style: int = 0, name: bytes = b"Times") -> bytes:
    return b"\x08" + bytes([tface & 0xFF, style & 0xFF]) + name + b"\x00"


def _size_record(*, lsize: int = 0, dsize: int = 0) -> bytes:
    return b"\x09" + bytes([lsize & 0xFF, (dsize + 128) & 0xFF])


def _ruler_record(*, tab_stops: list[tuple[int, int]] | None = None) -> bytes:
    tab_stops = tab_stops or []
    encoded = [bytes([stop_type & 0xFF, offset & 0xFF, (offset >> 8) & 0xFF]) for stop_type, offset in tab_stops]
    return b"\x07" + bytes([len(tab_stops) & 0xFF]) + b"".join(encoded)


def _parbox_with_formatting_records(selector: int, main: bytes, *, variation: int = 0) -> bytes:
    return (
        b"\x03"
        + bytes([selector, variation])
        + b"\x00"
        + _font_record()
        + _size_record()
        + _ruler_record()
        + _line(main)
        + b"\x00"
    )


def _parbox(selector: int, main: bytes, *, variation: int = 0, include_fence_chars: bool = False) -> bytes:
    fence_chars = b""
    if include_fence_chars:
        fence_chars = _char(ord("("), typeface=6, options=0) + _char(ord(")"), typeface=6, options=0)
    return b"\x03" + bytes([selector, variation]) + b"\x00" + b"\x01" + main + b"\x00" + fence_chars + b"\x00"


def _line(objects: bytes) -> bytes:
    return b"\x01" + objects + b"\x00"


def _pile(line_objects: list[bytes], *, h_just: int = 1, v_just: int = 1) -> bytes:
    return (
        b"\x04"
        + bytes([h_just, v_just])
        + b"".join(_line(objects) for objects in line_objects)
        + b"\x00"
    )


def _matrix(rows: int, cols: int, cell_lines: list[bytes]) -> bytes:
    assert rows * cols == len(cell_lines)
    row_parts = b"\x00" * (((rows + 1) * 2 + 7) // 8)
    col_parts = b"\x00" * (((cols + 1) * 2 + 7) // 8)
    return (
        b"\x05"
        + b"\x00"  # valign
        + b"\x01"  # h_just (left)
        + b"\x00"  # v_just
        + bytes([rows, cols])
        + row_parts
        + col_parts
        + b"".join(cell_lines)
        + b"\x00"
    )


def _null_line() -> bytes:
    return b"\x11"  # LINE record with xfNULL option set; object list omitted.


def _bigop(
    selector: int,
    variation: int,
    *,
    main: bytes,
    operator_codepoint: int,
    upper: bytes | None = None,
    lower: bytes | None = None,
) -> bytes:
    upper_record = _line(upper) if upper is not None else _null_line()
    lower_record = _line(lower) if lower is not None else _null_line()
    operator_record = _char(operator_codepoint, typeface=6, options=0)
    return (
        b"\x03"
        + bytes([selector, variation])
        + b"\x00"
        + _line(main)
        + upper_record
        + lower_record
        + operator_record
        + b"\x00"
    )


def _limit_template(
    variation: int,
    *,
    main: bytes,
    upper: bytes | None = None,
    lower: bytes | None = None,
) -> bytes:
    lower_record = _line(lower) if lower is not None else _null_line()
    upper_record = _line(upper) if upper is not None else _null_line()
    return b"\x03\x27" + bytes([variation]) + b"\x00" + _line(main) + lower_record + upper_record + b"\x00"


def _supported_equation_native_stream() -> bytes:
    expression = (
        b"\x0a"
        + b"\x01"
        + _char(ord("b"))
        + _subscript(_char(ord("k")))
        + b"\x0a"
        + _char(ord("="), typeface=6, options=0)
        + _char(ord("a"))
        + _subscript(_char(ord("k")))
        + b"\x00"
        + b"\x00"
    )
    return bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression


def _supported_mtef2_equation_native_stream_with_template_child_matrix() -> bytes:
    # MTEF v2 header (Mac) plus a ParBox template whose subobject list contains a MATRIX record directly.
    template = (
        b"\x03\x01\x00\x00"
        + _matrix(
            2,
            2,
            [
                _line(_char_v2(ord("a"))),
                _line(_char_v2(ord("b"))),
                _line(_char_v2(ord("c"))),
                _line(_char_v2(ord("d"))),
            ],
        )
        + b"\x00"
    )
    expression = b"\x0a" + _line(template) + b"\x00"
    return bytes(EQNOLEFILEHDR_SIZE) + b"\x02\x00\x01\x02\x01" + expression


def test_supported_mtef3_script_slice_converts_to_mathml() -> None:
    result = convert_equation_native_stream_to_mathml(_supported_equation_native_stream())
    root = ET.fromstring(result.mathml_text)

    assert local_name(root.tag) == "math"
    assert root.attrib["display"] == "block"
    assert [local_name(node.tag) for node in root.iter()].count("msub") == 2
    assert "".join(root.itertext()) == "bk=ak"
    assert result.mtef_version == 3
    assert result.record_counts["3"] == 2
    assert result.template_selector_counts["15:1:tmSUB"] == 2


def test_supported_mtef3_leading_subscript_template_converts_to_mmultiscripts() -> None:
    expression = b"\x0a" + b"\x01" + _char(ord("N")) + _leading_subscript(_char(ord("k"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    mmultiscripts = next(node for node in root.iter() if local_name(node.tag) == "mmultiscripts")

    assert [local_name(node.tag) for node in root.iter()].count("mmultiscripts") == 1
    assert [local_name(node.tag) for node in root.iter()].count("mprescripts") == 1
    assert mmultiscripts.attrib["data-equation3-script-position"] == "leading"
    assert [local_name(child.tag) for child in list(mmultiscripts)][1:4] == ["none", "none", "mprescripts"]
    assert "".join(root.itertext()) == "Nk"
    assert result.template_selector_counts["44:1:tmLSUB"] == 1


def test_supported_mtef3_superscript_template_converts_to_msup() -> None:
    expression = b"\x0a" + b"\x01" + _char(ord("x")) + _superscript(_char(ord("2"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("msup") == 1
    assert "".join(root.itertext()) == "x2"
    assert result.template_selector_counts["15:0:tmSUP"] == 1


def test_supported_mtef3_subsup_template_converts_to_msubsup() -> None:
    expression = b"\x0a" + b"\x01" + _char(ord("x")) + _subsup(_char(ord("i")), _char(ord("2"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("msubsup") == 1
    assert "".join(root.itertext()) == "xi2"
    assert result.template_selector_counts["15:2:tmSUBSUP"] == 1


def test_supported_mtef3_leading_superscript_template_converts_to_mmultiscripts() -> None:
    expression = b"\x0a" + b"\x01" + _char(ord("N")) + _leading_superscript(_char(ord("k"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    mmultiscripts = next(node for node in root.iter() if local_name(node.tag) == "mmultiscripts")

    assert mmultiscripts.attrib["data-equation3-script-position"] == "leading"
    assert "".join(root.itertext()) == "Nk"
    assert result.template_selector_counts["44:0:tmLSUPER"] == 1


def test_supported_mtef3_leading_subsup_template_converts_to_mmultiscripts() -> None:
    expression = (
        b"\x0a"
        + b"\x01"
        + _char(ord("N"))
        + _leading_subsup(_char(ord("i")), _char(ord("2")))
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("mmultiscripts") == 1
    assert "".join(root.itertext()) == "Ni2"
    assert result.template_selector_counts["44:2:tmLSUBSUP"] == 1


def test_supported_mtef2_template_child_matrix_converts_inside_fence() -> None:
    result = convert_equation_native_stream_to_mathml(_supported_mtef2_equation_native_stream_with_template_child_matrix())
    root = ET.fromstring(result.mathml_text)

    assert result.mtef_version == 2
    assert [local_name(node.tag) for node in root.iter()].count("mtable") == 1
    assert "".join(root.itertext()) == "(abcd)"
    assert result.template_selector_counts["1:0:tmPAREN"] == 1


def test_supported_mtef2_space_and_lower_greek_typefaces_convert_to_valid_mathml() -> None:
    expression = (
        b"\x01"
        + _char_v2(ord("p"), typeface=4, options=0)
        + _char_v2(2, typeface=24, options=0)
        + b"\x00\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x02\x00\x01\x02\x01" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == "\u03c0"
    assert result.typeface_counts["4:fnLCGREEK"] == 1
    assert result.typeface_counts["24:fnSPACE"] == 1
    mspace = next(node for node in root.iter() if local_name(node.tag) == "mspace")
    assert mspace.attrib["data-equation3-mtef-typeface"] == "fnSPACE"
    assert mspace.attrib["data-equation3-mtef-char-code"] == "2"


def test_supported_mtef3_allows_trailing_checksum_word_after_end_record() -> None:
    stream = _supported_equation_native_stream() + b"\xda\xb6"
    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == "bk=ak"
    assert result.mtef_payload_bytes - result.parsed_bytes == 2


@pytest.mark.parametrize(
    "footer",
    [
        b"\x7d",
        b"\x00\x7a\x00",
        b"\x74\x4a\x00",
        b"\xff\xff\xff",
        b"\xef\xef\xef",
        b"\x06\x00\x07",
        b"\x04\x02\x01",
        b"\x83\x0f\xa0",
        b"\x65\x77\x20",
        b"\x0a\x01\x03",
        b"\x0a\x1a\x06",
        b"\x00" * 8 + b"\x09\x00\x00\x00",
    ],
)
def test_supported_mtef3_allows_observed_legacy_footers_after_end_record(footer: bytes) -> None:
    stream = _supported_equation_native_stream() + footer
    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == "bk=ak"
    assert result.mtef_payload_bytes - result.parsed_bytes == len(footer)


def test_supported_mtef3_rejects_unexpected_trailing_bytes_after_end_record() -> None:
    stream = _supported_equation_native_stream() + b"\xda\xb6\x01"

    with pytest.raises(Equation3MtefError, match="Parser stopped with 3 trailing bytes"):
        convert_equation_native_stream_to_mathml(stream)


def test_supported_mtef3_parses_valid_continuation_after_first_top_level_end() -> None:
    stream = _supported_equation_native_stream() + b"\x11\x00\x0a" + _char(ord("*")) + _char(ord("1")) + b"\x00"

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == "bk=ak*1"
    assert result.parsed_bytes == result.mtef_payload_bytes


def test_supported_mtef3_fraction_template_converts_to_mathml() -> None:
    expression = b"\x01" + _fraction(_char(ord("a")), _char(ord("b"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("mfrac") == 1
    assert "".join(root.itertext()) == "ab"
    assert result.record_counts["3"] == 1
    assert result.template_selector_counts["14:0:tmFRACT"] == 1


def test_supported_mtef3_small_fraction_preserves_size_signal() -> None:
    expression = b"\x01" + _fraction(_char(ord("a")), _char(ord("b")), variation=1) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    fraction = next(node for node in root.iter() if local_name(node.tag) == "mfrac")

    assert fraction.attrib["data-equation3-fraction-size"] == "small"
    assert result.template_selector_counts["14:1:tmFRACT_SMALL"] == 1


def test_supported_mtef3_allows_nested_template_records_in_template_slot_lists() -> None:
    nested = _parbox(1, _char(ord("a")))
    fraction = b"\x03\x0e\x00\x00" + nested + b"\x01" + _char(ord("b")) + b"\x00" + b"\x00"
    expression = b"\x01" + fraction + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("mfrac") == 1
    assert "".join(root.itertext()) == "(a)b"
    assert result.template_selector_counts["14:0:tmFRACT"] == 1
    assert result.template_selector_counts["1:0:tmPAREN"] == 1


def test_supported_mtef3_allows_formatting_records_in_template_slot_lists() -> None:
    template = _parbox_with_formatting_records(1, _char(ord("x")))
    expression = b"\x01" + template + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == "(x)"
    assert result.template_selector_counts["1:0:tmPAREN"] == 1
    assert result.record_counts["7"] == 1
    assert result.record_counts["8"] == 1
    assert result.record_counts["9"] == 1


def test_supported_mtef3_slash_fraction_template_converts_to_bevelled_mathml() -> None:
    expression = b"\x01" + _slash_fraction(_char(ord("a")), _char(ord("b"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    fraction = next(node for node in root.iter() if local_name(node.tag) == "mfrac")

    assert fraction.attrib["bevelled"] == "true"
    assert "".join(root.itertext()) == "ab"
    assert result.template_selector_counts["41:0:tmSLFRACT"] == 1


def test_supported_mtef3_slash_fraction_variations_preserve_layout_signals() -> None:
    expression = (
        b"\x01"
        + _slash_fraction(_char(ord("a")), _char(ord("b")), variation=1)
        + _slash_fraction(_char(ord("c")), _char(ord("d")), variation=2)
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    fractions = [node for node in root.iter() if local_name(node.tag) == "mfrac"]

    assert fractions[0].attrib["bevelled"] == "true"
    assert fractions[0].attrib["data-equation3-slash-fraction-layout"] == "baseline"
    assert fractions[1].attrib["bevelled"] == "true"
    assert fractions[1].attrib["data-equation3-fraction-size"] == "small"
    assert "".join(root.itertext()) == "abcd"
    assert result.template_selector_counts["41:1:tmSLFRACT_BASELINE"] == 1
    assert result.template_selector_counts["41:2:tmSLFRACT_SMALL"] == 1


def test_supported_mtef3_square_root_template_converts_to_mathml() -> None:
    expression = b"\x01" + _square_root(_char(ord("x"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("msqrt") == 1
    assert "".join(root.itertext()) == "x"
    assert result.template_selector_counts["13:0:tmROOT"] == 1


def test_supported_mtef3_nth_root_template_converts_to_mathml() -> None:
    expression = b"\x01" + _nth_root(_char(ord("3")), _char(ord("x"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    nth_root = next(node for node in root.iter() if local_name(node.tag) == "mroot")
    children = list(nth_root)

    assert [local_name(node.tag) for node in root.iter()].count("mroot") == 1
    assert "".join(children[0].itertext()) == "x"
    assert "".join(children[1].itertext()) == "3"
    assert result.template_selector_counts["13:1:tmNTHROOT"] == 1


def test_supported_mtef3_underbar_template_converts_to_mathml() -> None:
    expression = b"\x01" + _underbar(_char(ord("x"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    underbar = next(node for node in root.iter() if local_name(node.tag) == "munder")

    assert [local_name(node.tag) for node in root.iter()].count("munder") == 1
    assert underbar.attrib["accentunder"] == "true"
    assert "".join(underbar.itertext()) == "x_"
    assert result.template_selector_counts["16:0:tmUBAR"] == 1


def test_supported_mtef3_double_underbar_preserves_count_signal() -> None:
    expression = b"\x01" + _underbar(_char(ord("x")), variation=1) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    underbar = next(node for node in root.iter() if local_name(node.tag) == "munder")

    assert underbar.attrib["data-equation3-bar-count"] == "2"
    assert "".join(underbar.itertext()) == "x__"
    assert result.template_selector_counts["16:1:tmUBAR_DOUBLE"] == 1


def test_supported_mtef3_overbar_template_converts_to_mathml() -> None:
    expression = b"\x01" + _overbar(_char(ord("x"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    overbar = next(node for node in root.iter() if local_name(node.tag) == "mover")

    assert [local_name(node.tag) for node in root.iter()].count("mover") == 1
    assert overbar.attrib["accent"] == "true"
    assert "".join(overbar.itertext()) == "x\u203e"
    assert result.template_selector_counts["17:0:tmOBAR"] == 1


def test_supported_mtef3_double_overbar_preserves_count_signal() -> None:
    expression = b"\x01" + _overbar(_char(ord("x")), variation=1) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    overbar = next(node for node in root.iter() if local_name(node.tag) == "mover")

    assert overbar.attrib["data-equation3-bar-count"] == "2"
    assert "".join(overbar.itertext()) == "x\u203e\u203e"
    assert result.template_selector_counts["17:1:tmOBAR_DOUBLE"] == 1


@pytest.mark.parametrize(
    ("variation", "expected_arrow", "selector_key"),
    [
        (0, "\u2190", "46:0:tmUARROW_LEFT"),
        (1, "\u2192", "46:1:tmUARROW_RIGHT"),
        (2, "\u2194", "46:2:tmUARROW_BOTH"),
    ],
)
def test_supported_mtef3_under_arrow_template_variations(
    variation: int, expected_arrow: str, selector_key: str
) -> None:
    expression = b"\x01" + _under_arrow(_char(ord("x")), variation=variation) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    under_arrow = next(node for node in root.iter() if local_name(node.tag) == "munder")

    assert under_arrow.attrib["accentunder"] == "true"
    assert "".join(under_arrow.itertext()) == f"x{expected_arrow}"
    assert result.template_selector_counts[selector_key] == 1


@pytest.mark.parametrize(
    ("variation", "expected_arrow", "selector_key"),
    [
        (0, "\u2190", "47:0:tmOARROW_LEFT"),
        (1, "\u2192", "47:1:tmOARROW_RIGHT"),
        (2, "\u2194", "47:2:tmOARROW_BOTH"),
    ],
)
def test_supported_mtef3_over_arrow_template_variations(
    variation: int, expected_arrow: str, selector_key: str
) -> None:
    expression = b"\x01" + _over_arrow(_char(ord("x")), variation=variation) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    over_arrow = next(node for node in root.iter() if local_name(node.tag) == "mover")

    assert over_arrow.attrib["accent"] == "true"
    assert "".join(over_arrow.itertext()) == f"x{expected_arrow}"
    assert result.template_selector_counts[selector_key] == 1


def test_supported_mtef3_over_arc_template_converts_to_mathml() -> None:
    expression = b"\x01" + _over_arc(_char(ord("x"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    over_arc = next(node for node in root.iter() if local_name(node.tag) == "mover")

    assert over_arc.attrib["accent"] == "true"
    assert "".join(over_arc.itertext()) == "x\u2312"
    assert result.template_selector_counts["48:0:tmOARC"] == 1


def test_supported_mtef3_parentheses_template_converts_to_mathml_fence_mrow() -> None:
    expression = b"\x01" + _parbox(1, _char(ord("x"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("mrow") >= 2
    assert "".join(root.itertext()) == "(x)"
    assert result.template_selector_counts["1:0:tmPAREN"] == 1


@pytest.mark.parametrize(
    ("selector", "variation", "expected_text", "selector_key"),
    [
        (0, 0, "\u27e8x\u27e9", "0:0:tmANGLE"),
        (0, 1, "\u27e8x", "0:1:tmANGLE_LEFT"),
        (0, 2, "x\u27e9", "0:2:tmANGLE_RIGHT"),
        (1, 1, "(x", "1:1:tmPAREN_LEFT"),
        (1, 2, "x)", "1:2:tmPAREN_RIGHT"),
        (2, 2, "x}", "2:2:tmBRACE_RIGHT"),
        (3, 1, "[x", "3:1:tmBRACK_LEFT"),
        (4, 0, "|x|", "4:0:tmBAR"),
        (4, 1, "|x", "4:1:tmBAR_LEFT"),
        (4, 2, "x|", "4:2:tmBAR_RIGHT"),
        (5, 0, "\u2016x\u2016", "5:0:tmDBAR"),
        (5, 1, "\u2016x", "5:1:tmDBAR_LEFT"),
        (5, 2, "x\u2016", "5:2:tmDBAR_RIGHT"),
    ],
)
def test_supported_mtef3_parbox_template_variations_render_expected_fences(
    selector: int,
    variation: int,
    expected_text: str,
    selector_key: str,
) -> None:
    expression = b"\x01" + _parbox(selector, _char(ord("x")), variation=variation) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == expected_text
    assert result.template_selector_counts[selector_key] == 1


def test_supported_mtef3_one_sided_bracket_template_preserves_side() -> None:
    expression = b"\x01" + _parbox(3, _char(ord("x")), variation=2) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == "x]"
    assert result.template_selector_counts["3:2:tmBRACK_RIGHT"] == 1


def test_supported_mtef3_one_sided_floor_template_preserves_side() -> None:
    left = b"\x01" + _parbox(6, _char(ord("x")), variation=1) + b"\x00" + b"\x00"
    right = b"\x01" + _parbox(6, _char(ord("x")), variation=2) + b"\x00" + b"\x00"

    left_stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + left
    right_stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + right

    left_result = convert_equation_native_stream_to_mathml(left_stream)
    left_root = ET.fromstring(left_result.mathml_text)
    assert "".join(left_root.itertext()) == "\u230ax"
    assert left_result.template_selector_counts["6:1:tmFLOOR_LEFT"] == 1

    right_result = convert_equation_native_stream_to_mathml(right_stream)
    right_root = ET.fromstring(right_result.mathml_text)
    assert "".join(right_root.itertext()) == "x\u230b"
    assert right_result.template_selector_counts["6:2:tmFLOOR_RIGHT"] == 1


def test_supported_mtef3_one_sided_ceiling_template_preserves_side() -> None:
    left = b"\x01" + _parbox(7, _char(ord("x")), variation=1) + b"\x00" + b"\x00"
    right = b"\x01" + _parbox(7, _char(ord("x")), variation=2) + b"\x00" + b"\x00"

    left_stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + left
    right_stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + right

    left_result = convert_equation_native_stream_to_mathml(left_stream)
    left_root = ET.fromstring(left_result.mathml_text)
    assert "".join(left_root.itertext()) == "\u2308x"
    assert left_result.template_selector_counts["7:1:tmCEILING_LEFT"] == 1

    right_result = convert_equation_native_stream_to_mathml(right_stream)
    right_root = ET.fromstring(right_result.mathml_text)
    assert "".join(right_root.itertext()) == "x\u2309"
    assert right_result.template_selector_counts["7:2:tmCEILING_RIGHT"] == 1


def test_supported_mtef3_parbox_accepts_explicit_fence_character_subobjects() -> None:
    expression = b"\x01" + _parbox(1, _char(ord("x")), include_fence_chars=True) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == "(x)"
    assert result.record_counts["2"] == 3
    assert result.template_selector_counts["1:0:tmPAREN"] == 1


def test_supported_mtef3_char_embellishment_prime_converts_to_mathml() -> None:
    expression = b"\x01" + _char_with_prime(ord("x")) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("msup") == 1
    assert "".join(root.itertext()) == "x\u2032"
    assert result.record_counts["6"] == 1


def test_supported_mtef3_char_embellishment_hat_converts_to_mathml() -> None:
    expression = b"\x01" + _char_with_hat(ord("x")) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("mover") == 1
    assert "".join(root.itertext()) == "x^"
    assert result.record_counts["6"] == 1


def test_supported_mtef3_char_embellishment_tilde_converts_to_mathml() -> None:
    expression = b"\x01" + _char_with_tilde(ord("x")) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("mover") == 1
    assert "".join(root.itertext()) == "x~"
    assert result.record_counts["6"] == 1


def test_supported_mtef3_char_embellishment_overbar_converts_to_mathml() -> None:
    expression = b"\x01" + _char_with_overbar(ord("x")) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("mover") == 1
    assert "".join(root.itertext()) == "x\u203e"
    assert result.record_counts["6"] == 1


@pytest.mark.parametrize(
    ("embell_id", "expected_accent", "stretchy"),
    [
        (2, "\u02d9", False),  # emb1DOT
        (3, "\u00a8", False),  # emb2DOT
        (11, "\u2192", True),  # embRARROW
        (12, "\u2190", True),  # embLARROW
        (13, "\u2194", True),  # embBARROW
        (14, "\u21c0", True),  # embR1ARROW
        (15, "\u21bc", True),  # embL1ARROW
        (19, "\u2322", True),  # embFROWN
        (20, "\u2323", True),  # embSMILE
    ],
)
def test_supported_mtef3_char_embellishment_accent_variants_convert_to_mathml(
    embell_id: int, expected_accent: str, stretchy: bool
) -> None:
    expression = b"\x01" + _char_with_embellishment(ord("x"), embell_id) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("mover") == 1
    mover = next(node for node in root.iter() if local_name(node.tag) == "mover")
    assert mover.attrib["accent"] == "true"

    accent = list(mover)[1]
    assert local_name(accent.tag) == "mo"
    assert accent.text == expected_accent
    if stretchy:
        assert accent.attrib["stretchy"] == "true"
    else:
        assert "stretchy" not in accent.attrib

    assert "".join(root.itertext()) == f"x{expected_accent}"
    assert result.record_counts["6"] == 1


def test_supported_mtef3_matrix_record_converts_to_mathml_table() -> None:
    expression = (
        _line(
            _matrix(
                2,
                2,
                [
                    _line(_char(ord("a"))),
                    _line(_char(ord("b"))),
                    _line(_char(ord("c"))),
                    _line(_char(ord("d"))),
                ],
            )
        )
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("mtable") == 1
    matrix = next(node for node in root.iter() if local_name(node.tag) == "mtable")
    assert matrix.attrib["data-equation3-matrix-rows"] == "2"
    assert matrix.attrib["data-equation3-matrix-cols"] == "2"
    assert "".join(root.itertext()) == "abcd"
    assert result.record_counts["5"] == 1


def test_supported_mtef3_pile_record_converts_to_mathml_table() -> None:
    expression = b"\x01" + _pile([_char(ord("a")), _char(ord("b"))]) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    pile = next(node for node in root.iter() if local_name(node.tag) == "mtable")

    assert pile.attrib["data-equation3-pile-rows"] == "2"
    assert pile.attrib["data-equation3-pile-hjust"] == "1"
    assert [local_name(node.tag) for node in pile.iter()].count("mtr") == 2
    assert "".join(root.itertext()) == "ab"
    assert result.record_counts["4"] == 1


def test_supported_mtef3_pile_record_accepts_direct_continuation_row() -> None:
    expression = b"\x01" + b"\x04\x01\x01" + _line(_char(ord("a"))) + _char(ord("b")) + b"\x00\x00\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    pile = next(node for node in root.iter() if local_name(node.tag) == "mtable")

    assert pile.attrib["data-equation3-pile-rows"] == "2"
    assert [local_name(node.tag) for node in pile.iter()].count("mtr") == 2
    assert "".join(root.itertext()) == "ab"
    assert result.record_counts["4"] == 1


def test_supported_mtef3_template_child_pile_converts_inside_fence() -> None:
    brace_left_with_pile = (
        b"\x03\x02\x01\x00" + _pile([_char(ord("x")), _char(ord("y"))]) + b"\x00"
    )
    expression = b"\x01" + brace_left_with_pile + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    pile = next(node for node in root.iter() if local_name(node.tag) == "mtable")

    assert pile.attrib["data-equation3-pile-rows"] == "2"
    assert "".join(root.itertext()) == "{xy"
    assert result.template_selector_counts["2:1:tmBRACE_LEFT"] == 1
    assert result.record_counts["4"] == 1


def test_malformed_matrix_records_still_block_instead_of_guessing() -> None:
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + b"\x01" + b"\x05\x00\x00\x00\x00\x02"

    with pytest.raises(Equation3MtefError, match="Matrix records with rows=0 cols=2"):
        convert_equation_native_stream_to_mathml(stream)


def test_supported_mtef3_sum_template_with_limits_converts_to_munderover() -> None:
    expression = (
        b"\x01"
        + _bigop(
            29,
            1,
            main=_char(ord("a")),
            upper=_char(ord("n")),
            lower=_char(ord("i")) + _char(ord("="), typeface=6, options=0) + _char(ord("1")),
            operator_codepoint=0x2211,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("munderover") == 1
    assert "".join(root.itertext()) == "\u2211i=1na"
    assert result.template_selector_counts["29:1:tmSUM_BOTH"] == 1


def test_supported_mtef3_observed_sum_both_without_limit_slots_keeps_largeop_evidence() -> None:
    expression = (
        b"\x01"
        + b"\x03\x1d\x01\x00"
        + _line(_char(ord("a")))
        + _char(0xEC07, typeface=22, options=0)
        + b"\x00"
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    operator = next(node for node in root.iter() if local_name(node.tag) == "mo" and node.text == "\u2211")

    assert operator.attrib["largeop"] == "true"
    assert operator.attrib["movablelimits"] == "true"
    assert operator.attrib["data-equation3-missing-limit-slots"] == "both"
    assert operator.attrib["data-equation3-operator-source-codepoint"] == "U+EC07"
    assert [local_name(node.tag) for node in root.iter()].count("munderover") == 0
    assert "".join(root.itertext()) == "\u2211a"
    assert result.template_selector_counts["29:1:tmSUM_BOTH"] == 1


def test_supported_mtef3_observed_sum_both_can_split_embedded_expand_operator() -> None:
    expression = (
        b"\x01"
        + b"\x03\x1d\x01\x00"
        + _line(_char(ord("a")) + _char(0xEC08, typeface=22, options=0))
        + b"\x00"
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    operator = next(node for node in root.iter() if local_name(node.tag) == "mo" and node.text == "\u2211")

    assert operator.attrib["data-equation3-operator-slot-shape"] == "embedded-in-main-line"
    assert operator.attrib["data-equation3-operator-source-codepoint"] == "U+EC08"
    assert operator.attrib["data-equation3-missing-limit-slots"] == "both"
    assert "".join(root.itertext()) == "\u2211a"
    assert result.template_selector_counts["29:1:tmSUM_BOTH"] == 1


def test_supported_mtef3_single_integral_lower_limit_converts_to_munder() -> None:
    expression = (
        b"\x01"
        + _bigop(
            21,
            1,
            main=_char(ord("x")),
            lower=_char(ord("0")),
            operator_codepoint=0x222B,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("munder") == 1
    assert "".join(root.itertext()) == "\u222b0x"
    assert result.template_selector_counts["21:1:tmSINT_LOWER"] == 1


@pytest.mark.parametrize(
    ("selector", "variation", "upper", "lower", "main", "operator_codepoint", "expected_text", "selector_key"),
    [
        (29, 0, None, _char(ord("i")), _char(ord("a")), 0x2211, "\u2211ia", "29:0:tmSUM_LOWER"),
        (29, 2, None, None, _char(ord("a")), 0x2211, "\u2211a", "29:2:tmSUM_NO_LIMITS"),
        (21, 0, None, None, _char(ord("x")), 0x222B, "\u222bx", "21:0:tmSINT_NO_LIMITS"),
        (21, 2, _char(ord("1")), _char(ord("0")), _char(ord("x")), 0x222B, "\u222b01x", "21:2:tmSINT_BOTH"),
    ],
)
def test_supported_mtef3_bigop_template_variations_without_existing_coverage(
    selector: int,
    variation: int,
    upper: bytes | None,
    lower: bytes | None,
    main: bytes,
    operator_codepoint: int,
    expected_text: str,
    selector_key: str,
) -> None:
    expression = (
        b"\x01"
        + _bigop(
            selector,
            variation,
            main=main,
            upper=upper,
            lower=lower,
            operator_codepoint=operator_codepoint,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == expected_text
    assert result.template_selector_counts[selector_key] == 1


def test_supported_mtef3_product_template_with_limits_converts_to_munderover() -> None:
    expression = (
        b"\x01"
        + _bigop(
            31,
            1,
            main=_char(ord("a")),
            upper=_char(ord("n")),
            lower=_char(ord("i")) + _char(ord("="), typeface=6, options=0) + _char(ord("1")),
            operator_codepoint=0x220F,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("munderover") == 1
    assert "".join(root.itertext()) == "\u220fi=1na"
    assert result.template_selector_counts["31:1:tmPROD_BOTH"] == 1


def test_supported_mtef3_product_template_with_lower_limit_converts_to_munder() -> None:
    expression = (
        b"\x01"
        + _bigop(
            31,
            0,
            main=_char(ord("a")),
            lower=_char(ord("i")),
            operator_codepoint=0x220F,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("munder") == 1
    assert "".join(root.itertext()) == "\u220fia"
    assert result.template_selector_counts["31:0:tmPROD_LOWER"] == 1


def test_supported_mtef3_product_template_without_limits_converts_to_large_operator() -> None:
    expression = (
        b"\x01"
        + _bigop(
            31,
            2,
            main=_char(ord("a")),
            operator_codepoint=0x220F,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("munder") == 0
    assert [local_name(node.tag) for node in root.iter()].count("munderover") == 0
    assert "".join(root.itertext()) == "\u220fa"
    assert result.template_selector_counts["31:2:tmPROD_NO_LIMITS"] == 1


def test_supported_mtef3_integral_style_sum_template_uses_side_limits() -> None:
    expression = (
        b"\x01"
        + _bigop(
            30,
            1,
            main=_char(ord("a")),
            upper=_char(ord("n")),
            lower=_char(ord("i")) + _char(ord("="), typeface=6, options=0) + _char(ord("1")),
            operator_codepoint=0x2211,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    side_limits = next(node for node in root.iter() if local_name(node.tag) == "msubsup")

    assert side_limits.attrib["data-equation3-limit-style"] == "integral"
    assert "".join(root.itertext()) == "\u2211i=1na"
    assert result.template_selector_counts["30:1:tmISUM_BOTH"] == 1


def test_supported_mtef3_integral_style_product_lower_limit_uses_side_limit() -> None:
    expression = (
        b"\x01"
        + _bigop(
            32,
            0,
            main=_char(ord("a")),
            lower=_char(ord("i")),
            operator_codepoint=0x220F,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    side_limit = next(node for node in root.iter() if local_name(node.tag) == "msub")

    assert side_limit.attrib["data-equation3-limit-style"] == "integral"
    assert "".join(root.itertext()) == "\u220fia"
    assert result.template_selector_counts["32:0:tmIPROD_LOWER"] == 1


def test_supported_mtef3_coproduct_template_without_limits_converts_to_large_operator() -> None:
    expression = (
        b"\x01"
        + _bigop(
            33,
            2,
            main=_char(ord("A")),
            operator_codepoint=0x2210,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    operator = next(node for node in root.iter() if local_name(node.tag) == "mo" and node.text == "\u2210")

    assert operator.attrib["largeop"] == "true"
    assert operator.attrib["movablelimits"] == "true"
    assert "".join(root.itertext()) == "\u2210A"
    assert result.template_selector_counts["33:2:tmCOPROD_NO_LIMITS"] == 1


def test_supported_mtef3_limit_template_with_lower_limit_converts_to_munder() -> None:
    expression = (
        b"\x01"
        + _limit_template(
            1,
            main=_char(ord("l")) + _char(ord("i")) + _char(ord("m")),
            lower=_char(ord("n")) + _char(ord("\u2192"), typeface=6, options=0) + _char(ord("0")),
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    limit = next(node for node in root.iter() if local_name(node.tag) == "munder")

    assert "".join(limit[0].itertext()) == "lim"
    assert "".join(limit[1].itertext()) == "n\u21920"
    assert result.template_selector_counts["39:1:tmLIM_LOWER"] == 1


def test_supported_mtef3_limit_template_with_upper_limit_converts_to_mover() -> None:
    expression = (
        b"\x01"
        + _limit_template(
            0,
            main=_char(ord("m")) + _char(ord("a")) + _char(ord("x")),
            upper=_char(ord("n")),
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    limit = next(node for node in root.iter() if local_name(node.tag) == "mover")

    assert "".join(limit[0].itertext()) == "max"
    assert "".join(limit[1].itertext()) == "n"
    assert result.template_selector_counts["39:0:tmLIM_UPPER"] == 1


def test_supported_mtef3_limit_template_with_both_limits_converts_to_munderover() -> None:
    expression = (
        b"\x01"
        + _limit_template(
            2,
            main=_char(ord("l")) + _char(ord("i")) + _char(ord("m")),
            upper=_char(ord("N")),
            lower=_char(ord("n")),
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    limit = next(node for node in root.iter() if local_name(node.tag) == "munderover")

    assert "".join(limit[0].itertext()) == "lim"
    assert "".join(limit[1].itertext()) == "n"
    assert "".join(limit[2].itertext()) == "N"
    assert result.template_selector_counts["39:2:tmLIM_BOTH"] == 1


def test_supported_mtef3_integral_operator_template_with_limits_uses_side_limits() -> None:
    expression = (
        b"\x01"
        + _bigop(
            42,
            2,
            main=_char(ord("f")),
            upper=_char(ord("1")),
            lower=_char(ord("0")),
            operator_codepoint=0x222B,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    side_limits = next(node for node in root.iter() if local_name(node.tag) == "msubsup")

    assert side_limits.attrib["data-equation3-limit-style"] == "integral"
    assert "".join(root.itertext()) == "\u222b01f"
    assert result.template_selector_counts["42:2:tmINTOP_BOTH"] == 1


def test_supported_mtef3_sum_operator_template_converts_observed_standalone_sumop() -> None:
    expression = b"\x0a\x01\x03\x2b\x00\x00\x11\x0b\x11\x01\x00\x0d\x01\x00\x00\x00\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    operator = next(node for node in root.iter() if local_name(node.tag) == "mo" and node.text == "\u2211")

    assert operator.attrib["largeop"] == "true"
    assert operator.attrib["movablelimits"] == "true"
    assert "".join(root.itertext()) == "\u2211"
    assert result.template_selector_counts["43:0:tmSUMOP"] == 1


@pytest.mark.parametrize(
    ("selector", "variation", "operator_codepoint", "upper", "lower", "expected_tag", "expected_text", "selector_key"),
    [
        (30, 0, 0x2211, None, _char(ord("i")), "msub", "\u2211ia", "30:0:tmISUM_LOWER"),
        (30, 1, 0x2211, _char(ord("n")), _char(ord("i")), "msubsup", "\u2211ina", "30:1:tmISUM_BOTH"),
        (32, 0, 0x220F, None, _char(ord("i")), "msub", "\u220fia", "32:0:tmIPROD_LOWER"),
        (32, 1, 0x220F, _char(ord("n")), _char(ord("i")), "msubsup", "\u220fina", "32:1:tmIPROD_BOTH"),
        (42, 0, 0x222B, _char(ord("1")), None, "msup", "\u222b1f", "42:0:tmINTOP_UPPER"),
        (42, 1, 0x222B, None, _char(ord("0")), "msub", "\u222b0f", "42:1:tmINTOP_LOWER"),
        (42, 2, 0x222B, _char(ord("1")), _char(ord("0")), "msubsup", "\u222b01f", "42:2:tmINTOP_BOTH"),
    ],
)
def test_supported_mtef3_integral_style_bigop_variations_use_side_limits(
    selector: int,
    variation: int,
    operator_codepoint: int,
    upper: bytes | None,
    lower: bytes | None,
    expected_tag: str,
    expected_text: str,
    selector_key: str,
) -> None:
    expression = (
        b"\x01"
        + _bigop(
            selector,
            variation,
            main=_char(ord("a")) if selector != 42 else _char(ord("f")),
            upper=upper,
            lower=lower,
            operator_codepoint=operator_codepoint,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    side_limit = next(node for node in root.iter() if local_name(node.tag) == expected_tag)

    assert side_limit.attrib["data-equation3-limit-style"] == "integral"
    assert "".join(root.itertext()) == expected_text
    assert result.template_selector_counts[selector_key] == 1


@pytest.mark.parametrize(
    ("variation", "upper", "lower", "expected_tag", "expected_text", "selector_key"),
    [
        (0, None, _char(ord("i")), "munder", "\u2210iA", "33:0:tmCOPROD_LOWER"),
        (1, _char(ord("n")), _char(ord("i")), "munderover", "\u2210inA", "33:1:tmCOPROD_BOTH"),
        (2, None, None, "mo", "\u2210A", "33:2:tmCOPROD_NO_LIMITS"),
    ],
)
def test_supported_mtef3_coproduct_template_variations(
    variation: int,
    upper: bytes | None,
    lower: bytes | None,
    expected_tag: str,
    expected_text: str,
    selector_key: str,
) -> None:
    expression = (
        b"\x01"
        + _bigop(
            33,
            variation,
            main=_char(ord("A")),
            upper=upper,
            lower=lower,
            operator_codepoint=0x2210,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)
    assert any(local_name(node.tag) == expected_tag for node in root.iter())
    assert "".join(root.itertext()) == expected_text
    assert result.template_selector_counts[selector_key] == 1


@pytest.mark.parametrize(
    ("variation", "upper", "lower", "expected_tag", "expected_text", "selector_key"),
    [
        (0, _char(ord("1")), None, "mover", "lim1", "39:0:tmLIM_UPPER"),
        (1, None, _char(ord("0")), "munder", "lim0", "39:1:tmLIM_LOWER"),
        (2, _char(ord("1")), _char(ord("0")), "munderover", "lim01", "39:2:tmLIM_BOTH"),
    ],
)
def test_supported_mtef3_limit_template_variations(
    variation: int,
    upper: bytes | None,
    lower: bytes | None,
    expected_tag: str,
    expected_text: str,
    selector_key: str,
) -> None:
    expression = (
        b"\x01"
        + _limit_template(
            variation,
            main=_char(ord("l")) + _char(ord("i")) + _char(ord("m")),
            upper=upper,
            lower=lower,
        )
        + b"\x00"
        + b"\x00"
    )
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert any(local_name(node.tag) == expected_tag for node in root.iter())
    assert "".join(root.itertext()) == expected_text
    assert result.template_selector_counts[selector_key] == 1
