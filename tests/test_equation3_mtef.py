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


def _char_with_prime(codepoint: int) -> bytes:
    # CHAR record with xfEMBELL (0x2), followed by an EMBELL list containing embPRIME (5) and END (0).
    return _char(codepoint, options=2) + b"\x06\x05\x00"


def _subscript(slot: bytes) -> bytes:
    return b"\x03\x0f\x01\x00" + b"\x0b" + b"\x01" + slot + b"\x00" + b"\x11" + b"\x00"


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


def _parbox(selector: int, main: bytes, *, variation: int = 0, include_fence_chars: bool = False) -> bytes:
    fence_chars = b""
    if include_fence_chars:
        fence_chars = _char(ord("("), typeface=6, options=0) + _char(ord(")"), typeface=6, options=0)
    return b"\x03" + bytes([selector, variation]) + b"\x00" + b"\x01" + main + b"\x00" + fence_chars + b"\x00"


def _line(objects: bytes) -> bytes:
    return b"\x01" + objects + b"\x00"


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


def test_supported_mtef3_allows_trailing_checksum_word_after_end_record() -> None:
    stream = _supported_equation_native_stream() + b"\xda\xb6"
    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == "bk=ak"
    assert result.mtef_payload_bytes - result.parsed_bytes == 2


def test_supported_mtef3_rejects_unexpected_trailing_bytes_after_end_record() -> None:
    stream = _supported_equation_native_stream() + b"\xda\xb6\x01"

    with pytest.raises(Equation3MtefError, match="Parser stopped with 3 trailing bytes"):
        convert_equation_native_stream_to_mathml(stream)


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


def test_supported_mtef3_parentheses_template_converts_to_mathml_fence_mrow() -> None:
    expression = b"\x01" + _parbox(1, _char(ord("x"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("mrow") >= 2
    assert "".join(root.itertext()) == "(x)"
    assert result.template_selector_counts["1:0:tmPAREN"] == 1


def test_supported_mtef3_one_sided_bracket_template_preserves_side() -> None:
    expression = b"\x01" + _parbox(3, _char(ord("x")), variation=2) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert "".join(root.itertext()) == "x]"
    assert result.template_selector_counts["3:2:tmBRACK_RIGHT"] == 1


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
