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


def _square_root(radicand: bytes) -> bytes:
    return b"\x03\x0d\x00\x00" + b"\x01" + radicand + b"\x00" + b"\x00"


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


def test_supported_mtef3_square_root_template_converts_to_mathml() -> None:
    expression = b"\x01" + _square_root(_char(ord("x"))) + b"\x00" + b"\x00"
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + expression

    result = convert_equation_native_stream_to_mathml(stream)
    root = ET.fromstring(result.mathml_text)

    assert [local_name(node.tag) for node in root.iter()].count("msqrt") == 1
    assert "".join(root.itertext()) == "x"
    assert result.template_selector_counts["13:0:tmROOT"] == 1


def test_matrix_records_stay_blocked_instead_of_guessed() -> None:
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + b"\x05"

    with pytest.raises(Equation3MtefError, match="Matrix records"):
        convert_equation_native_stream_to_mathml(stream)
