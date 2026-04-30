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


def test_matrix_records_stay_blocked_instead_of_guessed() -> None:
    stream = bytes(EQNOLEFILEHDR_SIZE) + b"\x03\x01\x01\x03\x00" + b"\x05"

    with pytest.raises(Equation3MtefError, match="Matrix records"):
        convert_equation_native_stream_to_mathml(stream)
