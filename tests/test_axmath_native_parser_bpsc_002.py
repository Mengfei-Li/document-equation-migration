import pytest

from document_equation_migration.axmath_native_parser import (
    AxMathBoundedStructuralParseError,
    AxMathBoundedStructuralParseResult,
    scan_record_grid_13_plus_20n,
)


FALSE_CLAIMS = {
    "native_parser_claim": False,
    "conversion_claim": False,
    "export_success_claim": False,
    "public_fixture_eligibility": False,
    "score_10_claim": False,
    "universal_axmath_support": False,
}


def _synthetic_contents(total_length: int) -> bytes:
    return bytes(total_length)


def _assert_false_claims(value: dict[str, object]) -> None:
    for claim_name, expected_value in FALSE_CLAIMS.items():
        assert value[claim_name] is expected_value


def test_scan_record_grid_accepts_573_length_family() -> None:
    result = scan_record_grid_13_plus_20n(_synthetic_contents(573))

    assert isinstance(result, AxMathBoundedStructuralParseResult)
    assert result.total_length == 573
    assert result.preface_length == 13
    assert result.record_cell_size == 20
    assert result.record_cell_count == 28
    assert result.preface_span == (0, 13)
    assert result.first_record_span == (13, 33)
    assert result.last_record_span == (553, 573)
    assert result.record_span(27) == (553, 573)
    assert result.semantic_record_cell_supported_count == 0
    assert result.semantic_record_cell_unknown_count == 28

    as_dict = result.to_dict()
    assert as_dict["status"] == "bounded_structural_grid_scanned"
    assert as_dict["conversion_attempted"] is False
    _assert_false_claims(as_dict)


def test_scan_record_grid_accepts_5533_length_family() -> None:
    result = scan_record_grid_13_plus_20n(_synthetic_contents(5533))

    assert result.record_cell_count == 276
    assert result.last_record_span == (5513, 5533)
    assert result.semantic_record_cell_supported_count == 0
    assert result.semantic_record_cell_unknown_count == 276
    assert result.semantic_record_family_supported_count == 0
    assert result.semantic_record_family_unknown_count is None
    _assert_false_claims(result.to_dict())


@pytest.mark.parametrize(
    ("total_length", "record_count", "last_record_span"),
    [
        (593, 29, (573, 593)),
        (713, 35, (693, 713)),
        (833, 41, (813, 833)),
        (1133, 56, (1113, 1133)),
        (1313, 65, (1293, 1313)),
        (1853, 92, (1833, 1853)),
        (4333, 216, (4313, 4333)),
    ],
)
def test_scan_record_grid_accepts_all_manifest_length_families(
    total_length: int,
    record_count: int,
    last_record_span: tuple[int, int],
) -> None:
    result = scan_record_grid_13_plus_20n(_synthetic_contents(total_length))

    assert result.record_cell_count == record_count
    assert result.last_record_span == last_record_span
    assert result.remainder_after_preface == 0
    assert result.unaccounted_grid_length == 0
    _assert_false_claims(result.to_dict())


def test_scan_record_grid_accepts_bytes_like_without_copying_to_output() -> None:
    source = bytearray(_synthetic_contents(573))
    result = scan_record_grid_13_plus_20n(memoryview(source))
    as_dict = result.to_dict()

    assert result.record_cell_count == 28
    assert "record_cells" not in as_dict
    assert "contents" not in as_dict
    assert "hex" not in as_dict
    assert "base64" not in as_dict
    _assert_false_claims(as_dict)


def test_scan_record_grid_rejects_length_before_preface() -> None:
    with pytest.raises(AxMathBoundedStructuralParseError) as exc_info:
        scan_record_grid_13_plus_20n(_synthetic_contents(12))

    assert exc_info.value.error_code == "malformed_length_before_preface"
    assert exc_info.value.total_length == 12
    assert exc_info.value.remainder_after_preface is None
    as_dict = exc_info.value.to_dict()
    assert as_dict["conversion_attempted"] is False
    _assert_false_claims(as_dict)


def test_scan_record_grid_rejects_non_integral_grid_length() -> None:
    with pytest.raises(AxMathBoundedStructuralParseError) as exc_info:
        scan_record_grid_13_plus_20n(_synthetic_contents(574))

    assert exc_info.value.error_code == "malformed_length_non_integral_record_grid"
    assert exc_info.value.total_length == 574
    assert exc_info.value.remainder_after_preface == 1
    _assert_false_claims(exc_info.value.to_dict())


def test_scan_record_grid_rejects_non_bytes_like_input() -> None:
    with pytest.raises(AxMathBoundedStructuralParseError) as exc_info:
        scan_record_grid_13_plus_20n("not-bytes-like")  # type: ignore[arg-type]

    assert exc_info.value.error_code == "input_not_bytes_like"
    assert exc_info.value.total_length is None
    _assert_false_claims(exc_info.value.to_dict())
