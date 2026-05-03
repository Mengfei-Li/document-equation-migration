from __future__ import annotations

import hashlib
import io
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    import olefile
except ImportError:  # pragma: no cover - package dependency is declared, this is defensive.
    olefile = None


EQNOLEFILEHDR_SIZE = 28
MATHML_NS = "http://www.w3.org/1998/Math/MathML"
CONTROL_STREAMS = {"\x01CompObj", "\x01Ole", "\x03ObjInfo"}
TEMPLATE_SELECTOR = {
    (0, 0): "tmANGLE",
    (0, 1): "tmANGLE_LEFT",
    (0, 2): "tmANGLE_RIGHT",
    (1, 0): "tmPAREN",
    (1, 1): "tmPAREN_LEFT",
    (1, 2): "tmPAREN_RIGHT",
    (2, 0): "tmBRACE",
    (2, 1): "tmBRACE_LEFT",
    (2, 2): "tmBRACE_RIGHT",
    (3, 0): "tmBRACK",
    (3, 1): "tmBRACK_LEFT",
    (3, 2): "tmBRACK_RIGHT",
    (4, 0): "tmBAR",
    (4, 1): "tmBAR_LEFT",
    (4, 2): "tmBAR_RIGHT",
    (5, 0): "tmDBAR",
    (5, 1): "tmDBAR_LEFT",
    (5, 2): "tmDBAR_RIGHT",
    (6, 0): "tmFLOOR",
    (6, 1): "tmFLOOR_LEFT",
    (6, 2): "tmFLOOR_RIGHT",
    (7, 0): "tmCEILING",
    (7, 1): "tmCEILING_LEFT",
    (7, 2): "tmCEILING_RIGHT",
    (13, 0): "tmROOT",
    (13, 1): "tmNTHROOT",
    (14, 0): "tmFRACT",
    (14, 1): "tmFRACT_SMALL",
    (15, 0): "tmSUP",
    (15, 1): "tmSUB",
    (15, 2): "tmSUBSUP",
    (16, 0): "tmUBAR",
    (16, 1): "tmUBAR_DOUBLE",
    (17, 0): "tmOBAR",
    (17, 1): "tmOBAR_DOUBLE",
    (21, 0): "tmSINT_NO_LIMITS",
    (21, 1): "tmSINT_LOWER",
    (21, 2): "tmSINT_BOTH",
    (29, 0): "tmSUM_LOWER",
    (29, 1): "tmSUM_BOTH",
    (29, 2): "tmSUM_NO_LIMITS",
    (30, 0): "tmISUM_LOWER",
    (30, 1): "tmISUM_BOTH",
    (31, 0): "tmPROD_LOWER",
    (31, 1): "tmPROD_BOTH",
    (31, 2): "tmPROD_NO_LIMITS",
    (32, 0): "tmIPROD_LOWER",
    (32, 1): "tmIPROD_BOTH",
    (33, 0): "tmCOPROD_LOWER",
    (33, 1): "tmCOPROD_BOTH",
    (33, 2): "tmCOPROD_NO_LIMITS",
    (39, 0): "tmLIM_UPPER",
    (39, 1): "tmLIM_LOWER",
    (39, 2): "tmLIM_BOTH",
    (41, 0): "tmSLFRACT",
    (41, 1): "tmSLFRACT_BASELINE",
    (41, 2): "tmSLFRACT_SMALL",
    (42, 0): "tmINTOP_UPPER",
    (42, 1): "tmINTOP_LOWER",
    (42, 2): "tmINTOP_BOTH",
    (43, 0): "tmSUMOP",
    (44, 0): "tmLSUPER",
    (44, 1): "tmLSUB",
    (44, 2): "tmLSUBSUP",
}
BASE_CONSUMING_TEMPLATES = {"tmSUP", "tmSUB", "tmSUBSUP", "tmLSUPER", "tmLSUB", "tmLSUBSUP"}
OPERATOR_CHARS = set("=+-*/(),[]{}") | {"\u2192", "\u2211", "\u222b", "\u220f", "\u2210"}  # →, ∑, ∫, ∏, ∐
BIGOP_TEMPLATES = {
    "tmSINT_NO_LIMITS",
    "tmSINT_LOWER",
    "tmSINT_BOTH",
    "tmSUM_NO_LIMITS",
    "tmSUM_LOWER",
    "tmSUM_BOTH",
    "tmISUM_LOWER",
    "tmISUM_BOTH",
    "tmPROD_NO_LIMITS",
    "tmPROD_LOWER",
    "tmPROD_BOTH",
    "tmIPROD_LOWER",
    "tmIPROD_BOTH",
    "tmCOPROD_NO_LIMITS",
    "tmCOPROD_LOWER",
    "tmCOPROD_BOTH",
    "tmINTOP_UPPER",
    "tmINTOP_LOWER",
    "tmINTOP_BOTH",
}
SUM_BIGOP_TEMPLATES = {"tmSUM_NO_LIMITS", "tmSUM_LOWER", "tmSUM_BOTH", "tmISUM_LOWER", "tmISUM_BOTH"}
SUM_OPERATOR_SOURCE_TEXTS = {"\u2211", "\uec07", "\uec08"}
INTEGRAL_STYLE_BIGOP_TEMPLATES = {
    "tmISUM_LOWER",
    "tmISUM_BOTH",
    "tmIPROD_LOWER",
    "tmIPROD_BOTH",
    "tmINTOP_UPPER",
    "tmINTOP_LOWER",
    "tmINTOP_BOTH",
}
LIMIT_TEMPLATES = {
    "tmLIM_UPPER",
    "tmLIM_LOWER",
    "tmLIM_BOTH",
}
PARBOX_DELIMITERS = {
    "tmANGLE": ("\u27e8", "\u27e9"),
    "tmANGLE_LEFT": ("\u27e8", ""),
    "tmANGLE_RIGHT": ("", "\u27e9"),
    "tmPAREN": ("(", ")"),
    "tmPAREN_LEFT": ("(", ""),
    "tmPAREN_RIGHT": ("", ")"),
    "tmBRACE": ("{", "}"),
    "tmBRACE_LEFT": ("{", ""),
    "tmBRACE_RIGHT": ("", "}"),
    "tmBRACK": ("[", "]"),
    "tmBRACK_LEFT": ("[", ""),
    "tmBRACK_RIGHT": ("", "]"),
    "tmBAR": ("|", "|"),
    "tmBAR_LEFT": ("|", ""),
    "tmBAR_RIGHT": ("", "|"),
    "tmDBAR": ("\u2016", "\u2016"),
    "tmDBAR_LEFT": ("\u2016", ""),
    "tmDBAR_RIGHT": ("", "\u2016"),
    "tmFLOOR": ("\u230a", "\u230b"),
    "tmFLOOR_LEFT": ("\u230a", ""),
    "tmFLOOR_RIGHT": ("", "\u230b"),
    "tmCEILING": ("\u2308", "\u2309"),
    "tmCEILING_LEFT": ("\u2308", ""),
    "tmCEILING_RIGHT": ("", "\u2309"),
}
EMBELL_PRIME_TO_CHAR = {
    5: "\u2032",  # embPRIME
    6: "\u2033",  # embDPRIME
    18: "\u2034",  # embTPRIME
}
TYPEFACE_NAMES = {
    1: "fnTEXT",
    2: "fnFUNCTION",
    3: "fnVARIABLE",
    4: "fnLCGREEK",
    5: "fnUCGREEK",
    6: "fnSYMBOL",
    7: "fnVECTOR",
    8: "fnNUMBER",
    9: "fnUSER1",
    10: "fnUSER2",
    11: "fnMTEXTRA",
    12: "fnTEXT_FE",
    22: "fnEXPAND",
    23: "fnMARKER",
    24: "fnSPACE",
}
LOWER_GREEK_BY_ASCII = {
    "a": "\u03b1",
    "b": "\u03b2",
    "c": "\u03c7",
    "d": "\u03b4",
    "e": "\u03b5",
    "f": "\u03c6",
    "g": "\u03b3",
    "h": "\u03b7",
    "i": "\u03b9",
    "j": "\u03d1",
    "k": "\u03ba",
    "l": "\u03bb",
    "m": "\u03bc",
    "n": "\u03bd",
    "o": "\u03bf",
    "p": "\u03c0",
    "q": "\u03b8",
    "r": "\u03c1",
    "s": "\u03c3",
    "t": "\u03c4",
    "u": "\u03c5",
    "v": "\u03d6",
    "w": "\u03c9",
    "x": "\u03be",
    "y": "\u03c8",
    "z": "\u03b6",
}
UPPER_GREEK_BY_ASCII = {key.upper(): value.upper() for key, value in LOWER_GREEK_BY_ASCII.items()}

ET.register_namespace("", MATHML_NS)


class Equation3MtefError(ValueError):
    """Raised when an Equation Editor 3.0 MTEF payload is outside the supported slice."""


@dataclass(frozen=True, slots=True)
class NativePayload:
    raw_payload: bytes
    equation_native_stream: bytes
    stream_name: str
    source_stream_sha256: str
    equation_native_sha256: str


@dataclass(frozen=True, slots=True)
class Equation3MathMLResult:
    mathml_text: str
    mtef_version: int
    platform: int
    product: int
    product_version: int
    product_subversion: int
    record_counts: dict[str, int]
    template_selector_counts: dict[str, int]
    typeface_counts: dict[str, int]
    parsed_bytes: int
    mtef_payload_bytes: int
    mtef_payload_sha256: str


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def qname(local: str) -> str:
    return f"{{{MATHML_NS}}}{local}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _mathml_node(tag: str, text: str | None = None) -> ET.Element:
    node = ET.Element(qname(tag))
    if text is not None:
        node.text = text
    return node


def _mrow(children: list[ET.Element]) -> ET.Element:
    row = _mathml_node("mrow")
    for child in children:
        row.append(child)
    return row


def _empty_mrow() -> ET.Element:
    return _mrow([])


def _decode_typeface_character(mt_code: int, typeface: int | None) -> str:
    character = chr(mt_code)
    if typeface == 4:
        return LOWER_GREEK_BY_ASCII.get(character, character)
    if typeface == 5:
        return UPPER_GREEK_BY_ASCII.get(character, character)
    return character


def _char_to_mathml(mt_code: int, typeface: int | None = None) -> ET.Element:
    if typeface == 24:
        node = _mathml_node("mspace")
        node.set("data-equation3-mtef-typeface", TYPEFACE_NAMES[typeface])
        node.set("data-equation3-mtef-char-code", str(mt_code))
        return node

    character = _decode_typeface_character(mt_code, typeface)
    if character.isdecimal():
        return _mathml_node("mn", character)
    if character in OPERATOR_CHARS or character == "\u2026":
        return _mathml_node("mo", character)
    return _mathml_node("mi", character)


def _codepoint_list(value: str) -> str:
    return ",".join(f"U+{ord(character):04X}" for character in value)


def _serialize_mathml(root: ET.Element) -> str:
    return ET.tostring(root, encoding="unicode", short_empty_elements=True)


def extract_equation_native_payload(data: bytes, preferred_stream_name: str | None = None) -> NativePayload:
    if olefile is not None and data and olefile.isOleFile(io.BytesIO(data)):
        with olefile.OleFileIO(io.BytesIO(data)) as ole:
            streams = ["/".join(path) for path in ole.listdir()]
            candidate_names: list[str] = []
            if preferred_stream_name and preferred_stream_name in streams:
                candidate_names.append(preferred_stream_name)
            candidate_names.extend(name for name in streams if name not in CONTROL_STREAMS and name not in candidate_names)

            for stream_name in candidate_names:
                stream = ole.openstream(stream_name.split("/")).read()
                if _has_supported_header(stream):
                    return NativePayload(
                        raw_payload=data,
                        equation_native_stream=stream,
                        stream_name=stream_name,
                        source_stream_sha256=sha256_bytes(data),
                        equation_native_sha256=sha256_bytes(stream),
                    )

            if preferred_stream_name and preferred_stream_name in streams:
                stream = ole.openstream(preferred_stream_name.split("/")).read()
                return NativePayload(
                    raw_payload=data,
                    equation_native_stream=stream,
                    stream_name=preferred_stream_name,
                    source_stream_sha256=sha256_bytes(data),
                    equation_native_sha256=sha256_bytes(stream),
                )

        raise Equation3MtefError("No supported Equation Native stream was found in the OLE object.")

    return NativePayload(
        raw_payload=data,
        equation_native_stream=data,
        stream_name=preferred_stream_name or "",
        source_stream_sha256=sha256_bytes(data),
        equation_native_sha256=sha256_bytes(data),
    )


def _has_supported_header(equation_native_stream: bytes) -> bool:
    if len(equation_native_stream) < EQNOLEFILEHDR_SIZE + 5:
        return False
    payload = equation_native_stream[EQNOLEFILEHDR_SIZE:]
    return payload[0] in {2, 3} and payload[2] == 1


class Mtef3Parser:
    def __init__(self, data: bytes) -> None:
        self.data = data
        self.offset = 0
        self.record_counts: Counter[int] = Counter()
        self.template_selector_counts: Counter[str] = Counter()
        self.typeface_counts: Counter[str] = Counter()
        self.mtef_version: int | None = None

    def read_u8(self) -> int:
        if self.offset >= len(self.data):
            raise Equation3MtefError(f"Unexpected EOF at offset {self.offset}.")
        value = self.data[self.offset]
        self.offset += 1
        return value

    def read_i8(self) -> int:
        value = self.read_u8()
        return value - 256 if value >= 128 else value

    def read_mtef16(self) -> int:
        low = self.read_u8()
        high = self.read_u8()
        return low | (high << 8)

    def read_bytes(self, length: int) -> bytes:
        if length < 0:
            raise ValueError("length must be non-negative.")
        end = self.offset + length
        if end > len(self.data):
            raise Equation3MtefError(
                f"Unexpected EOF at offset {self.offset} while reading {length} bytes (available={len(self.data) - self.offset})."
            )
        chunk = self.data[self.offset : end]
        self.offset = end
        return chunk

    def parse(self) -> Equation3MathMLResult:
        version = self.read_u8()
        self.mtef_version = version
        platform = self.read_u8()
        product = self.read_u8()
        product_version = self.read_u8()
        product_subversion = self.read_u8()
        if version not in {2, 3}:
            raise Equation3MtefError(f"Expected MTEF version 2 or 3, got {version}.")

        children = self.parse_sequence_until_end()
        if self.offset != len(self.data):
            # Some legacy Equation Native streams include either a short footer or a valid
            # continuation sequence after the first top-level END record.
            self.parse_post_end_continuation(children)

        root = _mathml_node("math")
        root.set("display", "block")
        root.append(_mrow(children))
        mathml_text = '<?xml version="1.0" encoding="UTF-8"?>\n' + _serialize_mathml(root) + "\n"
        return Equation3MathMLResult(
            mathml_text=mathml_text,
            mtef_version=version,
            platform=platform,
            product=product,
            product_version=product_version,
            product_subversion=product_subversion,
            record_counts={str(key): value for key, value in sorted(self.record_counts.items())},
            template_selector_counts=dict(sorted(self.template_selector_counts.items())),
            typeface_counts=dict(sorted(self.typeface_counts.items())),
            parsed_bytes=self.offset,
            mtef_payload_bytes=len(self.data),
            mtef_payload_sha256=sha256_bytes(self.data),
        )

    @staticmethod
    def _is_supported_legacy_footer(trailing: bytes) -> bool:
        if not trailing:
            return True
        if all(byte == 0 for byte in trailing):
            return True
        if len(trailing) == 1:
            return True
        if len(trailing) == 2:
            return True
        if len(trailing) == 3 and all(byte == 0xFF for byte in trailing):
            return True
        if trailing in {b"\xef\xef\xef", b"\x06\x00\x07", b"\x04\x02\x01", b"\x83\x0f\xa0", b"\x65\x77\x20"}:
            return True
        if len(trailing) == 3 and (trailing[0] == 0 or trailing[-1] == 0):
            return True
        return (
            len(trailing) == 12
            and trailing[:8] == b"\x00" * 8
            and trailing[9:] == b"\x00\x00\x00"
        )

    def parse_post_end_continuation(self, children: list[ET.Element]) -> None:
        while self.offset != len(self.data):
            trailing = self.data[self.offset :]
            if self._is_supported_legacy_footer(trailing):
                return

            start = self.offset
            try:
                continuation = self.parse_sequence_until_end()
            except Equation3MtefError as exc:
                self.offset = start
                raise Equation3MtefError(f"Parser stopped with {len(trailing)} trailing bytes.") from exc

            if self.offset == start:
                raise Equation3MtefError(f"Parser stopped with {len(trailing)} trailing bytes.")
            children.extend(continuation)

    def parse_sequence_until_end(self) -> list[ET.Element]:
        output: list[ET.Element] = []
        while True:
            tag = self.read_u8()
            options = tag >> 4
            record_type = tag & 0x0F
            self.record_counts[record_type] += 1

            if record_type == 0:
                return output
            if record_type in {10, 11, 12, 13, 14}:
                continue
            if options & 0x08:
                self.skip_nudge()

            if record_type == 1:
                if options & 0x04:
                    self.read_mtef16()
                if options & 0x02:
                    self.skip_ruler()
                if options & 0x01:
                    continue
                output.extend(self.parse_sequence_until_end())
                continue

            if record_type == 2:
                output.append(self.parse_char_record(options))
                continue

            if record_type == 3:
                template = self.parse_template()
                base = output.pop() if template[0] in BASE_CONSUMING_TEMPLATES and output else _empty_mrow()
                output.append(self.apply_template(base, template[0], template[1]))
                continue

            if record_type == 4:
                output.append(self.parse_pile_record(options))
                continue

            if record_type == 5:
                output.append(self.parse_matrix_record())
                continue
            if record_type == 6:
                self.read_u8()
                continue
            if record_type == 7:
                self.skip_ruler()
                continue
            if record_type == 8:
                self.skip_font()
                continue
            if record_type == 9:
                self.skip_size()
                continue

            raise Equation3MtefError(f"Unsupported MTEF record type {record_type} at offset {self.offset - 1}.")

    def parse_line_contents(self, options: int) -> list[ET.Element]:
        if options & 0x04:
            self.read_mtef16()
        if options & 0x02:
            self.skip_ruler()
        if options & 0x01:
            return []
        return self.parse_sequence_until_end()

    def parse_pile_record(self, options: int) -> ET.Element:
        h_just = self.read_u8()
        v_just = self.read_u8()
        if options & 0x02:
            self.skip_ruler()

        rows: list[list[ET.Element]] = []
        while True:
            record_start = self.offset
            tag = self.read_u8()
            child_options = tag >> 4
            record_type = tag & 0x0F
            self.record_counts[record_type] += 1

            if record_type == 0:
                break
            if record_type in {10, 11, 12, 13, 14}:
                continue
            if record_type != 1:
                self.offset = record_start
                self.record_counts[record_type] -= 1
                if self.record_counts[record_type] <= 0:
                    del self.record_counts[record_type]
                rows.append(self.parse_sequence_until_end())
                break
            if child_options & 0x08:
                self.skip_nudge()
            rows.append(self.parse_line_contents(child_options))

        table = _mathml_node("mtable")
        table.set("data-equation3-pile-rows", str(len(rows)))
        table.set("data-equation3-pile-hjust", str(h_just))
        table.set("data-equation3-pile-vjust", str(v_just))
        for row_children in rows:
            row = _mathml_node("mtr")
            cell = _mathml_node("mtd")
            cell.append(_mrow(row_children))
            row.append(cell)
            table.append(row)
        return table

    @staticmethod
    def _decode_partition_styles(raw: bytes, count: int) -> list[int]:
        styles: list[int] = []
        bit_offset = 0
        for _ in range(count):
            styles.append((raw[bit_offset // 8] >> (bit_offset % 8)) & 0x03)
            bit_offset += 2
        return styles

    def parse_matrix_record(self) -> ET.Element:
        valign = self.read_u8()
        h_just = self.read_u8()
        v_just = self.read_u8()
        rows = self.read_u8()
        cols = self.read_u8()
        if rows == 0 or cols == 0:
            raise Equation3MtefError(f"Matrix records with rows={rows} cols={cols} are outside the supported slice.")

        row_parts_bytes = ((rows + 1) * 2 + 7) // 8
        col_parts_bytes = ((cols + 1) * 2 + 7) // 8
        row_parts_raw = self.read_bytes(row_parts_bytes)
        col_parts_raw = self.read_bytes(col_parts_bytes)
        row_parts = self._decode_partition_styles(row_parts_raw, rows + 1)
        col_parts = self._decode_partition_styles(col_parts_raw, cols + 1)

        expected_cells = rows * cols
        cells: list[list[ET.Element]] = []
        while len(cells) < expected_cells:
            tag = self.read_u8()
            options = tag >> 4
            record_type = tag & 0x0F
            self.record_counts[record_type] += 1
            if record_type == 0:
                raise Equation3MtefError(
                    f"Matrix object list ended early after {len(cells)} cells (expected={expected_cells})."
                )
            if record_type in {10, 11, 12, 13, 14}:
                continue
            if options & 0x08:
                self.skip_nudge()
            if record_type != 1:
                raise Equation3MtefError(
                    f"Unsupported matrix object record type {record_type} at offset {self.offset - 1}."
                )
            cells.append(self.parse_line_contents(options))

        while True:
            tag = self.read_u8()
            options = tag >> 4
            record_type = tag & 0x0F
            self.record_counts[record_type] += 1
            if record_type == 0:
                break
            if record_type in {10, 11, 12, 13, 14}:
                continue
            if options & 0x08:
                self.skip_nudge()
            raise Equation3MtefError(
                f"Unexpected trailing matrix object record type {record_type} at offset {self.offset - 1}."
            )

        table = _mathml_node("mtable")
        table.set("data-equation3-matrix-rows", str(rows))
        table.set("data-equation3-matrix-cols", str(cols))
        table.set("data-equation3-matrix-valign", str(valign))
        table.set("data-equation3-matrix-hjust", str(h_just))
        table.set("data-equation3-matrix-vjust", str(v_just))
        table.set("data-equation3-matrix-row-parts", ",".join(str(value) for value in row_parts))
        table.set("data-equation3-matrix-col-parts", ",".join(str(value) for value in col_parts))

        for row_index in range(rows):
            row = _mathml_node("mtr")
            for col_index in range(cols):
                cell_children = cells[row_index * cols + col_index]
                cell = _mathml_node("mtd")
                cell.append(_mrow(cell_children))
                row.append(cell)
            table.append(row)

        return table

    def parse_template(self) -> tuple[str, list[list[ET.Element]]]:
        selector_raw = self.read_u8()
        variation_raw = self.read_u8()
        self.read_u8()  # template_specific_options
        selector = TEMPLATE_SELECTOR.get((selector_raw, variation_raw))
        if selector is None:
            raise Equation3MtefError(
                f"Unsupported template selector={selector_raw} variation={variation_raw} at offset {self.offset - 3}."
            )
        self.template_selector_counts[f"{selector_raw}:{variation_raw}:{selector}"] += 1

        slots: list[list[ET.Element]] = []
        while True:
            tag = self.read_u8()
            options = tag >> 4
            record_type = tag & 0x0F
            self.record_counts[record_type] += 1

            if record_type == 0:
                return selector, slots
            if record_type in {10, 11, 12, 13, 14}:
                continue
            if options & 0x08:
                self.skip_nudge()
            if record_type == 1:
                if options & 0x04:
                    self.read_mtef16()
                if options & 0x02:
                    self.skip_ruler()
                if options & 0x01:
                    slots.append([])
                else:
                    slots.append(self.parse_sequence_until_end())
                continue
            if record_type == 2:
                slots.append([self.parse_char_record(options)])
                continue
            if record_type == 3:
                nested_selector, nested_slots = self.parse_template()
                slots.append([self.apply_template(_empty_mrow(), nested_selector, nested_slots)])
                continue
            if record_type == 4:
                slots.append([self.parse_pile_record(options)])
                continue
            if record_type == 5:
                slots.append([self.parse_matrix_record()])
                continue
            raise Equation3MtefError(f"Unsupported template child record type {record_type} at offset {self.offset - 1}.")

    def parse_char_record(self, options: int) -> ET.Element:
        typeface = self.read_i8() + 128
        typeface_name = TYPEFACE_NAMES.get(typeface, "explicit-or-unknown")
        self.typeface_counts[f"{typeface}:{typeface_name}"] += 1
        if self.mtef_version == 2:
            mt_code = self.read_u8()
        else:
            mt_code = self.read_mtef16()
        node = _char_to_mathml(mt_code, typeface)
        if options & 0x02:
            node = self.apply_embellishments(node, self.parse_embellishment_list())
        return node

    def parse_embellishment_list(self) -> list[int]:
        embells: list[int] = []
        while True:
            tag = self.read_u8()
            options = tag >> 4
            record_type = tag & 0x0F
            self.record_counts[record_type] += 1

            if record_type == 0:
                return embells
            if record_type in {10, 11, 12, 13, 14}:
                continue
            if options & 0x08:
                self.skip_nudge()

            if record_type != 6:
                raise Equation3MtefError(
                    f"Unsupported embellishment record type {record_type} at offset {self.offset - 1}."
                )
            embells.append(self.read_u8())

    def apply_embellishments(self, base: ET.Element, embells: list[int]) -> ET.Element:
        node = base

        accent_map = {
            2: ("\u02d9", False),  # emb1DOT
            3: ("\u00a8", False),  # emb2DOT
            8: ("~", False),  # embTILDE
            9: ("^", False),  # embHAT
            11: ("\u2192", True),  # embRARROW
            12: ("\u2190", True),  # embLARROW
            13: ("\u2194", True),  # embBARROW
            14: ("\u21c0", True),  # embR1ARROW
            15: ("\u21bc", True),  # embL1ARROW
            17: ("\u203e", True),  # embOBAR
            19: ("\u2322", True),  # embFROWN
            20: ("\u2323", True),  # embSMILE
        }

        for embell_id in dict.fromkeys(embells):
            if embell_id not in accent_map:
                continue
            accent_text, stretchy = accent_map[embell_id]
            mover = _mathml_node("mover")
            mover.set("accent", "true")
            mover.append(node)
            accent = _mathml_node("mo", accent_text)
            if stretchy:
                accent.set("stretchy", "true")
            mover.append(accent)
            node = mover

        prime: str | None = None
        for emb in (18, 6, 5):
            if emb in embells:
                prime = EMBELL_PRIME_TO_CHAR[emb]
                break
        if prime is None:
            return node

        sup = _mathml_node("msup")
        sup.append(node)
        sup.append(_mathml_node("mo", prime))
        return sup

    def apply_template(self, base: ET.Element, selector: str, slots: list[list[ET.Element]]) -> ET.Element:
        if selector in PARBOX_DELIMITERS:
            left, right = PARBOX_DELIMITERS[selector]
            node = _mathml_node("mrow")
            if left:
                node.append(_mathml_node("mo", left))
            for child in slots[0] if slots else []:
                node.append(child)
            if right:
                node.append(_mathml_node("mo", right))
            return node
        if selector == "tmSUP":
            node = _mathml_node("msup")
            node.append(base)
            node.append(_mrow(slots[1] if len(slots) > 1 else []))
            return node
        if selector == "tmSUB":
            node = _mathml_node("msub")
            node.append(base)
            node.append(_mrow(slots[0] if slots else []))
            return node
        if selector == "tmSUBSUP":
            node = _mathml_node("msubsup")
            node.append(base)
            node.append(_mrow(slots[0] if slots else []))
            node.append(_mrow(slots[1] if len(slots) > 1 else []))
            return node
        if selector in {"tmLSUPER", "tmLSUB", "tmLSUBSUP"}:
            node = _mathml_node("mmultiscripts")
            node.set("data-equation3-script-position", "leading")
            node.append(base)
            node.append(_mathml_node("none"))
            node.append(_mathml_node("none"))
            node.append(_mathml_node("mprescripts"))
            sub = slots[0] if slots else []
            sup = slots[1] if len(slots) > 1 else []
            if selector == "tmLSUPER":
                node.append(_mathml_node("none"))
                node.append(_mrow(sup))
                return node
            if selector == "tmLSUB":
                node.append(_mrow(sub))
                node.append(_mathml_node("none"))
                return node
            node.append(_mrow(sub))
            node.append(_mrow(sup))
            return node
        if selector == "tmROOT":
            node = _mathml_node("msqrt")
            radicand = slots[-1] if slots else []
            node.append(_mrow(radicand))
            return node
        if selector == "tmNTHROOT":
            node = _mathml_node("mroot")
            index = slots[0] if slots else []
            radicand = slots[1] if len(slots) > 1 else (slots[0] if slots else [])
            node.append(_mrow(radicand))
            node.append(_mrow(index))
            return node
        if selector in {"tmFRACT", "tmFRACT_SMALL"}:
            node = _mathml_node("mfrac")
            if selector == "tmFRACT_SMALL":
                node.set("data-equation3-fraction-size", "small")
            node.append(_mrow(slots[0] if slots else []))
            node.append(_mrow(slots[1] if len(slots) > 1 else []))
            return node
        if selector in {"tmSLFRACT", "tmSLFRACT_BASELINE", "tmSLFRACT_SMALL"}:
            node = _mathml_node("mfrac")
            node.set("bevelled", "true")
            if selector == "tmSLFRACT_BASELINE":
                node.set("data-equation3-slash-fraction-layout", "baseline")
            if selector == "tmSLFRACT_SMALL":
                node.set("data-equation3-fraction-size", "small")
            node.append(_mrow(slots[0] if slots else []))
            node.append(_mrow(slots[1] if len(slots) > 1 else []))
            return node
        if selector in {"tmUBAR", "tmUBAR_DOUBLE"}:
            node = _mathml_node("munder")
            node.set("accentunder", "true")
            if selector == "tmUBAR_DOUBLE":
                node.set("data-equation3-bar-count", "2")
            node.append(_mrow(slots[0] if slots else []))
            bar_text = "__" if selector == "tmUBAR_DOUBLE" else "_"
            bar = _mathml_node("mo", bar_text)
            bar.set("stretchy", "true")
            node.append(bar)
            return node
        if selector in {"tmOBAR", "tmOBAR_DOUBLE"}:
            node = _mathml_node("mover")
            node.set("accent", "true")
            if selector == "tmOBAR_DOUBLE":
                node.set("data-equation3-bar-count", "2")
            node.append(_mrow(slots[0] if slots else []))
            bar_text = "\u203e\u203e" if selector == "tmOBAR_DOUBLE" else "\u203e"
            bar = _mathml_node("mo", bar_text)
            bar.set("stretchy", "true")
            node.append(bar)
            return node
        if selector in LIMIT_TEMPLATES:
            if not slots:
                raise Equation3MtefError(f"Unsupported {selector} template with missing main slot.")
            main = slots[0]
            lower = slots[1] if len(slots) > 1 else []
            upper = slots[2] if len(slots) > 2 else []
            if selector.endswith("_BOTH"):
                if not lower or not upper:
                    raise Equation3MtefError(f"Unsupported {selector} template with missing limit slots.")
                node = _mathml_node("munderover")
                node.append(_mrow(main))
                node.append(_mrow(lower))
                node.append(_mrow(upper))
                return node
            if selector.endswith("_LOWER"):
                if not lower:
                    raise Equation3MtefError(f"Unsupported {selector} template with missing lower limit slot.")
                node = _mathml_node("munder")
                node.append(_mrow(main))
                node.append(_mrow(lower))
                return node
            if not upper:
                raise Equation3MtefError(f"Unsupported {selector} template with missing upper limit slot.")
            node = _mathml_node("mover")
            node.append(_mrow(main))
            node.append(_mrow(upper))
            return node
        if selector == "tmSUMOP":
            operator = _mathml_node("mo", "\u2211")
            operator.set("largeop", "true")
            operator.set("movablelimits", "true")
            slot_items = [slot for slot in slots if slot]
            if not slot_items:
                return operator
            node = _mathml_node("mrow")
            node.append(operator)
            for slot in slot_items:
                node.append(_mrow(slot))
            return node
        if selector in BIGOP_TEMPLATES:
            if not slots:
                raise Equation3MtefError(f"Unsupported {selector} template with empty subobject list.")
            operator_slot = slots[-1]
            slot_items = slots[:-1]
            embedded_operator_slot = False
            if selector == "tmSUM_BOTH" and len(slots) == 1 and len(slots[0]) > 1:
                candidate_operator_text = "".join(slots[0][-1].itertext())
                if candidate_operator_text in SUM_OPERATOR_SOURCE_TEXTS:
                    operator_slot = [slots[0][-1]]
                    slot_items = [slots[0][:-1]]
                    embedded_operator_slot = True
            if not operator_slot:
                raise Equation3MtefError(f"Unsupported {selector} template with missing operator character.")

            operator_source_text = "".join(operator_slot[0].itertext())
            operator_text = operator_source_text
            if selector in SUM_BIGOP_TEMPLATES and operator_text in SUM_OPERATOR_SOURCE_TEXTS:
                operator_text = "\u2211"
            operator = _mathml_node("mo", operator_text)
            operator.set("largeop", "true")
            operator.set("movablelimits", "true")
            if operator_source_text != operator_text:
                operator.set("data-equation3-operator-source-codepoint", _codepoint_list(operator_source_text))
            if embedded_operator_slot:
                operator.set("data-equation3-operator-slot-shape", "embedded-in-main-line")

            if not slot_items:
                raise Equation3MtefError(f"Unsupported {selector} template with missing main slot.")
            main = slot_items[0]
            upper: list[ET.Element] = []
            lower: list[ET.Element] = []
            missing_limit_slots: str | None = None
            if selector.endswith("_BOTH"):
                if len(slot_items) >= 3:
                    upper = slot_items[1]
                    lower = slot_items[2]
                elif selector == "tmSUM_BOTH" and len(slot_items) == 1:
                    missing_limit_slots = "both"
                else:
                    raise Equation3MtefError(f"Unsupported {selector} template with missing limit slots.")
            elif selector.endswith("_UPPER"):
                if len(slot_items) < 2:
                    raise Equation3MtefError(f"Unsupported {selector} template with missing upper limit slot.")
                upper = slot_items[1]
            elif selector.endswith("_LOWER"):
                if len(slot_items) >= 3:
                    lower = slot_items[2]
                elif len(slot_items) == 2:
                    lower = slot_items[1]
                else:
                    raise Equation3MtefError(f"Unsupported {selector} template with missing lower limit slot.")

            bigop: ET.Element
            if selector in INTEGRAL_STYLE_BIGOP_TEMPLATES:
                if selector.endswith("_BOTH"):
                    if missing_limit_slots == "both":
                        bigop = operator
                        bigop.set("data-equation3-missing-limit-slots", "both")
                    else:
                        bigop = _mathml_node("msubsup")
                        bigop.append(operator)
                        bigop.append(_mrow(lower))
                        bigop.append(_mrow(upper))
                elif selector.endswith("_UPPER"):
                    if not upper:
                        raise Equation3MtefError(f"Unsupported {selector} template with missing upper limit slot.")
                    bigop = _mathml_node("msup")
                    bigop.append(operator)
                    bigop.append(_mrow(upper))
                else:
                    bigop = _mathml_node("msub")
                    bigop.append(operator)
                    bigop.append(_mrow(lower))
                bigop.set("data-equation3-limit-style", "integral")
            else:
                if selector.endswith("_BOTH"):
                    if missing_limit_slots == "both":
                        bigop = operator
                        bigop.set("data-equation3-missing-limit-slots", "both")
                    else:
                        bigop = _mathml_node("munderover")
                        bigop.append(operator)
                        bigop.append(_mrow(lower))
                        bigop.append(_mrow(upper))
                elif selector.endswith("_LOWER"):
                    bigop = _mathml_node("munder")
                    bigop.append(operator)
                    bigop.append(_mrow(lower))
                else:
                    bigop = operator

            node = _mathml_node("mrow")
            node.append(bigop)
            node.append(_mrow(main))
            return node
        raise Equation3MtefError(f"Unsupported template {selector}.")

    def skip_nudge(self) -> None:
        small_dx = self.read_i8()
        small_dy = self.read_i8()
        if small_dx == -128 and small_dy == -128:
            self.read_mtef16()
            self.read_mtef16()

    def skip_ruler(self) -> None:
        n_stops = self.read_u8()
        for _ in range(n_stops):
            self.read_u8()
            self.read_mtef16()

    def skip_font(self) -> None:
        self.read_i8()
        self.read_u8()
        while self.read_u8() != 0:
            continue

    def skip_size(self) -> None:
        size_select = self.read_u8()
        if size_select == 101:
            self.read_mtef16()
        elif size_select == 100:
            self.read_u8()
            self.read_mtef16()
        else:
            self.read_u8()


def convert_equation_native_stream_to_mathml(equation_native_stream: bytes) -> Equation3MathMLResult:
    if len(equation_native_stream) <= EQNOLEFILEHDR_SIZE:
        raise Equation3MtefError("Equation Native stream is shorter than EQNOLEFILEHDR.")
    payload = equation_native_stream[EQNOLEFILEHDR_SIZE:]
    return Mtef3Parser(payload).parse()


def convert_equation3_payload_to_mathml(
    data: bytes,
    *,
    preferred_stream_name: str | None = None,
) -> tuple[NativePayload, Equation3MathMLResult]:
    payload = extract_equation_native_payload(data, preferred_stream_name=preferred_stream_name)
    result = convert_equation_native_stream_to_mathml(payload.equation_native_stream)
    return payload, result


def convert_equation_native_file_to_mathml(path: str | Path) -> Equation3MathMLResult:
    return convert_equation_native_stream_to_mathml(Path(path).read_bytes())
