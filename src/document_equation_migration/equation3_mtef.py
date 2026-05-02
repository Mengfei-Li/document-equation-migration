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
    (7, 0): "tmCEILING",
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
}
BASE_CONSUMING_TEMPLATES = {"tmSUP", "tmSUB", "tmSUBSUP"}
OPERATOR_CHARS = set("=+-*/(),[]{}") | {"\u2211", "\u222b"}  # ∑, ∫
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
    "tmCEILING": ("\u2308", "\u2309"),
}
EMBELL_PRIME_TO_CHAR = {
    5: "\u2032",  # embPRIME
    6: "\u2033",  # embDPRIME
    18: "\u2034",  # embTPRIME
}

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


def _char_to_mathml(mt_code: int) -> ET.Element:
    character = chr(mt_code)
    if character.isdecimal():
        return _mathml_node("mn", character)
    if character in OPERATOR_CHARS or character == "\u2026":
        return _mathml_node("mo", character)
    return _mathml_node("mi", character)


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
    return payload[0] == 3 and payload[2] == 1


class Mtef3Parser:
    def __init__(self, data: bytes) -> None:
        self.data = data
        self.offset = 0
        self.record_counts: Counter[int] = Counter()
        self.template_selector_counts: Counter[str] = Counter()

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
        platform = self.read_u8()
        product = self.read_u8()
        product_version = self.read_u8()
        product_subversion = self.read_u8()
        if version != 3:
            raise Equation3MtefError(f"Expected MTEF version 3, got {version}.")

        children = self.parse_sequence_until_end()
        if self.offset != len(self.data):
            raise Equation3MtefError(f"Parser stopped with {len(self.data) - self.offset} trailing bytes.")

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
            parsed_bytes=self.offset,
            mtef_payload_bytes=len(self.data),
            mtef_payload_sha256=sha256_bytes(self.data),
        )

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
                self.read_i8() + 128
                mt_code = self.read_mtef16()
                node = _char_to_mathml(mt_code)
                if options & 0x02:
                    node = self.apply_embellishments(node, self.parse_embellishment_list())
                output.append(node)
                continue

            if record_type == 3:
                template = self.parse_template()
                base = output.pop() if template[0] in BASE_CONSUMING_TEMPLATES and output else _empty_mrow()
                output.append(self.apply_template(base, template[0], template[1]))
                continue

            if record_type == 4:
                self.read_u8()
                self.read_u8()
                if options & 0x02:
                    self.skip_ruler()
                output.extend(self.parse_sequence_until_end())
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

            raise Equation3MtefError(f"Unsupported MTEF3 record type {record_type} at offset {self.offset - 1}.")

    def parse_line_contents(self, options: int) -> list[ET.Element]:
        if options & 0x04:
            self.read_mtef16()
        if options & 0x02:
            self.skip_ruler()
        if options & 0x01:
            return []
        return self.parse_sequence_until_end()

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
                self.read_i8() + 128
                mt_code = self.read_mtef16()
                node = _char_to_mathml(mt_code)
                if options & 0x02:
                    node = self.apply_embellishments(node, self.parse_embellishment_list())
                slots.append([node])
                continue
            raise Equation3MtefError(f"Unsupported template child record type {record_type} at offset {self.offset - 1}.")

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
        prime: str | None = None
        for emb in (18, 6, 5):
            if emb in embells:
                prime = EMBELL_PRIME_TO_CHAR[emb]
                break
        if prime is None:
            return base

        node = _mathml_node("msup")
        node.append(base)
        node.append(_mathml_node("mo", prime))
        return node

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
        if selector in {
            "tmSINT_NO_LIMITS",
            "tmSINT_LOWER",
            "tmSINT_BOTH",
            "tmSUM_NO_LIMITS",
            "tmSUM_LOWER",
            "tmSUM_BOTH",
        }:
            if not slots:
                raise Equation3MtefError(f"Unsupported {selector} template with empty subobject list.")
            if not slots[-1]:
                raise Equation3MtefError(f"Unsupported {selector} template with missing operator character.")

            operator_text = "".join(slots[-1][0].itertext())
            operator = _mathml_node("mo", operator_text)
            operator.set("largeop", "true")
            operator.set("movablelimits", "true")

            slot_items = slots[:-1]
            if not slot_items:
                raise Equation3MtefError(f"Unsupported {selector} template with missing main slot.")
            main = slot_items[0]
            upper: list[ET.Element] = []
            lower: list[ET.Element] = []
            if selector.endswith("_BOTH"):
                if len(slot_items) < 3:
                    raise Equation3MtefError(f"Unsupported {selector} template with missing limit slots.")
                upper = slot_items[1]
                lower = slot_items[2]
            elif selector.endswith("_LOWER"):
                if len(slot_items) >= 3:
                    lower = slot_items[2]
                elif len(slot_items) == 2:
                    lower = slot_items[1]
                else:
                    raise Equation3MtefError(f"Unsupported {selector} template with missing lower limit slot.")

            bigop: ET.Element
            if selector.endswith("_BOTH"):
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
