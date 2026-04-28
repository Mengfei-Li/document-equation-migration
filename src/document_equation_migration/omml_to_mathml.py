from __future__ import annotations

from copy import deepcopy
from xml.etree import ElementTree as ET


OMML_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MATHML_NAMESPACE = "http://www.w3.org/1998/Math/MathML"

ET.register_namespace("math", MATHML_NAMESPACE)


def _qname(local_name: str) -> str:
    return f"{{{MATHML_NAMESPACE}}}{local_name}"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _mml(local_name: str, text: str | None = None, children: list[ET.Element] | None = None) -> ET.Element:
    element = ET.Element(_qname(local_name))
    if text is not None:
        element.text = text
    for child in children or []:
        element.append(child)
    return element


def _flatten_children(element: ET.Element) -> list[ET.Element]:
    converted: list[ET.Element] = []
    for child in list(element):
        item = _convert_node(child)
        if item is None:
            continue
        converted.append(item)
    return converted


def _row(children: list[ET.Element]) -> ET.Element:
    if len(children) == 1:
        return children[0]
    return _mml("mrow", children=children)


def _first_child(element: ET.Element, local_name: str) -> ET.Element | None:
    for child in list(element):
        if _local_name(child.tag) == local_name:
            return child
    return None


def _converted_child(element: ET.Element, local_name: str) -> ET.Element:
    child = _first_child(element, local_name)
    if child is None:
        return _mml("mrow")
    return _row(_flatten_children(child))


def _token_for_text(text: str) -> ET.Element:
    stripped = text.strip()
    if not stripped:
        return _mml("mtext", "")
    if stripped.replace(".", "", 1).isdigit():
        return _mml("mn", stripped)
    if all(char in "+-=*/^_()[]{}<>|,.;:!?" for char in stripped):
        return _mml("mo", stripped)
    return _mml("mi", stripped)


def _convert_run(element: ET.Element) -> ET.Element | None:
    tokens: list[ET.Element] = []
    for child in list(element):
        if _local_name(child.tag) != "t":
            converted = _convert_node(child)
            if converted is not None:
                tokens.append(converted)
            continue
        if child.text:
            tokens.append(_token_for_text(child.text))
    if not tokens:
        return None
    return _row(tokens)


def _convert_fraction(element: ET.Element) -> ET.Element:
    return _mml(
        "mfrac",
        children=[
            _converted_child(element, "num"),
            _converted_child(element, "den"),
        ],
    )


def _convert_script(element: ET.Element, tag: str) -> ET.Element:
    base = _converted_child(element, "e")
    if tag == "sSup":
        return _mml("msup", children=[base, _converted_child(element, "sup")])
    if tag == "sSub":
        return _mml("msub", children=[base, _converted_child(element, "sub")])
    return _mml(
        "msubsup",
        children=[
            base,
            _converted_child(element, "sub"),
            _converted_child(element, "sup"),
        ],
    )


def _convert_radical(element: ET.Element) -> ET.Element:
    base = _converted_child(element, "e")
    degree = _first_child(element, "deg")
    if degree is None or not list(degree):
        return _mml("msqrt", children=[base])
    return _mml("mroot", children=[base, _row(_flatten_children(degree))])


def _convert_delimiter(element: ET.Element) -> ET.Element:
    base = _converted_child(element, "e")
    fenced = _mml("mfenced", children=[base])
    delimiter_properties = _first_child(element, "dPr")
    if delimiter_properties is None:
        return fenced
    begin = delimiter_properties.find(f".//{{{OMML_NAMESPACE}}}begChr")
    end = delimiter_properties.find(f".//{{{OMML_NAMESPACE}}}endChr")
    if begin is not None:
        value = begin.attrib.get(f"{{{OMML_NAMESPACE}}}val")
        if value:
            fenced.set("open", value)
    if end is not None:
        value = end.attrib.get(f"{{{OMML_NAMESPACE}}}val")
        if value:
            fenced.set("close", value)
    return fenced


def _convert_nary(element: ET.Element) -> ET.Element:
    nary_properties = _first_child(element, "naryPr")
    operator = "\u2211"
    if nary_properties is not None:
        chr_node = nary_properties.find(f".//{{{OMML_NAMESPACE}}}chr")
        if chr_node is not None:
            operator = chr_node.attrib.get(f"{{{OMML_NAMESPACE}}}val", operator)
    children = [_mml("mo", operator)]
    sub = _first_child(element, "sub")
    sup = _first_child(element, "sup")
    expression = _first_child(element, "e")
    if sub is not None and sup is not None:
        children[0] = _mml("munderover", children=[children[0], _row(_flatten_children(sub)), _row(_flatten_children(sup))])
    elif sub is not None:
        children[0] = _mml("munder", children=[children[0], _row(_flatten_children(sub))])
    elif sup is not None:
        children[0] = _mml("mover", children=[children[0], _row(_flatten_children(sup))])
    if expression is not None:
        children.append(_row(_flatten_children(expression)))
    return _row(children)


def _convert_eq_array(element: ET.Element) -> ET.Element:
    rows = []
    for item in list(element):
        if _local_name(item.tag) != "e":
            continue
        rows.append(_mml("mtr", children=[_mml("mtd", children=[_row(_flatten_children(item))])]))
    return _mml("mtable", children=rows)


def _convert_matrix(element: ET.Element) -> ET.Element:
    rows = []
    for row in list(element):
        if _local_name(row.tag) != "mr":
            continue
        cells = [
            _mml("mtd", children=[_row(_flatten_children(cell))])
            for cell in list(row)
            if _local_name(cell.tag) == "e"
        ]
        if cells:
            rows.append(_mml("mtr", children=cells))
    return _mml("mtable", children=rows)


def _omml_property_value(element: ET.Element, property_name: str, default: str) -> str:
    property_node = element.find(f".//{{{OMML_NAMESPACE}}}{property_name}")
    if property_node is None:
        return default
    return property_node.attrib.get(f"{{{OMML_NAMESPACE}}}val", default)


def _convert_accent(element: ET.Element) -> ET.Element:
    operator = _omml_property_value(element, "chr", "^")
    mover = _mml("mover", children=[_converted_child(element, "e"), _mml("mo", operator)])
    mover.set("accent", "true")
    return mover


def _convert_bar(element: ET.Element) -> ET.Element:
    operator = "\u00af"
    base = _converted_child(element, "e")
    position = _omml_property_value(element, "pos", "top")
    if position == "bot":
        return _mml("munder", children=[base, _mml("mo", operator)])
    return _mml("mover", children=[base, _mml("mo", operator)])


def _convert_function(element: ET.Element) -> ET.Element:
    return _row([
        _converted_child(element, "fName"),
        _mml("mo", "\u2061"),
        _converted_child(element, "e"),
    ])


def _convert_limit(element: ET.Element, tag: str) -> ET.Element:
    base = _converted_child(element, "e")
    limit = _converted_child(element, "lim")
    if tag == "limLow":
        return _mml("munder", children=[base, limit])
    return _mml("mover", children=[base, limit])


def _unsupported_row(element: ET.Element) -> ET.Element:
    children = _flatten_children(element)
    row = _mml("mrow", children=children)
    row.set("data-omml-unsupported", _local_name(element.tag))
    return row


def _convert_node(element: ET.Element) -> ET.Element | None:
    local = _local_name(element.tag)
    if local in {"oMath", "oMathPara", "num", "den", "e", "sup", "sub", "deg", "fName", "lim"}:
        return _row(_flatten_children(element))
    if local == "r":
        return _convert_run(element)
    if local == "t":
        return _token_for_text(element.text or "")
    if local == "f":
        return _convert_fraction(element)
    if local in {"sSup", "sSub", "sSubSup"}:
        return _convert_script(element, local)
    if local == "rad":
        return _convert_radical(element)
    if local == "d":
        return _convert_delimiter(element)
    if local == "nary":
        return _convert_nary(element)
    if local == "eqArr":
        return _convert_eq_array(element)
    if local == "m":
        return _convert_matrix(element)
    if local == "acc":
        return _convert_accent(element)
    if local == "bar":
        return _convert_bar(element)
    if local == "func":
        return _convert_function(element)
    if local in {"limLow", "limUpp"}:
        return _convert_limit(element, local)
    if local.endswith("Pr"):
        return None
    if element.text and not list(element):
        return _token_for_text(element.text)
    if list(element):
        return _unsupported_row(element)
    return None


def omml_fragment_to_mathml(omml_xml: str | bytes) -> str:
    root = ET.fromstring(omml_xml)
    math = _mml("math")
    converted = _convert_node(root)
    if converted is not None:
        math.append(deepcopy(converted))
    return ET.tostring(math, encoding="unicode")
