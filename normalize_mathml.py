import sys
from pathlib import Path
from xml.etree import ElementTree as ET


sys.stdout.reconfigure(encoding="utf-8")


MATHML_NS = "http://www.w3.org/1998/Math/MathML"
ET.register_namespace("", MATHML_NS)
KNOWN_FUNCTIONS = (
    "sin",
    "cos",
    "tan",
    "cot",
    "sec",
    "csc",
    "log",
    "ln",
)
SCRIPT_TAGS = {"msup", "msub", "msubsup"}


def qname(local: str) -> str:
    return f"{{{MATHML_NS}}}{local}"


def local_name(tag: str) -> str:
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def is_empty_node(node: ET.Element) -> bool:
    if (node.text or "").strip():
        return False
    return len(list(node)) == 0


def clear_children(node: ET.Element):
    for child in list(node):
        node.remove(child)


def ensure_child_count(node: ET.Element, expected: int):
    while len(list(node)) < expected:
        node.append(ET.Element(qname("mrow")))


def replace_children(node: ET.Element, children: list[ET.Element]):
    clear_children(node)
    for child in children:
        node.append(child)


def make_text_node(tag: str, text: str) -> ET.Element:
    node = ET.Element(qname(tag))
    node.text = text
    return node


def get_single_ascii_letter(node: ET.Element | None) -> str | None:
    if node is None:
        return None
    if local_name(node.tag) not in {"mi", "mn"}:
        return None
    if len(list(node)) != 0:
        return None

    text = (node.text or "").strip()
    if len(text) != 1 or not text.isascii() or not text.isalpha():
        return None
    return text


def fix_binary_script(parent: ET.Element, index: int, node: ET.Element, script_kind: str) -> bool:
    if index <= 0:
        return False

    ensure_child_count(node, 2)
    first, second = list(node)[:2]
    prev = list(parent)[index - 1]

    if script_kind == "msup" and is_empty_node(first) and not is_empty_node(second):
        parent.remove(prev)
        replace_children(node, [prev, second])
        return True

    if script_kind == "msub" and not is_empty_node(first) and is_empty_node(second):
        parent.remove(prev)
        replace_children(node, [prev, first])
        return True

    return False


def fix_ternary_script(parent: ET.Element, index: int, node: ET.Element) -> bool:
    if index <= 0:
        return False

    ensure_child_count(node, 3)
    first, second, third = list(node)[:3]
    prev = list(parent)[index - 1]

    if is_empty_node(first) and (not is_empty_node(second) or not is_empty_node(third)):
        parent.remove(prev)
        replace_children(node, [prev, second, third])
        return True

    return False


def normalize_scripts(parent: ET.Element):
    index = 0
    while index < len(list(parent)):
        child = list(parent)[index]
        tag = local_name(child.tag)
        fixed = False

        if tag == "msup":
            fixed = fix_binary_script(parent, index, child, "msup")
        elif tag == "msub":
            fixed = fix_binary_script(parent, index, child, "msub")
        elif tag == "msubsup":
            fixed = fix_ternary_script(parent, index, child)

        if fixed:
            index = max(0, index - 1)
        else:
            index += 1


def normalize_function_sequences(parent: ET.Element):
    index = 0
    while index < len(list(parent)):
        children = list(parent)
        child = children[index]
        tag = local_name(child.tag)
        fixed = False

        if tag in SCRIPT_TAGS and len(children) > 0:
            base = children[index][0] if len(children[index]) > 0 else None
            base_letter = get_single_ascii_letter(base)
            if base_letter is not None:
                previous_letters: list[tuple[int, str]] = []
                back = index - 1
                max_prev = max(len(name) for name in KNOWN_FUNCTIONS) - 1
                while back >= 0 and len(previous_letters) < max_prev:
                    letter = get_single_ascii_letter(children[back])
                    if letter is None:
                        break
                    previous_letters.append((back, letter))
                    back -= 1
                previous_letters.reverse()

                for function_name in sorted(KNOWN_FUNCTIONS, key=len, reverse=True):
                    needed_prev = len(function_name) - 1
                    if base_letter.lower() != function_name[-1]:
                        continue
                    if len(previous_letters) < needed_prev:
                        continue

                    candidate_letters = [letter for _, letter in previous_letters[-needed_prev:]] + [base_letter]
                    if "".join(candidate_letters).lower() != function_name:
                        continue

                    remove_indexes = [position for position, _ in previous_letters[-needed_prev:]]
                    for remove_index in reversed(remove_indexes):
                        parent.remove(children[remove_index])
                    new_base = make_text_node("mi", function_name)
                    children[index].remove(base)
                    children[index].insert(0, new_base)
                    fixed = True
                    break

        if not fixed:
            children = list(parent)
            sequence: list[tuple[int, str]] = []
            max_length = max(len(name) for name in KNOWN_FUNCTIONS)
            probe = index
            while probe < len(children) and len(sequence) < max_length:
                letter = get_single_ascii_letter(children[probe])
                if letter is None:
                    break
                sequence.append((probe, letter))
                probe += 1

            for function_name in sorted(KNOWN_FUNCTIONS, key=len, reverse=True):
                needed = len(function_name)
                if len(sequence) < needed:
                    continue
                candidate_letters = [letter for _, letter in sequence[:needed]]
                if "".join(candidate_letters).lower() != function_name:
                    continue

                new_node = make_text_node("mi", function_name)
                parent.remove(children[index])
                parent.insert(index, new_node)
                for remove_index, _ in reversed(sequence[1:needed]):
                    parent.remove(children[remove_index])
                fixed = True
                break

        if fixed:
            index = max(0, index - 1)
        else:
            index += 1


def normalize_tree(node: ET.Element):
    for child in list(node):
        normalize_tree(child)
    normalize_scripts(node)
    normalize_function_sequences(node)


def main():
    if len(sys.argv) != 2:
        print("Usage: python normalize_mathml.py <mathml.xml>")
        sys.exit(1)

    path = Path(sys.argv[1]).resolve()
    tree = ET.parse(path)
    root = tree.getroot()
    normalize_tree(root)
    tree.write(path, encoding="utf-8", xml_declaration=True)


if __name__ == "__main__":
    main()
