import csv
import json
import re
import sys
from pathlib import Path


sys.stdout.reconfigure(encoding="utf-8")


MANUAL_PATTERNS = [
    (re.compile(r"\\frac\{\}\{"), "empty_numerator"),
    (re.compile(r"\\frac\{[^}]*\}\{\}"), "empty_denominator"),
    (re.compile(r"\{\}\^\{\}"), "empty_superscript"),
    (re.compile(r"\^\{\}"), "empty_superscript"),
    (re.compile(r"_\{\}"), "empty_subscript"),
    (re.compile(r"\\left\{\s*\\right\."), "suspicious_empty_piecewise"),
    (re.compile(r"\\left\{\s*\{\s*\}\s*\\right\."), "empty_piecewise_body"),
    (re.compile(r"[=+\-*/]\s*\)"), "operator_missing_rhs"),
    (re.compile(r"\(\s*[+\-*/=]"), "operator_missing_lhs"),
    (re.compile(r"(?:=|\+|-|/|\\leq|\\geq|<|>|//)\s*(?:\\\]|\\\)|$)"), "trailing_operator_or_relation"),
]

SPOT_CHECK_TOKENS = [
    "\\left",
    "\\right",
    "\\overrightarrow",
    "\\bigtriangleup",
    "\\cup",
    "\\cap",
    "\\subset",
    "\\lbrack",
    "\\rbrack",
    "\\alpha",
    "\\beta",
    "\\Omega",
    "\\log",
    "\\sin",
    "\\cos",
    "\\tan",
    "\\parallel",
    "\\perp",
    "\\sum",
    "\\prod",
    "\\int",
]


def load_csv(path: Path):
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def load_map(path: Path):
    data = json.loads(path.read_text(encoding="utf-8"))
    result = {}
    for item in data.get("items", []):
        ole_target = item.get("ole_target", "")
        if not ole_target:
            continue
        stem = Path(ole_target).stem
        result[stem] = item
    return result


def has_blank_placeholder(text: str) -> bool:
    stripped = (text or "").strip()
    return bool(re.match(r"^_{3,}[.]*$", stripped))


def classify_row(row: dict, map_item: dict | None = None):
    preview = (row.get("latex_preview") or "").strip()
    reasons = []
    map_item = map_item or {}

    if row.get("status") != "ok":
        reasons.append("status_not_ok")
    if row.get("omml_exists") != "True":
        reasons.append("missing_omml")
    if row.get("tex_exists") != "True":
        reasons.append("missing_latex")
    if not preview:
        reasons.append("empty_latex_preview")

    for pattern, label in MANUAL_PATTERNS:
        if pattern.search(preview):
            reasons.append(label)

    normalized = preview.replace(" ", "")
    if "{}^{}" in normalized:
        reasons.append("empty_superscript")
    if "_{}" in normalized:
        reasons.append("empty_subscript")
    if "x+)" in normalized or "x-)" in normalized:
        reasons.append("missing_operand_inside_parentheses")
    if "\\left\\{{}\\right." in normalized:
        reasons.append("empty_piecewise_body")

    deduped_reasons = sorted(set(reasons))
    if deduped_reasons:
        tail_operator = re.search(r"(?:=|\+|-|/|\\leq|\\geq|<|>|//)\s*(?:\\\]|\\\)|$)", preview)
        if tail_operator and has_blank_placeholder(map_item.get("text_after", "")):
            return "spot_check", ["formula_followed_by_blank"]
        return "manual_review", deduped_reasons

    spot_check_reasons = []
    if len(preview) >= 28:
        spot_check_reasons.append("long_formula")
    if any(token in preview for token in SPOT_CHECK_TOKENS):
        spot_check_reasons.append("has_complex_commands")
    if "_" in preview or "^" in preview:
        spot_check_reasons.append("has_scripts")
    if "|" in preview or "//" in preview:
        spot_check_reasons.append("has_set_or_geometry_relation")

    if spot_check_reasons:
        return "spot_check", sorted(set(spot_check_reasons))

    if preview.startswith("{") and preview.endswith("}"):
        return "spot_check", ["grouped_output_needs_review"]

    return "auto_replace", []


def build_record(row: dict, map_item: dict):
    name = row.get("name", "")
    stem = Path(name).stem
    category, reasons = classify_row(row, map_item)
    return {
        "name": name,
        "stem": stem,
        "category": category,
        "reasons": reasons,
        "latex_preview": row.get("latex_preview", ""),
        "paragraph_index": map_item.get("paragraph_index"),
        "run_index": map_item.get("run_index"),
        "text_before": map_item.get("text_before", ""),
        "text_after": map_item.get("text_after", ""),
        "paragraph_text": map_item.get("paragraph_text", ""),
    }


def write_summary_text(output_path: Path, counts: dict, manual_items: list, spot_items: list):
    lines = [
        f"auto_replace={counts['auto_replace']}",
        f"spot_check={counts['spot_check']}",
        f"manual_review={counts['manual_review']}",
        "",
        "[manual_reason_counts]",
    ]

    for reason, count in counts["manual_reason_counts"].items():
        lines.append(f"{reason}={count}")

    lines.extend(["", "[manual_review]"])

    for item in manual_items:
        lines.extend(
            [
                f"- {item['name']}",
                f"  category={item['category']}",
                f"  reasons={'; '.join(item['reasons'])}",
                f"  paragraph_index={item['paragraph_index']}",
                f"  latex_preview={item['latex_preview']}",
                f"  text_before={item['text_before']}",
                f"  text_after={item['text_after']}",
            ]
        )

    lines.append("")
    lines.append("[spot_check_top30]")
    for item in spot_items[:30]:
        lines.extend(
            [
                f"- {item['name']}",
                f"  reasons={'; '.join(item['reasons'])}",
                f"  paragraph_index={item['paragraph_index']}",
                f"  latex_preview={item['latex_preview']}",
            ]
        )

    output_path.write_text("\n".join(lines), encoding="utf-8")


def main():
    if len(sys.argv) != 5:
        print(
            "Usage: python analyze_formula_risks.py <summary.csv> <original_ole_map.json> <output.json> <output.txt>"
        )
        sys.exit(1)

    csv_path = Path(sys.argv[1]).resolve()
    map_path = Path(sys.argv[2]).resolve()
    output_json = Path(sys.argv[3]).resolve()
    output_txt = Path(sys.argv[4]).resolve()

    rows = load_csv(csv_path)
    ole_map = load_map(map_path)

    records = []
    for row in rows:
        stem = Path(row.get("name", "")).stem
        map_item = ole_map.get(stem, {})
        records.append(build_record(row, map_item))

    counts = {
        "auto_replace": sum(1 for item in records if item["category"] == "auto_replace"),
        "spot_check": sum(1 for item in records if item["category"] == "spot_check"),
        "manual_review": sum(1 for item in records if item["category"] == "manual_review"),
    }
    reason_counts = {}
    for item in records:
        if item["category"] != "manual_review":
            continue
        for reason in item["reasons"]:
            reason_counts[reason] = reason_counts.get(reason, 0) + 1
    counts["manual_reason_counts"] = dict(sorted(reason_counts.items(), key=lambda pair: (-pair[1], pair[0])))

    records.sort(key=lambda item: int(re.search(r"\d+", item["stem"]).group()) if re.search(r"\d+", item["stem"]) else 10**9)
    manual_items = [item for item in records if item["category"] == "manual_review"]
    spot_items = [item for item in records if item["category"] == "spot_check"]

    output = {
        "summary_csv": str(csv_path),
        "original_ole_map": str(map_path),
        "counts": counts,
        "manual_review": manual_items,
        "spot_check": spot_items,
        "auto_replace": [item for item in records if item["category"] == "auto_replace"],
    }
    output_json.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    write_summary_text(output_txt, counts, manual_items, spot_items)


if __name__ == "__main__":
    main()
