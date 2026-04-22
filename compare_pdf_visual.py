import json
import math
import sys
from pathlib import Path

import fitz
from PIL import Image, ImageChops, ImageDraw, ImageStat


sys.stdout.reconfigure(encoding="utf-8")


def ensure_dir(path: Path):
    path.mkdir(parents=True, exist_ok=True)


def render_pdf(pdf_path: Path, output_dir: Path, dpi: int):
    ensure_dir(output_dir)
    document = fitz.open(pdf_path)
    page_paths = []
    try:
        scale = dpi / 72.0
        matrix = fitz.Matrix(scale, scale)
        for index, page in enumerate(document):
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            page_path = output_dir / f"page_{index + 1:03d}.png"
            pixmap.save(page_path)
            page_paths.append(page_path)
    finally:
        document.close()
    return page_paths


def open_rgb(path: Path) -> Image.Image:
    return Image.open(path).convert("RGB")


def pad_to_match(left: Image.Image, right: Image.Image):
    width = max(left.width, right.width)
    height = max(left.height, right.height)

    if left.size != (width, height):
        canvas = Image.new("RGB", (width, height), "white")
        canvas.paste(left, (0, 0))
        left = canvas

    if right.size != (width, height):
        canvas = Image.new("RGB", (width, height), "white")
        canvas.paste(right, (0, 0))
        right = canvas

    return left, right


def build_diff_mask(diff_image: Image.Image, threshold: int = 24):
    grayscale = diff_image.convert("L")
    return grayscale.point(lambda value: 255 if value > threshold else 0)


def bbox_to_dict(bbox):
    if bbox is None:
        return None
    return {"left": bbox[0], "top": bbox[1], "right": bbox[2], "bottom": bbox[3]}


def make_highlight(base: Image.Image, mask: Image.Image):
    base = base.copy()
    overlay = Image.new("RGBA", base.size, (255, 0, 0, 0))
    overlay_draw = ImageDraw.Draw(overlay)
    bbox = mask.getbbox()
    if bbox is not None:
        overlay_draw.rectangle(bbox, outline=(255, 0, 0, 220), width=6)
    tinted = Image.new("RGBA", base.size, (255, 0, 0, 90))
    overlay.paste(tinted, mask=mask)
    return Image.alpha_composite(base.convert("RGBA"), overlay).convert("RGB")


def add_label(image: Image.Image, label: str):
    labeled = Image.new("RGB", (image.width, image.height + 36), "white")
    labeled.paste(image, (0, 36))
    draw = ImageDraw.Draw(labeled)
    draw.text((12, 10), label, fill="black")
    return labeled


def stack_horizontal(images):
    gap = 20
    width = sum(image.width for image in images) + gap * (len(images) - 1)
    height = max(image.height for image in images)
    canvas = Image.new("RGB", (width, height), "white")
    x = 0
    for image in images:
        canvas.paste(image, (x, 0))
        x += image.width + gap
    return canvas


def stack_vertical(images, max_width: int | None = None):
    gap = 24
    resized = []
    for image in images:
        if max_width is not None and image.width > max_width:
            ratio = max_width / image.width
            resized.append(image.resize((max_width, int(image.height * ratio)), Image.LANCZOS))
        else:
            resized.append(image)

    width = max(image.width for image in resized)
    height = sum(image.height for image in resized) + gap * (len(resized) - 1)
    canvas = Image.new("RGB", (width, height), "white")
    y = 0
    for image in resized:
        canvas.paste(image, (0, y))
        y += image.height + gap
    return canvas


def compare_page(original_path: Path, converted_path: Path, page_index: int, output_dir: Path):
    original = open_rgb(original_path)
    converted = open_rgb(converted_path)
    original, converted = pad_to_match(original, converted)

    diff = ImageChops.difference(original, converted)
    diff_mask = build_diff_mask(diff)
    diff_bbox = diff_mask.getbbox()
    changed_pixels = sum(1 for value in diff_mask.getdata() if value)
    total_pixels = diff_mask.width * diff_mask.height
    changed_ratio = changed_pixels / total_pixels if total_pixels else 0.0

    stat = ImageStat.Stat(diff)
    mean_abs_diff = sum(stat.mean) / len(stat.mean)
    rms_diff = math.sqrt(sum(value * value for value in stat.rms) / len(stat.rms))

    diff_enhanced = make_highlight(converted, diff_mask)
    labeled = stack_horizontal(
        [
            add_label(original, f"Original P{page_index}"),
            add_label(converted, f"Converted P{page_index}"),
            add_label(diff_enhanced, f"Diff P{page_index}"),
        ]
    )

    compare_path = output_dir / "compare" / f"page_{page_index:03d}_compare.png"
    diff_path = output_dir / "diff" / f"page_{page_index:03d}_diff.png"
    ensure_dir(compare_path.parent)
    ensure_dir(diff_path.parent)
    labeled.save(compare_path)
    diff_enhanced.save(diff_path)

    return {
        "page": page_index,
        "original_image": str(original_path),
        "converted_image": str(converted_path),
        "compare_image": str(compare_path),
        "diff_image": str(diff_path),
        "size": {"width": original.width, "height": original.height},
        "changed_pixels": changed_pixels,
        "changed_ratio": changed_ratio,
        "mean_abs_diff": mean_abs_diff,
        "rms_diff": rms_diff,
        "diff_bbox": bbox_to_dict(diff_bbox),
    }


def write_summary(report, output_dir: Path):
    json_path = output_dir / "visual_compare_summary.json"
    txt_path = output_dir / "visual_compare_summary.txt"

    json_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    lines = [
        f"original_pdf={report['original_pdf']}",
        f"converted_pdf={report['converted_pdf']}",
        f"page_count_original={report['page_count_original']}",
        f"page_count_converted={report['page_count_converted']}",
        f"page_count_compared={report['page_count_compared']}",
        f"unmatched_original_pages={report['unmatched_original_pages']}",
        f"unmatched_converted_pages={report['unmatched_converted_pages']}",
        f"max_changed_ratio={report['max_changed_ratio']:.6f}",
        f"average_changed_ratio={report['average_changed_ratio']:.6f}",
        f"max_mean_abs_diff={report['max_mean_abs_diff']:.6f}",
        f"max_rms_diff={report['max_rms_diff']:.6f}",
        "",
    ]

    for item in report["pages"]:
        lines.extend(
            [
                f"[page_{item['page']:03d}]",
                f"changed_pixels={item['changed_pixels']}",
                f"changed_ratio={item['changed_ratio']:.6f}",
                f"mean_abs_diff={item['mean_abs_diff']:.6f}",
                f"rms_diff={item['rms_diff']:.6f}",
                f"diff_bbox={item['diff_bbox']}",
                f"compare_image={item['compare_image']}",
                f"diff_image={item['diff_image']}",
                "",
            ]
        )

    txt_path.write_text("\n".join(lines), encoding="utf-8")
    return json_path, txt_path


def build_contact_sheet(compare_dir: Path, output_dir: Path):
    compare_images = sorted(compare_dir.glob("page_*_compare.png"))
    if not compare_images:
        return None
    sheet = stack_vertical([open_rgb(path) for path in compare_images], max_width=1800)
    contact_path = output_dir / "all_pages_compare.png"
    sheet.save(contact_path)
    return contact_path


def build_labeled_page_sheet(page_paths: list[Path], output_path: Path, label_prefix: str):
    if not page_paths:
        return None
    images = [add_label(open_rgb(path), f"{label_prefix} P{index + 1}") for index, path in enumerate(page_paths)]
    sheet = stack_vertical(images, max_width=1400)
    sheet.save(output_path)
    return output_path


def main():
    if len(sys.argv) != 4:
        print("Usage: python compare_pdf_visual.py <original.pdf> <converted.pdf> <output_dir>")
        sys.exit(1)

    original_pdf = Path(sys.argv[1]).resolve()
    converted_pdf = Path(sys.argv[2]).resolve()
    output_dir = Path(sys.argv[3]).resolve()

    ensure_dir(output_dir)
    original_pages_dir = output_dir / "pages" / "original"
    converted_pages_dir = output_dir / "pages" / "converted"

    original_pages = render_pdf(original_pdf, original_pages_dir, dpi=180)
    converted_pages = render_pdf(converted_pdf, converted_pages_dir, dpi=180)

    page_count = min(len(original_pages), len(converted_pages))
    page_reports = []
    for index in range(page_count):
        page_reports.append(compare_page(original_pages[index], converted_pages[index], index + 1, output_dir))

    changed_ratios = [item["changed_ratio"] for item in page_reports]
    mean_abs_diffs = [item["mean_abs_diff"] for item in page_reports]
    rms_diffs = [item["rms_diff"] for item in page_reports]

    report = {
        "original_pdf": str(original_pdf),
        "converted_pdf": str(converted_pdf),
        "page_count_original": len(original_pages),
        "page_count_converted": len(converted_pages),
        "page_count_compared": page_count,
        "unmatched_original_pages": max(0, len(original_pages) - page_count),
        "unmatched_converted_pages": max(0, len(converted_pages) - page_count),
        "max_changed_ratio": max(changed_ratios, default=0.0),
        "average_changed_ratio": sum(changed_ratios) / len(changed_ratios) if changed_ratios else 0.0,
        "max_mean_abs_diff": max(mean_abs_diffs, default=0.0),
        "max_rms_diff": max(rms_diffs, default=0.0),
        "pages": page_reports,
    }

    json_path, txt_path = write_summary(report, output_dir)
    contact_path = build_contact_sheet(output_dir / "compare", output_dir)
    original_sheet = build_labeled_page_sheet(original_pages, output_dir / "all_pages_original.png", "Original")
    converted_sheet = build_labeled_page_sheet(converted_pages, output_dir / "all_pages_converted.png", "Converted")

    print(str(json_path))
    print(str(txt_path))
    if contact_path is not None:
        print(str(contact_path))
    if original_sheet is not None:
        print(str(original_sheet))
    if converted_sheet is not None:
        print(str(converted_sheet))


if __name__ == "__main__":
    main()
