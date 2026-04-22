import json
import sys
import zipfile
from pathlib import Path

import olefile


sys.stdout.reconfigure(encoding="utf-8")


BASE_DIR = Path(__file__).resolve().parent
SAMPLE_DIR = BASE_DIR / "sample_unzipped"
OUTPUT_TXT = BASE_DIR / "ole_streams_report.txt"
OUTPUT_JSON = BASE_DIR / "ole_streams_report.json"


def inspect_bin(docx_path: Path, member_name: str):
    with zipfile.ZipFile(docx_path) as zf:
        data = zf.read(member_name)

    result = {
        "docx": str(docx_path),
        "member_name": member_name,
        "size_bytes": len(data),
        "is_ole": olefile.isOleFile(data),
        "streams": [],
    }

    if not result["is_ole"]:
        return result

    ole = olefile.OleFileIO(data)
    try:
        for stream_path in ole.listdir():
            joined = "/".join(stream_path)
            stream_data = ole.openstream(stream_path).read()
            result["streams"].append(
                {
                    "name": joined,
                    "size_bytes": len(stream_data),
                    "ascii_preview": "".join(
                        chr(b) if 32 <= b <= 126 else "."
                        for b in stream_data[:96]
                    ),
                    "hex_prefix": stream_data[:32].hex(),
                }
            )
    finally:
        ole.close()

    return result


def main():
    reports = []
    for docx_path in sorted(SAMPLE_DIR.glob("*.docx")):
        with zipfile.ZipFile(docx_path) as zf:
            names = sorted(
                name for name in zf.namelist() if name.startswith("word/embeddings/oleObject")
            )
        for member_name in names[:5]:
            reports.append(inspect_bin(docx_path, member_name))

    OUTPUT_JSON.write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")

    lines = []
    for report in reports:
        lines.append(f"Document: {report['docx']}")
        lines.append(f"Object: {report['member_name']}")
        lines.append(f"Size: {report['size_bytes']}")
        lines.append(f"Is OLE: {report['is_ole']}")
        for stream in report["streams"]:
            lines.append(
                f"  - Stream: {stream['name']} | Size: {stream['size_bytes']} | HEX prefix: {stream['hex_prefix']} | Preview: {stream['ascii_preview']}"
            )
        lines.append("")

    OUTPUT_TXT.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    main()
