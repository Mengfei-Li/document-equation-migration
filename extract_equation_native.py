import json
import re
import sys
import zipfile
from pathlib import Path

import olefile


sys.stdout.reconfigure(encoding="utf-8")


BASE_DIR = Path(__file__).resolve().parent
SAMPLE_DIR = BASE_DIR / "sample_unzipped"
OUTPUT_DIR = BASE_DIR / "extracted_equation_native"
INDEX_JSON = OUTPUT_DIR / "index.json"


def slugify(text: str) -> str:
    text = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", text)
    return text.strip("_") or "document"


def ensure_parent(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)


def extract_streams(docx_path: Path):
    doc_slug = slugify(docx_path.stem)
    doc_out_dir = OUTPUT_DIR / doc_slug
    doc_out_dir.mkdir(parents=True, exist_ok=True)
    index_rows = []

    with zipfile.ZipFile(docx_path) as zf:
        members = sorted(
            name for name in zf.namelist() if name.startswith("word/embeddings/oleObject")
        )
        for member_name in members:
            bin_bytes = zf.read(member_name)
            member_slug = Path(member_name).stem
            bin_path = doc_out_dir / f"{member_slug}.bin"
            bin_path.write_bytes(bin_bytes)

            row = {
                "docx": str(docx_path),
                "ole_member": member_name,
                "bin_path": str(bin_path),
                "is_ole": False,
                "streams": [],
            }

            if olefile.isOleFile(bin_bytes):
                row["is_ole"] = True
                ole = olefile.OleFileIO(bin_bytes)
                try:
                    for stream_path in ole.listdir():
                        stream_name = "/".join(stream_path)
                        stream_bytes = ole.openstream(stream_path).read()
                        stream_ext = ".bin"
                        if stream_name == "Equation Native":
                            stream_ext = ".eqn"
                        stream_path_out = doc_out_dir / f"{member_slug}__{slugify(stream_name)}{stream_ext}"
                        ensure_parent(stream_path_out)
                        stream_path_out.write_bytes(stream_bytes)
                        row["streams"].append(
                            {
                                "name": stream_name,
                                "size_bytes": len(stream_bytes),
                                "path": str(stream_path_out),
                            }
                        )
                finally:
                    ole.close()

            index_rows.append(row)

    return index_rows


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    all_rows = []
    for docx_path in sorted(SAMPLE_DIR.glob("*.docx")):
        all_rows.extend(extract_streams(docx_path))
    INDEX_JSON.write_text(json.dumps(all_rows, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
