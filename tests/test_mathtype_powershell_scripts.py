import csv
import json
import shutil
import subprocess
import zipfile
from pathlib import Path

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]
PROBE_SCRIPT = REPO_ROOT / "probe_formula_pipeline.ps1"
WRAPPER_SCRIPT = REPO_ROOT / "run_docx_open_source_pipeline.ps1"


DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8"?>
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <w:r>
        <w:object>
          <o:OLEObject r:id="rId1" ProgID="Equation.DSMT4" />
        </w:object>
      </w:r>
    </w:p>
  </w:body>
</w:document>
"""


RELS_XML = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship
      Id="rId1"
      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
      Target="embeddings/oleObject001.bin" />
</Relationships>
"""


def _powershell() -> str:
    executable = shutil.which("powershell") or shutil.which("pwsh")
    if executable is None:
        pytest.skip("PowerShell is required for script-level tests.")
    return executable


def _write_minimal_mathtype_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", DOCUMENT_XML)
        archive.writestr("word/_rels/document.xml.rels", RELS_XML)
        archive.writestr("word/embeddings/oleObject001.bin", b"fake ole payload")


def test_probe_skip_existing_uses_existing_omml_without_external_dependencies(tmp_path: Path) -> None:
    input_dir = tmp_path / "ole"
    output_dir = tmp_path / "converted"
    input_dir.mkdir()
    output_dir.mkdir()
    for index in range(1, 4):
        (input_dir / f"oleObject{index:03d}.bin").write_bytes(b"fake ole payload")
    (output_dir / "oleObject002.omml.xml").write_text("<m:oMath />", encoding="utf-8")

    completed = subprocess.run(
        [
            _powershell(),
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(PROBE_SCRIPT),
            "-InputDir",
            str(input_dir),
            "-OutputDir",
            str(output_dir),
            "-Limit",
            "0",
            "-StartIndex",
            "2",
            "-EndIndex",
            "2",
            "-SkipExisting",
            "-SkipLatexPreview",
            "-MathtypeExtensionDir",
            str(tmp_path / "missing-mathtype-extension"),
            "-MathTypeToMathMlDir",
            str(tmp_path / "missing-mathtype-to-mathml"),
            "-Mml2OmmlXsl",
            str(tmp_path / "missing-mml2omml.xsl"),
        ],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )

    assert completed.returncode == 0, completed.stdout + completed.stderr
    summary_path = output_dir / "summary.csv"
    assert summary_path.exists()

    with summary_path.open(newline="", encoding="utf-8-sig") as handle:
        rows = list(csv.DictReader(handle))
    assert rows == [
        {
            "name": "oleObject002.bin",
            "status": "ok",
            "xml_exists": "False",
            "mathml_exists": "False",
            "omml_exists": "True",
            "tex_exists": "False",
            "skipped_existing": "True",
            "latex_preview": "",
            "mathml_preview": "",
            "error": "",
        }
    ]


def test_wrapper_resume_chunk_reuses_existing_omml_and_replaces_docx(tmp_path: Path) -> None:
    input_docx = tmp_path / "input.docx"
    output_dir = tmp_path / "out"
    converted_dir = output_dir / "converted"
    output_docx = output_dir / "output.docx"
    _write_minimal_mathtype_docx(input_docx)
    converted_dir.mkdir(parents=True)
    (converted_dir / "oleObject001.omml.xml").write_text(
        '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        "<m:r><m:t>x</m:t></m:r></m:oMath>",
        encoding="utf-8",
    )

    completed = subprocess.run(
        [
            _powershell(),
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(WRAPPER_SCRIPT),
            "-InputDocx",
            str(input_docx),
            "-OutputDir",
            str(output_dir),
            "-OutputDocx",
            str(output_docx),
            "-Resume",
            "-StartIndex",
            "1",
            "-EndIndex",
            "1",
            "-SkipLatexPreview",
        ],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )

    assert completed.returncode == 0, completed.stdout + completed.stderr
    assert output_docx.exists()
    summary = json.loads((output_dir / "pipeline_summary.json").read_text(encoding="utf-8-sig"))
    assert summary["resume"] is True
    assert summary["start_index"] == 1
    assert summary["end_index"] == 1
    assert summary["attempted_count"] == 1
    assert summary["converted_ok_count"] == 1
    assert summary["converted_available_count"] == 1
    assert summary["replaced_count"] == 1
    assert summary["xml_counts"]["ObjectCount"] == 0
    assert summary["xml_counts"]["OMathCount"] == 1

    with (converted_dir / "summary.csv").open(newline="", encoding="utf-8-sig") as handle:
        rows = list(csv.DictReader(handle))
    assert rows[0]["skipped_existing"] == "True"
