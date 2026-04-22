import base64
import json
import os
import shutil
import subprocess
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest


ROOT = Path(__file__).resolve().parents[1]
FIXTURE_ROOT = Path(__file__).resolve().parent / "fixtures" / "mathtype_ole"
LIVE_CONTROL_ROOT = FIXTURE_ROOT / "live_control"
PAYLOAD_PATH = LIVE_CONTROL_ROOT / "word" / "embeddings" / "oleObject1.bin.b64"
EXPECTED_PAYLOAD_SHA256 = (
    "9f53c650efc68c5c94952892a5432a7bbc6966558a5cc7de6f7c0581ead14d4e"
)

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}


def require_external_mathtype_environment() -> dict[str, str]:
    if os.environ.get("DEM_RUN_EXTERNAL_MATHTYPE_TESTS") != "1":
        pytest.skip(
            "set DEM_RUN_EXTERNAL_MATHTYPE_TESTS=1 to run external MathType tests"
        )

    required_vars = {
        "JAVA_EXE": "java.exe",
        "JAVAC_EXE": "javac.exe",
        "MML2OMML_XSL": "MML2OMML.XSL",
        "MATHTYPE_EXTENSION_DIR": "transpect/mathtype-extension directory",
        "MATHTYPE_TO_MATHML_DIR": "jure/mathtype_to_mathml directory",
    }
    values: dict[str, str] = {}
    missing = []

    for var_name, label in required_vars.items():
        value = os.environ.get(var_name, "").strip()
        if not value:
            missing.append(f"{var_name} ({label})")
            continue
        path = Path(value)
        if not path.exists():
            missing.append(f"{var_name}={value}")
            continue
        values[var_name] = str(path)

    if missing:
        pytest.skip("missing external MathType prerequisites: " + ", ".join(missing))

    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        pytest.skip("PowerShell is required to run the MathType pipeline script")
    values["POWERSHELL"] = powershell
    return values


def build_live_control_docx(tmp_path: Path) -> Path:
    docx_path = tmp_path / "public_live_control_source.docx"
    payload = base64.b64decode(PAYLOAD_PATH.read_text(encoding="ascii"))
    files = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
</Relationships>
""",
        "word/document.xml": (LIVE_CONTROL_ROOT / "word" / "document.xml").read_text(
            encoding="utf-8"
        ),
        "word/_rels/document.xml.rels": (
            LIVE_CONTROL_ROOT / "word" / "_rels" / "document.xml.rels"
        ).read_text(encoding="utf-8"),
        "word/embeddings/oleObject1.bin": payload,
    }

    with zipfile.ZipFile(docx_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for arcname, data in files.items():
            zf.writestr(arcname, data)

    return docx_path


def count_docx_nodes(docx_path: Path) -> tuple[int, int, int]:
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    object_count = len(root.findall(".//w:object", NS))
    omath_count = len(root.findall(".//m:oMath", NS))
    omath_para_count = len(root.findall(".//m:oMathPara", NS))
    return object_count, omath_count, omath_para_count


def test_live_control_fixture_runs_external_mathtype_pipeline(tmp_path: Path) -> None:
    env = require_external_mathtype_environment()
    input_docx = build_live_control_docx(tmp_path)
    output_dir = tmp_path / "pipeline"
    output_docx = tmp_path / "public_live_control_source.omml.docx"
    script = ROOT / "run_docx_open_source_pipeline.ps1"

    completed = subprocess.run(
        [
            env["POWERSHELL"],
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
            "-InputDocx",
            str(input_docx),
            "-OutputDir",
            str(output_dir),
            "-OutputDocx",
            str(output_docx),
            "-SkipLatexPreview",
            "-JavaExe",
            env["JAVA_EXE"],
            "-JavacExe",
            env["JAVAC_EXE"],
            "-MathtypeExtensionDir",
            env["MATHTYPE_EXTENSION_DIR"],
            "-MathTypeToMathMlDir",
            env["MATHTYPE_TO_MATHML_DIR"],
            "-Mml2OmmlXsl",
            env["MML2OMML_XSL"],
        ],
        cwd=str(ROOT),
        text=True,
        capture_output=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stdout + completed.stderr
    summary_path = output_dir / "pipeline_summary.json"
    assert summary_path.exists()
    summary = json.loads(summary_path.read_text(encoding="utf-8-sig"))
    assert summary["source_ole_count"] == 1
    assert summary["attempted_count"] == 1
    assert summary["converted_ok_count"] == 1
    assert summary["converted_error_count"] == 0
    assert summary["converted_available_count"] == 1
    assert summary["replaced_count"] == 1
    assert summary["xml_counts"] == {
        "ObjectCount": 0,
        "OMathCount": 1,
        "OMathParaCount": 0,
    }

    assert output_docx.exists()
    assert count_docx_nodes(output_docx) == (0, 1, 0)
