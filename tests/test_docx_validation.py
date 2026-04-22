import json
import subprocess
import zipfile
from pathlib import Path

import document_equation_migration.docx_validation as validation_module
from document_equation_migration.docx_validation import validate_docx_artifact


DOCX_XML = b"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>
    </w:p>
  </w:body>
</w:document>
"""


def make_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("word/document.xml", DOCX_XML)


def test_validate_docx_artifact_reports_research_only_without_word_export(tmp_path: Path) -> None:
    target_docx = tmp_path / "sample.docx"
    output_dir = tmp_path / "validation"
    make_docx(target_docx)

    report = validate_docx_artifact(
        target_docx=str(target_docx),
        output_dir=str(output_dir),
        provider="omml",
        source_family="omml-native",
    )

    assert report["conclusion"] == "research-only"
    assert report["checks"]["target_docx"]["status"] == "passed"
    assert report["checks"]["word_export"]["status"] == "skipped"
    assert report["checks"]["visual_compare"]["status"] == "skipped"
    assert Path(report["report_json_path"]).exists()
    assert Path(report["report_txt_path"]).exists()
    assert report["target_resolution_source"] == "direct"


def test_validate_docx_artifact_resolves_target_from_metadata(tmp_path: Path) -> None:
    target_docx = tmp_path / "sample.docx"
    output_dir = tmp_path / "validation"
    metadata_path = tmp_path / "execution-metadata.json"
    make_docx(target_docx)
    metadata_path.write_text(
        json.dumps({"validation_target_docx": str(target_docx)}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    report = validate_docx_artifact(
        target_docx="",
        target_from_metadata=str(metadata_path),
        output_dir=str(output_dir),
        provider="omml",
        source_family="omml-native",
    )

    assert report["checks"]["target_docx"]["status"] == "passed"
    assert report["target_docx"] == str(target_docx.resolve())
    assert report["target_resolution_source"] == str(metadata_path.resolve())


def test_validate_docx_artifact_runs_word_export_and_visual_compare(tmp_path: Path, monkeypatch) -> None:
    target_docx = tmp_path / "sample.docx"
    reference_pdf = tmp_path / "reference.pdf"
    output_dir = tmp_path / "validation"
    make_docx(target_docx)
    reference_pdf.write_bytes(b"%PDF-1.4\n")

    def fake_run(argv, *, cwd, stdout, stderr, check, text=False):
        argv = tuple(argv)
        if argv[0] == "powershell":
            output_pdf = Path(argv[argv.index("-OutputPdf") + 1])
            output_pdf.parent.mkdir(parents=True, exist_ok=True)
            output_pdf.write_bytes(b"%PDF-1.4\n")
            return subprocess.CompletedProcess(argv, 0, stdout="word ok\n", stderr="")

        compare_dir = Path(argv[-1])
        compare_dir.mkdir(parents=True, exist_ok=True)
        (compare_dir / "visual_compare_summary.json").write_text(
            json.dumps(
                {
                    "page_count_original": 1,
                    "page_count_converted": 1,
                    "page_count_compared": 1,
                    "unmatched_original_pages": 0,
                    "unmatched_converted_pages": 0,
                    "max_changed_ratio": 0.01,
                    "average_changed_ratio": 0.01,
                    "pages": [{"page": 1, "changed_ratio": 0.01}],
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        (compare_dir / "visual_compare_summary.txt").write_text(
            "page_count_compared=1\n",
            encoding="utf-8",
        )
        return subprocess.CompletedProcess(argv, 0, stdout="compare ok\n", stderr="")

    monkeypatch.setattr("document_equation_migration.docx_validation.subprocess.run", fake_run)

    report = validate_docx_artifact(
        target_docx=str(target_docx),
        output_dir=str(output_dir),
        provider="omml",
        source_family="omml-native",
        reference_pdf=str(reference_pdf),
        allow_word_export=True,
        compare_visual=True,
    )

    assert report["conclusion"] == "deliverable-ready"
    assert report["checks"]["word_export"]["status"] == "passed"
    assert Path(report["checks"]["word_export"]["output_pdf"]).exists()
    assert report["checks"]["visual_compare"]["status"] == "passed"
    assert Path(report["checks"]["visual_compare"]["summary_json_path"]).exists()
    assert Path(report["checks"]["visual_compare"]["summary_txt_path"]).exists()
    assert report["checks"]["visual_compare"]["gate"]["reasons"] == []


def test_validate_docx_artifact_review_gates_when_visual_thresholds_fail(tmp_path: Path, monkeypatch) -> None:
    target_docx = tmp_path / "sample.docx"
    reference_pdf = tmp_path / "reference.pdf"
    output_dir = tmp_path / "validation"
    make_docx(target_docx)
    reference_pdf.write_bytes(b"%PDF-1.4\n")

    def fake_run(argv, *, cwd, stdout, stderr, check, text=False):
        argv = tuple(argv)
        if argv[0] == "powershell":
            output_pdf = Path(argv[argv.index("-OutputPdf") + 1])
            output_pdf.parent.mkdir(parents=True, exist_ok=True)
            output_pdf.write_bytes(b"%PDF-1.4\n")
            return subprocess.CompletedProcess(argv, 0, stdout="word ok\n", stderr="")

        compare_dir = Path(argv[-1])
        compare_dir.mkdir(parents=True, exist_ok=True)
        (compare_dir / "visual_compare_summary.json").write_text(
            json.dumps(
                {
                    "page_count_original": 5,
                    "page_count_converted": 4,
                    "page_count_compared": 4,
                    "unmatched_original_pages": 1,
                    "unmatched_converted_pages": 0,
                    "max_changed_ratio": 0.08,
                    "average_changed_ratio": 0.06,
                    "pages": [{"page": 1, "changed_ratio": 0.08}],
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        (compare_dir / "visual_compare_summary.txt").write_text(
            "page_count_compared=4\n",
            encoding="utf-8",
        )
        return subprocess.CompletedProcess(argv, 0, stdout="compare ok\n", stderr="")

    monkeypatch.setattr("document_equation_migration.docx_validation.subprocess.run", fake_run)

    report = validate_docx_artifact(
        target_docx=str(target_docx),
        output_dir=str(output_dir),
        provider="mathtype",
        source_family="mathtype-ole",
        reference_pdf=str(reference_pdf),
        allow_word_export=True,
        compare_visual=True,
    )

    assert report["conclusion"] == "review-gated"
    assert report["checks"]["word_export"]["status"] == "passed"
    assert report["checks"]["visual_compare"]["status"] == "review-gated"
    assert report["checks"]["visual_compare"]["summary_metrics"]["unmatched_pages_total"] == 1
    assert len(report["checks"]["visual_compare"]["gate"]["reasons"]) == 2
    assert any("Visual compare gate:" in item for item in report["residual_risks"])


def test_validate_docx_artifact_shortens_long_word_export_paths(tmp_path: Path, monkeypatch) -> None:
    target_docx = tmp_path / "sample.docx"
    output_dir = tmp_path / "validation"
    make_docx(target_docx)
    captured: dict[str, object] = {}

    def fake_run(argv, *, cwd, stdout, stderr, check, text=False):
        argv = tuple(argv)
        output_pdf = Path(argv[argv.index("-OutputPdf") + 1])
        output_pdf.parent.mkdir(parents=True, exist_ok=True)
        output_pdf.write_bytes(b"%PDF-1.4\n")
        captured["argv"] = argv
        return subprocess.CompletedProcess(argv, 0, stdout="word ok\n", stderr="")

    monkeypatch.setattr(validation_module, "_WORD_COM_SAFE_PATH_LIMIT", 10)
    monkeypatch.setattr("document_equation_migration.docx_validation.subprocess.run", fake_run)

    report = validate_docx_artifact(
        target_docx=str(target_docx),
        output_dir=str(output_dir),
        provider="omml",
        source_family="omml-native",
        allow_word_export=True,
    )

    argv = captured["argv"]
    input_docx = Path(argv[argv.index("-InputDocx") + 1])
    output_pdf = Path(argv[argv.index("-OutputPdf") + 1])
    assert input_docx.name.endswith(".docx")
    assert input_docx.name != target_docx.name
    assert "_word-path-safe" in str(input_docx)
    assert output_pdf.name.endswith(".pdf")
    assert report["checks"]["word_export"]["status"] == "passed"
    assert report["checks"]["word_export"]["staged_input_docx"] == str(input_docx)
    assert any("short staged path" in item for item in report["checks"]["word_export"]["notes"])
