from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

DEFAULT_VISUAL_MAX_CHANGED_RATIO_PER_PAGE = 0.02
DEFAULT_VISUAL_MAX_UNMATCHED_PAGES = 0
_WORD_COM_SAFE_PATH_LIMIT = 240


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _write_json(path: Path, payload: dict[str, object]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _write_text(path: Path, payload: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(payload, encoding="utf-8")
    return path


def _read_json(path: Path) -> dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def _resolve_target_docx(
    *,
    target_docx: str,
    target_from_metadata: str,
) -> tuple[Path, str]:
    if target_docx:
        return Path(target_docx).resolve(), "direct"
    metadata_path = Path(target_from_metadata).resolve()
    payload = _read_json(metadata_path)
    candidates = [
        payload.get("validation_target_docx"),
        payload.get("target_docx"),
    ]
    artifacts = payload.get("artifacts")
    if isinstance(artifacts, dict):
        validation_target = artifacts.get("validation_target")
        if isinstance(validation_target, dict):
            candidates.append(validation_target.get("validation_target_docx"))
    for candidate in candidates:
        if candidate:
            return Path(str(candidate)).resolve(), str(metadata_path)
    raise ValueError(f"No validation target DOCX path found in metadata: {metadata_path}")


def _decode_output(payload: bytes | str | None) -> str:
    if payload is None:
        return ""
    if isinstance(payload, str):
        return payload
    for encoding in ("utf-8", "gbk", "cp936"):
        try:
            return payload.decode(encoding)
        except UnicodeDecodeError:
            continue
    return payload.decode("utf-8", errors="replace")


def _run_logged_process(
    argv: tuple[str, ...],
    *,
    cwd: Path,
    stdout_path: Path,
    stderr_path: Path,
) -> subprocess.CompletedProcess[bytes]:
    completed = subprocess.run(
        argv,
        cwd=str(cwd),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    _write_text(stdout_path, _decode_output(completed.stdout))
    _write_text(stderr_path, _decode_output(completed.stderr))
    return completed


def _word_export(
    *,
    target_docx: Path,
    output_dir: Path,
) -> dict[str, object]:
    script_path = _repo_root() / "export_word_pdf.ps1"
    staged_target_docx = target_docx
    output_pdf = output_dir / "word-export" / f"{target_docx.stem}.pdf"
    notes = []
    if len(str(target_docx)) > _WORD_COM_SAFE_PATH_LIMIT or len(str(output_pdf)) > _WORD_COM_SAFE_PATH_LIMIT:
        safe_dir = output_dir / "_word-path-safe"
        safe_dir.mkdir(parents=True, exist_ok=True)
        digest = hashlib.sha1(str(target_docx).encode("utf-8")).hexdigest()[:10]
        staged_target_docx = safe_dir / f"{digest}.docx"
        shutil.copy2(target_docx, staged_target_docx)
        output_pdf = output_dir / "word-export" / f"{digest}.pdf"
        notes.append("Word export used a short staged path to avoid COM false negatives on long paths.")
    stdout_path = output_dir / "logs" / "word-export.stdout.txt"
    stderr_path = output_dir / "logs" / "word-export.stderr.txt"

    if not script_path.exists():
        return {
            "status": "failed",
            "runner": "powershell",
            "script_path": str(script_path),
            "output_pdf": str(output_pdf),
            "stdout_path": str(stdout_path),
            "stderr_path": str(stderr_path),
            "notes": ["Word export script does not exist."],
        }

    completed = _run_logged_process(
        (
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
            "-InputDocx",
            str(staged_target_docx),
            "-OutputPdf",
            str(output_pdf),
        ),
        cwd=_repo_root(),
        stdout_path=stdout_path,
        stderr_path=stderr_path,
    )

    status = "passed" if completed.returncode == 0 and output_pdf.exists() else "failed"
    if status == "passed":
        notes.append("Word opened the target DOCX and exported PDF.")
    else:
        notes.append("Word export did not complete successfully.")

    return {
        "status": status,
        "runner": "powershell",
        "script_path": str(script_path),
        "argv": [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
            "-InputDocx",
            str(staged_target_docx),
            "-OutputPdf",
            str(output_pdf),
        ],
        "output_pdf": str(output_pdf),
        "staged_input_docx": str(staged_target_docx) if staged_target_docx != target_docx else "",
        "stdout_path": str(stdout_path),
        "stderr_path": str(stderr_path),
        "exit_code": completed.returncode,
        "notes": notes,
    }


def _visual_compare(
    *,
    reference_pdf: Path,
    candidate_pdf: Path,
    output_dir: Path,
    max_changed_ratio_per_page: float,
    max_unmatched_pages: int,
) -> dict[str, object]:
    script_path = _repo_root() / "compare_pdf_visual.py"
    compare_dir = output_dir / "visual-compare"
    stdout_path = output_dir / "logs" / "visual-compare.stdout.txt"
    stderr_path = output_dir / "logs" / "visual-compare.stderr.txt"
    summary_json_path = compare_dir / "visual_compare_summary.json"
    summary_txt_path = compare_dir / "visual_compare_summary.txt"

    if not script_path.exists():
        return {
            "status": "failed",
            "runner": "python",
            "script_path": str(script_path),
            "summary_json_path": str(summary_json_path),
            "summary_txt_path": str(summary_txt_path),
            "stdout_path": str(stdout_path),
            "stderr_path": str(stderr_path),
            "notes": ["Visual compare script does not exist."],
        }

    completed = _run_logged_process(
        (
            sys.executable,
            str(script_path),
            str(reference_pdf),
            str(candidate_pdf),
            str(compare_dir),
        ),
        cwd=_repo_root(),
        stdout_path=stdout_path,
        stderr_path=stderr_path,
    )

    status = "passed" if completed.returncode == 0 and summary_json_path.exists() else "failed"
    notes = []
    if status != "passed":
        notes.append("Visual PDF comparison did not complete successfully.")
        return {
            "status": status,
            "execution_status": status,
            "runner": "python",
            "script_path": str(script_path),
            "argv": [sys.executable, str(script_path), str(reference_pdf), str(candidate_pdf), str(compare_dir)],
            "summary_json_path": str(summary_json_path),
            "summary_txt_path": str(summary_txt_path),
            "stdout_path": str(stdout_path),
            "stderr_path": str(stderr_path),
            "exit_code": completed.returncode,
            "notes": notes,
        }

    summary = _read_json(summary_json_path)
    unmatched_pages = int(summary.get("unmatched_original_pages", 0)) + int(summary.get("unmatched_converted_pages", 0))
    observed_max_changed_ratio = float(summary.get("max_changed_ratio", 0.0))
    gate_reasons: list[str] = []
    if unmatched_pages > max_unmatched_pages:
        gate_reasons.append(
            f"Unmatched page count {unmatched_pages} exceeds threshold {max_unmatched_pages}."
        )
    if observed_max_changed_ratio > max_changed_ratio_per_page:
        gate_reasons.append(
            "Max changed ratio "
            f"{observed_max_changed_ratio:.6f} exceeds threshold {max_changed_ratio_per_page:.6f}."
        )

    status = "passed" if not gate_reasons else "review-gated"
    notes.append("Visual PDF comparison summary was produced.")
    if status == "passed":
        notes.append("Visual compare met the current gate thresholds.")
    else:
        notes.append("Visual compare exceeded the current gate thresholds and needs review.")

    return {
        "status": status,
        "execution_status": "passed",
        "runner": "python",
        "script_path": str(script_path),
        "argv": [sys.executable, str(script_path), str(reference_pdf), str(candidate_pdf), str(compare_dir)],
        "summary_json_path": str(summary_json_path),
        "summary_txt_path": str(summary_txt_path),
        "stdout_path": str(stdout_path),
        "stderr_path": str(stderr_path),
        "exit_code": completed.returncode,
        "summary_metrics": {
            "page_count_original": int(summary.get("page_count_original", 0)),
            "page_count_converted": int(summary.get("page_count_converted", 0)),
            "page_count_compared": int(summary.get("page_count_compared", 0)),
            "unmatched_original_pages": int(summary.get("unmatched_original_pages", 0)),
            "unmatched_converted_pages": int(summary.get("unmatched_converted_pages", 0)),
            "unmatched_pages_total": unmatched_pages,
            "max_changed_ratio": observed_max_changed_ratio,
            "average_changed_ratio": float(summary.get("average_changed_ratio", 0.0)),
        },
        "gate": {
            "status": status,
            "max_changed_ratio_per_page_threshold": max_changed_ratio_per_page,
            "max_unmatched_pages_threshold": max_unmatched_pages,
            "reasons": gate_reasons,
        },
        "notes": notes,
    }


def _report_lines(report: dict[str, object]) -> str:
    checks = report["checks"]
    lines = [
        f"artifact_type={report['artifact_type']}",
        f"conclusion={report['conclusion']}",
        f"target_docx={report['target_docx']}",
        f"provider={report['provider']}",
        f"source_family={report['source_family']}",
        f"target_exists={checks['target_docx']['status']}",
        f"word_export={checks['word_export']['status']}",
        f"visual_compare={checks['visual_compare']['status']}",
    ]
    visual_gate = checks["visual_compare"].get("gate")
    if visual_gate:
        lines.extend(
            [
                f"visual_threshold_max_changed_ratio={visual_gate['max_changed_ratio_per_page_threshold']}",
                f"visual_threshold_max_unmatched_pages={visual_gate['max_unmatched_pages_threshold']}",
            ]
        )
    visual_metrics = checks["visual_compare"].get("summary_metrics")
    if visual_metrics:
        lines.extend(
            [
                f"visual_page_count_original={visual_metrics['page_count_original']}",
                f"visual_page_count_converted={visual_metrics['page_count_converted']}",
                f"visual_unmatched_pages_total={visual_metrics['unmatched_pages_total']}",
                f"visual_max_changed_ratio={visual_metrics['max_changed_ratio']}",
            ]
        )
    if report["reference_pdf"]:
        lines.append(f"reference_pdf={report['reference_pdf']}")
    lines.append("")
    lines.append("residual_risks:")
    for item in report["residual_risks"]:
        lines.append(f"- {item}")
    return "\n".join(lines) + "\n"


def validate_docx_artifact(
    *,
    target_docx: str,
    target_from_metadata: str = "",
    output_dir: str,
    provider: str = "",
    source_family: str = "",
    execution_plan_path: str = "",
    reference_pdf: str = "",
    evidence_paths: tuple[str, ...] = (),
    allow_word_export: bool = False,
    compare_visual: bool = False,
    visual_max_changed_ratio_per_page: float = DEFAULT_VISUAL_MAX_CHANGED_RATIO_PER_PAGE,
    visual_max_unmatched_pages: int = DEFAULT_VISUAL_MAX_UNMATCHED_PAGES,
    output_path: str = "",
) -> dict[str, object]:
    target_path, target_resolution_source = _resolve_target_docx(
        target_docx=target_docx,
        target_from_metadata=target_from_metadata,
    )
    output_root = Path(output_dir).resolve()
    report_json_path = Path(output_path).resolve() if output_path else output_root / "validation-report.json"
    report_txt_path = output_root / "validation-report.txt"

    report: dict[str, object] = {
        "artifact_type": "docx-validation-report",
        "provider": provider,
        "source_family": source_family,
        "execution_plan_path": execution_plan_path,
        "target_docx": str(target_path),
        "target_resolution_source": target_resolution_source,
        "reference_pdf": str(Path(reference_pdf).resolve()) if reference_pdf else "",
        "output_dir": str(output_root),
        "report_json_path": str(report_json_path),
        "report_txt_path": str(report_txt_path),
        "evidence_paths": [str(Path(path)) for path in evidence_paths if path],
        "visual_thresholds": {
            "max_changed_ratio_per_page": visual_max_changed_ratio_per_page,
            "max_unmatched_pages": visual_max_unmatched_pages,
        },
        "checks": {
            "target_docx": {
                "status": "passed" if target_path.exists() else "failed",
                "path": str(target_path),
            },
            "word_export": {
                "status": "skipped",
                "notes": ["Word export was not requested."],
            },
            "visual_compare": {
                "status": "skipped",
                "notes": ["Visual PDF comparison was not requested."],
            },
        },
        "residual_risks": [],
        "conclusion": "research-only",
    }

    if not target_path.exists():
        report["checks"]["word_export"] = {
            "status": "skipped",
            "notes": ["Target DOCX is missing, so Word export could not run."],
        }
        report["checks"]["visual_compare"] = {
            "status": "skipped",
            "notes": ["Target DOCX is missing, so visual comparison could not run."],
        }
        report["residual_risks"].append("Target DOCX does not exist.")
        report["conclusion"] = "blocked"
        _write_json(report_json_path, report)
        _write_text(report_txt_path, _report_lines(report))
        return report

    if allow_word_export:
        report["checks"]["word_export"] = _word_export(target_docx=target_path, output_dir=output_root)
        if report["checks"]["word_export"]["status"] != "passed":
            report["residual_risks"].append("Word PDF export failed; output is not deliverable-ready.")
            report["conclusion"] = "blocked"
        else:
            report["conclusion"] = "deliverable-ready"
    else:
        report["residual_risks"].append("Word export not run; current result is research-only.")

    if compare_visual:
        if not reference_pdf:
            report["checks"]["visual_compare"] = {
                "status": "skipped",
                "notes": ["Visual comparison was requested but no reference PDF was provided."],
            }
            report["residual_risks"].append("Visual comparison requested without reference PDF.")
        elif report["checks"]["word_export"]["status"] != "passed":
            report["checks"]["visual_compare"] = {
                "status": "skipped",
                "notes": ["Visual comparison requires a successfully exported target PDF."],
            }
        else:
            compare_result = _visual_compare(
                reference_pdf=Path(reference_pdf).resolve(),
                candidate_pdf=Path(report["checks"]["word_export"]["output_pdf"]),
                output_dir=output_root,
                max_changed_ratio_per_page=visual_max_changed_ratio_per_page,
                max_unmatched_pages=visual_max_unmatched_pages,
            )
            report["checks"]["visual_compare"] = compare_result
            if compare_result["status"] == "failed":
                report["residual_risks"].append("Visual PDF comparison failed; inspect logs before claiming visual parity.")
                report["conclusion"] = "blocked"
            elif compare_result["status"] == "review-gated":
                for reason in compare_result["gate"]["reasons"]:
                    report["residual_risks"].append(f"Visual compare gate: {reason}")
                report["conclusion"] = "review-gated"
    else:
        report["residual_risks"].append("Visual PDF comparison not run.")

    _write_json(report_json_path, report)
    _write_text(report_txt_path, _report_lines(report))
    return report
