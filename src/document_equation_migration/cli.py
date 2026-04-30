from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import __version__
from .container_scan import scan_container
from .detectors.base import DetectorContext
from .docx_validation import validate_docx_artifact
from .detectors.registry import DetectorRegistry
from .executor import build_dry_run_execution_report, build_execution_report, load_execution_plan
from .execution_plan import build_execution_plan
from .manifest import Manifest
from .routing import build_routing_report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Detector-first formula source discovery CLI.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    scan_parser = subparsers.add_parser("scan", help="Scan a document and emit a manifest JSON.")
    scan_parser.add_argument("input", help="Path to .doc, .docx, .odt, or .fodt input.")
    scan_parser.add_argument("-o", "--output", help="Write manifest JSON to a file instead of stdout.")
    scan_parser.add_argument(
        "--routing",
        help="Write a document-level routing report JSON to a file. Use '-' to write routing JSON to stderr.",
    )
    scan_parser.add_argument(
        "--summary",
        help="Write a human-readable scan summary to a file. Use '-' to write the summary to stderr.",
    )
    scan_parser.add_argument(
        "--execution-plan",
        help="Write a document-level execution plan JSON to a file. Requires routing metadata and will auto-build it.",
    )
    scan_parser.add_argument("--indent", type=int, default=2, help="JSON indent level.")

    run_plan_parser = subparsers.add_parser(
        "run-plan",
        help="Preview an execution plan with dry-run command bindings.",
    )
    run_plan_parser.add_argument("plan", help="Path to execution-plan JSON.")
    run_plan_parser.add_argument(
        "-o",
        "--output",
        help="Write execution report JSON to a file instead of stdout.",
    )
    mode_group = run_plan_parser.add_mutually_exclusive_group(required=True)
    mode_group.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview commands without executing them.",
    )
    mode_group.add_argument(
        "--execute",
        action="store_true",
        help="Run supported execution bindings. External tools remain blocked unless explicitly allowed.",
    )
    run_plan_parser.add_argument(
        "--output-dir",
        help="Directory for execution artifacts. Defaults to ./out/<document-id> under the repository root.",
    )
    run_plan_parser.add_argument(
        "--allow-external-tools",
        action="store_true",
        help="Allow guarded external scripts such as the MathType PowerShell pipeline during --execute.",
    )
    run_plan_parser.add_argument("--indent", type=int, default=2, help="JSON indent level.")

    validate_parser = subparsers.add_parser(
        "validate-docx",
        help="Validate DOCX deliverability and optionally export PDF / run visual comparison.",
    )
    validate_parser.add_argument("target_docx", nargs="?", default="", help="Path to the target DOCX to validate.")
    validate_parser.add_argument("--output-dir", required=True, help="Directory for validation artifacts and logs.")
    validate_parser.add_argument("-o", "--output", help="Write validation report JSON to a file.")
    validate_parser.add_argument("--provider", default="", help="Provider name for the report metadata.")
    validate_parser.add_argument("--source-family", default="", help="Source family for the report metadata.")
    validate_parser.add_argument("--execution-plan", default="", help="Execution plan path for traceability.")
    validate_parser.add_argument(
        "--target-from-metadata",
        default="",
        help="Resolve the target DOCX from execution metadata / validation evidence JSON.",
    )
    validate_parser.add_argument("--reference-pdf", default="", help="Reference PDF for visual comparison.")
    validate_parser.add_argument(
        "--evidence-path",
        action="append",
        default=[],
        help="Additional evidence artifact path to record in the validation report.",
    )
    validate_parser.add_argument(
        "--allow-word-export",
        action="store_true",
        help="Run the Word COM export script to confirm the target DOCX can export PDF.",
    )
    validate_parser.add_argument(
        "--visual-compare",
        action="store_true",
        help="Run visual PDF comparison when a reference PDF is available.",
    )
    validate_parser.add_argument(
        "--visual-max-changed-ratio-per-page",
        type=float,
        default=0.02,
        help="Maximum per-page changed-ratio accepted by the shared visual gate.",
    )
    validate_parser.add_argument(
        "--visual-max-unmatched-pages",
        type=int,
        default=0,
        help="Maximum unmatched pages accepted by the shared visual gate.",
    )
    return parser


def build_summary(manifest: Manifest) -> str:
    document = manifest.document
    lines = [
        "Document Equation Migration scan summary",
        f"input: {document.input_path}",
        f"format: {document.container_format}",
        f"detector_version: {document.detector_version}",
        f"formula_count: {len(manifest.formulas)}",
    ]
    if document.source_counts:
        lines.append("source_counts:")
        for source_family, count in sorted(document.source_counts.items()):
            lines.append(f"  {source_family}: {count}")
    else:
        lines.append("source_counts: none")
    if document.notes:
        lines.append("notes:")
        for note in document.notes:
            lines.append(f"  - {note}")
    return "\n".join(lines) + "\n"


def _write_text(path_or_dash: str, payload: str, *, stderr_when_dash: bool = False) -> None:
    if path_or_dash == "-":
        stream = sys.stderr if stderr_when_dash else sys.stdout
        stream.write(payload)
        if not payload.endswith("\n"):
            stream.write("\n")
        return
    path = Path(path_or_dash)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(payload, encoding="utf-8")


def run_scan(
    input_path: str,
    output_path: str | None,
    routing_path: str | None,
    summary_path: str | None,
    execution_plan_path: str | None,
    indent: int,
) -> int:
    scan_result = scan_container(input_path)
    registry, discovery = DetectorRegistry.autodiscover()
    context = DetectorContext(scan_result=scan_result, detector_version=__version__)
    formulas = registry.run(context)
    notes = []
    if discovery.loaded_modules:
        notes.append(f"Loaded detector modules: {', '.join(discovery.loaded_modules)}")
    if discovery.warnings:
        notes.extend(f"Detector discovery warning: {item}" for item in discovery.warnings)
    if registry.run_warnings:
        notes.extend(f"Detector runtime warning: {item}" for item in registry.run_warnings)
    if not formulas:
        notes.append("Manifest currently contains only shared-core scan results.")
    manifest = Manifest.from_scan(
        scan_result=scan_result,
        detector_version=__version__,
        formulas=formulas,
        notes=notes,
    )
    payload = manifest.to_json(indent=indent)
    if output_path:
        _write_text(output_path, payload)
    else:
        sys.stdout.write(payload)
        sys.stdout.write("\n")
    routing_report = None
    if routing_path or execution_plan_path:
        routing_report = build_routing_report(manifest)
    if routing_path and routing_report is not None:
        routing_payload = json.dumps(routing_report, ensure_ascii=False, indent=indent)
        _write_text(routing_path, routing_payload, stderr_when_dash=True)
    if execution_plan_path and routing_report is not None:
        execution_plan_payload = json.dumps(
            build_execution_plan(routing_report).to_dict(),
            ensure_ascii=False,
            indent=indent,
        )
        _write_text(execution_plan_path, execution_plan_payload, stderr_when_dash=True)
    if summary_path:
        _write_text(summary_path, build_summary(manifest), stderr_when_dash=True)
    return 0


def run_plan(
    plan_path: str,
    output_path: str | None,
    dry_run: bool,
    execute: bool,
    output_dir: str | None,
    allow_external_tools: bool,
    indent: int,
) -> int:
    execution_plan = load_execution_plan(plan_path)
    resolved_plan_path = str(Path(plan_path).resolve())
    if dry_run:
        report = build_dry_run_execution_report(
            execution_plan,
            execution_plan_path=resolved_plan_path,
        )
    elif execute:
        report = build_execution_report(
            execution_plan,
            execution_plan_path=resolved_plan_path,
            output_dir=output_dir,
            allow_external_tools=allow_external_tools,
        )
    else:
        raise ValueError("Use either --dry-run or --execute with `run-plan`.")
    payload = json.dumps(report.to_dict(), ensure_ascii=False, indent=indent)
    if output_path:
        _write_text(output_path, payload)
    else:
        sys.stdout.write(payload)
        sys.stdout.write("\n")
    return 0


def run_validate_docx(
    target_docx: str,
    target_from_metadata: str,
    output_dir: str,
    output_path: str | None,
    provider: str,
    source_family: str,
    execution_plan_path: str,
    reference_pdf: str,
    evidence_paths: list[str],
    allow_word_export: bool,
    compare_visual: bool,
    visual_max_changed_ratio_per_page: float,
    visual_max_unmatched_pages: int,
) -> int:
    report = validate_docx_artifact(
        target_docx=target_docx,
        target_from_metadata=target_from_metadata,
        output_dir=output_dir,
        provider=provider,
        source_family=source_family,
        execution_plan_path=execution_plan_path,
        reference_pdf=reference_pdf,
        evidence_paths=tuple(evidence_paths),
        allow_word_export=allow_word_export,
        compare_visual=compare_visual,
        visual_max_changed_ratio_per_page=visual_max_changed_ratio_per_page,
        visual_max_unmatched_pages=visual_max_unmatched_pages,
        output_path=output_path or "",
    )
    if not output_path:
        sys.stdout.write(json.dumps(report, ensure_ascii=False, indent=2))
        sys.stdout.write("\n")
    return 0


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "scan":
        return run_scan(
            args.input,
            args.output,
            args.routing,
            args.summary,
            args.execution_plan,
            args.indent,
        )
    if args.command == "run-plan":
        return run_plan(
            args.plan,
            args.output,
            args.dry_run,
            args.execute,
            args.output_dir,
            args.allow_external_tools,
            args.indent,
        )
    if args.command == "validate-docx":
        return run_validate_docx(
            args.target_docx,
            args.target_from_metadata,
            args.output_dir,
            args.output,
            args.provider,
            args.source_family,
            args.execution_plan,
            args.reference_pdf,
            args.evidence_path,
            args.allow_word_export,
            args.visual_compare,
            args.visual_max_changed_ratio_per_page,
            args.visual_max_unmatched_pages,
        )
    raise ValueError(f"Unsupported command: {args.command}")


if __name__ == "__main__":
    raise SystemExit(main())
