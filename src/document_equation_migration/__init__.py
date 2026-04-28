"""Shared core for detector-first formula source discovery."""

from .docx_validation import validate_docx_artifact
from .canonical_target import canonical_mathml_contract_for_source_family
from .executor import build_dry_run_execution_report, build_execution_report, load_execution_plan
from .execution_plan import build_execution_plan
from .manifest import Manifest
from .routing import build_execution_plan_report, build_routing_report
from .source_taxonomy import SourceFamily, SourceRole

__all__ = [
    "Manifest",
    "SourceFamily",
    "SourceRole",
    "build_execution_plan",
    "build_dry_run_execution_report",
    "build_execution_report",
    "build_routing_report",
    "build_execution_plan_report",
    "canonical_mathml_contract_for_source_family",
    "load_execution_plan",
    "validate_docx_artifact",
]

__version__ = "0.1.0"
