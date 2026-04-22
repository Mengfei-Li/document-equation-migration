from __future__ import annotations

from .axmath import build_axmath_dry_run_reports, execute_axmath_step
from .base import DryRunBindingBuilder, ExecutionBindingRunner
from .equation3 import build_equation3_dry_run_reports, execute_equation3_step
from .mathtype import build_mathtype_dry_run_reports, execute_mathtype_step
from .odf import build_odf_dry_run_reports, execute_odf_step
from .omml import build_omml_dry_run_reports, execute_omml_step

DRY_RUN_BINDING_REGISTRY: dict[str, DryRunBindingBuilder] = {
    "axmath": build_axmath_dry_run_reports,
    "equation3": build_equation3_dry_run_reports,
    "mathtype": build_mathtype_dry_run_reports,
    "odf": build_odf_dry_run_reports,
    "omml": build_omml_dry_run_reports,
}

EXECUTION_BINDING_REGISTRY: dict[str, ExecutionBindingRunner] = {
    "axmath": execute_axmath_step,
    "equation3": execute_equation3_step,
    "mathtype": execute_mathtype_step,
    "odf": execute_odf_step,
    "omml": execute_omml_step,
}
