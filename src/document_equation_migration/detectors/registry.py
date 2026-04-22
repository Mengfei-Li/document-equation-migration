from __future__ import annotations

import importlib
import inspect
import pkgutil
from dataclasses import dataclass, field

from ..source_taxonomy import normalize_source_family
from .base import DetectorContext, FormulaDetector, FunctionDetector


@dataclass(slots=True)
class DiscoveryResult:
    loaded_modules: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


class DetectorRegistry:
    def __init__(self) -> None:
        self._detectors: dict[str, FormulaDetector] = {}
        self.run_warnings: list[str] = []

    def register(self, detector: FormulaDetector) -> None:
        key = detector.source_family.value
        if key in self._detectors:
            raise ValueError(f"Detector already registered for source family: {key}")
        self._detectors[key] = detector

    def available_source_families(self) -> list[str]:
        return sorted(self._detectors)

    def iter_detectors(self) -> list[FormulaDetector]:
        return sorted(self._detectors.values(), key=lambda item: (item.priority, item.source_family.value))

    def run(self, context: DetectorContext) -> list:
        self.run_warnings = []
        formulas = []
        for detector in self.iter_detectors():
            if not detector.supports(context.scan_result):
                continue
            try:
                formulas.extend(detector.detect(context))
            except Exception as exc:  # pragma: no cover - defensive runtime isolation
                self.run_warnings.append(f"{detector.name}: {exc}")
        return formulas

    @classmethod
    def autodiscover(cls) -> tuple["DetectorRegistry", DiscoveryResult]:
        registry = cls()
        discovery = DiscoveryResult()
        package = importlib.import_module("document_equation_migration.detectors")
        for module_info in pkgutil.iter_modules(package.__path__):
            if module_info.name in {"base", "registry"}:
                continue
            module_name = f"{package.__name__}.{module_info.name}"
            try:
                module = importlib.import_module(module_name)
                discovery.loaded_modules.append(module_name)
                registry._register_module_detectors(module)
            except Exception as exc:  # pragma: no cover - warning path
                discovery.warnings.append(f"{module_name}: {exc}")
        return registry, discovery

    def _register_module_detectors(self, module: object) -> None:
        if hasattr(module, "DETECTOR"):
            detector = getattr(module, "DETECTOR")
            if isinstance(detector, FormulaDetector):
                self.register(detector)
                return
        if hasattr(module, "build_detector"):
            detector = getattr(module, "build_detector")()
            if isinstance(detector, FormulaDetector):
                self.register(detector)
                return
        for _, member in inspect.getmembers(module):
            if isinstance(member, FormulaDetector):
                self.register(member)
                return
        module_source_family = normalize_source_family(module.__name__.split(".")[-1].replace("_", "-"))
        for name, member in inspect.getmembers(module, inspect.isfunction):
            if name.startswith("detect_"):
                self.register(
                    FunctionDetector(
                        source_family=module_source_family,
                        name=f"{module.__name__}.{name}",
                        handler=member,
                    )
                )
                return
