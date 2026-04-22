from document_equation_migration.execution_plan import build_execution_plan
from document_equation_migration.execution_plan.model import ExecutionPlan


def test_build_execution_plan_uses_source_line_providers() -> None:
    routing_report = {
        "document_id": "sample",
        "input_path": "sample.docx",
        "detector_version": "0.1.0",
        "formula_count": 6,
        "recommended_sequence": [
            "mathtype-ole",
            "omml-native",
            "equation-editor-3-ole",
            "axmath-ole",
            "odf-native",
            "libreoffice-transformed",
        ],
        "route_plan": [
            {
                "source_family": "mathtype-ole",
                "formula_count": 1,
                "route_kind": "primary-source-first",
                "priority": 10,
                "next_action": "run-mathtype-source-first-pipeline",
                "confidence_policy": "high",
                "requires_manual_review": False,
            },
            {
                "source_family": "omml-native",
                "formula_count": 1,
                "route_kind": "primary-source-first",
                "priority": 20,
                "next_action": "run-omml-native-pipeline",
                "confidence_policy": "high",
                "requires_manual_review": False,
            },
            {
                "source_family": "equation-editor-3-ole",
                "formula_count": 1,
                "route_kind": "primary-candidate",
                "priority": 30,
                "next_action": "run-equation3-probe-and-conversion",
                "confidence_policy": "medium",
                "requires_manual_review": True,
            },
            {
                "source_family": "axmath-ole",
                "formula_count": 1,
                "route_kind": "export-assisted",
                "priority": 50,
                "next_action": "run-axmath-export-assisted-pipeline",
                "confidence_policy": "medium",
                "requires_manual_review": True,
            },
            {
                "source_family": "odf-native",
                "formula_count": 1,
                "route_kind": "primary-source-first",
                "priority": 40,
                "next_action": "run-odf-native-pipeline",
                "confidence_policy": "medium",
                "requires_manual_review": False,
            },
            {
                "source_family": "libreoffice-transformed",
                "formula_count": 1,
                "route_kind": "bridge-source",
                "priority": 60,
                "next_action": "run-libreoffice-bridge-review-pipeline",
                "confidence_policy": "low",
                "requires_manual_review": True,
            },
        ],
    }

    plan = build_execution_plan(routing_report).to_dict()

    assert plan["document_id"] == "sample"
    assert plan["formula_count"] == 6
    assert plan["manual_review_required"] is True
    assert len(plan["steps"]) == 6
    assert [item["provider"] for item in plan["steps"]] == [
        "mathtype",
        "omml",
        "equation3",
        "axmath",
        "odf",
        "odf",
    ]
    assert all(item["actions"] for item in plan["steps"])


def test_build_execution_plan_falls_back_for_unknown_source() -> None:
    routing_report = {
        "document_id": "unknown",
        "input_path": "unknown.docx",
        "detector_version": "0.1.0",
        "formula_count": 1,
        "recommended_sequence": ["unknown-ole"],
        "route_plan": [
            {
                "source_family": "unknown-ole",
                "formula_count": 1,
                "route_kind": "manual-classification",
                "priority": 80,
                "next_action": "manual-classification-required",
                "confidence_policy": "low",
                "requires_manual_review": True,
            }
        ],
    }

    plan = build_execution_plan(routing_report).to_dict()
    step = plan["steps"][0]

    assert step["provider"] == "default"
    assert step["requires_manual_review"] is True
    assert step["actions"][0]["action_id"] == "manual-triage"


def test_build_execution_plan_falls_back_for_unregistered_family_text() -> None:
    routing_report = {
        "document_id": "future",
        "input_path": "future.docx",
        "detector_version": "0.1.0",
        "formula_count": 1,
        "recommended_sequence": ["future-source-family"],
        "route_plan": [
            {
                "source_family": "future-source-family",
                "formula_count": 1,
                "route_kind": "experimental",
                "next_action": "investigate",
                "confidence_policy": "low",
                "requires_manual_review": True,
            }
        ],
    }

    plan = build_execution_plan(routing_report).to_dict()
    step = plan["steps"][0]

    assert step["source_family"] == "future-source-family"
    assert step["provider"] == "default"
    assert step["actions"][0]["action_id"] == "manual-triage"


def test_build_execution_plan_round_trips_mathtype_experimental_options() -> None:
    routing_report = {
        "document_id": "sample",
        "input_path": "sample.docx",
        "detector_version": "0.1.0",
        "formula_count": 1,
        "recommended_sequence": ["mathtype-ole"],
        "route_plan": [
            {
                "source_family": "mathtype-ole",
                "formula_count": 1,
                "route_kind": "primary-source-first",
                "next_action": "run-mathtype-source-first-pipeline",
                "confidence_policy": "high",
                "requires_manual_review": False,
                "experimental_options": {
                    "preserve_mathtype_layout": True,
                    "mathtype_layout_factor": "1.02",
                    "resume_mathtype_pipeline": "true",
                    "mathtype_start_index": "216",
                    "mathtype_end_index": 238,
                },
            }
        ],
    }

    plan = build_execution_plan(routing_report)
    step = plan.steps[0]

    assert step.metadata == {
        "experimental_options": {
            "preserve_mathtype_layout": True,
            "mathtype_layout_factor": 1.02,
            "resume_mathtype_pipeline": True,
            "mathtype_start_index": 216,
            "mathtype_end_index": 238,
        }
    }

    round_trip = ExecutionPlan.from_dict(plan.to_dict())
    assert round_trip.steps[0].metadata == step.metadata
