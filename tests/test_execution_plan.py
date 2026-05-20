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
    equation3_step = next(item for item in plan["steps"] if item["provider"] == "equation3")
    assert "BigOp (sum/integral/product/coproduct/integral-op)" in "\n".join(equation3_step["notes"])
    axmath_step = next(item for item in plan["steps"] if item["provider"] == "axmath")
    assert "classify-axmath-structural-metadata" not in {
        action["action_id"] for action in axmath_step["actions"]
    }


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


def test_build_execution_plan_adds_axmath_structural_metadata_action_only_when_explicit() -> None:
    routing_report = {
        "document_id": "axmath-structural",
        "input_path": "sample.docx",
        "detector_version": "0.1.0",
        "formula_count": 1,
        "recommended_sequence": ["axmath-ole"],
        "route_plan": [
            {
                "source_family": "axmath-ole",
                "formula_count": 1,
                "route_kind": "export-assisted",
                "next_action": "run-axmath-export-assisted-pipeline",
                "confidence_policy": "medium",
                "requires_manual_review": True,
                "axmath_structural_metadata": {
                    "source_family": "axmath-ole",
                    "rows": [
                        {
                            "row_id": 1,
                            "object_id": "synthetic-object-1",
                            "stream_locator_id": "locator-1",
                            "stream_length": 573,
                            "unique_stream_id": "U1",
                            "tail_token_class_id": "TT01",
                            "tail_window_class": "small-tail",
                            "tail_start_candidate_offset": 339,
                            "tail_length": 234,
                            "trailer_length": 8,
                            "parse_status": "candidate_count_relation_pass",
                            "unsupported_boundary": "none",
                        }
                    ],
                },
            }
        ],
    }

    plan = build_execution_plan(routing_report).to_dict()
    step = plan["steps"][0]

    assert step["provider"] == "axmath"
    assert step["route_kind"] == "export-assisted"
    assert step["next_action"] == "run-axmath-export-assisted-pipeline"
    assert [action["action_id"] for action in step["actions"]] == [
        "classify-axmath-object",
        "classify-axmath-structural-metadata",
        "export-assisted-conversion",
        "import-converted-math",
        "manual-spot-check",
    ]
    metadata = step["actions"][1]["metadata"]
    assert metadata["status"] == "metadata-only-structural-status-recorded"
    assert metadata["records"][0]["structural_status"] == "metadata_relation_consistent"
    assert metadata["records"][0]["blocker_ids"] == []
    assert metadata["claim_boundaries"]["conversion_claim"] is False
    assert metadata["claim_boundaries"]["native_parser_claim"] is False


def test_build_execution_plan_adds_axmath_controlled_acceptance_contract_only_when_explicit() -> None:
    routing_report = {
        "document_id": "axmath-controlled-acceptance",
        "input_path": "sample.docx",
        "detector_version": "0.1.0",
        "formula_count": 1,
        "recommended_sequence": ["axmath-ole"],
        "route_plan": [
            {
                "source_family": "axmath-ole",
                "formula_count": 1,
                "route_kind": "export-assisted",
                "next_action": "run-axmath-export-assisted-pipeline",
                "confidence_policy": "medium",
                "requires_manual_review": True,
                "axmath_controlled_acceptance": {},
            }
        ],
    }

    plan = build_execution_plan(routing_report).to_dict()
    step = plan["steps"][0]

    assert [action["action_id"] for action in step["actions"]] == [
        "classify-axmath-object",
        "record-axmath-controlled-acceptance-contract",
        "export-assisted-conversion",
        "import-converted-math",
        "manual-spot-check",
    ]
    metadata = step["actions"][1]["metadata"]
    assert metadata["decision_enum"] == "axmath_angle_direct_native_acceptance_refresh_passed"
    assert metadata["controlled_sample_count"] == 355
    assert metadata["accepted_count"] == 355
    assert metadata["remaining_fail_closed_count"] == 0
    assert metadata["claim_boundaries"]["conversion_claim"] is False
    assert metadata["claim_boundaries"]["public_native_parser_binding"] is False
    assert any("contract-only" in note for note in step["notes"])


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
