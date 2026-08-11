import json
from copy import deepcopy

import table_schema_mapper
from scripts.promote_schema_candidates import review_candidates


def test_llm_schema_candidate_is_recorded_without_mutating_config(tmp_path, monkeypatch):
    config_path = tmp_path / "mappings.json"
    config = {
        "version": 1,
        "schemas": {
            "surface_water_plan": {
                "required_fields": ["point_code"],
                "aliases": {
                    "point_code": ["code"],
                    "river_name": ["river_name"],
                },
            }
        },
        "candidates": [],
    }
    config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
    before = config_path.read_bytes()

    monkeypatch.setattr(table_schema_mapper, "DEFAULT_CONFIG_PATH", config_path)
    monkeypatch.setenv("EIA_LLM_API_KEY", "test-only")
    monkeypatch.setattr(
        table_schema_mapper,
        "call_schema_llm",
        lambda _payload: {"mapping": {"point_code": "section_code"}},
    )

    result = table_schema_mapper.resolve_table_schema(
        "surface_water_plan",
        ["section_code", "river_name"],
        sample_rows=[["WJ1", "sample river"]],
        output_dir=tmp_path / "output",
        source_table="plan.docx:table:1",
        job_id="job-1",
        enable_llm=True,
    )

    assert result["validation"]["valid"] is True
    assert result["mapping_sources"]["point_code"] == "llm"
    assert config_path.read_bytes() == before
    candidates = json.loads(
        (tmp_path / "output" / "debug_tables" / "table_schema_candidates.json").read_text(
            encoding="utf-8"
        )
    )
    assert candidates == [
        {
            "schema": "surface_water_plan",
            "field": "point_code",
            "header": "section_code",
            "source_headers": ["section_code", "river_name"],
            "source_table": "plan.docx:table:1",
            "job_id": "job-1",
            "validation_status": "accepted",
            "seen_at": candidates[0]["seen_at"],
        }
    ]


def test_candidate_review_accepts_valid_mapping():
    config = {
        "schemas": {
            "surface_water_plan": {
                "aliases": {
                    "point_code": ["code"],
                    "river_name": ["river_name"],
                }
            }
        }
    }
    candidates = [
        {
            "schema": "surface_water_plan",
            "field": "point_code",
            "header": "section_code",
            "source_headers": ["section_code", "river_name"],
            "validation_status": "accepted",
        }
    ]

    updated, summary = review_candidates(deepcopy(config), candidates)

    assert summary["can_apply"] is True
    assert summary["accepted"][0]["header"] == "section_code"
    assert updated["schemas"]["surface_water_plan"]["aliases"]["point_code"] == [
        "code",
        "section_code",
    ]


def test_candidate_review_rejects_conflict_as_all_or_nothing():
    config = {
        "schemas": {
            "surface_water_plan": {
                "aliases": {
                    "point_code": ["code"],
                    "river_name": ["river_name"],
                }
            }
        }
    }
    candidates = [
        {
            "schema": "surface_water_plan",
            "field": "point_code",
            "header": "river_name",
            "source_headers": ["river_name"],
            "validation_status": "accepted",
        }
    ]

    _updated, summary = review_candidates(config, candidates)

    assert summary["can_apply"] is False
    assert "already belongs" in summary["rejected"][0]["reason"]
