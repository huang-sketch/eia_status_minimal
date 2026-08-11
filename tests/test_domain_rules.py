import pytest

from noise_section_generator import (
    infer_noise_frequency_from_records,
    clean_result_position,
    compose_plan_monitor_position,
    extract_orientation_context,
    monitor_point_display_position,
    result_position_for_record,
    normalize_floor_position_display,
    split_sensitive_records_by_flow_unit,
)
from surface_water_pipeline import calc_do_index, calc_ph_index


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (6.0, 1.0),
        (7.0, 0.0),
        (9.0, 1.0),
        (7.5, 0.25),
    ],
)
def test_ph_index_boundaries(value, expected):
    assert calc_ph_index(value, 6.0, 9.0) == pytest.approx(expected)


def test_do_index_uses_same_day_temperature():
    assert calc_do_index(6.0, 5.0, 20.0) <= 1.0
    assert calc_do_index(2.5, 5.0, 20.0) > 1.0


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("10", "10\u5c42"),
        ("3\u5ba4\u5916", "3\u5c42\u5ba4\u5916"),
        ("5\u697c", "5\u5c42"),
        ("\u9762\u5411\u672c\u9879\u76ee1", "\u9762\u5411\u672c\u9879\u76ee1\u5c42"),
    ],
)
def test_floor_position_normalization(raw, expected):
    assert normalize_floor_position_display(raw) == expected


def test_noise_frequency_prefers_result_evidence():
    records = [{"noise_type": "area_noise"}]
    assert infer_noise_frequency_from_records(records, ["\u76d1\u6d4b1\u5929"]) == "\u76d1\u6d4b2\u5929\uff0c\u663c\u3001\u591c\u95f4\u5404\u76d1\u6d4b1\u6b21\u3002"


def test_sensitive_records_split_train_flow_from_road_flow():
    records = [
        {"traffic_flow": "120", "traffic_flow_unit": "\u8f86/20min"},
        {"traffic_flow": "8", "traffic_flow_unit": "\u5217/60min"},
        {"traffic_flow": "-", "traffic_flow_unit": ""},
    ]
    groups = split_sensitive_records_by_flow_unit("area_traffic_noise", records)
    assert [name for name, _items in groups] == [
        "area_traffic_noise",
        "area_traffic_train_flow_noise",
    ]
    assert len(groups[0][1]) == 2
    assert len(groups[1][1]) == 1

@pytest.mark.parametrize(
    "position",
    [
        "\u9762\u5411\u82cf\u5dde\u7ed5\u57ce\u9ad8\u901f\u9996\u6392",
        "\u9762\u5411\u82cf\u5dde\u7ed5\u57ce\u9ad8\u901f\u4e00\u6392",
        "\u9762\u5411\u82cf\u5dde\u7ed5\u57ce\u9ad8\u901f\u7b2c\u4e00\u6392",
        "\u9762\u5411\u82cf\u5dde\u7ed5\u57ce\u9ad8\u901f\u4e8c\u6392",
        "\u9762\u5411\u82cf\u5dde\u7ed5\u57ce\u9ad8\u901f\u7b2c\u4e8c\u6392",
        "\u9762\u5411\u82cf\u5dde\u7ed5\u57ce\u9ad8\u901f\u4e09\u6392",
        "\u9762\u5411\u82cf\u5dde\u7ed5\u57ce\u9ad8\u901f\u7b2c\u4e09\u6392",
    ],
)
def test_monitor_position_preserves_plan_row_wording(position):
    report_text = f"NJ1 \u6d4b\u8bd5\u70b9{position}2\u5c42"

    assert compose_plan_monitor_position(position, report_text) == f"{position}2\u5c42"
    assert extract_orientation_context(position) == position


def test_table_21_and_result_table_keep_second_row_context():
    plan = {
        "point_code": "NJ1-2",
        "point_name": "\u845b\u5bb6\u5df7",
        "position": "\u9762\u5411\u82cf\u5dde\u7ed5\u57ce\u9ad8\u901f\u4e8c\u6392",
        "station": "K1+000",
        "is_attenuation": False,
    }
    record = {
        "raw_point_code": "NJ1-2",
        "point_text": "NJ1-2 \u845b\u5bb6\u5df7\u9762\u5411\u82cf\u5dde\u7ed5\u57ce\u9ad8\u901f\u4e8c\u63922\u5c42",
        "raw": {},
        "source_headers": [],
        "laeq_key": "",
    }
    expected = "\u9762\u5411\u82cf\u5dde\u7ed5\u57ce\u9ad8\u901f\u4e8c\u63922\u5c42"

    assert monitor_point_display_position(plan, record) == expected
    result_position = result_position_for_record(record, plan)
    assert result_position == expected
    assert clean_result_position(result_position, record, plan) == expected


@pytest.mark.parametrize(
    ("plan_position", "report_text", "expected"),
    [
        ("2\u5c42", "NJ1 \u6d4b\u8bd5\u70b92\u5c42", "2\u5c42"),
        ("\u5ba4\u5916", "NJ1 \u6d4b\u8bd5\u70b92\u5c42", "2\u5c42\u5ba4\u5916"),
        ("\u80cc\u5411\u672c\u9879\u76ee\u4e8c\u6392", "NJ1 \u6d4b\u8bd5\u70b94\u5c42", "\u80cc\u5411\u672c\u9879\u76ee\u4e8c\u63924\u5c42"),
        ("\u4e34\u672c\u9879\u76ee\u4fa7", "NJ1 \u6d4b\u8bd5\u70b96\u5c42", "\u4e34\u672c\u9879\u76ee\u4fa76\u5c42"),
    ],
)
def test_monitor_position_composition_regressions(plan_position, report_text, expected):
    assert compose_plan_monitor_position(plan_position, report_text) == expected
