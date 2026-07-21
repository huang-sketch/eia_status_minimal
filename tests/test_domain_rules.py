import pytest

from noise_section_generator import (
    infer_noise_frequency_from_records,
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
