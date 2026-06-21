"""Tests against a *persistent* checked-in external-sample .pptx fixture.

The fixture lives at ``tests/fixtures/external_sample.pptx`` and is generated
by ``tests/fixtures/_generate_external_sample.py``.

Why a *persistent* fixture (vs. the inline-generated ones in
``test_external_chart.py``):

- byte-level changes catch python-pptx / OOXML behaviour drift across versions
- simulates the real "user-uploaded template" scenario — a persistent artifact
  authored outside the engine, not freshly minted by the test
- supports multi-chart-per-slide and multi-slide layouts more naturally than
  inline construction

Together with ``test_external_chart.py`` they bracket the contract:
live generation (drift detector for python-pptx native API) + persistent
artifact (drift detector for the .pptx file format itself).
"""

from __future__ import annotations

from pathlib import Path

import pytest

from ablechart import (
    SeriesData,
    inspect_pptx_charts,
    replace_pptx_chart_data,
)


FIXTURE = Path(__file__).parent / "fixtures" / "external_sample.pptx"


@pytest.fixture
def fixture_path():
    if not FIXTURE.exists():
        pytest.skip(
            f"fixture not found: {FIXTURE}\n"
            f"Run: python tests/fixtures/_generate_external_sample.py"
        )
    return str(FIXTURE)


# ---------------------------------------------------------------------------
# Inventory facts (multi-slide, multi-chart-per-slide, mixed types)
# ---------------------------------------------------------------------------


def test_external_sample_inventory_count_and_order(fixture_path):
    """4 charts total: slide 0 (bar), slide 1 (line + scatter), slide 2 (pie)."""
    inv = inspect_pptx_charts(fixture_path)
    assert len(inv) == 4

    order = [
        (i.selector.slide_index, i.chart_index_on_slide, i.chart_type)
        for i in inv
    ]
    assert order == [
        (0, 0, "bar"),
        (1, 0, "line"),
        (1, 1, "scatter"),
        (2, 0, "pie"),
    ], f"got unexpected ordering: {order}"


def test_external_sample_explicit_business_name_extracted(fixture_path):
    """Slide 0 bar was renamed to 'quarterly_performance_chart' in the fixture;
    inspect must surface it via selector.explicit_name (ADR-0006 §1 priority 1)."""
    inv = inspect_pptx_charts(fixture_path)
    bar = inv[0]
    assert bar.shape_name == "quarterly_performance_chart"
    assert bar.selector.explicit_name == "quarterly_performance_chart"


def test_external_sample_non_chart_shapes_skipped(fixture_path):
    """Slide 2 has a text box alongside the pie chart — inspect must skip it
    (only chart shapes appear in inventory)."""
    inv = inspect_pptx_charts(fixture_path)
    slide_2_items = [i for i in inv if i.selector.slide_index == 2]
    assert len(slide_2_items) == 1, "slide 2 should report only the pie chart"
    assert slide_2_items[0].chart_type == "pie"


def test_external_sample_all_charts_replaceable(fixture_path):
    """All 4 fixture charts are in ADR-0006 §3 first-batch and have embedded workbook."""
    inv = inspect_pptx_charts(fixture_path)
    for item in inv:
        assert item.replaceable, \
            f"expected replaceable: {item.chart_type} @ slide={item.selector.slide_index}"
        assert item.has_embedded_workbook
        assert item.warnings == []


# ---------------------------------------------------------------------------
# Read-only contract (no fixture mutation under any inspect call)
# ---------------------------------------------------------------------------


def test_external_sample_fixture_byte_identical_after_inspect(fixture_path):
    before = open(fixture_path, "rb").read()
    inspect_pptx_charts(fixture_path)
    after = open(fixture_path, "rb").read()
    assert before == after, "inspect_pptx_charts must not mutate the fixture file"


# ---------------------------------------------------------------------------
# Surgical replace: change one chart, leave the other 3 alone
# ---------------------------------------------------------------------------


def test_external_sample_replace_one_chart_leaves_others_intact(fixture_path, tmp_path):
    """Replace the line chart on slide 1 chart_idx 0; bar / scatter / pie keep
    their shape_id and chart_part (selectors stable)."""
    inv_before = inspect_pptx_charts(fixture_path)
    line_selector = next(
        i.selector for i in inv_before
        if i.selector.slide_index == 1 and i.chart_index_on_slide == 0
    )

    others_before = {
        (i.selector.slide_index, i.chart_index_on_slide):
            (i.selector.shape_id, i.selector.chart_part, i.chart_type)
        for i in inv_before
        if not (i.selector.slide_index == 1 and i.chart_index_on_slide == 0)
    }

    output_pptx = str(tmp_path / "updated.pptx")
    result = replace_pptx_chart_data(
        input_pptx=fixture_path,
        output_pptx=output_pptx,
        selector=line_selector,
        categories=["2025-01", "2025-02", "2025-03"],
        series=[
            SeriesData(name="Index A", values=[120.0, 125.0, 130.0]),
            SeriesData(name="Index B", values=[110.0, 112.0, 115.0]),
        ],
    )
    assert result.status == "ok", f"{result.error_code}: {result.error_detail}"

    inv_after = inspect_pptx_charts(output_pptx)
    assert len(inv_after) == 4

    for item in inv_after:
        key = (item.selector.slide_index, item.chart_index_on_slide)
        if key == (1, 0):
            # the replaced one
            assert item.chart_type == "line", "chart_type must survive replace"
            assert item.category_count == 3, "new categories should be visible"
            assert "Index A" in item.series_names
            assert "Index B" in item.series_names
        else:
            shape_id_before, chart_part_before, type_before = others_before[key]
            assert item.selector.shape_id == shape_id_before, \
                f"unrelated chart {key} shape_id changed across surgical replace"
            assert item.selector.chart_part == chart_part_before
            assert item.chart_type == type_before


# ---------------------------------------------------------------------------
# Replace via explicit_name-based lookup also works
# ---------------------------------------------------------------------------


def test_external_sample_replace_chart_with_business_name_preserved(fixture_path, tmp_path):
    """The bar chart was tagged 'quarterly_performance_chart'. Replacing it
    must preserve that explicit_name (shape.name unchanged)."""
    inv_before = inspect_pptx_charts(fixture_path)
    bar_selector = inv_before[0].selector  # slide 0 chart 0 = bar

    output_pptx = str(tmp_path / "updated.pptx")
    result = replace_pptx_chart_data(
        input_pptx=fixture_path,
        output_pptx=output_pptx,
        selector=bar_selector,
        categories=["FY24", "FY25"],
        series=[
            SeriesData(name="Revenue", values=[300.0, 350.0]),
            SeriesData(name="Costs", values=[180.0, 200.0]),
        ],
    )
    assert result.status == "ok"

    inv_after = inspect_pptx_charts(output_pptx)
    bar_after = inv_after[0]
    assert bar_after.shape_name == "quarterly_performance_chart", \
        "shape.name (business tag) must survive replace"
    assert bar_after.selector.explicit_name == "quarterly_performance_chart"
