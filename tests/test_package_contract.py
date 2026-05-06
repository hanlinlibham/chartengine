import pandas as pd
import pytest
from pptx import Presentation

from pptchartengine import (
    BIWEEKLY_TICKS,
    DAILY_TICKS,
    MONTHLY_TICKS,
    QUARTERLY_TICKS,
    WEEKLY_TICKS,
    YEARLY_TICKS,
    CHART_PRESET_FUNCTIONS,
    ChartLayoutConfig,
    ChartParser,
    DateAxisConfig,
    FINANCE_PRESET_FUNCTIONS,
    StyleConfig,
    WaterfallParseResult,
    ScatterParseResult,
    create_bubble_chart,
    create_combo_chart,
    create_scatter_chart,
    create_waterfall_chart,
    get_waterfall_spec,
    get_chart1_config,
    get_chart4_config,
    parse_all_charts_from_pptx,
    parse_bubble_from_pptx,
    parse_chart_from_pptx,
    parse_scatter_from_pptx,
    parse_waterfall_from_pptx,
    prepare_waterfall_dataframe,
    restore_waterfall_dataframe,
)


def _sample_finance_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "分类": pd.date_range("2024-01-01", periods=14, freq="D"),
            "沪深300指数(收盘价)": [3500 + i * 10 for i in range(14)],
            "组合收益率（左轴）": [0.01 + i * 0.001 for i in range(14)],
            "组合规模(万元)": [100000 + i * 2000 for i in range(14)],
            "累计收益率": [0.10 + i * 0.002 for i in range(14)],
            "权益仓位（人社部口径）": [0.18 + i * 0.003 for i in range(14)],
            "中债新综合总财富指数(收盘价)": [210.5 + i * 0.8 for i in range(14)],
            "久期": [0.45 + i * 0.02 for i in range(14)],
        }
    )


def test_public_contract_exports():
    assert callable(create_combo_chart)
    assert callable(parse_chart_from_pptx)
    assert callable(parse_all_charts_from_pptx)
    assert ChartParser.__name__ == "ChartParser"
    assert StyleConfig.__name__ == "StyleConfig"
    assert ChartLayoutConfig.__name__ == "ChartLayoutConfig"
    assert DateAxisConfig.__name__ == "DateAxisConfig"
    assert WaterfallParseResult.__name__ == "WaterfallParseResult"
    assert ScatterParseResult.__name__ == "ScatterParseResult"


def test_date_axis_presets_exposed():
    presets = [
        DAILY_TICKS,
        WEEKLY_TICKS,
        BIWEEKLY_TICKS,
        MONTHLY_TICKS,
        QUARTERLY_TICKS,
        YEARLY_TICKS,
    ]
    assert all(isinstance(preset, DateAxisConfig) for preset in presets)


def test_financial_chart1_preset_shape():
    config = get_chart1_config(_sample_finance_df())

    assert config["categories_col"] == "分类"
    assert [series["type"] for series in config["series_config"]] == ["bar", "line"]
    assert [series["axis"] for series in config["series_config"]] == ["secondary", "primary"]
    assert config["style_config"].color_scheme == "aim00"
    assert config["layout_config"].title == "组合2024年以来收益率走势图"
    assert config["layout_config"].secondary_value_axis_config.number_format == "#,##0"
    assert config["layout_config"].date_axis_config.number_format == "yyyy/mm"


def test_financial_chart4_preset_uses_dynamic_secondary_axis_range():
    df = _sample_finance_df()
    config = get_chart4_config(df)
    axis = config["layout_config"].secondary_value_axis_config
    bond_series = df["中债新综合总财富指数(收盘价)"]
    bond_range = bond_series.max() - bond_series.min()
    bond_unit = int(bond_range / 6)

    assert axis.min_value == int(bond_series.min() - (bond_range / 6))
    assert axis.max_value == int(bond_series.max() + (bond_range / 6))
    assert axis.major_unit == bond_unit


def test_finance_registry_aliases_match():
    assert FINANCE_PRESET_FUNCTIONS is CHART_PRESET_FUNCTIONS
    assert set(CHART_PRESET_FUNCTIONS) == {
        "图表1 - 当年收益率走势",
        "图表2 - 成立以来收益率",
        "图表3 - 权益仓位走势",
        "图表4 - 久期走势",
    }


def test_prepare_waterfall_dataframe():
    df = pd.DataFrame(
        {
            "项目": ["期初", "权益贡献", "债券贡献", "汇率拖累", "期末"],
            "值": [100, 20, 10, -15, 115],
            "measure": ["total", "relative", "relative", "relative", "total"],
        }
    )

    waterfall = prepare_waterfall_dataframe(df, "项目", "值", measure_col="measure")
    assert list(waterfall.columns) == ["项目", "__base__", "__increase__", "__decrease__", "__total__"]
    assert waterfall.iloc[0]["__total__"] == 100
    assert waterfall.iloc[1]["__base__"] == 100
    assert waterfall.iloc[1]["__increase__"] == 20
    assert waterfall.iloc[3]["__base__"] == 115
    assert waterfall.iloc[3]["__decrease__"] == 15
    assert waterfall.iloc[4]["__total__"] == 115


def test_prepare_waterfall_dataframe_supports_total_categories_without_measure_col():
    df = pd.DataFrame(
        {
            "项目": ["起点", "经营改善", "融资拖累", "终点"],
            "值": [80, 25, -10, 95],
        }
    )

    waterfall = prepare_waterfall_dataframe(
        df,
        "项目",
        "值",
        total_categories=["起点", "终点"],
    )

    assert waterfall.iloc[0]["__total__"] == 80
    assert waterfall.iloc[1]["__base__"] == 80
    assert waterfall.iloc[2]["__base__"] == 95
    assert waterfall.iloc[2]["__decrease__"] == 10
    assert waterfall.iloc[3]["__total__"] == 95


def test_round_trip_recovers_categories_col_and_series_keys(tmp_path):
    df = pd.DataFrame(
        {
            "年份": [2021, 2022, 2023, 2024],
            "revenue": [100.0, 110.0, 120.0, 140.0],
            "profit": [10.0, 12.5, 15.0, 18.0],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    create_combo_chart(
        slide=slide,
        df=df,
        categories_col="年份",
        series_config=[
            {"key": "revenue", "name": "营收(亿元)", "type": "bar", "axis": "primary"},
            {"key": "profit", "name": "净利润(亿元)", "type": "line", "axis": "secondary"},
        ],
    )

    output = tmp_path / "roundtrip.pptx"
    prs.save(output)

    series_config, parsed_df, categories_col, layout_info = parse_chart_from_pptx(str(output))

    assert categories_col == "年份"
    assert [series["key"] for series in series_config] == ["revenue", "profit"]
    assert [series["name"] for series in series_config] == ["营收(亿元)", "净利润(亿元)"]
    assert [series["axis"] for series in series_config] == ["primary", "secondary"]
    assert list(parsed_df.columns) == ["年份", "营收(亿元)", "净利润(亿元)"]
    assert isinstance(layout_info, dict)


def test_round_trip_recovers_categories_col_with_ascii_header(tmp_path):
    # Pins the metadata fallback for the case where the embedded workbook
    # leaves the first header cell blank — the failure mode previously
    # surfaced only because every other combo bar/line round-trip test
    # happened to use a non-ASCII categories_col.
    df = pd.DataFrame(
        {
            "year": [2021, 2022, 2023, 2024],
            "revenue": [100.0, 110.0, 120.0, 140.0],
            "profit": [10.0, 12.5, 15.0, 18.0],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    create_combo_chart(
        slide=slide,
        df=df,
        categories_col="year",
        series_config=[
            {"key": "revenue", "name": "Revenue", "type": "bar", "axis": "primary"},
            {"key": "profit", "name": "Profit", "type": "line", "axis": "secondary"},
        ],
    )

    output = tmp_path / "roundtrip-ascii.pptx"
    prs.save(output)

    series_config, parsed_df, categories_col, _ = parse_chart_from_pptx(str(output))

    assert categories_col == "year"
    assert [series["key"] for series in series_config] == ["revenue", "profit"]
    assert list(parsed_df.columns) == ["year", "Revenue", "Profit"]


def test_round_trip_recovers_stacked_grouping(tmp_path):
    df = pd.DataFrame(
        {
            "年份": [2021, 2022, 2023],
            "domestic": [40.0, 42.0, 44.0],
            "overseas": [60.0, 58.0, 56.0],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    create_combo_chart(
        slide=slide,
        df=df,
        categories_col="年份",
        series_config=[
            {"key": "domestic", "name": "国内收入", "type": "bar", "axis": "primary", "grouping": "stacked"},
            {"key": "overseas", "name": "海外收入", "type": "bar", "axis": "primary", "grouping": "stacked"},
        ],
    )

    output = tmp_path / "stacked-roundtrip.pptx"
    prs.save(output)

    series_config, parsed_df, categories_col, layout_info = parse_chart_from_pptx(str(output))

    assert categories_col == "年份"
    assert all(series["grouping"] == "stacked" for series in series_config)
    assert [series["key"] for series in series_config] == ["domestic", "overseas"]


def test_round_trip_recovers_percent_stacked_grouping(tmp_path):
    df = pd.DataFrame(
        {
            "年份": [2021, 2022, 2023],
            "equity": [55.0, 52.0, 50.0],
            "bond": [45.0, 48.0, 50.0],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    create_combo_chart(
        slide=slide,
        df=df,
        categories_col="年份",
        series_config=[
            {"key": "equity", "name": "权益占比", "type": "bar", "axis": "primary", "grouping": "percent_stacked"},
            {"key": "bond", "name": "固收占比", "type": "bar", "axis": "primary", "grouping": "percent_stacked"},
        ],
    )

    output = tmp_path / "percent-stacked-roundtrip.pptx"
    prs.save(output)

    series_config, parsed_df, categories_col, layout_info = parse_chart_from_pptx(str(output))

    assert categories_col == "年份"
    assert all(series["grouping"] == "percent_stacked" for series in series_config)
    assert [series["key"] for series in series_config] == ["equity", "bond"]


def test_create_waterfall_chart_round_trip(tmp_path):
    df = pd.DataFrame(
        {
            "项目": ["期初", "权益贡献", "债券贡献", "汇率拖累", "期末"],
            "值": [100, 20, 10, -15, 115],
            "measure": ["total", "relative", "relative", "relative", "total"],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    create_waterfall_chart(
        slide=slide,
        df=df,
        categories_col="项目",
        value_col="值",
        measure_col="measure",
    )

    output = tmp_path / "waterfall.pptx"
    prs.save(output)
    series_config, parsed_df, categories_col, layout_info = parse_chart_from_pptx(str(output))

    assert categories_col == "项目"
    assert all(series["grouping"] == "stacked" for series in series_config)
    assert [series["key"] for series in series_config] == [
        "__base__",
        "__increase__",
        "__decrease__",
        "__total__",
    ]
    assert parsed_df.shape == (5, 5)
    assert layout_info["chart_family"] == "waterfall"
    assert get_waterfall_spec(layout_info)["value_col"] == "值"

    restored = restore_waterfall_dataframe(parsed_df, layout_info)
    assert restored.to_dict(orient="list") == df.to_dict(orient="list")


def test_parse_waterfall_from_pptx_returns_semantic_dataframe(tmp_path):
    df = pd.DataFrame(
        {
            "项目": ["期初", "股票贡献", "债券贡献", "期末"],
            "值": [100, 15, -5, 110],
            "measure": ["total", "relative", "relative", "absolute"],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    create_waterfall_chart(
        slide=slide,
        df=df,
        categories_col="项目",
        value_col="值",
        measure_col="measure",
    )

    output = tmp_path / "waterfall-semantic-roundtrip.pptx"
    prs.save(output)

    result = parse_waterfall_from_pptx(str(output))

    assert result.categories_col == "项目"
    assert result.value_col == "值"
    assert result.measure_col == "measure"
    assert result.df.to_dict(orient="list") == df.to_dict(orient="list")
    assert result.raw_chart_df.columns.tolist() == [
        "项目",
        "__base__",
        "__increase__",
        "__decrease__",
        "__total__",
    ]

    assert result.series_config[0]["key"] == "__base__"
    assert result.layout_info["chart_family"] == "waterfall"


def test_create_scatter_chart_round_trip(tmp_path):
    df = pd.DataFrame(
        {
            "volatility": [8.1, 9.2, 7.4],
            "return": [10.5, 12.0, 8.8],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    create_scatter_chart(
        slide=slide,
        df=df,
        x_col="volatility",
        y_col="return",
        series_name="Risk/Return",
    )
    output = tmp_path / "scatter.pptx"
    prs.save(output)

    result = parse_scatter_from_pptx(str(output))
    assert result.chart_family == "scatter"
    assert result.x_col == "volatility"
    assert result.y_col == "return"
    assert result.size_col is None
    assert result.df.to_dict(orient="list") == df.to_dict(orient="list")


def test_create_bubble_chart_round_trip(tmp_path):
    df = pd.DataFrame(
        {
            "volatility": [8.1, 9.2, 7.4],
            "return": [10.5, 12.0, 8.8],
            "aum": [120.0, 210.0, 95.0],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    create_bubble_chart(
        slide=slide,
        df=df,
        x_col="volatility",
        y_col="return",
        size_col="aum",
        series_name="Risk/Return/AUM",
    )
    output = tmp_path / "bubble.pptx"
    prs.save(output)

    result = parse_bubble_from_pptx(str(output))
    assert result.chart_family == "bubble"
    assert result.x_col == "volatility"
    assert result.y_col == "return"
    assert result.size_col == "aum"
    assert result.df.to_dict(orient="list") == df.to_dict(orient="list")


def test_create_combo_chart_supports_scatter_round_trip(tmp_path):
    df = pd.DataFrame(
        {
            "volatility": [8.1, 9.2, 7.4],
            "return_a": [10.5, 12.0, 8.8],
            "return_b": [9.8, 11.4, 8.1],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    create_combo_chart(
        slide=slide,
        df=df,
        categories_col="volatility",
        series_config=[
            {"key": "return_a", "name": "Risk/Return A", "type": "scatter", "axis": "primary", "x_key": "volatility"},
            {"key": "return_b", "name": "Risk/Return B", "type": "scatter", "axis": "primary", "x_key": "volatility"},
        ],
    )
    output = tmp_path / "combo-scatter.pptx"
    prs.save(output)

    series_config, parsed_df, categories_col, _layout_info = parse_chart_from_pptx(str(output))
    assert categories_col == "volatility"
    assert [series["type"] for series in series_config] == ["scatter", "scatter"]
    assert [series["x_key"] for series in series_config] == ["volatility", "volatility"]
    assert parsed_df.to_dict(orient="list") == df.to_dict(orient="list")


def test_create_combo_chart_supports_bubble_round_trip(tmp_path):
    df = pd.DataFrame(
        {
            "volatility": [8.1, 9.2, 7.4],
            "return": [10.5, 12.0, 8.8],
            "aum": [120.0, 210.0, 95.0],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    create_combo_chart(
        slide=slide,
        df=df,
        categories_col="volatility",
        series_config=[
            {
                "key": "return",
                "name": "Risk/Return/AUM",
                "type": "bubble",
                "axis": "primary",
                "x_key": "volatility",
                "size_key": "aum",
            }
        ],
    )
    output = tmp_path / "combo-bubble.pptx"
    prs.save(output)

    series_config, parsed_df, categories_col, _layout_info = parse_chart_from_pptx(str(output))
    assert categories_col == "volatility"
    assert series_config[0]["type"] == "bubble"
    assert series_config[0]["x_key"] == "volatility"
    assert series_config[0]["size_key"] == "aum"
    assert parsed_df.to_dict(orient="list") == df.to_dict(orient="list")


def test_create_combo_chart_rejects_xy_and_category_mix():
    df = pd.DataFrame(
        {
            "volatility": [8.1, 9.2, 7.4],
            "return": [10.5, 12.0, 8.8],
            "aum": [120.0, 210.0, 95.0],
        }
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    with pytest.raises(ValueError, match="混搭"):
        create_combo_chart(
            slide=slide,
            df=df,
            categories_col="volatility",
            series_config=[
                {"key": "return", "name": "Scatter", "type": "scatter", "axis": "primary", "x_key": "volatility"},
                {"key": "aum", "name": "Column", "type": "bar", "axis": "primary"},
            ],
        )
