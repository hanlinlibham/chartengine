"""pptchartengine

Core editable PowerPoint chart engine extracted from `ppt-st`.

Scope:
- native editable combo charts in `.pptx`
- dual-axis financial time-series charts
- date-axis presets
- chart parsing and cleanup helpers
- finance preset configs for common pension/investment charts
"""

from .api import create_combo_chart
from .cleaner import ChartJunkCleaner, clean_chart
from .date_axis import (
    BIWEEKLY_TICKS,
    DAILY_TICKS,
    MONTHLY_TICKS,
    QUARTERLY_TICKS,
    WEEKLY_TICKS,
    YEARLY_TICKS,
    DateAxisConfig,
)
from .layout import (
    DEFAULT_CATEGORY_AXIS_CONFIG,
    DEFAULT_LEGEND_CONFIG,
    DEFAULT_VALUE_AXIS_CONFIG,
    CategoryAxisConfig,
    ChartLayoutConfig,
    LegendConfig,
    ValueAxisConfig,
)
from .parser import ChartParser, parse_all_charts_from_pptx, parse_chart_from_pptx
from .presets import (
    CHART_PRESET_FUNCTIONS,
    FINANCE_PRESET_FUNCTIONS,
    get_chart1_config,
    get_chart2_config,
    get_chart3_config,
    get_chart4_config,
    get_chart_config,
)
from .styles import (
    COLOR_SCHEMES,
    DEFAULT_STYLE_CONFIG,
    DARK_BLUE,
    DARK_GRAY,
    DARK_ORANGE,
    DARK_RED,
    LIGHT_BLUE,
    LIGHT_GRAY,
    LIGHT_ORANGE,
    LIGHT_RED,
    StyleConfig,
)
from .scatter import (
    ScatterParseResult,
    create_bubble_chart,
    create_scatter_chart,
    parse_bubble_chart,
    parse_bubble_from_pptx,
    parse_scatter_chart,
    parse_scatter_from_pptx,
)
from .waterfall import (
    WaterfallParseResult,
    build_waterfall_spec,
    create_waterfall_chart,
    get_waterfall_spec,
    parse_waterfall_chart,
    parse_waterfall_from_pptx,
    prepare_waterfall_dataframe,
    restore_waterfall_dataframe,
)

__version__ = "0.1.0"

__all__ = [
    "BIWEEKLY_TICKS",
    "CHART_PRESET_FUNCTIONS",
    "COLOR_SCHEMES",
    "ChartJunkCleaner",
    "ChartLayoutConfig",
    "ChartParser",
    "CategoryAxisConfig",
    "DAILY_TICKS",
    "DEFAULT_CATEGORY_AXIS_CONFIG",
    "DEFAULT_LEGEND_CONFIG",
    "DEFAULT_STYLE_CONFIG",
    "DEFAULT_VALUE_AXIS_CONFIG",
    "DARK_BLUE",
    "DARK_GRAY",
    "DARK_ORANGE",
    "DARK_RED",
    "DateAxisConfig",
    "FINANCE_PRESET_FUNCTIONS",
    "LIGHT_BLUE",
    "LIGHT_GRAY",
    "LIGHT_ORANGE",
    "LIGHT_RED",
    "LegendConfig",
    "MONTHLY_TICKS",
    "QUARTERLY_TICKS",
    "StyleConfig",
    "ValueAxisConfig",
    "WEEKLY_TICKS",
    "YEARLY_TICKS",
    "clean_chart",
    "create_bubble_chart",
    "create_combo_chart",
    "create_scatter_chart",
    "get_chart1_config",
    "get_chart2_config",
    "get_chart3_config",
    "get_chart4_config",
    "get_chart_config",
    "parse_bubble_chart",
    "parse_bubble_from_pptx",
    "parse_all_charts_from_pptx",
    "parse_chart_from_pptx",
    "parse_scatter_chart",
    "parse_scatter_from_pptx",
    "ScatterParseResult",
    "build_waterfall_spec",
    "create_waterfall_chart",
    "get_waterfall_spec",
    "prepare_waterfall_dataframe",
    "restore_waterfall_dataframe",
    "parse_waterfall_chart",
    "parse_waterfall_from_pptx",
    "WaterfallParseResult",
]
