"""Unit tests for ChartParser internal helpers.

These pin contracts that are not directly observable through the public
round-trip surface but whose breakage would silently disable the
metadata-driven recovery path.
"""

import math

import pandas as pd
import pytest

from pptchartengine import ChartParser


@pytest.mark.parametrize(
    "value",
    [
        None,
        float("nan"),
        math.nan,
        pd.NA,
        "",
        "   ",
        "nan",
        "NaN",
        "NAN",
        "None",
        "<NA>",
        "<na>",
        "Unnamed: 0",
        "Unnamed: 12",
        "unnamed: 3",
    ],
)
def test_is_missing_header_treats_null_like_values_as_missing(value):
    assert ChartParser._is_missing_header(value) is True


@pytest.mark.parametrize(
    "value",
    [
        "year",
        "年份",
        "项目",
        "Revenue (USD)",
        "  trimmed  ",
        0,
        2024,
        "0",
    ],
)
def test_is_missing_header_treats_real_headers_as_present(value):
    assert ChartParser._is_missing_header(value) is False
