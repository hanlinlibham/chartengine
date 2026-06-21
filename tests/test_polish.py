"""报告级收尾层（polish）单元测试。"""

import pytest

from ablechart.polish import (
    _choose_tick_count,
    nice_range,
    strftime_from_excel,
)


class TestNiceRange:
    def test_zero_based_bar_range(self):
        # 140 在 6 格下应收敛到 150 而不是 5 格的 250
        assert nice_range(0, 140, 6, include_zero="always") == (0.0, 150.0, 25.0)

    def test_auto_zoom_for_offset_data(self):
        # 净值 1.0~1.238：数据远离 0，自动收窄
        nmin, nmax, unit = nice_range(1.0, 1.238, 5, include_zero="auto")
        assert nmin == 1.0 and nmax >= 1.238
        assert nmax <= 1.3

    def test_auto_includes_zero_when_close(self):
        nmin, _nmax, _unit = nice_range(20, 140, 5, include_zero="auto")
        assert nmin == 0.0

    def test_negative_range(self):
        nmin, nmax, _unit = nice_range(-95779, 91872, 5, include_zero="always")
        assert nmin <= -95779 and nmax >= 91872

    def test_exact_tick_count(self):
        nmin, nmax, unit = nice_range(0, 17, 4, include_zero="always")
        assert (nmax - nmin) / unit == pytest.approx(4)


class TestChooseTickCount:
    def test_picks_best_fill(self):
        # max=140: 6 格 (0..150) 明显优于 5 格 (0..250)
        n = _choose_tick_count([(0.0, 140.0, "always")])
        nmin, nmax, _ = nice_range(0, 140, n, include_zero="always")
        assert nmax <= 200

    def test_dual_axis_shares_tick_count(self):
        n = _choose_tick_count([(0.0, 140.0, "always"), (0.0, 0.17, "always")])
        assert n in (4, 5, 6)


class TestStrftimeFromExcel:
    @pytest.mark.parametrize("excel,expected", [
        ("yyyy/mm", "%Y/%m"),
        ("yyyy-mm-dd", "%Y-%m-%d"),
        ("yyyy/mm/dd", "%Y/%m/%d"),
        ("mm月", "%m月"),
        ("yy-m-d", "%y-%m-%d"),
        (None, "%Y/%m"),
    ])
    def test_no_cascade_corruption(self, excel, expected):
        # 回归：级联 replace 曾把 mm→%m 的结果再替换成 %%mm，
        # 导致轴标签渲染出字面量 "2024/%m"
        assert strftime_from_excel(excel) == expected
