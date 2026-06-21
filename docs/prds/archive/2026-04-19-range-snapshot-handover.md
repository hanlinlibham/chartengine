# Range Snapshot Handover

## 背景

这一轮工作的主线有两段：

1. 先把 `jp_demo.pdf` 里的图做了 chart family 建模，明确 `range_snapshot_chart` 是一类独立 family。
2. 然后围绕这个 family 做了能力实现、PowerPoint 实机验证、以及 vertical valuation 页的 page-specific presets。

后续用户明确指出：当前 `range_snapshot` 图质量明显低于仓库里原有的 combo 图。

这是一个正确判断。

## 当前结论

### 1. 学习目标确实一度偏了

前半段我把重心放在了：

- family 分类
- standalone / composer 打通
- axis break 功能闭环

这些是能力建设，但不是质量建设。

而用户真正要的，是接近 `jp_demo` 样张的视觉质量。

### 2. 原有 combo 的质量明显更高

我已经把原来的 4 个 combo preset 跑出来，并用真实 PowerPoint 导出了预览图。

总览：

- [原 combo 总览](pptfi/output/combo-quality-previews/contact-sheet.png)
- [range snapshot 总览](pptfi/output/vertical-valuation-previews/contact-sheet.png)

对比后可以确认：

- combo 图几乎完全落在 Office 原生图表的舒适区
- `range_snapshot` 虽然功能通了，但仍然是 `native chart + custom overlay` 的复合体
- 所以视觉质感、轴系统的一体感、标签自然度，都明显弱于 combo

### 3. `axis_break` 已经从视觉标记升级为“真正影响坐标”的分段压缩

在 `range_snapshot` 里，`axis_break` 现在不只是两条斜杠：

- bar 长度会按 piecewise transform 改变
- average/current marker 位置也会跟着改
- axis tick / gridline 走同一个变换

但它仍然不是 PowerPoint 原生 broken axis，而是我在 chart 之上实现的一套自定义轴 overlay。

这就是为什么它“功能可用”，但还没达到原 combo 的质感。

## 关键文件

### `pptchartengine`

- [range_snapshot.py](../../../src/pptchartengine/range_snapshot.py)
- [presets.py](../../../src/pptchartengine/presets.py)
- [__init__.py](../../../src/pptchartengine/__init__.py)
- [test_package_contract.py](../../../tests/test_package_contract.py)

### `pptfi`

- [operations.py](pptfi/pptfi/operations.py)
- [cli.py](pptfi/pptfi/cli.py)
- [range_snapshot composer layout](pptfi/pptfi/composer/layouts/range_snapshot.py)
- [layouts/__init__.py](pptfi/pptfi/composer/layouts/__init__.py)
- [README.md](pptfi/README.md)
- [test_cli_sdk_contract.py](pptfi/tests/test_cli_sdk_contract.py)
- [test_system_contract.py](pptfi/tests/test_system_contract.py)

### 建模 / 文档

- [jp demo chart catalog](pptfi/reference/jp_demo_chart_catalog.json)
- [jp demo atlas html](pptfi/reference/jp-demo-chart-atlas.html)
- [range snapshot presets 文档](pptfi/reference/range-snapshot-presets.md)

### 新增数据 / demo

- [range_snapshot_demo.json](pptfi/range_snapshot_demo.json)
- [job_range_snapshot_demo.json](pptfi/job_range_snapshot_demo.json)
- [range_snapshot_sector_demo.json](pptfi/range_snapshot_sector_demo.json)
- [job_range_snapshot_sector_demo.json](pptfi/job_range_snapshot_sector_demo.json)

- [range_snapshot_valuation.csv](pptfi/data/range_snapshot_valuation.csv)
- [range_snapshot_sector_valuation.csv](pptfi/data/range_snapshot_sector_valuation.csv)
- [asx200_sector_valuation_snapshot.csv](pptfi/data/asx200_sector_valuation_snapshot.csv)
- [sp500_sector_valuation_snapshot.csv](pptfi/data/sp500_sector_valuation_snapshot.csv)
- [msci_emu_sector_valuation_snapshot.csv](pptfi/data/msci_emu_sector_valuation_snapshot.csv)
- [msci_japan_sector_valuation_snapshot.csv](pptfi/data/msci_japan_sector_valuation_snapshot.csv)

### PowerPoint 实机导出脚本

- [export_powerpoint_pdf.py](pptfi/scripts/export_powerpoint_pdf.py)
- [render_vertical_valuation_previews.py](pptfi/scripts/render_vertical_valuation_previews.py)

## 现成产物

### 1. vertical valuation 预览

目录：

- [vertical valuation previews](pptfi/output/vertical-valuation-previews)

关键文件：

- [总览图](pptfi/output/vertical-valuation-previews/contact-sheet.png)
- [ASX 200 预览](pptfi/output/vertical-valuation-previews/asx200-png/slide-1.png)
- [S&P 500 预览](pptfi/output/vertical-valuation-previews/sp500-png/slide-1.png)
- [MSCI EMU 预览](pptfi/output/vertical-valuation-previews/msci_emu-png/slide-1.png)
- [MSCI Japan 预览](pptfi/output/vertical-valuation-previews/msci_japan-png/slide-1.png)

### 2. combo 预览

目录：

- [combo quality previews](pptfi/output/combo-quality-previews)

关键文件：

- [总览图](pptfi/output/combo-quality-previews/contact-sheet.png)

## 当前验证状态

已通过：

- `pptchartengine`: `24 passed`
- `pptfi`: `26 passed`

此外已做过 PowerPoint 实机导出验证，而不只是单测。

## 当前实现边界

### 已经做完

- `range_snapshot` vertical / horizontal
- `axis_break` piecewise compression
- `tick_values` 支持
- standalone / composer / CLI 打通
- 4 个 page-specific vertical valuation presets
- PowerPoint 预览脚本

### 还没有做好的

- `range_snapshot` 的整体视觉质量仍明显低于原 combo
- current label 的摆放仍偏机械
- plot area / bar gap / 轴密度 仍然不够像 `jp_demo`
- custom axis overlay 虽然可用，但“原生感”不够
- 我没有尝试还原图表外实现的标注点，这一条是按用户要求故意不做的

## 对下一线程的建议

如果用户换目标，这一轮工作建议到此为止，不要再在 `range_snapshot` 上继续补功能。

如果后面又回到这条线，建议只做一类事情：

### 纯质量精修，不再扩功能

优先顺序：

1. 逐页精修 vertical valuation 的图表内视觉
2. 不再新增 family
3. 不再新增 workflow

具体调参对象：

- plot area size / margin
- bar gap / bar width
- tick label font size / density / position
- average tick width
- current label offset
- axis break 的位置与尺寸

换句话说，后续应该进入“设计调优模式”，不是“能力建设模式”。

## 一个很重要的判断

如果后面用户再次问“为什么原 combo 好这么多”，不要从“参数还没调好”开始答。

更准确的回答是：

- combo 图本质上是 `Office native chart quality`
- range snapshot 目前仍然是 `native chart + custom overlay quality`

这是结构性差异，不是单个参数问题。

这个判断已经有真实 PowerPoint 导出的证据支撑。
