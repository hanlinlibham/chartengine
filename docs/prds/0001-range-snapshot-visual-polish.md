# PRD-0001: Range Snapshot 视觉质量精修

- Status: Deferred to Slice 3 (v0.3.0)
- Created: 2026-04-19（来源：原 range-snapshot-handover 文档）
- Converted to PRD: 2026-05-03
- Target: v0.3.0
- Related: ADR-0001, ADR-0002, ADR-0003, ADR-0004

## 问题陈述

`range_snapshot` family 当前已功能完整，但视觉质量明显低于同仓库的 combo 系列图。

**结构性原因（非参数问题）**：

- combo 是 Office native chart quality
- range_snapshot 是 native chart + custom overlay quality

axis break / current label / average tick 等能力 PowerPoint 不原生支持，所以引擎选择了 chart + overlay shape 的复合实现。这导致：

- 视觉质感弱于纯 native
- 轴系统的一体感不够
- 标签摆放偏机械
- 整体"原生感"明显低于 combo

如果将来用户问"为什么 combo 好这么多"，答案不是"参数还没调好"，而是这个结构性差异。已有 PowerPoint 实机导出证据支撑。

## 影响范围

| 维度 | 影响 |
|---|---|
| ADR-0001 "Native editable first" | range_snapshot 当前实现是已知妥协（native + overlay），未来如果出现类似妥协需要明确边界 |
| ADR-0003 valuation_snapshot golden | 该 golden 以 range_snapshot 为核心 family，本 PRD 不解决则 v0.3.0 视觉验收过不了 |
| ADR-0004 round-trip | 不受影响，metadata 闭环已完整 |
| Slice 3 排期 | v0.3.0 发版前必须完成本 PRD |

## 解决方案概要

进入"设计调优模式"——不再扩功能、不再加 family、不再加 workflow。

调参对象按优先级：

1. plot area size / margin
2. bar gap / bar width
3. tick label font size / density / position
4. average tick width
5. current label offset
6. axis break 位置与尺寸

视觉目标：达到 `jp_demo.pdf` 样张水准。注意"达到"= 看起来同档次，不是"完全 native"——把整个实现重写为 native 成本远超收益。

## 测试策略

按 [ADR-0002](../adr/0002-vertical-first-quality-gates.md)：

- **L2**：`range_snapshot` family 的视觉规范断言（颜色 / 字号 / 轴范围 / tick 密度）
- **L3**：`valuation_snapshot` golden 的 visual diff < 5% 阈值
- **真 PowerPoint 半自动验收**：渲染 → 人眼对比 jp_demo 样张

## 参考材料

- 原 handover 全文：[archive/2026-04-19-range-snapshot-handover.md](archive/2026-04-19-range-snapshot-handover.md)
- 引擎实机预览：`pptfi/output/vertical-valuation-previews/contact-sheet.png`
- combo 对比基线：`pptfi/output/combo-quality-previews/contact-sheet.png`
- 视觉目标样张：`ppt-project/jp_demo.pdf`
- 关键文件：
  - `pptchartengine/src/pptchartengine/range_snapshot.py`
  - `pptchartengine/src/pptchartengine/presets.py`（vertical valuation presets）
  - `pptfi/pptfi/composer/layouts/range_snapshot.py`

## 完成定义

- [ ] L2 range_snapshot 视觉规范断言全绿
- [ ] L3 `valuation_snapshot` golden visual diff < 5% 阈值
- [ ] PowerPoint 实机导出与 jp_demo 样张比对，主观判断"达到样张同档次"
- [ ] 本 PRD 中"结构性 vs 参数性"判断回写到 ADR-0001 评论或新增 ADR

## 显式不做的事

- 还原 jp_demo 中"图表外的标注点"（按用户原意故意不做）
- 把 range_snapshot 从 overlay 实现改为完全 native（成本远超收益；目标是"看起来像 native"）
