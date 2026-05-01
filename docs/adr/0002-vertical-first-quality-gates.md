# ADR-0002: 发版前的垂直质量闸门

- Status: Accepted
- Date: 2026-05-01
- Related: ADR-0003

## Context

L1 契约测试（当前 27 个）覆盖了：公开 API 导出、metadata round-trip、preset 形态、基础 family round-trip。

但实际"专业金融 PPT"的退化往往出现在以下维度，L1 全绿不能覆盖：

- 颜色串了 / 字号违规 / legend 顺序错
- 轴范围异常 / 断轴位置奇怪 / 日期轴 tick 错位或重叠
- PowerPoint 打开报错或显示警告
- 双击图表后无法继续编辑（embedded workbook formula 损坏 / external link 异常）
- 标签遮挡 / overlay 错位 / 文字溢出图区

这些问题如果带到下游 `pptfi` 或最终用户那边再发现，修复成本大幅升高。需要在引擎发版前就拦下。

## Decision

定义三层测试金字塔，所有 pip tag 发版必须满足全部闸门：

| 层级 | 范围 | 触发时机 |
|------|------|---------|
| L1 契约 | 公开 API + 单 family round-trip | 每次 PR |
| L2 垂直场景 | 每个 family happy path + edge case + 规范断言 | 改动 family 代码的 PR + nightly |
| L3 端到端 | golden reference reports 再生 + 视觉 diff | nightly + tag |
| Skill dogfood | Claude 用 skill 产出 vs reference | nightly + tag |

**发版闸门**：

- PR merge 闸门：L1 全绿 + 改动 family 的 L2 全绿
- tag 闸门：L1 全绿 + 该版本声明支持的 family 的 L2 全绿 + 该版本声明支持的 golden 的 L3 全绿 + 这些 family 的 skill dogfood 全绿
- 声明支持范围必须在 README 与 CHANGELOG 中明示，未声明的 family 不阻断发版
- nightly 跑全量并把 baseline diff 推到当日 review

注：本仓库采用垂直切片发版（详见 ROADMAP），首版仅声明支持 1 份 golden 涉及的 family。"该版本声明支持的"是显式作用域，而非"所有 public API"。

**L2 必须覆盖的规范断言**（每个 family 都要过）：

- 颜色：每个图表颜色必须从 `COLOR_SCHEMES` 取，不能硬编码
- 字号：legend / axis / data label 字号在允许集合内
- legend：顺序与 `series_config` 一致
- 数值轴：min/max 与数据实际范围匹配（含负值、跨零、单点）
- 日期轴：tick 数量在合理区间，不密集到重叠
- round-trip：create → parse → 语义等价（已有，并入 L2 统一管理）

## Consequences

正面：

- 退化在 CI 拦截，不会带到下游或最终用户
- "可发版"有客观判定，不靠人工感觉
- 视觉回归基线提供长期 reference

负面：

- 开发循环变长（nightly L3 跑得慢，估计 5-15 分钟）
- CI 基础设施投入（headless PowerPoint 渲染、视觉 diff 工具链）
- baseline 维护需要明确 owner 和更新流程
- 字体 / OS 渲染差异需要阈值调试，初期会有假阳性

## Alternatives Considered

**仅靠 L1 + manual review**：当前模式。
否决：人工 review 不 scalable；颜色 / 字号 / legend 类细节退化容易漏检。

**完全靠真 PowerPoint 自动化**：osascript 驱动 macOS PowerPoint 全量验收。
否决：跑得太慢、CI 跑不了、不能 cross-platform。保留为季度手动验收即可。

**只跑 nightly 不卡 PR**：
否决：会让 PR 作者发现退化的反馈周期变成 1 天，定位成本高。
