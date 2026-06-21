# ADR-0003: Golden reference reports 作为质量北极星

- Status: Accepted
- Date: 2026-05-01
- Updated: 2026-05-03（锁定 3 份 golden 与命名原则）
- Related: ADR-0002

## Context

"专业金融 PPT" 是主观定义，必须被操作化为可机器验证的标准，否则引擎演进会逐渐偏离垂直目标。

ppt-project 中已有素材作为分析类型抽取的原始材料（**仅作为灵感来源，不直接作为 golden 命名或内容**）：

- `ppt-project/demo01.html`
- `ppt-project/portfolio.html`
- `ppt-project/jp_demo.pdf`
- `ppt-project/financial_charts_template_sample.pptx`
- `pptfi/scripts/gen_*_report.py`

## Decision

### 1. 命名原则：以分析类型而非主题命名

为避免特定标的、第三方品牌、第三方研究报告带来的版权 / 商标 / 数据使用风险，所有 golden 都按**分析类型**命名，不引用任何具体公司名、基金名、机构名、研究品牌。

- ❌ `moutai_quarterly` / `anker_q3` / `jp_demo` —— 含主题或来源
- ✅ `performance_attribution` / `valuation_snapshot` / `factor_style_analysis` —— 仅描述分析类型

输入数据必须脱敏：保留行业代码 / 资产类别 / 时间序列形态等结构性信息，隐去具体公司名、基金代码、机构名。

### 2. 三份首发 golden（按 slice 顺序）

| Slice | Golden | 分析类型 | 涉及 family |
|-------|--------|---------|------------|
| Slice 1 (v0.1.0) | `performance_attribution` | 业绩归因：基金 / 组合收益来源拆解 | `performance_compare` (combo) + `attribution_decomposition` (waterfall) + `regime_table_panel` |
| Slice 2 (v0.2.0) | `factor_style_analysis` | 因子风格分析：组合在风格 / 因子上的暴露 | `style_box` (scatter) + `factor_exposure` (combo) + `style_allocation` + `score_overlay` |
| Slice 3 (v0.3.0) | `valuation_snapshot` | 估值快照：市场 / 行业当前估值在历史区间的位置 | `range_snapshot` + `score_overlay` (复用) + `heatmap_matrix` |

**排序考量**：

- Slice 1 选业绩归因：涉及的 family 都是纯 native chart（combo / waterfall / table），视觉质量结构性最强，是稳健的 v0.1.0
- Slice 2 引入 scatter family 端到端验证；combo 在 slice 1 已跑通可以复用
- Slice 3 放估值快照在最后：range_snapshot 当前是 native + custom overlay 形态（详见 [PRD-0001](../prds/0001-range-snapshot-visual-polish.md)），视觉精修依赖前两个 slice 沉淀的 L2 / L3 工具

### 3. 目录结构与托管位置

`goldens/` 目录托管在 **pptfi** 仓库——依赖方向上 goldens 需要 pptfi 的 job spec + 输入数据。ablechart 仅持有 visual diff harness 与 fixtures 接入逻辑。

```
goldens/   (in pptfi repo)
├── reports/
│   ├── performance_attribution/
│   │   ├── input.parquet           # 完整可重跑的脱敏输入数据
│   │   ├── job.json                # pptfi job spec
│   │   ├── reference.pptx          # 人工验收过的 PPT
│   │   ├── reference_pages/        # 每页 PNG baseline
│   │   └── assertions.yaml         # 关键质量断言
│   ├── factor_style_analysis/
│   └── valuation_snapshot/
└── README.md
```

每份 golden 必须包含：

1. 完整可重跑的脱敏输入数据
2. 认可的 reference.pptx（人工验收过、标记为"达到专业水准"）
3. 每页渲染的 PNG baseline
4. 关键断言（YAML 形式）：页数、每页 chart family、关键数值（轴 max 等）、颜色规范命中

### 4. baseline 更新流程

- 改动会导致 baseline 变化的 PR 必须显式更新 baseline
- baseline diff 走 PR review，不允许 CI 自动 accept
- 每季度 review 一次，确认 baseline 仍代表"专业"
- baseline 更新需要 PM 或资深分析师签字

## Consequences

正面：

- "专业"有明确锚点，新人可快速理解
- 视觉退化可被检测
- 命名按分析类型避免侵权 / 商标风险，未来开源 goldens/ 时无障碍
- 同一分析类型可换底层数据快速验证泛化性
- 给 skill dogfood 提供天然评测集

负面：

- baseline 维护持续成本
- 字体 / 渲染差异需要阈值容忍
- 真实数据脱敏流程需要建立
- 分析类型比主题更抽象，每类是什么需要在 README 明确解释

## Alternatives Considered

**按主题命名（moutai / anker 等）**：
否决。涉及具体公司 / 基金 / 第三方研究报告时，存在版权、商标、数据使用条款风险，且未来开源 goldens/ 时会成为重大障碍。

**只依赖 unit test 断言**：足够防数据错误，但发现不了"不专业"——比如颜色不和谐、字号不统一、视觉重心偏。
否决。

**雇人定期人工评审**：成本高、不可重复、无法 CI 化。
保留为季度 review，不作为主要质量手段。

**完全合成数据**：避免合规问题但失真，可能掩盖真实场景的边界条件。
折中方案：脱敏后的真实数据 + 部分边界用合成 case 补足。
