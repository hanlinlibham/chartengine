# ADR-0003: Golden reference reports 作为质量北极星

- Status: Accepted
- Date: 2026-05-01
- Related: ADR-0002

## Context

"专业金融 PPT" 是主观定义，必须被操作化为可机器验证的标准，否则引擎演进会逐渐偏离垂直目标。

ppt-project 中已有候选金标准素材：

- `ppt-project/demo01.html`（10MB 金融分析看板）
- `ppt-project/portfolio.html`（6MB 持仓分析）
- `ppt-project/jp_demo.pdf`（日本市场样例）
- `ppt-project/financial_charts_template_sample.pptx`
- `pptfi/scripts/gen_moutai_report.py`（端到端报告生成）
- `pptfi/scripts/gen_anker_report.py`

如果不把"什么叫合格"固化为一组可重跑的 reference，所有质量讨论都会停留在主观争论。

## Decision

建立 `goldens/` 目录作为质量锚点，托管在 pptchartengine 与 pptfi 中较合适的一方（按依赖方向初步定为 pptfi，引擎只负责 visual diff harness）。

目录结构：

```
goldens/
├── reports/
│   ├── moutai_quarterly/
│   │   ├── input.parquet           # 完整可重跑的输入数据
│   │   ├── job.json                # pptfi job spec
│   │   ├── reference.pptx          # 人工验收过的 PPT
│   │   ├── reference_pages/        # 每页 PNG baseline
│   │   └── assertions.yaml         # 关键质量断言
│   └── ...
└── README.md
```

**初期固化 3 份 reference**（具体选哪 3 份待 M1 决定，候选见上）。

每份 golden 必须包含：

1. **完整可重跑的输入数据**（脱敏后版本）
2. **认可的 reference.pptx**（人工验收过、标记为"达到专业水准"）
3. **每页渲染的 PNG baseline**
4. **关键断言**（YAML 形式）：
   - 页数、每页 chart family
   - 关键数值（如某 chart 的轴 max、特定 series 的值）
   - 颜色规范命中情况

**baseline 更新流程**：

- 改动会导致 baseline 变化的 PR 必须显式更新 baseline
- baseline diff 走 PR review，不允许在 CI 里自动 accept
- 每季度 review 一次，确认 baseline 仍代表"专业"
- baseline 更新需要 PM 或资深分析师签字

## Consequences

正面：

- "专业"有明确锚点，新人可快速理解
- 视觉退化可被检测
- 后续可作为 demo / 营销素材
- 给 skill dogfood 提供天然评测集

负面：

- baseline 维护持续成本
- 字体 / 渲染差异需要阈值容忍
- 真实金融数据涉及合规，需要脱敏版本与脱敏流程
- 季度 PM review 需要稳定的人力承诺

## Alternatives Considered

**只依赖 unit test 断言**：足够防数据错误，但发现不了"不专业"——比如颜色不和谐、字号不统一、视觉重心偏。
否决。

**雇人定期人工评审**：成本高、不可重复、无法 CI 化。
保留为季度 review，不作为主要质量手段。

**用合成数据而非真实报告**：避免合规问题但失真。
折中方案：脱敏后的真实数据，保留行业代码、隐去具体公司名。
