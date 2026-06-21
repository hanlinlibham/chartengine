# ADR-0008: 单人项目的轻量测试纪律

- Status: Proposed
- Date: 2026-05-24
- Related:
  - ADR-0002
  - ADR-0004
  - ADR-0005
  - ADR-0006
  - ADR-0007

## Context

ADR-0005 把 TDD 定为新 family / 新 public API 的默认流程:

```text
先写测试 -> 确认失败 -> 实现 -> 重构
```

这个规则的出发点是对的:AI 协作下需要可验证契约,否则容易写出"看起来对但不可审计"的代码。

但 `ablechart` 目前主要由一个开发者推进。过细的 TDD 分档、流程证明和文档标注会把单人项目拖成团队合规流程。PowerPoint / OOXML / `python-pptx` 又经常需要先探索对象模型,不适合每个小改动都强制红绿循环。

因此本 ADR 弱化 ADR-0005 §1:保留"完成前必须有证据"这个质量底线,取消一刀切 TDD 和繁琐分档。

## Decision

ADR-0005 §1 更新为:

> 测试纪律服务于开发速度和信心,不是流程本身。新能力完成前必须留下足够证据;是否严格 TDD 由开发者按风险判断。

实际执行采用四条轻量规则。

### 1. 高风险改动优先测试先行

以下情况建议先写测试或最小复现:

- 新 public API
- metadata schema / source of truth 变化
- parser / generator / replace 的核心语义变化
- bug 修复
- 会影响多个 chart family 的 shared path
- 会改变 ADR-0004 round-trip 或 ADR-0006 replace invariant

"建议"不是强制流程门。如果上下文已经清楚,也可以边实现边补测试;但最终必须有能证明行为的测试或实证。

### 2. 同构扩展不用机械红绿

已有框架下新增 chart type、错误码、fixture、matrix case,可以直接实现 + 补合同测试。

这类工作重点不是证明"红过",而是证明:

- 覆盖了新增行为
- 不破坏既有 baseline
- 对应的 round-trip / replace / inspect 契约仍成立

### 3. 未知外部行为先探索

遇到外部 `.pptx`、PowerPoint 兼容性、`python-pptx` 边界、custom XML、external workbook、脏模板等未知行为,允许先写探索代码或 characterization test。

探索结束后只需要沉淀其中一个结果:

- test
- fixture
- issue
- PRD / ADR 补充
- final report 中的明确限制说明

不要求探索阶段遵守 TDD。

### 4. 提交前看证据

完成声明只看证据,不看形式:

- 跑过哪些测试?
- 哪些契约被覆盖?
- 哪些风险没测?
- 如果没测,原因是什么?

对于单人开发,这比记录"是否红绿循环"更有用。

## Guardrails

轻量不等于无约束。以下底线仍然保留:

- 不能在没有测试或可复现实证的情况下宣称关键能力完成。
- 新 public API 最终必须有 contract test 或等价验收样例。
- 新 chart family 最终必须回到 ADR-0004 round-trip 约束。
- replace / inspect 能力最终必须回到 ADR-0006 fail-loud 和 invariant。
- 没测的风险要在 final report、issue 或 PRD 中说清楚。

## Consequences

正面:

- 单人开发节奏更轻,不会为流程服务。
- 仍保留可验证证据,不牺牲核心质量。
- 更适合 PowerPoint / OOXML 这种需要探索的工程对象。
- 后续如果项目进入多人协作,可以再收紧为更明确的 review gate。

负面:

- 质量更多依赖开发者判断,不像严格 TDD 那样机械可检查。
- 如果 final report 不诚实记录未测风险,后续 agent 仍可能误判完成度。
- 多人协作前需要重新审视测试纪律。

## Alternatives Considered

### A. 保持 ADR-0005 原样

否决。单人项目不需要每个新 public API / family 都走严格红绿流程。

### B. 使用严格 T0 / T1 / T2 分档

否决。分档比一刀切合理,但仍然太像团队流程。现在更需要轻量判断原则。

### C. 完全取消测试纪律

否决。`ablechart` 的价值来自 PowerPoint 原生可编辑、round-trip、replace invariant 和 fail-loud;没有证据就不能形成可信内核。
